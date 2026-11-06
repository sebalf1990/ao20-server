Attribute VB_Name = "modDenuncias"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
' modDenuncias (plan 2026-10-29-seguridad-antitrampas, etapa E3: /denunciar)
'
' 1. Buffer circular en memoria por slot de UserList: las ultimas DenunciaBufferAcciones acciones del
'    jugador (pasos, ataques, hechizos, usar/tirar/agarrar, cambios de oro y de inventario, eventos del
'    log antitrampa) con tick, mapa y posicion. Tamanio fijo por slot, reservado al arrancar (no crece).
'    Acciones iguales dentro de FUSION_MS se suman en una sola entrada. Se vacia campo por campo al
'    conectar/desconectar (TCP.ResetContadores -> AntitrampaResetUser), nunca copiando el UDT entero.
' 2. /DENUNCIAR <nick> [motivo] (paquete eDenounce, que ya existia: el cliente manda el texto crudo).
'    El denunciado tiene que estar online: si no, la respuesta es la misma de siempre (Msg532) y no se
'    guarda nada (no revela si esta conectado). No vale uno mismo ni un GM. Limites por denunciante (Id de
'    personaje: sobreviven al relogueo, no a un reinicio): DenunciasPorHora y DenunciasMismoObjetivoPorHora
'    (<= 0 = sin limite).
' 3. Snapshot: Logs\Denuncias\Denuncia_NNNNNN.jsonl (cabecera, acciones del denunciado de los ultimos
'    DenunciaVentanaMs y su chat reciente). Indice persistente Logs\Denuncias\Denuncias_indice.tsv (solo se
'    agrega: nueva/cerrar/borrar) que se relee al arrancar. Todo se encola y lo baja a disco DenunciasTick
'    desde el loop principal, nunca el handler del paquete.
' 4. Aviso a todos los GMs (ToAdmins, Consejero incluido): uno inmediato; las que llegan dentro de
'    DenunciaAvisoGmCadaMs se suman en un aviso agregado que sale al vencer la ventana.
' 5. ~DENUNCIAS (duenio 2026-10-29: "las ven todos los GMs y las cierran los Admins"):
'    todos los GMs: lista [todas] | ver <id>; solo Admin: cerrar <id> [nota] | borrar <id> | reset_limites.
'    ver/cerrar/borrar/reset quedan en Logs\Auditoria. Borrar mueve el archivo a Logs\Denuncias\Borradas
'    (no lo elimina). Nada sanciona: la denuncia es evidencia para un GM.
Option Explicit

Public Enum e_DenunciaAccion
    eDenPaso = 1
    eDenAtaque = 2
    eDenHechizo = 3
    eDenUsar = 4
    eDenTirar = 5
    eDenAgarrar = 6
    eDenOro = 7
    eDenItem = 8
    eDenAntitrampa = 9
End Enum

Private Const ACCION_MAX As Long = 9
Private Const FUSION_MS As Long = 2000
Private Const BUFFER_TOPE As Long = 1024
Private Const HORA_MS As Long = 3600000
Private Const LIMITE_ENTRADAS_MAX As Long = 100
Private Const LISTA_MAX As Long = 15
Private Const VER_LINEAS_MAX As Long = 15
Private Const PENDIENTES_MAX As Long = 64
Private Const FLUSH_FALLOS_MAX As Long = 3
Private Const MOTIVO_MAX As Long = 120
Private Const PALABRAS_NICK_MAX As Long = 4
Private Const CANTIDAD_TOPE As Double = 1000000000#
Private Const CARPETA_DENUNCIAS As String = "Denuncias"
Private Const ARCHIVO_INDICE As String = "Denuncias_indice.tsv"

Private Const ESTADO_ABIERTA As Byte = 1
Private Const ESTADO_CERRADA As Byte = 2
Private Const ESTADO_BORRADA As Byte = 3

Private Type t_DenAccion
    Tick As Long
    Clave As Long
    Cantidad As Long
    Total As Long
    Mapa As Integer
    TipoAccion As Byte
    PosX As Byte
    PosY As Byte
End Type

Private Type t_DenBuffer
    Acciones() As t_DenAccion
    Cabeza As Long
    Ocupadas As Long
    OroConocido As Long
    OroValido As Boolean
    InvObj(1 To MAX_INVENTORY_SLOTS) As Integer
    InvCant(1 To MAX_INVENTORY_SLOTS) As Long
End Type

Private Type t_DenIndice
    NumeroId As Long
    Fecha As String
    Denunciante As String
    Denunciado As String
    Motivo As String
    Resumen As String
    Estado As Byte
    CerradaPor As String
End Type

Private m_Buf() As t_DenBuffer
Private m_BufMax As Long
Private m_Tam As Long
Private m_VentanaMs As Long
Private m_PorHora As Long
Private m_MismoObjetivoPorHora As Long
Private m_AvisoCadaMs As Long

Private m_Indice() As t_DenIndice
Private m_IndiceCant As Long
Private m_ProximoId As Long

Private m_Limites As Dictionary               ' Id del denunciante -> "tick:IdDenunciado;" de la ultima hora

Private m_PendRuta() As String
Private m_PendTexto() As String
Private m_PendCant As Long
Private m_FlushFallos As Long
Private m_CarpetaLista As Boolean

Private m_UltimoAviso As Long
Private m_AvisosPendientes As Long
Private m_UltimoPendienteId As Long

' ---------------------------------------------------------------------------------------------
' Arranque y reset por jugador
' ---------------------------------------------------------------------------------------------

' La llama General.Main despues de InitAntitrampa (MaxUsers ya esta cargado).
Public Sub InitDenuncias()
    On Error GoTo InitDenuncias_Err
    Dim i As Long
    Set m_Limites = New Dictionary
    ReDim m_PendRuta(1 To PENDIENTES_MAX)
    ReDim m_PendTexto(1 To PENDIENTES_MAX)
    m_PendCant = 0
    m_Tam = SvrConfig.GetValue("DenunciaBufferAcciones")
    If m_Tam < 0 Then m_Tam = 0
    If m_Tam > BUFFER_TOPE Then m_Tam = BUFFER_TOPE
    m_VentanaMs = SvrConfig.GetValue("DenunciaVentanaMs")
    If m_VentanaMs < 1000 Then m_VentanaMs = 1000
    m_PorHora = SvrConfig.GetValue("DenunciasPorHora")
    m_MismoObjetivoPorHora = SvrConfig.GetValue("DenunciasMismoObjetivoPorHora")
    m_AvisoCadaMs = SvrConfig.GetValue("DenunciaAvisoGmCadaMs")
    If m_AvisoCadaMs < 0 Then m_AvisoCadaMs = 0
    m_BufMax = 0
    If MaxUsers > 0 Then
        ReDim m_Buf(1 To MaxUsers)
        If m_Tam > 0 Then
            For i = 1 To MaxUsers
                ReDim m_Buf(i).Acciones(0 To m_Tam - 1)
            Next i
        End If
        m_BufMax = MaxUsers
    End If
    Call CargarIndice
    Call LogThis(0, "Denuncias: buffer=" & m_Tam & " acciones ventana_ms=" & m_VentanaMs & " por_hora=" & m_PorHora & _
            " mismo_objetivo_por_hora=" & m_MismoObjetivoPorHora & " aviso_gm_cada_ms=" & m_AvisoCadaMs & _
            " en_indice=" & m_IndiceCant & " proximo_id=" & m_ProximoId, vbLogEventTypeInformation)
    Exit Sub
InitDenuncias_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.InitDenuncias", Erl)
End Sub

' Campo por campo (como el resto de TCP.ResetContadores), sin asignar el UDT entero.
Public Sub DenunciasResetUser(ByVal UserIndex As Integer)
    On Error GoTo DenunciasResetUser_Err
    Dim i As Long
    If UserIndex < 1 Or UserIndex > m_BufMax Then Exit Sub
    With m_Buf(UserIndex)
        .Cabeza = 0
        .Ocupadas = 0
        .OroConocido = 0
        .OroValido = False
        For i = 1 To MAX_INVENTORY_SLOTS
            .InvObj(i) = 0
            .InvCant(i) = 0
        Next i
    End With
    Exit Sub
DenunciasResetUser_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciasResetUser", Erl)
End Sub

Private Function TickNoCero(ByVal nowRaw As Long) As Long
    TickNoCero = nowRaw
    If TickNoCero = 0 Then TickNoCero = 1
End Function

' ---------------------------------------------------------------------------------------------
' Buffer de acciones (hot path: solo escribe en un arreglo ya reservado)
' ---------------------------------------------------------------------------------------------

Private Sub RegistrarAccion(ByVal UserIndex As Integer, ByVal nTipo As Long, ByVal nClave As Long, ByVal nCantidad As Long, ByVal nTotal As Long)
    Dim nowRaw As Long, ultimo As Long, suma As Double
    If m_Tam <= 0 Then Exit Sub
    If UserIndex < 1 Or UserIndex > m_BufMax Then Exit Sub
    nowRaw = GetTickCountRaw()
    With m_Buf(UserIndex)
        If .Ocupadas > 0 Then
            ultimo = (.Cabeza + m_Tam - 1) Mod m_Tam
            If .Acciones(ultimo).TipoAccion = nTipo And .Acciones(ultimo).Clave = nClave Then
                If nTipo <> eDenItem Or .Acciones(ultimo).Total = nTotal Then
                    If TicksElapsed(.Acciones(ultimo).Tick, nowRaw) < FUSION_MS Then
                        suma = CDbl(.Acciones(ultimo).Cantidad) + CDbl(nCantidad)
                        If suma > CANTIDAD_TOPE Then suma = CANTIDAD_TOPE
                        If suma < -CANTIDAD_TOPE Then suma = -CANTIDAD_TOPE
                        .Acciones(ultimo).Cantidad = CLng(suma)
                        If nTipo = eDenOro Then .Acciones(ultimo).Total = nTotal
                        .Acciones(ultimo).Mapa = UserList(UserIndex).pos.Map
                        .Acciones(ultimo).PosX = UserList(UserIndex).pos.x
                        .Acciones(ultimo).PosY = UserList(UserIndex).pos.y
                        Exit Sub
                    End If
                End If
            End If
        End If
        ultimo = .Cabeza
        .Acciones(ultimo).Tick = TickNoCero(nowRaw)
        .Acciones(ultimo).TipoAccion = nTipo
        .Acciones(ultimo).Clave = nClave
        .Acciones(ultimo).Cantidad = nCantidad
        .Acciones(ultimo).Total = nTotal
        .Acciones(ultimo).Mapa = UserList(UserIndex).pos.Map
        .Acciones(ultimo).PosX = UserList(UserIndex).pos.x
        .Acciones(ultimo).PosY = UserList(UserIndex).pos.y
        .Cabeza = (.Cabeza + 1) Mod m_Tam
        If .Ocupadas < m_Tam Then .Ocupadas = .Ocupadas + 1
    End With
End Sub

' Pasos, ataques, hechizos, usar/tirar/agarrar (los llaman los handlers de Protocol.bas).
Public Sub DenunciaAccion(ByVal UserIndex As Integer, ByVal nTipo As e_DenunciaAccion, ByVal nClave As Long, ByVal nCantidad As Long)
    On Error GoTo DenunciaAccion_Err
    Call RegistrarAccion(UserIndex, nTipo, nClave, nCantidad, 0)
    Exit Sub
DenunciaAccion_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciaAccion", Erl)
End Sub

' Desde WriteUpdateGold: registra la diferencia contra el ultimo oro conocido. Antes de que el personaje
' este logueado solo toma la base (el oro del login no es un cambio).
Public Sub DenunciaOro(ByVal UserIndex As Integer)
    On Error GoTo DenunciaOro_Err
    Dim oro As Long, delta As Double
    If UserIndex < 1 Or UserIndex > m_BufMax Then Exit Sub
    oro = UserList(UserIndex).Stats.GLD
    If m_Buf(UserIndex).OroValido And UserList(UserIndex).flags.UserLogged Then
        If oro <> m_Buf(UserIndex).OroConocido Then
            delta = CDbl(oro) - CDbl(m_Buf(UserIndex).OroConocido)
            If delta > CANTIDAD_TOPE Then delta = CANTIDAD_TOPE
            If delta < -CANTIDAD_TOPE Then delta = -CANTIDAD_TOPE
            Call RegistrarAccion(UserIndex, eDenOro, 0, CLng(delta), oro)
        End If
    End If
    m_Buf(UserIndex).OroConocido = oro
    m_Buf(UserIndex).OroValido = True
    Exit Sub
DenunciaOro_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciaOro", Erl)
End Sub

' Desde WriteChangeInventorySlot: registra lo que entro o salio del slot (obj y diferencia de cantidad).
' Antes del login solo toma la base (el envio inicial del inventario no es un cambio).
Public Sub DenunciaItemSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo DenunciaItemSlot_Err
    Dim obj As Integer, cant As Long
    If UserIndex < 1 Or UserIndex > m_BufMax Then Exit Sub
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    obj = UserList(UserIndex).invent.Object(Slot).ObjIndex
    cant = UserList(UserIndex).invent.Object(Slot).amount
    With m_Buf(UserIndex)
        If UserList(UserIndex).flags.UserLogged Then
            If obj <> .InvObj(Slot) Then
                If .InvObj(Slot) > 0 And .InvCant(Slot) > 0 Then Call RegistrarAccion(UserIndex, eDenItem, .InvObj(Slot), -.InvCant(Slot), Slot)
                If obj > 0 And cant > 0 Then Call RegistrarAccion(UserIndex, eDenItem, obj, cant, Slot)
            ElseIf obj > 0 And cant <> .InvCant(Slot) Then
                Call RegistrarAccion(UserIndex, eDenItem, obj, cant - .InvCant(Slot), Slot)
            End If
        End If
        .InvObj(Slot) = obj
        .InvCant(Slot) = cant
    End With
    Exit Sub
DenunciaItemSlot_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciaItemSlot", Erl)
End Sub

Private Function CodigoEventoAntitrampa(ByVal evento As String) As Long
    Select Case evento
        Case "speedhack": CodigoEventoAntitrampa = 1
        Case "intervalo_rechazado": CodigoEventoAntitrampa = 2
        Case "contador_repetido": CodigoEventoAntitrampa = 3
        Case "kill_misma_ip": CodigoEventoAntitrampa = 4
        Case "kill_misma_pc": CodigoEventoAntitrampa = 5
        Case "kill_tope_diario": CodigoEventoAntitrampa = 6
        Case "reto_misma_ip": CodigoEventoAntitrampa = 7
        Case "reto_misma_pc": CodigoEventoAntitrampa = 8
        Case "denunciado": CodigoEventoAntitrampa = 9
        Case Else: CodigoEventoAntitrampa = 0
    End Select
End Function

Private Function NombreEventoAntitrampa(ByVal codigo As Long) As String
    Select Case codigo
        Case 1: NombreEventoAntitrampa = "speedhack"
        Case 2: NombreEventoAntitrampa = "intervalo_rechazado"
        Case 3: NombreEventoAntitrampa = "contador_repetido"
        Case 4: NombreEventoAntitrampa = "kill_misma_ip"
        Case 5: NombreEventoAntitrampa = "kill_misma_pc"
        Case 6: NombreEventoAntitrampa = "kill_tope_diario"
        Case 7: NombreEventoAntitrampa = "reto_misma_ip"
        Case 8: NombreEventoAntitrampa = "reto_misma_pc"
        Case 9: NombreEventoAntitrampa = "denunciado"
        Case Else: NombreEventoAntitrampa = "otro"
    End Select
End Function

' Desde modAntitrampa.AntitrampaRegistrar: el evento queda tambien en el buffer del jugador.
Public Sub DenunciaAntitrampa(ByVal UserIndex As Integer, ByVal evento As String)
    On Error GoTo DenunciaAntitrampa_Err
    Call RegistrarAccion(UserIndex, eDenAntitrampa, CodigoEventoAntitrampa(evento), 1, 0)
    Exit Sub
DenunciaAntitrampa_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciaAntitrampa", Erl)
End Sub

Private Function NombreAccion(ByVal nTipo As Long) As String
    Select Case nTipo
        Case eDenPaso: NombreAccion = "paso"
        Case eDenAtaque: NombreAccion = "ataque"
        Case eDenHechizo: NombreAccion = "hechizo"
        Case eDenUsar: NombreAccion = "usar"
        Case eDenTirar: NombreAccion = "tirar"
        Case eDenAgarrar: NombreAccion = "agarrar"
        Case eDenOro: NombreAccion = "oro"
        Case eDenItem: NombreAccion = "item"
        Case eDenAntitrampa: NombreAccion = "antitrampa"
        Case Else: NombreAccion = "otra"
    End Select
End Function

Private Function NombreResumen(ByVal nTipo As Long) As String
    Select Case nTipo
        Case eDenPaso: NombreResumen = "pasos"
        Case eDenAtaque: NombreResumen = "ataques"
        Case eDenHechizo: NombreResumen = "hechizos"
        Case eDenUsar: NombreResumen = "usos"
        Case eDenTirar: NombreResumen = "tirados"
        Case eDenAgarrar: NombreResumen = "agarrados"
        Case eDenOro: NombreResumen = "oro"
        Case eDenItem: NombreResumen = "items"
        Case Else: NombreResumen = "antitrampa"
    End Select
End Function

' ---------------------------------------------------------------------------------------------
' Snapshot
' ---------------------------------------------------------------------------------------------

Private Function JsonTexto(ByVal texto As String) As String
    Dim i As Long, c As Long, salida As String
    salida = """"
    For i = 1 To Len(texto)
        c = AscW(Mid$(texto, i, 1))
        Select Case c
            Case 34
                salida = salida & "\"""
            Case 92
                salida = salida & "\\"
            Case 0 To 31
                salida = salida & "\u" & Right$("000" & Hex$(c), 4)
            Case Else
                salida = salida & Mid$(texto, i, 1)
        End Select
    Next i
    JsonTexto = salida & """"
End Function

Private Function JsonPersona(ByVal UserIndex As Integer, ByVal conEstado As Boolean) As String
    With UserList(UserIndex)
        JsonPersona = "{""pj"":" & .Id & ",""nick"":" & JsonTexto(.name) & ",""cuenta"":" & .AccountID & _
                ",""ip"":" & JsonTexto(.ConnectionDetails.IP) & ",""mapa"":" & .pos.Map & ",""x"":" & .pos.x & ",""y"":" & .pos.y
        If conEstado Then
            JsonPersona = JsonPersona & ",""nivel"":" & .Stats.ELV & ",""oro"":" & .Stats.GLD & ",""vida"":" & .Stats.MinHp & _
                    ",""vida_max"":" & .Stats.MaxHp & ",""muerto"":" & .flags.Muerto
        End If
        JsonPersona = JsonPersona & "}"
    End With
End Function

' Arma el contenido del archivo de la denuncia. Devuelve tambien el resumen y la cantidad de acciones.
Private Function ArmarSnapshot(ByVal numero As Long, ByVal fecha As String, ByVal denunciante As Integer, ByVal denunciado As Integer, _
                               ByVal Motivo As String, ByRef Resumen As String, ByRef cantAcciones As Long) As String
    Dim nowRaw As Long, ahora As Date, i As Long, idx As Long, inicio As Long, hace As Double
    Dim cuerpo As String, linea As String, chat As String
    Dim totales(1 To ACCION_MAX) As Double
    nowRaw = GetTickCountRaw()
    ahora = Now
    cantAcciones = 0
    If m_Tam > 0 And denunciado >= 1 And denunciado <= m_BufMax Then
        With m_Buf(denunciado)
            inicio = (.Cabeza - .Ocupadas + m_Tam) Mod m_Tam
            For i = 0 To .Ocupadas - 1
                idx = (inicio + i) Mod m_Tam
                hace = TicksElapsed(.Acciones(idx).Tick, nowRaw)
                If hace <= m_VentanaMs Then
                    linea = "{""tipo"":""accion"",""ts"":" & JsonTexto(Format$(ahora - hace / 86400000#, "yyyy-mm-dd hh:nn:ss")) & _
                            ",""hace_s"":" & CStr(Fix(hace / 1000#)) & ",""accion"":" & JsonTexto(NombreAccion(.Acciones(idx).TipoAccion)) & _
                            ",""mapa"":" & .Acciones(idx).Mapa & ",""x"":" & .Acciones(idx).PosX & ",""y"":" & .Acciones(idx).PosY
                    Select Case .Acciones(idx).TipoAccion
                        Case eDenHechizo
                            linea = linea & ",""hechizo"":" & .Acciones(idx).Clave & ",""cantidad"":" & .Acciones(idx).Cantidad
                        Case eDenUsar, eDenTirar, eDenAgarrar
                            linea = linea & ",""obj"":" & .Acciones(idx).Clave & ",""cantidad"":" & .Acciones(idx).Cantidad
                        Case eDenOro
                            linea = linea & ",""delta"":" & .Acciones(idx).Cantidad & ",""total"":" & .Acciones(idx).Total
                        Case eDenItem
                            linea = linea & ",""obj"":" & .Acciones(idx).Clave & ",""delta"":" & .Acciones(idx).Cantidad & ",""slot"":" & .Acciones(idx).Total
                        Case eDenAntitrampa
                            linea = linea & ",""evento"":" & JsonTexto(NombreEventoAntitrampa(.Acciones(idx).Clave)) & ",""cantidad"":" & .Acciones(idx).Cantidad
                        Case Else
                            linea = linea & ",""cantidad"":" & .Acciones(idx).Cantidad
                    End Select
                    cuerpo = cuerpo & linea & "}" & vbCrLf
                    cantAcciones = cantAcciones + 1
                    If .Acciones(idx).TipoAccion >= 1 And .Acciones(idx).TipoAccion <= ACCION_MAX Then
                        If .Acciones(idx).TipoAccion = eDenItem Then
                            totales(eDenItem) = totales(eDenItem) + 1
                        Else
                            totales(.Acciones(idx).TipoAccion) = totales(.Acciones(idx).TipoAccion) + .Acciones(idx).Cantidad
                        End If
                    End If
                End If
            Next i
        End With
    End If
    Resumen = "acciones=" & cantAcciones
    For i = 1 To ACCION_MAX
        If totales(i) <> 0 Then
            If i = eDenOro And totales(i) > 0 Then
                Resumen = Resumen & " " & NombreResumen(i) & "=+" & CStr(totales(i))
            Else
                Resumen = Resumen & " " & NombreResumen(i) & "=" & CStr(totales(i))
            End If
        End If
    Next i
    With UserList(denunciado)
        For i = LBound(.flags.ChatHistory) To UBound(.flags.ChatHistory)
            If LenB(.flags.ChatHistory(i)) <> 0 Then
                chat = chat & "{""tipo"":""chat"",""texto"":" & JsonTexto(.flags.ChatHistory(i)) & "}" & vbCrLf
            End If
        Next i
    End With
    ArmarSnapshot = "{""tipo"":""denuncia"",""id"":" & numero & ",""ts"":" & JsonTexto(fecha) & ",""motivo"":" & JsonTexto(Motivo) & _
            ",""ventana_s"":" & CStr(m_VentanaMs \ 1000) & ",""acciones"":" & cantAcciones & ",""resumen"":" & JsonTexto(Resumen) & _
            ",""denunciante"":" & JsonPersona(denunciante, False) & ",""denunciado"":" & JsonPersona(denunciado, True) & "}" & vbCrLf & _
            cuerpo & chat
End Function

' ---------------------------------------------------------------------------------------------
' Archivos: cola que baja a disco DenunciasTick (loop principal)
' ---------------------------------------------------------------------------------------------

Private Function CarpetaDenuncias() As String
    CarpetaDenuncias = App.Path & "\Logs\" & CARPETA_DENUNCIAS
End Function

Private Function RutaDenuncia(ByVal numero As Long) As String
    RutaDenuncia = CarpetaDenuncias() & "\Denuncia_" & Format$(numero, "000000") & ".jsonl"
End Function

Private Function RutaBorrada(ByVal numero As Long) As String
    RutaBorrada = CarpetaDenuncias() & "\Borradas\Denuncia_" & Format$(numero, "000000") & ".jsonl"
End Function

Private Function RutaIndice() As String
    RutaIndice = CarpetaDenuncias() & "\" & ARCHIVO_INDICE
End Function

Private Sub AsegurarCarpeta()
    If m_CarpetaLista Then Exit Sub
    If LenB(Dir$(App.Path & "\Logs", vbDirectory)) = 0 Then MkDir App.Path & "\Logs"
    If LenB(Dir$(CarpetaDenuncias(), vbDirectory)) = 0 Then MkDir CarpetaDenuncias()
    If LenB(Dir$(CarpetaDenuncias() & "\Borradas", vbDirectory)) = 0 Then MkDir CarpetaDenuncias() & "\Borradas"
    m_CarpetaLista = True
End Sub

Private Sub Encolar(ByVal ruta As String, ByVal texto As String)
    If m_PendCant >= PENDIENTES_MAX Then Call DenunciasFlush
    If m_PendCant >= PENDIENTES_MAX Then
        Call LogThis(0, "Denuncias: cola llena, se descarta una escritura en " & ruta, vbLogEventTypeWarning)
        Exit Sub
    End If
    m_PendCant = m_PendCant + 1
    m_PendRuta(m_PendCant) = ruta
    m_PendTexto(m_PendCant) = texto
End Sub

' Baja la cola en orden. Si falla, lo que no se escribio queda para la proxima vuelta; despues de
' FLUSH_FALLOS_MAX fallos seguidos se descarta (no crece sin limite) y queda una linea en Logs\Operativo.
Public Sub DenunciasFlush()
    Dim fn As Integer, i As Long, hechos As Long
    On Error GoTo ErrHandler
    If m_PendCant = 0 Then Exit Sub
    Call AsegurarCarpeta
    For i = 1 To m_PendCant
        fn = FreeFile
        Open m_PendRuta(i) For Append As #fn
        Print #fn, m_PendTexto(i);
        Close #fn
        fn = 0
        hechos = i
    Next i
    For i = 1 To m_PendCant
        m_PendRuta(i) = vbNullString
        m_PendTexto(i) = vbNullString
    Next i
    m_PendCant = 0
    m_FlushFallos = 0
    Exit Sub
ErrHandler:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    m_CarpetaLista = False
    m_FlushFallos = m_FlushFallos + 1
    ' Saca de la cola lo que ya se escribio.
    For i = hechos + 1 To m_PendCant
        m_PendRuta(i - hechos) = m_PendRuta(i)
        m_PendTexto(i - hechos) = m_PendTexto(i)
    Next i
    For i = m_PendCant - hechos + 1 To m_PendCant
        m_PendRuta(i) = vbNullString
        m_PendTexto(i) = vbNullString
    Next i
    m_PendCant = m_PendCant - hechos
    If m_FlushFallos >= FLUSH_FALLOS_MAX Then
        Call LogThis(0, "Denuncias: no se pudo escribir en " & CarpetaDenuncias() & "; se descartan " & m_PendCant & " escrituras", vbLogEventTypeError)
        For i = 1 To m_PendCant
            m_PendRuta(i) = vbNullString
            m_PendTexto(i) = vbNullString
        Next i
        m_PendCant = 0
        m_FlushFallos = 0
    End If
End Sub

' Una vez por vuelta del loop principal (General.Main): barato si no hay nada pendiente.
Public Sub DenunciasTick()
    On Error GoTo DenunciasTick_Err
    If m_PendCant > 0 Then Call DenunciasFlush
    If m_AvisosPendientes > 0 Then
        If TicksElapsed(m_UltimoAviso, GetTickCountRaw()) >= m_AvisoCadaMs Then Call AvisoAgregado
    End If
    Exit Sub
DenunciasTick_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciasTick", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Indice persistente
' ---------------------------------------------------------------------------------------------

Private Function Limpiar(ByVal texto As String) As String
    Limpiar = Replace$(Replace$(Replace$(texto, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Function LineaIndice(ByVal op As String, ByVal numero As Long, ByVal fecha As String, ByVal denunciante As String, _
                             ByVal denunciado As String, ByVal Motivo As String, ByVal Resumen As String, ByVal por As String) As String
    LineaIndice = op & vbTab & numero & vbTab & Limpiar(fecha) & vbTab & Limpiar(denunciante) & vbTab & Limpiar(denunciado) & vbTab & _
            Limpiar(Motivo) & vbTab & Limpiar(Resumen) & vbTab & Limpiar(por) & vbCrLf
End Function

Private Sub AgregarIndice(ByVal numero As Long, ByVal fecha As String, ByVal denunciante As String, ByVal denunciado As String, _
                          ByVal Motivo As String, ByVal Resumen As String)
    If m_IndiceCant = 0 Then
        ReDim m_Indice(1 To 16)
    ElseIf m_IndiceCant >= UBound(m_Indice) Then
        ReDim Preserve m_Indice(1 To UBound(m_Indice) * 2)
    End If
    m_IndiceCant = m_IndiceCant + 1
    With m_Indice(m_IndiceCant)
        .NumeroId = numero
        .Fecha = fecha
        .Denunciante = denunciante
        .Denunciado = denunciado
        .Motivo = Motivo
        .Resumen = Resumen
        .Estado = ESTADO_ABIERTA
        .CerradaPor = vbNullString
    End With
End Sub

Private Function BuscarIndice(ByVal numero As Long) As Long
    Dim i As Long
    For i = m_IndiceCant To 1 Step -1
        If m_Indice(i).NumeroId = numero Then
            BuscarIndice = i
            Exit Function
        End If
    Next i
End Function

Private Sub CargarIndice()
    On Error GoTo CargarIndice_Err
    Dim fn As Integer, linea As String, campos() As String, numero As Long, pos As Long
    m_IndiceCant = 0
    m_ProximoId = 1
    If Not FileExist(RutaIndice()) Then Exit Sub
    fn = FreeFile
    Open RutaIndice() For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, linea
        campos = Split(linea, vbTab)
        If UBound(campos) >= 7 Then
            numero = CLng(val(campos(1)))
            If numero > 0 Then
                Select Case campos(0)
                    Case "nueva"
                        If BuscarIndice(numero) = 0 Then Call AgregarIndice(numero, campos(2), campos(3), campos(4), campos(5), campos(6))
                    Case "cerrar"
                        pos = BuscarIndice(numero)
                        If pos > 0 Then
                            m_Indice(pos).Estado = ESTADO_CERRADA
                            m_Indice(pos).CerradaPor = campos(7)
                        End If
                    Case "borrar"
                        pos = BuscarIndice(numero)
                        If pos > 0 Then m_Indice(pos).Estado = ESTADO_BORRADA
                End Select
                If numero >= m_ProximoId Then m_ProximoId = numero + 1
            End If
        End If
    Loop
    Close #fn
    Exit Sub
CargarIndice_Err:
    If fn <> 0 Then Close #fn
    Call TraceError(Err.Number, Err.Description, "modDenuncias.CargarIndice", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Limites por denunciante
' ---------------------------------------------------------------------------------------------

' Limpia lo viejo (mas de una hora) y cuenta las denuncias de la ultima hora: en total y contra objetivoId.
Private Sub ContarLimites(ByVal denuncianteId As Long, ByVal objetivoId As Long, ByVal nowRaw As Long, ByRef total As Long, ByRef mismo As Long)
    Dim clave As String, lista As String, partes() As String, campos() As String, i As Long, salida As String
    total = 0
    mismo = 0
    If m_Limites Is Nothing Then Set m_Limites = New Dictionary
    clave = CStr(denuncianteId)
    If Not m_Limites.Exists(clave) Then Exit Sub
    lista = m_Limites.Item(clave)
    partes = Split(lista, ";")
    For i = 0 To UBound(partes)
        campos = Split(partes(i), ":")
        If UBound(campos) = 1 Then
            If TicksElapsed(CLng(val(campos(0))), nowRaw) < HORA_MS Then
                total = total + 1
                If CLng(val(campos(1))) = objetivoId Then mismo = mismo + 1
                salida = salida & partes(i) & ";"
            End If
        End If
    Next i
    If LenB(salida) = 0 Then
        Call m_Limites.Remove(clave)
    Else
        m_Limites.Item(clave) = salida
    End If
End Sub

Private Sub AgregarLimite(ByVal denuncianteId As Long, ByVal objetivoId As Long, ByVal nowRaw As Long)
    Dim clave As String, total As Long, mismo As Long
    Call ContarLimites(denuncianteId, objetivoId, nowRaw, total, mismo)
    If total >= LIMITE_ENTRADAS_MAX Then Exit Sub
    clave = CStr(denuncianteId)
    If m_Limites.Exists(clave) Then
        m_Limites.Item(clave) = m_Limites.Item(clave) & CStr(TickNoCero(nowRaw)) & ":" & CStr(objetivoId) & ";"
    Else
        m_Limites.Item(clave) = CStr(TickNoCero(nowRaw)) & ":" & CStr(objetivoId) & ";"
    End If
End Sub

' ---------------------------------------------------------------------------------------------
' /DENUNCIAR (Protocol_GmCommands.HandleDenounce)
' ---------------------------------------------------------------------------------------------

Private Function UnirPalabras(ByRef palabras() As String, ByVal desde As Long, ByVal hasta As Long) As String
    Dim i As Long
    For i = desde To hasta
        If i > desde Then UnirPalabras = UnirPalabras & " "
        UnirPalabras = UnirPalabras & palabras(i)
    Next i
End Function

Public Sub DenunciarJugador(ByVal UserIndex As Integer, ByVal texto As String)
    On Error GoTo DenunciarJugador_Err
    Dim palabras() As String, limpias() As String, cant As Long, i As Long, k As Long
    Dim candidato As String, Motivo As String, objetivo As Integer, tUser As t_UserReference
    Dim nowRaw As Long, total As Long, mismo As Long
    texto = Trim$(Limpiar(texto))
    If LenB(texto) = 0 Then Exit Sub
    palabras = Split(texto, " ")
    ReDim limpias(0 To UBound(palabras))
    For i = 0 To UBound(palabras)
        If LenB(palabras(i)) > 0 Then
            limpias(cant) = palabras(i)
            cant = cant + 1
        End If
    Next i
    If cant = 0 Then Exit Sub
    ' Los nicks pueden tener espacios: se prueba el prefijo mas largo que sea un personaje conectado.
    k = cant
    If k > PALABRAS_NICK_MAX Then k = PALABRAS_NICK_MAX
    Do While k >= 1
        candidato = UnirPalabras(limpias, 0, k - 1)
        If EsGmChar(candidato) Then
            ' Msg531=No podes denunciar a un administrador.
            Call WriteLocaleMsg(UserIndex, MSG_NO_PODES_DENUNCIAR_ADMINISTRADOR, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
            Exit Sub
        End If
        tUser = NameIndex(candidato)
        If IsValidUserRef(tUser) Then
            objetivo = tUser.ArrayIndex
            Motivo = Left$(UnirPalabras(limpias, k, cant - 1), MOTIVO_MAX)
            Exit Do
        End If
        k = k - 1
    Loop
    If objetivo = 0 Then
        ' Desconectado o inexistente: la respuesta de siempre (no revela si esta online) y no se guarda nada.
        ' Msg532=Tu denuncia fue recibida por el equipo de soporte.
        Call WriteLocaleMsg(UserIndex, MSG_DENUNCIA_FUE_RECIBIDA_EQUIPO_SOPORTE, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
        Exit Sub
    End If
    If objetivo = UserIndex Then
        ' Msg9016=No podes denunciarte a vos mismo.
        Call WriteLocaleMsg(UserIndex, MSG_DENUNCIA_A_VOS_MISMO, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
        Exit Sub
    End If
    ' Con el bridge los rangos vienen de la cuenta (EsGmChar solo mira Server.ini): se mira el conectado.
    If EsGM(objetivo) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_PODES_DENUNCIAR_ADMINISTRADOR, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
        Exit Sub
    End If
    nowRaw = GetTickCountRaw()
    Call ContarLimites(UserList(UserIndex).Id, UserList(objetivo).Id, nowRaw, total, mismo)
    If m_PorHora > 0 And total >= m_PorHora Then
        ' Msg9017=Llegaste al limite de 1 denuncias por hora.
        Call WriteLocaleMsg(UserIndex, MSG_DENUNCIA_LIMITE_POR_HORA, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER, CStr(m_PorHora))
        Exit Sub
    End If
    If m_MismoObjetivoPorHora > 0 And mismo >= m_MismoObjetivoPorHora Then
        ' Msg9018=Ya denunciaste a 1 hace menos de una hora.
        Call WriteLocaleMsg(UserIndex, MSG_DENUNCIA_MISMO_OBJETIVO, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER, UserList(objetivo).name)
        Exit Sub
    End If
    Call CrearDenuncia(UserIndex, objetivo, Motivo, nowRaw)
    Exit Sub
DenunciarJugador_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.DenunciarJugador", Erl)
End Sub

Private Sub CrearDenuncia(ByVal UserIndex As Integer, ByVal objetivo As Integer, ByVal Motivo As String, ByVal nowRaw As Long)
    Dim numero As Long, fecha As String, contenido As String, Resumen As String, cantAcciones As Long
    Dim chatTexto As String, i As Long, nickDenunciante As String, nickDenunciado As String
    nickDenunciante = UserList(UserIndex).name
    nickDenunciado = UserList(objetivo).name
    numero = m_ProximoId
    If numero < 1 Then numero = 1
    ' Si se perdio el indice, no pisar un archivo que ya existe.
    Do While FileExist(RutaDenuncia(numero)) Or FileExist(RutaBorrada(numero))
        numero = numero + 1
    Loop
    m_ProximoId = numero + 1
    fecha = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    contenido = ArmarSnapshot(numero, fecha, UserIndex, objetivo, Motivo, Resumen, cantAcciones)
    Call Encolar(RutaDenuncia(numero), contenido)
    Call AgregarIndice(numero, fecha, nickDenunciante, nickDenunciado, Motivo, Resumen)
    Call Encolar(RutaIndice(), LineaIndice("nueva", numero, fecha, nickDenunciante, nickDenunciado, Motivo, Resumen, vbNullString))
    Call AgregarLimite(UserList(UserIndex).Id, UserList(objetivo).Id, nowRaw)
    Call AntitrampaRegistrar(objetivo, "denunciado", "log", "id=" & numero & " denunciante_pj=" & UserList(UserIndex).Id & _
            " denunciante_nick=" & nickDenunciante & " motivo=" & Motivo)
    ' Compatibilidad con el panel de consultas (SOS) de los GMs: sigue recibiendo el chat del denunciado.
    With UserList(objetivo)
        For i = LBound(.flags.ChatHistory) To UBound(.flags.ChatHistory)
            If LenB(.flags.ChatHistory(i)) <> 0 Then chatTexto = chatTexto & .flags.ChatHistory(i) & vbNewLine
        Next i
    End With
    If LenB(chatTexto) <> 0 Then
        Call Ayuda.Push(nickDenunciante, "[Ultimos mensajes de " & nickDenunciado & "]" & vbNewLine & chatTexto, "Denuncia #" & numero & " a " & nickDenunciado)
    End If
    Call LogConsulta(nickDenunciante & " (Denuncia #" & numero & " a " & nickDenunciado & ") motivo=" & Motivo & " " & Resumen)
    Call AvisarNueva(numero, nickDenunciante, nickDenunciado, Motivo)
    ' Msg534=Tu denuncia fue recibida por el equipo de soporte.
    Call WriteLocaleMsg(UserIndex, MSG_DENUNCIA_FUE_RECIBIDA_EQUIPO_SOPORTE_534, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
End Sub

' ---------------------------------------------------------------------------------------------
' Avisos a GMs (agregados: riesgo "ruido a los GMs" del plan)
' ---------------------------------------------------------------------------------------------

Private Sub AvisarGms(ByVal texto As String)
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(texto, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER))
End Sub

Private Sub AvisarNueva(ByVal numero As Long, ByVal denunciante As String, ByVal denunciado As String, ByVal Motivo As String)
    Dim nowRaw As Long, texto As String
    nowRaw = GetTickCountRaw()
    If m_UltimoAviso <> 0 And m_AvisoCadaMs > 0 Then
        If TicksElapsed(m_UltimoAviso, nowRaw) < m_AvisoCadaMs Then
            m_AvisosPendientes = m_AvisosPendientes + 1
            m_UltimoPendienteId = numero
            Exit Sub
        End If
    End If
    texto = "Denuncia #" & numero & ": " & denunciante & " denuncio a " & denunciado
    If LenB(Motivo) > 0 Then texto = texto & " (" & Motivo & ")"
    texto = texto & ". Ver: ~DENUNCIAS ver " & numero
    If m_AvisosPendientes > 0 Then texto = texto & " (+" & m_AvisosPendientes & " sin avisar: ~DENUNCIAS)"
    m_AvisosPendientes = 0
    m_UltimoAviso = TickNoCero(nowRaw)
    Call AvisarGms(texto)
End Sub

Private Sub AvisoAgregado()
    Call AvisarGms("Denuncias: " & m_AvisosPendientes & " nuevas sin avisar (ultima #" & m_UltimoPendienteId & "). Ver: ~DENUNCIAS")
    m_AvisosPendientes = 0
    m_UltimoAviso = TickNoCero(GetTickCountRaw())
End Sub

' ---------------------------------------------------------------------------------------------
' Comando de GM "~DENUNCIAS" (viaja como chat, igual que ~ANTITRAMPA; lo deja pasar HandleTalk a todo GM)
' ---------------------------------------------------------------------------------------------

Private Function RangoTexto(ByVal UserIndex As Integer) As String
    With UserList(UserIndex).flags
        If .Privilegios And e_PlayerType.Admin Then
            RangoTexto = "admin"
        ElseIf .Privilegios And e_PlayerType.Dios Then
            RangoTexto = "dios"
        ElseIf .Privilegios And e_PlayerType.SemiDios Then
            RangoTexto = "semidios"
        ElseIf .Privilegios And e_PlayerType.Consejero Then
            RangoTexto = "consejero"
        Else
            RangoTexto = "usuario"
        End If
    End With
End Function

Private Function PuedeCerrar(ByVal UserIndex As Integer) As Boolean
    PuedeCerrar = (UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin) <> 0
End Function

Private Sub Responder(ByVal UserIndex As Integer, ByVal texto As String)
    Call WriteConsoleMsg(UserIndex, texto, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EstadoTexto(ByVal Estado As Byte) As String
    Select Case Estado
        Case ESTADO_ABIERTA: EstadoTexto = "abierta"
        Case ESTADO_CERRADA: EstadoTexto = "cerrada"
        Case Else: EstadoTexto = "borrada"
    End Select
End Function

Private Function LineaLista(ByVal pos As Long) As String
    With m_Indice(pos)
        LineaLista = "Denuncia #" & .NumeroId & " [" & EstadoTexto(.Estado) & "] " & .Fecha & " " & .Denunciante & " -> " & .Denunciado
        If LenB(.Motivo) > 0 Then LineaLista = LineaLista & ": " & .Motivo
        LineaLista = LineaLista & " | " & .Resumen
        If .Estado = ESTADO_CERRADA Then LineaLista = LineaLista & " | cerrada por " & .CerradaPor
    End With
End Function

Public Function HandleDenunciasChatCommand(ByVal UserIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo HandleDenunciasChatCommand_Err
    Dim partes() As String, opcion As String, numero As Long, nota As String, i As Long
    chat = Trim$(chat)
    Do While InStr(chat, "  ") > 0
        chat = Replace$(chat, "  ", " ")
    Loop
    partes = Split(chat, " ")
    ' Split de "" devuelve un arreglo vacio (UBound = -1).
    If UBound(partes) < 0 Then Exit Function
    If UCase$(partes(0)) <> "~DENUNCIAS" Then Exit Function
    HandleDenunciasChatCommand = True
    If UBound(partes) >= 1 Then opcion = LCase$(partes(1))
    If UBound(partes) >= 2 Then numero = CLng(val(partes(2)))
    Select Case opcion
        Case vbNullString, "lista"
            Call ComandoLista(UserIndex, UBound(partes) >= 2 And LCase$(partes(UBound(partes))) = "todas")
        Case "ver"
            Call ComandoVer(UserIndex, numero)
        Case "cerrar", "borrar", "reset_limites"
            If Not PuedeCerrar(UserIndex) Then
                Call Responder(UserIndex, "Denuncias: solo un Admin puede cerrar o borrar denuncias (tu rango: " & RangoTexto(UserIndex) & ").")
                Exit Function
            End If
            Select Case opcion
                Case "cerrar"
                    For i = 3 To UBound(partes)
                        nota = nota & " " & partes(i)
                    Next i
                    Call ComandoCerrar(UserIndex, numero, Trim$(nota))
                Case "borrar"
                    Call ComandoBorrar(UserIndex, numero)
                Case Else
                    If m_Limites Is Nothing Then Set m_Limites = New Dictionary
                    Call m_Limites.RemoveAll
                    m_UltimoAviso = 0
                    m_AvisosPendientes = 0
                    Call LogAudit("DENUNCIA", "GM " & UserList(UserIndex).name & " reinicio los limites por hora y los avisos de denuncias")
                    Call Responder(UserIndex, "Denuncias: limites por hora y avisos reiniciados.")
            End Select
        Case Else
            Call Responder(UserIndex, "Denuncias: uso ~DENUNCIAS [lista [todas] | ver <id> | cerrar <id> [nota] | borrar <id> | reset_limites]")
    End Select
    Exit Function
HandleDenunciasChatCommand_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.HandleDenunciasChatCommand", Erl)
End Function

Private Sub ComandoLista(ByVal UserIndex As Integer, ByVal todas As Boolean)
    Dim i As Long, abiertas As Long, cerradas As Long, mostradas As Long
    For i = 1 To m_IndiceCant
        If m_Indice(i).Estado = ESTADO_ABIERTA Then abiertas = abiertas + 1
        If m_Indice(i).Estado = ESTADO_CERRADA Then cerradas = cerradas + 1
    Next i
    Call Responder(UserIndex, "Denuncias: abiertas=" & abiertas & " cerradas=" & cerradas & " rango=" & RangoTexto(UserIndex) & _
            " (cerrar/borrar: solo Admin). Uso: ~DENUNCIAS [lista [todas] | ver <id> | cerrar <id> [nota] | borrar <id> | reset_limites]")
    For i = m_IndiceCant To 1 Step -1
        If mostradas >= LISTA_MAX Then Exit For
        If m_Indice(i).Estado = ESTADO_ABIERTA Or (todas And m_Indice(i).Estado = ESTADO_CERRADA) Then
            Call Responder(UserIndex, LineaLista(i))
            mostradas = mostradas + 1
        End If
    Next i
End Sub

Private Sub ComandoVer(ByVal UserIndex As Integer, ByVal numero As Long)
    On Error GoTo ComandoVer_Err
    Dim pos As Long, fn As Integer, linea As String, ultimas(1 To VER_LINEAS_MAX) As String
    Dim cantAcc As Long, i As Long, idx As Long, chats As Long
    pos = BuscarIndice(numero)
    If pos = 0 Then
        Call Responder(UserIndex, "Denuncias: no existe la denuncia #" & numero & ".")
        Exit Sub
    End If
    If m_Indice(pos).Estado = ESTADO_BORRADA Then
        Call Responder(UserIndex, "Denuncias: no existe la denuncia #" & numero & ".")
        Exit Sub
    End If
    Call LogAudit("DENUNCIA", "GM " & UserList(UserIndex).name & " abrio la denuncia #" & numero)
    Call Responder(UserIndex, LineaLista(pos))
    ' Lo encolado se baja ya, asi el archivo esta completo.
    Call DenunciasFlush
    If FileExist(RutaDenuncia(numero)) Then
        fn = FreeFile
        Open RutaDenuncia(numero) For Input As #fn
        Do While Not EOF(fn)
            Line Input #fn, linea
            If InStr(1, linea, """tipo"":""accion""", vbBinaryCompare) > 0 Then
                cantAcc = cantAcc + 1
                ultimas(((cantAcc - 1) Mod VER_LINEAS_MAX) + 1) = linea
            ElseIf InStr(1, linea, """tipo"":""chat""", vbBinaryCompare) > 0 Then
                chats = chats + 1
            End If
        Loop
        Close #fn
        fn = 0
        Call Responder(UserIndex, "Denuncia #" & numero & " acciones=" & cantAcc & " chat=" & chats & " (ultimas " & IIf(cantAcc < VER_LINEAS_MAX, cantAcc, VER_LINEAS_MAX) & "):")
        For i = 1 To VER_LINEAS_MAX
            If i > cantAcc Then Exit For
            If cantAcc <= VER_LINEAS_MAX Then
                idx = i
            Else
                idx = ((cantAcc + i - 1) Mod VER_LINEAS_MAX) + 1
            End If
            Call Responder(UserIndex, "Denuncia #" & numero & " " & ultimas(idx))
        Next i
    End If
    Call Responder(UserIndex, "Denuncia #" & numero & " archivo: Logs\" & CARPETA_DENUNCIAS & "\Denuncia_" & Format$(numero, "000000") & ".jsonl")
    Exit Sub
ComandoVer_Err:
    If fn <> 0 Then Close #fn
    Call TraceError(Err.Number, Err.Description, "modDenuncias.ComandoVer", Erl)
End Sub

Private Sub ComandoCerrar(ByVal UserIndex As Integer, ByVal numero As Long, ByVal nota As String)
    Dim pos As Long
    pos = BuscarIndice(numero)
    If pos = 0 Then
        Call Responder(UserIndex, "Denuncias: no existe la denuncia #" & numero & ".")
        Exit Sub
    End If
    If m_Indice(pos).Estado <> ESTADO_ABIERTA Then
        Call Responder(UserIndex, "Denuncias: la denuncia #" & numero & " no esta abierta (" & EstadoTexto(m_Indice(pos).Estado) & ").")
        Exit Sub
    End If
    m_Indice(pos).Estado = ESTADO_CERRADA
    m_Indice(pos).CerradaPor = UserList(UserIndex).name
    With m_Indice(pos)
        Call Encolar(RutaIndice(), LineaIndice("cerrar", numero, Format$(Now, "yyyy-mm-dd hh:nn:ss"), .Denunciante, .Denunciado, nota, .Resumen, UserList(UserIndex).name))
    End With
    Call LogAudit("DENUNCIA", "GM " & UserList(UserIndex).name & " cerro la denuncia #" & numero & " nota=" & nota)
    Call Responder(UserIndex, "Denuncia #" & numero & " cerrada por " & UserList(UserIndex).name & ".")
End Sub

Private Sub ComandoBorrar(ByVal UserIndex As Integer, ByVal numero As Long)
    On Error GoTo ComandoBorrar_Err
    Dim pos As Long
    pos = BuscarIndice(numero)
    If pos = 0 Then
        Call Responder(UserIndex, "Denuncias: no existe la denuncia #" & numero & ".")
        Exit Sub
    End If
    If m_Indice(pos).Estado = ESTADO_BORRADA Then
        Call Responder(UserIndex, "Denuncias: no existe la denuncia #" & numero & ".")
        Exit Sub
    End If
    m_Indice(pos).Estado = ESTADO_BORRADA
    With m_Indice(pos)
        Call Encolar(RutaIndice(), LineaIndice("borrar", numero, Format$(Now, "yyyy-mm-dd hh:nn:ss"), .Denunciante, .Denunciado, .Motivo, .Resumen, UserList(UserIndex).name))
    End With
    ' Primero se baja la cola (el snapshot puede estar pendiente), despues se mueve el archivo.
    Call DenunciasFlush
    Call AsegurarCarpeta
    If FileExist(RutaDenuncia(numero)) Then Name RutaDenuncia(numero) As RutaBorrada(numero)
    Call LogAudit("DENUNCIA", "GM " & UserList(UserIndex).name & " borro la denuncia #" & numero & " (archivo movido a Logs\" & CARPETA_DENUNCIAS & "\Borradas)")
    Call Responder(UserIndex, "Denuncia #" & numero & " borrada (archivo movido a Logs\" & CARPETA_DENUNCIAS & "\Borradas).")
    Exit Sub
ComandoBorrar_Err:
    Call TraceError(Err.Number, Err.Description, "modDenuncias.ComandoBorrar", Erl)
End Sub
