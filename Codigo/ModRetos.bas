Attribute VB_Name = "ModRetos"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Private Const APUESTA_MAXIMA = 100000000
' Plan 27.001 ola 6 (05.002): tabla del enfriamiento anti win-trading. La clave es el CONJUNTO
' exacto de participantes: si cambia uno solo es otro enfrentamiento y arranca de cero.
' Es tabla fija a proposito: 256 entradas sobran para una ventana de 45 minutos.
Private Const MAX_HISTORIAL_RETOS As Integer = 256

Private Type t_HistorialReto
    Clave        As String
    Repeticiones As Integer
    UltimoTick   As Long
End Type

Public Retos          As t_Retos
Private ListaDeEspera As New Dictionary
Private HistorialRetos(0 To MAX_HISTORIAL_RETOS - 1) As t_HistorialReto

Public Sub CargarInfoRetos()
    Dim File As clsIniManager
    Set File = New clsIniManager
    Call File.Initialize(DatPath & "Retos.dat")
    With Retos
        .TamañoMaximoEquipo = val(File.GetValue("Retos", "MaximoEquipo"))
        .ApuestaMinima = val(File.GetValue("Retos", "ApuestaMinima"))
        .ImpuestoApuesta = val(File.GetValue("Retos", "ImpuestoApuesta"))
        .DuracionMaxima = val(File.GetValue("Retos", "DuracionMaxima"))
        .TiempoGuardarItems = LeerRetoAcotado(File, "TiempoGuardarItems", 60, 32767)
        .FactorK = LeerRetoAcotado(File, "FactorK", 24, 32767)
        .NivelMinimoELO = LeerRetoAcotado(File, "NivelMinimoELO", 35, 255)
        .MaximoRetosSeguidos = LeerRetoAcotado(File, "MaximoRetosSeguidos", 3, 32767)
        .MinutosEnfriamiento = LeerRetoAcotado(File, "MinutosEnfriamiento", 45, 32767)
        ' Defaults obligatorios (en LeerRetoAcotado): el binario y el .dat se despliegan por separado. Con un
        ' .dat viejo FactorK quedaria en 0 y el ELO de todo el servidor se congelaria sin
        ' que nadie se entere.
        #If DEBUGGING Then
            .TiempoConteo = 3
        #Else
            .TiempoConteo = val(File.GetValue("Retos", "TiempoConteo"))
        #End If
        .TotalSalas = val(File.GetValue("Salas", "Cantidad"))
        If .TotalSalas <= 0 Then Exit Sub
        ReDim .Salas(1 To .TotalSalas)
        .SalasLibres = .TotalSalas
        .AnchoSala = val(File.GetValue("Salas", "Ancho"))
        .AltoSala = val(File.GetValue("Salas", "Alto"))
        Dim Sala As Integer, SalaStr As String
        For Sala = 1 To .TotalSalas
            SalaStr = "Sala" & Sala
            With .Salas(Sala)
                .PosIzquierda.Map = val(File.GetValue(SalaStr, "Mapa"))
                .PosIzquierda.x = val(File.GetValue(SalaStr, "X"))
                .PosIzquierda.y = val(File.GetValue(SalaStr, "Y"))
                .PosDerecha.Map = .PosIzquierda.Map
                .PosDerecha.x = .PosIzquierda.x + Retos.AnchoSala - 1
                .PosDerecha.y = .PosIzquierda.y + Retos.AltoSala - 1
            End With
        Next
    End With
    Set File = Nothing
End Sub

' Lee un entero de Retos.dat acotado a 1..Maximo; fuera de rango (o 0/vacio) cae al default.
' val() devuelve Double: asignarlo directo a un Byte/Integer con un valor enorme da error 6.
Private Function LeerRetoAcotado(ByVal File As clsIniManager, ByVal Clave As String, ByVal ValorDefault As Long, ByVal Maximo As Long) As Long
    Dim Valor As Double
    Valor = val(File.GetValue("Retos", Clave))
    If Valor < 1# Or Valor > CDbl(Maximo) Then
        LeerRetoAcotado = ValorDefault
    Else
        LeerRetoAcotado = CLng(Fix(Valor))
    End If
End Function

Public Sub CrearReto(ByVal UserIndex As Integer, JugadoresStr As String, ByVal Apuesta As Long, ByVal PocionesMaximas As Integer, Optional ByVal CaenItems As Boolean = False)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
        Dim TamanoReal As Byte: TamanoReal = Retos.TamañoMaximoEquipo * 2 - 1
        If LenB(JugadoresStr) <= 0 Then Exit Sub
        Dim Jugadores() As String: Jugadores = Split(JugadoresStr, ";", TamanoReal)
        If UBound(Jugadores) > TamanoReal - 1 Or UBound(Jugadores) Mod 2 = 1 Then Exit Sub
        Dim MaxIndexEquipo As Integer: MaxIndexEquipo = UBound(Jugadores) \ 2
        If Apuesta < Retos.ApuestaMinima Or Apuesta > APUESTA_MAXIMA Then
            Call WriteLocaleMsg(UserIndex, MSG_APUESTA_MINIMA_MONEDAS_ORO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PonerPuntos(Retos.ApuestaMinima)) ' Msg1958=La apuesta mínima es de ¬1 monedas de oro.
            Exit Sub
        End If
        If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
        If .Stats.GLD < Apuesta Then
            ' Msg588=No tienes el oro suficiente.
            Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ORO_SUFICIENTE, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If PocionesMaximas >= 0 Then
            If TieneObjetos(38, PocionesMaximas + 1, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, MSG_TIENES_DEMASIADAS_POCIONES_ROJAS_CANTIDAD_MAXIMA_1443, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PocionesMaximas) ' Msg1443=Tienes demasiadas pociones rojas (Cantidad máxima: ¬1).
                Exit Sub
            End If
        End If
        With .flags.SolicitudReto
            .Apuesta = Apuesta
            .PocionesMaximas = PocionesMaximas
            .CaenItems = CaenItems
            ReDim .Jugadores(0 To UBound(Jugadores))
            Dim i       As Integer, tIndex As t_UserReference
            Dim Equipo1 As String, Equipo2 As String
            Equipo1 = UserList(UserIndex).name
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    If EsGmChar(Jugadores(i)) Then
                        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_JUGAR_RETOS_ADMINISTRADORES, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg1444=¡No puedes jugar retos con administradores!
                        Exit Sub
                    End If
                    tIndex = NameIndex(Jugadores(i))
                    If Not IsValidUserRef(tIndex) Then
                        Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, Jugadores(i)) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
                        Exit Sub
                    End If
                    If Not PuedeReto(tIndex.ArrayIndex) Then
                        Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, UserList(tIndex.ArrayIndex).name) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
                        Exit Sub
                    End If
                    .CurIndex = tIndex
                    .nombre = UserList(.CurIndex.ArrayIndex).name
                    .Aceptado = False
                    If i Mod 2 Then
                        Equipo1 = Equipo1 & IIf((i + 1) \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                    Else
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf(i \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                        Else
                            Equipo2 = .nombre
                        End If
                    End If
                End With
            Next
            ' Freno anti win-trading EN LA PUERTA: con los nombres ya resueltos pero antes de
            ' mandar una sola invitacion, y solo cuando el reto moveria ELO.
            If RetoOtorgaElo(UserIndex) Then
                Dim MinutosFaltan As Long
                MinutosFaltan = MinutosParaRepetir(ClaveEnfrentamiento(UserIndex))
                If MinutosFaltan > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Ya jugaste " & Retos.MaximoRetosSeguidos & " retos seguidos contra ese mismo equipo. Faltan " & MinutosFaltan & " minutos para poder repetirlo.", e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                    Exit Sub
                End If
            End If
            Dim Texto1 As String, Texto2 As String, Texto3 As String
            Texto1 = UserList(UserIndex).name & "(" & UserList(UserIndex).Stats.ELV & ") te invita a jugar el siguiente reto:"
            Texto2 = Equipo1 & " vs " & Equipo2 & ". Apuesta: " & PonerPuntos(Apuesta) & " monedas de oro" & IIf(CaenItems, " y los items.", ".")
            Texto3 = "Escribe /ACEPTAR " & UCase$(UserList(UserIndex).name) & " para participar en el reto."
            If PocionesMaximas >= 0 Then
                Texto2 = Texto2 & " Máximo " & PocionesMaximas & " pociones rojas."
            End If
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto1, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto2, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto3, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                End With
            Next
            .Estado = e_SolicitudRetoEstado.Enviada
        End With
        Call WriteLocaleMsg(UserIndex, MSG_ENVIADO_SOLICITUD_SIGUIENTE_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg1446=Has enviado una solicitud para el siguiente reto:
        Call WriteConsoleMsg(UserIndex, Texto2, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
        Call WriteLocaleMsg(UserIndex, MSG_ESCRIBE_CANCELAR_ANULAR_SOLICITUD, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg1447=Escribe /CANCELAR para anular la solicitud.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CrearReto", Erl)
End Sub

Public Sub AceptarReto(ByVal UserIndex As Integer, OferenteName As String)
    On Error GoTo ErrHandler
    If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
    End With
    If EsGmChar(OferenteName) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_JUGAR_RETOS_ADMINISTRADORES, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1959=¡No puedes jugar retos con administradores!
        Exit Sub
    End If
    Dim Oferente As t_UserReference
    Oferente = NameIndex(OferenteName)
    If Not IsValidUserRef(Oferente) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_CONECTADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, OferenteName) ' Msg1960=¬1 no está conectado.
        Exit Sub
    End If
    With UserList(Oferente.ArrayIndex).flags.SolicitudReto
        Dim JugadorIndex As Integer
        JugadorIndex = IndiceJugadorEnSolicitud(UserIndex, Oferente.ArrayIndex)
        If JugadorIndex < 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_INVITADO_NINGUN_RETO_SIDO_CANCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, UserList(Oferente.ArrayIndex).name) ' Msg1961=¬1 no te ha invitado a ningún reto o ha sido cancelado.
            Exit Sub
        End If
        If UserList(UserIndex).Stats.GLD < .Apuesta Then
            Call WriteLocaleMsg(UserIndex, MSG_NECESITAS_MONEDAS_ORO_ACEPTAR_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PonerPuntos(.Apuesta)) ' Msg1962=Necesitas al menos ¬1 monedas de oro para aceptar este reto.
            Exit Sub
        End If
        If .PocionesMaximas >= 0 Then
            If TieneObjetos(38, .PocionesMaximas + 1, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, MSG_TIENES_DEMASIADAS_POCIONES_ROJAS_CANTIDAD_MAXIMA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, .PocionesMaximas) ' Msg1963=Tienes demasiadas pociones rojas (Cantidad máxima: ¬1).
                Exit Sub
            End If
        End If
        Call MensajeATodosSolicitud(Oferente.ArrayIndex, UserList(UserIndex).name & " ha aceptado el reto.", e_FontTypeNames.FONTTYPE_INFO)
        .Jugadores(JugadorIndex).Aceptado = True
        Call SetUserRef(.Jugadores(JugadorIndex).CurIndex, UserIndex)
        UserList(UserIndex).flags.AceptoReto = Oferente
        Call WriteLocaleMsg(UserIndex, MSG_HAS_ACEPTADO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, UserList(Oferente.ArrayIndex).name) ' Msg1964=Has aceptado el reto de ¬1.
        Dim FaltanAceptar As String
        Dim i             As Integer
        For i = 0 To UBound(.Jugadores)
            If Not .Jugadores(i).Aceptado Then
                FaltanAceptar = FaltanAceptar & .Jugadores(i).nombre & " - "
            End If
        Next
        If LenB(FaltanAceptar) > 0 Then
            FaltanAceptar = Left$(FaltanAceptar, Len(FaltanAceptar) - 3)
            Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Faltan aceptar: " & FaltanAceptar, e_FontTypeNames.FONTTYPE_New_Gris)
            Exit Sub
        End If
        Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Todos los jugadores han aceptado el reto. Buscando sala...", e_FontTypeNames.FONTTYPE_New_Gris)
        Call BuscarSala(Oferente.ArrayIndex)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AceptarReto", Erl)
End Sub

Public Sub CancelarSolicitudReto(ByVal Oferente As Integer, mensaje As String)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If .Estado = e_SolicitudRetoEstado.EnCola Then
            If ListaDeEspera.Exists(Oferente) Then Call ListaDeEspera.Remove(Oferente)
        End If
        .Estado = e_SolicitudRetoEstado.Libre
        Dim i As Integer, tUser As t_UserReference
        ' Enviamos a los invitados
        For i = 0 To UBound(.Jugadores)
            tUser = NameIndex(.Jugadores(i).nombre)
            If IsValidUserRef(tUser) Then
                Call WriteConsoleMsg(tUser.ArrayIndex, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                Call WriteLocaleMsg(tUser.ArrayIndex, MSG_RETO_SIDO_CANCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1965=El reto ha sido cancelado.
                If .Jugadores(i).Aceptado Then
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.AceptoReto, 0)
                End If
            End If
        Next
        ' Y al oferente por separado
        Call WriteConsoleMsg(Oferente, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
        Call WriteLocaleMsg(Oferente, MSG_RETO_SIDO_CANCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1965=El reto ha sido cancelado.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CancelarSolicitudReto", Erl)
End Sub

Private Sub BuscarSala(ByVal Oferente As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Retos.SalasLibres <= 0 Then
            ' Plan 05.002 ola 1: el estado EnCola no se asignaba NUNCA, asi que el guard de
            ' CancelarSolicitudReto no disparaba y el oferente quedaba en ListaDeEspera para
            ' siempre: al liberarse una sala se le arrancaba un reto que ya habia cancelado.
            .Estado = e_SolicitudRetoEstado.EnCola
            Call ListaDeEspera.Add(Oferente, 0)
            Call MensajeATodosSolicitud(Oferente, "No hay salas disponibles. El reto comenzará cuando se desocupe una sala.", e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        Dim Sala As Integer, SalaAleatoria As Integer
        SalaAleatoria = RandomNumber(1, Retos.SalasLibres)
        For Sala = 1 To Retos.TotalSalas
            If Not Retos.Salas(Sala).EnUso Then
                SalaAleatoria = SalaAleatoria - 1
                If SalaAleatoria = 0 Then Exit For
            End If
        Next
        Call IniciarReto(Oferente, Sala)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.BuscarSala", Erl)
End Sub

Private Sub IniciarReto(ByVal Oferente As Integer, ByVal Sala As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        ' Última comprobación de si todos pueden entrar/pagar
        If Not TodosPuedenReto(Oferente) Then Exit Sub
        Dim Apuesta As Long, ApuestaStr As String
        Apuesta = .Apuesta
        ApuestaStr = PonerPuntos(Apuesta)
        ' Calculamos el tamaño del equipo
        Retos.Salas(Sala).TamañoEquipoIzq = UBound(.Jugadores) \ 2 + 1
        Retos.Salas(Sala).TamañoEquipoDer = Retos.Salas(Sala).TamañoEquipoIzq
        ' Reservamos espacio para los jugadores (incluyendo al oferente)
        ReDim Retos.Salas(Sala).Jugadores(0 To UBound(.Jugadores) + 1)
        ' Tiramos una moneda (50-50) y decidimos si agregar al oferente al inicio o al final de la lista
        Dim Moneda As Byte
        Moneda = RandomNumber(0, 1)
        Dim CurIndex As Integer
        If Moneda = 0 Then
            ' Agregamos al oferente al inicio (su equipo juega a la izquierda)
            Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
            CurIndex = CurIndex + 1
        End If
        Dim i As Integer
        ' Agregamos los jugadores alternando 1 y 1 (en los índices pares está el equipo izquierdo y en los impares el derecho - el array empieza en cero)
        For i = 0 To UBound(.Jugadores)
            Retos.Salas(Sala).Jugadores(CurIndex) = .Jugadores(i).CurIndex
            CurIndex = CurIndex + 1
            ' Reset flag
            Call SetUserRef(UserList(.Jugadores(i).CurIndex.ArrayIndex).flags.AceptoReto, 0)
        Next
        If Moneda = 1 Then
            ' Agregamos al oferente al final (su equipo juega a la derecha)
            Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
        End If
        ' Reset estado de la solicitud, ya que no la necesitamos más
        .Estado = e_SolicitudRetoEstado.Libre
    End With
    With Retos.Salas(Sala)
        .EnUso = True
        .Finalizado = False
        .Puntaje = 0
        .Ronda = 1
        .Apuesta = Apuesta
        .TiempoRestante = Retos.DuracionMaxima
        .CaenItems = UserList(Oferente).flags.SolicitudReto.CaenItems
        ' Foto del ELO y del plantel. Todo lo que el rating necesita se congela aca: si se
        ' leyera al final, el que abandona le cambiaria el rating al rival y el gate de
        ' nivel se evaluaria sobre los sobrevivientes.
        ReDim .EloSaldado(0 To UBound(.Jugadores))
        .TamañoInicialIzq = .TamañoEquipoIzq
        .TamañoInicialDer = .TamañoEquipoDer
        .EloHabilitado = True
        Dim EloSumaIzq As Long, EloSumaDer As Long
        .EloPromedioIzq = 0
        .EloPromedioDer = 0
        For i = 0 To UBound(.Jugadores)
            .EloSaldado(i) = False
            If IsValidUserRef(.Jugadores(i)) Then
                If UserList(.Jugadores(i).ArrayIndex).Stats.ELV < Retos.NivelMinimoELO Then .EloHabilitado = False
                If i Mod 2 = 0 Then
                    EloSumaIzq = EloSumaIzq + UserList(.Jugadores(i).ArrayIndex).Stats.ELO
                Else
                    EloSumaDer = EloSumaDer + UserList(.Jugadores(i).ArrayIndex).Stats.ELO
                End If
            Else
                ' Falta alguien antes de empezar: mejor no mover rating que inventarlo.
                .EloHabilitado = False
            End If
        Next i
        If .TamañoInicialIzq > 0 Then .EloPromedioIzq = EloSumaIzq \ CLng(.TamañoInicialIzq)
        If .TamañoInicialDer > 0 Then .EloPromedioDer = EloSumaDer \ CLng(.TamañoInicialDer)
        ' Recien aca se registra el enfrentamiento para el enfriamiento: una solicitud que
        ' se cancela o que nunca arranca no gasta uno de los retos permitidos.
        If .EloHabilitado Then Call RegistrarEnfrentamiento(ClaveEnfrentamiento(Oferente))
        Dim tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            ' Le cobramos
            UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD - Apuesta
            Call WriteUpdateGold(tUser.ArrayIndex)
            Call WriteLocaleMsg(tUser.ArrayIndex, MSG_OTORGAS_MONEDAS_ORO_POZO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, ApuestaStr) ' Msg1966=Otorgas ¬1 monedas de oro al pozo del reto.
            ' Desmontamos
            If UserList(tUser.ArrayIndex).flags.Montado <> 0 Then
                Call DoMontar(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedSaddleObjIndex), UserList(tUser.ArrayIndex).invent.EquippedSaddleSlot)
            End If
            ' Dejamos de navegar
            If UserList(tUser.ArrayIndex).flags.Nadando <> 0 Or UserList(tUser.ArrayIndex).flags.Navegando <> 0 Then
                Call DoNavega(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedShipObjIndex), UserList(tUser.ArrayIndex).invent.EquippedShipSlot)
            End If
            ' Asignamos flags
            With UserList(tUser.ArrayIndex).flags
                .EnReto = True
                .EquipoReto = IIf(i Mod 2, e_EquipoReto.Derecha, e_EquipoReto.Izquierda)
                .SalaReto = Sala
                ' Guardar posición
                .LastPos = UserList(tUser.ArrayIndex).pos
            End With
            Call WriteLocaleMsg(tUser.ArrayIndex, MSG_COMENZADO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1967=¡Ha comenzado el reto!
            Call WriteLocaleMsg(tUser.ArrayIndex, MSG_ADMITIR_DERROTA_ESCRIBE_ABANDONAR, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1968=Para admitir la derrota escribe /ABANDONAR.
        Next
    End With
    Retos.SalasLibres = Retos.SalasLibres - 1
    Call iniciarRonda(Sala)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.IniciarReto", Erl)
End Sub

Private Sub iniciarRonda(ByVal Sala As Integer)
    With Retos.Salas(Sala)
        Dim i As Integer, tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If tUser.ArrayIndex <> 0 Then
                Call RevivirYLimpiar(tUser.ArrayIndex)
                ' Usando el número de ronda y el índice, decidimos el lado al que corresponde
                If (.Ronda + i) Mod 2 = 1 Then
                    ' Lado izquierdo
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                Else
                    ' Lado derecho
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosDerecha.Map, .PosDerecha.x, .PosDerecha.y, True)
                End If
                ' Si usamos el conteo
                If Retos.TiempoConteo > 0 Then
                    ' Le ponemos el conteo
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = Retos.TiempoConteo
                    ' Lo stoppeamos
                    Call WriteStopped(tUser.ArrayIndex, True)
                End If
                Call WriteLocaleMsg(tUser.ArrayIndex, MSG_COMIENZA_RONDA_N, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, .Ronda) ' Msg1969=Comienza la ronda Nº¬1
            End If
        Next
    End With
End Sub

Public Sub MuereEnReto(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler
    Dim Sala As Integer, Equipo As e_EquipoReto
    With UserList(UserIndex)
        Sala = .flags.SalaReto
        Equipo = .flags.EquipoReto
    End With
    With Retos.Salas(Sala)
        Dim CurIndex As Integer
        ' El equipo derecho está en índices pares
        If Equipo = e_EquipoReto.Derecha Then CurIndex = 1
        For CurIndex = CurIndex To UBound(.Jugadores) Step 2
            If .Jugadores(CurIndex).ArrayIndex <> 0 Then
                ' Si todavía hay alguno vivo del equipo
                If UserList(.Jugadores(CurIndex).ArrayIndex).flags.Muerto = 0 Then
                    Exit Sub
                End If
            End If
        Next
        ' Están todos muertos, ganó el equipo contrario
        Call ProcesarRondaGanada(Sala, EquipoContrario(Equipo))
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.MuereEnReto", Erl)
End Sub

Private Sub ProcesarRondaGanada(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto)
    ' El label ErrorHandler estaba escrito pero no habia On Error que lo activara: cualquier
    ' error de este Sub subia al handler de MuereEnReto con el origen mal.
    On Error GoTo ErrorHandler
    With Retos.Salas(Sala)
        ' Sumamos puntaje o restamos según el equipo
        If Equipo = e_EquipoReto.Derecha Then
            .Puntaje = .Puntaje + 1
        Else
            .Puntaje = .Puntaje - 1
        End If
        ' Si terminó la tercer ronda o bien algún equipo obtuvo 2 victorias seguidas
        If .Ronda >= 3 Or Abs(.Puntaje) >= 2 Then
            Call FinalizarReto(Sala)
            Exit Sub
        End If
        ' Aumentamos el número de ronda
        .Ronda = .Ronda + 1
        ' Nombres del equipo ganador. El bucle viejo iba hasta 2 * TamañoEquipo - 1 con el
        ' tamano YA descontado por los abandonos, pero los slots del array no se compactan:
        ' salteaba sobrevivientes en indices altos y podia leer el slot vaciado del que se fue.
        ' Ahora se recorre el array entero filtrando por paridad (pares = izquierda) y por
        ' referencia viva.
        Dim i As Integer, nombres As String
        For i = IIf(Equipo = e_EquipoReto.Izquierda, 0, 1) To UBound(.Jugadores) Step 2
            If IsValidUserRef(.Jugadores(i)) Then
                If LenB(nombres) > 0 Then nombres = nombres & ", "
                nombres = nombres & UserList(.Jugadores(i).ArrayIndex).name
            End If
        Next
        ' Informamos el ganador de esta ronda
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex <> 0 Then
                Call WriteLocaleMsg(.Jugadores(i).ArrayIndex, MSG_RONDA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, nombres) ' Msg1970=Esta ronda es para ¬1.
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, "", e_TextChannel.TEXTCHANNEL_EVENT, 0) ' Dejamos un espacio vertical
            End If
        Next
        ' Iniciamos la próxima ronda
        Call iniciarRonda(Sala)
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.ProcesarRondaGanada", Erl)
End Sub

Public Sub FinalizarReto(ByVal Sala As Integer, Optional ByVal TiempoAgotado As Boolean)
    On Error GoTo ErrorHandler
    With Retos.Salas(Sala)
        ' Plan 05.002 ola 1 (patron de cierre, paso 1): marcar terminado ANTES de pagar.
        ' El timer de 10s (frmMain.TiempoRetos_Timer) vuelve a entrar aca mientras la sala
        ' siga EnUso, y en un reto con CaenItems la sala NO se libera al finalizar (queda
        ' el minuto de looteo): sin este guard el pozo se pagaba una y otra vez.
        If .Finalizado Then Exit Sub
        .Finalizado = True
        ' Calculamos el oro total del premio
        Dim OroTotal As Long, Oro As Long, OroStr As String
        OroTotal = .Apuesta * (UBound(.Jugadores) + 1)
        ' Descontamos el impuesto
        OroTotal = OroTotal * (1 - Retos.ImpuestoApuesta)
        Dim i As Integer, tUser As t_UserReference, Equipo1 As String, Equipo2 As String
        Dim Presentes As Integer
        ' Sa es el puntaje del equipo IZQUIERDO: 1 gano, 0.5 empate, 0 perdio.
        Dim Sa As Double
        If .Puntaje = 0 Then
            Sa = 0.5
        ElseIf .Puntaje < 0 Then
            Sa = 1
        Else
            Sa = 0
        End If
        ' El ELO se liquida ACA, antes de pagar el pozo y antes de tocar la sala. Antes se
        ' hacia al final del Sub, DESPUES de SalaLiberada/IniciarDepositoItems, y
        ' SalaLiberada arranca el reto que estaba esperando en la cola: el ReDim de
        ' IniciarReto pisaba .Jugadores y el rating terminaba cayendole a los jugadores del
        ' reto SIGUIENTE.
        Call AplicarEloReto(Sala, Sa)
        ' Empate
        If .Puntaje = 0 Then
            ' Pagamos a todos los que no abandonaron. Antes se dividia por la cantidad
            ' ORIGINAL de jugadores pero solo cobraban los presentes: la parte del que
            ' abandonaba no la cobraba nadie y desaparecia del pozo.
            For i = 0 To UBound(.Jugadores)
                If IsValidUserRef(.Jugadores(i)) Then Presentes = Presentes + 1
            Next i
            If Presentes > 0 Then Oro = OroTotal \ Presentes
            OroStr = PonerPuntos(Oro)
            For i = 0 To UBound(.Jugadores)
                tUser = .Jugadores(i)
                If IsValidUserRef(tUser) Then
                    UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + Oro
                    Call WriteUpdateGold(tUser.ArrayIndex)
                    Call WriteLocaleMsg(tUser.ArrayIndex, MSG_HAS_GANADO_MONEDAS_DE_ORO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, OroStr) ' Has ganado X monedas de oro
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    Call DevolverPosAnterior(tUser.ArrayIndex)
                    ' Reset flags
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    UserList(tUser.ArrayIndex).flags.EnReto = False
                    ' Nombres
                    If i Mod 2 Then
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo2 = UserList(tUser.ArrayIndex).name
                        End If
                    Else
                        If LenB(Equipo1) > 0 Then
                            Equipo1 = Equipo1 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo1 = UserList(tUser.ArrayIndex).name
                        End If
                    End If
                End If
            Next
            ' Anuncio global
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_VS_NINGUNO_PUDO_VENCER_RIVAL, Equipo1 & "¬" & Equipo2, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1670=Retos » ¬1 vs ¬2. Ninguno pudo vencer a su rival.
            Call SalaLiberada(Sala)
            ' Hubo un ganador
        Else
            Dim Ganador As e_EquipoReto
            If .Puntaje < 0 Then
                Ganador = e_EquipoReto.Izquierda
            Else
                Ganador = e_EquipoReto.Derecha
            End If
            ' Pagamos a los ganadores que no abandonaron
            Presentes = ObtenerTamañoEquipo(Sala, Ganador)
            If Presentes > 0 Then Oro = OroTotal \ Presentes
            OroStr = PonerPuntos(Oro)
            For i = 0 To UBound(.Jugadores)
                tUser = .Jugadores(i)
                If IsValidUserRef(tUser) Then
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    If UserList(tUser.ArrayIndex).flags.EquipoReto = Ganador Then
                        UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + Oro
                        Call WriteUpdateGold(tUser.ArrayIndex)
                        Call WriteLocaleMsg(tUser.ArrayIndex, MSG_HAS_GANADO_MONEDAS_DE_ORO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, OroStr) ' Has ganado X monedas de oro
                        If .CaenItems Then
                            Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                        Else
                            UserList(tUser.ArrayIndex).flags.EnReto = False
                            Call DevolverPosAnterior(tUser.ArrayIndex)
                        End If
                    Else
                        If .CaenItems Then
                            Call TirarItemsEnPos(tUser.ArrayIndex, ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x, ((.PosDerecha.y - .PosIzquierda.y) \ 2) + _
                                    .PosIzquierda.y)
                        End If
                        UserList(tUser.ArrayIndex).flags.EnReto = False
                        Call DevolverPosAnterior(tUser.ArrayIndex)
                    End If
                    ' Reset flags
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    If TiempoAgotado Then
                        Call WriteLocaleMsg(tUser.ArrayIndex, MSG_AGOTADO_TIEMPO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1971=Se ha agotado el tiempo del reto.
                    End If
                    ' Nombres
                    If i Mod 2 Then
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo2 = UserList(tUser.ArrayIndex).name
                        End If
                    Else
                        If LenB(Equipo1) > 0 Then
                            Equipo1 = Equipo1 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo1 = UserList(tUser.ArrayIndex).name
                        End If
                    End If
                End If
            Next
            Dim equipoGanador As String, equipoPerdedor As String
            equipoGanador = IIf(Ganador = e_EquipoReto.Izquierda, Equipo1, Equipo2)
            equipoPerdedor = IIf(Ganador = e_EquipoReto.Izquierda, Equipo2, Equipo1)
            ' Anuncio global
            If UBound(.Jugadores) > 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_EQUIPO_VENCIO_EQUIPO_QUEDO_BOTIN_MONEDAS, equipoGanador & "¬" & equipoPerdedor & "¬" & PonerPuntos(OroTotal), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1671=Retos » El equipo ¬1 venció al equipo ¬2 y se quedó con el botín de: ¬3 monedas de oro.
            Else ' 1 vs 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_VENCIO_QUEDO_BOTIN_MONEDAS_ORO, equipoGanador & "¬" & equipoPerdedor & "¬" & PonerPuntos(OroTotal), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1672=Retos » ¬1 venció a ¬2 y se quedó con el botín de: ¬3 monedas de oro.
            End If
            If .CaenItems Then
                Call IniciarDepositoItems(Sala)
            Else
                Call SalaLiberada(Sala)
            End If
        End If
        ' El calculo del ELO se movio ARRIBA, a AplicarEloReto: corria despues de
        ' SalaLiberada / IniciarDepositoItems, o sea despues de que la sala pudiera haber
        ' sido reasignada al reto que esperaba en la cola.
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.FinalizarReto", Erl)
End Sub

Public Sub TirarItemsEnPos(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo TirarItemsEnPos_Err
    Dim i         As Byte
    Dim NuevaPos  As t_WorldPos
    Dim MiObj     As t_Obj
    Dim ItemIndex As Integer
    Dim posItems  As t_WorldPos
    With UserList(UserIndex)
        posItems.Map = .pos.Map
        posItems.x = x
        posItems.y = y
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And Not ItemNewbieProtegidoAlMorir(UserIndex, ItemIndex) Then
                    NuevaPos.x = 0
                    NuevaPos.y = 0
                    MiObj.amount = .invent.Object(i).amount
                    MiObj.ObjIndex = ItemIndex
                    MiObj.ElementalTags = .invent.Object(i).ElementalTags
                    Call Tilelibre(posItems, NuevaPos, MiObj, True, True, False)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                        Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                        '  Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                    Else
                        Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
                        Call UpdateUserInv(False, UserIndex, i)
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
TirarItemsEnPos_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarItemsEnPos", Erl)
End Sub

Public Sub IniciarDepositoItems(ByVal Sala As Integer)
    ' Este Sub no tenia On Error propio y el For de abajo indexaba UserList(0), que no
    ' existe (el array arranca en 1) cuando alguien habia abandonado y su slot quedaba en
    ' cero. El error subia a FinalizarReto, que abortaba SIN spawnear el banquero y SIN
    ' setear TiempoItems: la sala quedaba EnUso para siempre.
    On Error GoTo ErrHandler
    Dim i       As Byte
    Dim Ganador As e_EquipoReto
    With Retos.Salas(Sala)
        If .Puntaje < 0 Then
            Ganador = e_EquipoReto.Izquierda
        Else
            Ganador = e_EquipoReto.Derecha
        End If
        For i = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(i)) Then
                If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Ganador Then
                    Call WriteLocaleMsg(.Jugadores(i).ArrayIndex, MSG_TIENES_MINUTO_LEVANTAR_ITEMS_PISO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1972=Tienes 1 minuto para levantar los items del piso.
                End If
            End If
        Next i
        Dim pos As t_WorldPos
        pos.Map = .PosIzquierda.Map
        pos.x = ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x
        pos.y = ((.PosDerecha.y - .PosIzquierda.y) \ 2) + .PosIzquierda.y
        'Spawneo un banquero.
        .IndexBanquero = SpawnNpc(3, pos, True, False)
        ' Retos.dat ya traia TiempoGuardarItems y el codigo lo ignoraba. El #If DEBUGGING que
        ' estaba aca dejaba el fix INERTE: Server.VBP declara DEBUGGING = 1 (y VB6 /make
        ' ignora el /d de la linea de comandos), asi que se compilaba el 20 hardcodeado.
        ' Ahora el .dat es la unica verdad.
        .TiempoItems = Retos.TiempoGuardarItems
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.IniciarDepositoItems", Erl)
    ' Pase lo que pase la sala tiene que quedar en un estado que el timer pueda cerrar.
    If Retos.Salas(Sala).TiempoItems <= 0 Then Retos.Salas(Sala).TiempoItems = Retos.TiempoGuardarItems
End Sub

Public Sub TerminarTiempoAgarrarItems(ByVal Sala As Integer)
    ' Plan 05.002 ola 1: esta Sub NO tenia On Error. Si QuitarNPC fallaba (banquero ya
    ' muerto o indice rancio) el error subia al timer de retos y abortaba el tick de TODAS
    ' las salas restantes, dejando esta sala EnUso para siempre.
    On Error GoTo ErrHandler
    Dim Ganador As e_EquipoReto
    With Retos.Salas(Sala)
        ' Idempotente: se puede llegar aca por el timer y por AbandonarReto en el mismo tick.
        If Not .EnUso Then Exit Sub
        'Mato al banquero (solo si sigue vivo; despues se olvida el indice para no
        'reintentar sobre un NPC ajeno que herede ese slot)
        If .IndexBanquero > 0 Then
            Call QuitarNPC(.IndexBanquero, eChallenge)
            .IndexBanquero = 0
        End If
        If .Puntaje < 0 Then
            Ganador = e_EquipoReto.Izquierda
        Else
            Ganador = e_EquipoReto.Derecha
        End If
        Dim i As Byte
        For i = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(i)) Then
                If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Ganador Then
                    UserList(.Jugadores(i).ArrayIndex).flags.EnReto = False
                    Call DevolverPosAnterior(.Jugadores(i).ArrayIndex)
                End If
            End If
        Next i
        .TiempoItems = 0
        Dim x As Byte
        Dim y As Byte
        For x = .PosIzquierda.x To .PosDerecha.x
            For y = .PosIzquierda.y To .PosDerecha.y
                Call EraseObj(GetMaxInvOBJ(), .PosIzquierda.Map, x, y)
            Next y
        Next x
    End With
    Call SalaLiberada(Sala)
    Exit Sub
ErrHandler:
    ' Aun si algo falla a mitad, la sala se libera: una sala colgada EnUso no se recupera
    ' sin reiniciar el server.
    Call TraceError(Err.Number, Err.Description, "ModRetos.TerminarTiempoAgarrarItems", Erl)
    Call SalaLiberada(Sala)
End Sub

Public Sub AbandonarReto(ByVal UserIndex As Integer, Optional ByVal Desconexion As Boolean)
    Dim Sala As Integer, Equipo As e_EquipoReto
    With UserList(UserIndex)
        Sala = .flags.SalaReto
        Equipo = .flags.EquipoReto
        .Counters.CuentaRegresiva = -1
        .flags.EnReto = False
    End With
    With Retos.Salas(Sala)
        ' El proxy viejo era Abs(.Puntaje) >= 2, que NO cubre el reto ganado por tiempo
        ' agotado (termina 1-0): un ganador que se desconectaba durante el minuto de looteo
        ' entraba por la rama de abajo, descontaba mal el equipo y dejaba el looteo abierto.
        ' .Finalizado dice exactamente "este reto ya se cerro".
        If .CaenItems And .Finalizado Then
            ' Fase de looteo: solo cuenta que se vayan los GANADORES (son los unicos que
            ' pueden levantar). Antes se descontaba al equipo ganador aunque el que se iba
            ' fuera del equipo perdedor, y el contador podia irse a negativo y re-disparar
            ' el cierre. Plan 05.002 ola 1.
            Dim GanadorLooteo As e_EquipoReto
            GanadorLooteo = IIf(.Puntaje < 0, e_EquipoReto.Izquierda, e_EquipoReto.Derecha)
            If Equipo = GanadorLooteo Then
                If GanadorLooteo = e_EquipoReto.Izquierda Then
                    If .TamañoEquipoIzq > 0 Then .TamañoEquipoIzq = .TamañoEquipoIzq - 1
                    If .TamañoEquipoIzq <= 0 Then Call TerminarTiempoAgarrarItems(Sala)
                Else
                    If .TamañoEquipoDer > 0 Then .TamañoEquipoDer = .TamañoEquipoDer - 1
                    If .TamañoEquipoDer <= 0 Then Call TerminarTiempoAgarrarItems(Sala)
                End If
            End If
            Exit Sub
        End If
        ' El que se va paga la derrota, y se le cobra ACA, antes de cualquier bifurcacion:
        ' en 1v1 la rama de abajo llama a FinalizarReto y sale, y en equipos la de mas abajo
        ' le borra la referencia del array; en ambos casos despues ya no hay a quien cobrarle.
        ' Vale para abandono voluntario y desconexion (TCP.ClearAndSaveUser entra por aca
        ' antes de SaveUser).
        Call SaldarEloAbandono(Sala, UserIndex, Equipo)
        If Not Desconexion Then
            Call WriteLocaleMsg(UserIndex, MSG_HAS_ABANDONADO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1973=Has abandonado el reto.
        End If
        ' Restamos un miembro al equipo y si llega a cero entonces procesamos la derrota
        If Equipo = e_EquipoReto.Izquierda Then
            If .TamañoEquipoIzq > 1 Then
                .TamañoEquipoIzq = .TamañoEquipoIzq - 1
            Else
                .Puntaje = 123 ' Forzamos puntaje positivo
                Call FinalizarReto(Sala)
                Exit Sub
            End If
        Else
            If .TamañoEquipoDer > 1 Then
                .TamañoEquipoDer = .TamañoEquipoDer - 1
            Else
                .Puntaje = -123 ' Forzamos puntaje negativo
                Call FinalizarReto(Sala)
                Exit Sub
            End If
        End If
        Call RevivirYLimpiar(UserIndex)
        Call DevolverPosAnterior(UserIndex)
        Dim texto As String
        If Desconexion Then
            texto = UserList(UserIndex).name & " es descalificado por desconectarse."
        Else
            texto = UserList(UserIndex).name & " ha abandonado el reto."
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex = UserIndex Then
                Call SetUserRef(.Jugadores(i), 0)
            Else
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, texto, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            End If
        Next
    End With
End Sub

Private Sub SalaLiberada(ByVal Sala As Integer)
    On Error GoTo ErrHandler
    ' Plan 05.002 ola 1: doble liberacion inflaba SalasLibres y podia arrancar un reto de
    ' la cola sobre una sala que otro flujo ya estaba reusando.
    If Not Retos.Salas(Sala).EnUso Then Exit Sub
    Retos.Salas(Sala).EnUso = False
    Retos.Salas(Sala).TiempoItems = 0
    Retos.SalasLibres = Retos.SalasLibres + 1
    ' Plan 05.002 ola 1: descartar de la cola a los que ya no pueden arrancar (se
    ' desconectaron o cancelaron). Antes se le iniciaba el reto igual.
    Do While ListaDeEspera.count > 0
        Dim Oferente As Integer
        Oferente = ListaDeEspera.Keys(0)
        Call ListaDeEspera.Remove(Oferente)
        If Oferente > 0 Then
            If UserList(Oferente).flags.UserLogged Then
                If UserList(Oferente).flags.SolicitudReto.Estado = e_SolicitudRetoEstado.EnCola Then
                    UserList(Oferente).flags.SolicitudReto.Estado = e_SolicitudRetoEstado.Libre
                    Call IniciarReto(Oferente, Sala)
                    Exit Do
                End If
            End If
        End If
    Loop
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.SalaLiberada", Erl)
End Sub

Public Function PuedeReto(ByVal UserIndex As Integer) As Boolean
    'Mismas reglas que PuedeRetoConMensaje (gate final consistente con el de entrada):
    'esCiudad en vez de Seguro + jugando_captura. Asi un reto aceptado no se cancela
    'silenciosamente por un criterio de mapa distinto.
    With UserList(UserIndex)
        If .flags.EnReto Then Exit Function
        If .flags.EnConsulta Then Exit Function
        If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then Exit Function
        If .flags.jugando_captura = 1 Then Exit Function
        If Not esCiudad(.pos.Map) Then Exit Function
        If .flags.EnTorneo Then Exit Function
        If MapData(.pos.x, .pos.y, .pos.Map).trigger = CARCEL Then Exit Function
    End With
    PuedeReto = True
End Function

Public Function PuedeRetoConMensaje(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        ' 07.001: paridad con PuedeReto, que si valida la posicion. Un UserIndex sin
        ' posicion valida no puede retar (y ademas evita indexar MapData con mapa 0).
        If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then Exit Function
        If .flags.EnReto Then
            Call WriteLocaleMsg(UserIndex, MSG_ENCUENTRAS_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1974=Ya te encuentras en un reto.
            Exit Function
        End If
        If .flags.EnConsulta Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_ACCEDER_RETO_CONSULTA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1975=No puedes acceder a un reto si estás en consulta.
            Exit Function
        End If
        If .flags.jugando_captura = 1 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_JUGAR_RETO_ESTANDO_EVENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1976=No puedes jugar un reto estando en un evento.
            Exit Function
        End If
        If Not esCiudad(.pos.Map) Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_PARTICIPAR_RETO_MAPA_INSEGURO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1977=No puedes participar de un reto en un mapa inseguro.
            Exit Function
        End If
        If .flags.EnTorneo Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_IR_RETO_PARTICIPAS_TORNEO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1978=No puedes ir a un reto si participas de un torneo.
            Exit Function
        End If
        If MapData(.pos.x, .pos.y, .pos.Map).trigger = CARCEL Then
            Call WriteLocaleMsg(UserIndex, MSG_ENCARCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1979=¡Estás encarcelado!
            Exit Function
        End If
    End With
    PuedeRetoConMensaje = True
End Function

Private Function IndiceJugadorEnSolicitud(ByVal UserIndex As Integer, ByVal Oferente As Integer) As Integer
    With UserList(Oferente).flags.SolicitudReto
        IndiceJugadorEnSolicitud = -1
        If .Estado <> e_SolicitudRetoEstado.Enviada Then Exit Function
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).nombre = UserList(UserIndex).name Then
                IndiceJugadorEnSolicitud = i
                Exit Function
            End If
        Next
    End With
End Function

Private Sub MensajeATodosSolicitud(ByVal Oferente As Integer, mensaje As String, ByVal Fuente As e_FontTypeNames)
    With UserList(Oferente).flags.SolicitudReto
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).Aceptado Then
                Call WriteConsoleMsg(.Jugadores(i).CurIndex.ArrayIndex, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, Fuente)
            End If
        Next
        Call WriteConsoleMsg(Oferente, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, Fuente)
    End With
End Sub

Private Function TodosPuedenReto(ByVal Oferente As Integer) As Boolean
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Not PuedeReto(Oferente) Then
            Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " no puede entrar al reto en este momento.")
            Exit Function
        ElseIf UserList(Oferente).Stats.GLD < .Apuesta Then
            Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " no tiene las monedas de oro suficientes.")
            Exit Function
        ElseIf .PocionesMaximas >= 0 Then
            If TieneObjetos(38, .PocionesMaximas + 1, Oferente) Then
                Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                Exit Function
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If Not PuedeReto(.Jugadores(i).CurIndex.ArrayIndex) Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no puede entrar al reto en este momento.")
                Exit Function
            ElseIf UserList(.Jugadores(i).CurIndex.ArrayIndex).Stats.GLD < .Apuesta Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no tiene las monedas de oro suficientes.")
                Exit Function
            ElseIf .PocionesMaximas >= 0 Then
                If TieneObjetos(38, .PocionesMaximas + 1, .Jugadores(i).CurIndex.ArrayIndex) Then
                    Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & _
                            .PocionesMaximas & ").")
                    Exit Function
                End If
            End If
        Next
        TodosPuedenReto = True
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.TodosPuedenReto", Erl)
End Function

Private Function EquipoContrario(ByVal Equipo As e_EquipoReto) As e_EquipoReto
    If Equipo = e_EquipoReto.Izquierda Then
        EquipoContrario = e_EquipoReto.Derecha
    Else
        EquipoContrario = e_EquipoReto.Izquierda
    End If
End Function

Private Function ObtenerTamañoEquipo(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto) As Integer
    If Equipo = e_EquipoReto.Izquierda Then
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoIzq
    Else
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoDer
    End If
End Function

Private Sub RevivirYLimpiar(ByVal UserIndex As Integer)
    Call WriteStopped(UserIndex, False)
    ' 06.002 Ola 1: los DoTs/slows del motor y los caches de veneno no pasan de una ronda a otra
    ' (LimpiarEstadosAlterados solo cubre los estados legacy). Los encantamientos de arma/flechas
    ' se conservan: son preparacion del jugador, igual que sus pociones.
    Call ClearEffectList(UserList(UserIndex).EffectOverTime, e_EffectType.eAny, False)
    Call ClearModifiers(UserList(UserIndex).Modifiers)
    Call ClearPoisonEffectCaches(UserIndex)
    ' Si está vivo
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call LimpiarEstadosAlterados(UserIndex)
    End If
    ' Si está muerto lo revivimos, sino lo curamos
    Call RevivirUsuario(UserIndex)
End Sub

' ---------------------------------------------------------------------------------------
' ELO (plan 27.001 ola 6 / 05.002)
' ---------------------------------------------------------------------------------------
' La formula vieja no era Elo: eloDiff = wins * (eloTotalRival * 0.1). Ahora es Elo
' clasico con K fijo (Retos.FactorK): el delta se calcula UNA sola vez para el equipo
' izquierdo y se niega para el derecho, asi ningun redondeo rompe la simetria.

' Delta ENTERO que le corresponde al equipo IZQUIERDO completo.
' Sa: 1 gano la izquierda, 0.5 empate, 0 perdio la izquierda.
Private Function DeltaEloEquipoIzq(ByVal Sala As Integer, ByVal Sa As Double) As Long
    Dim Exponente As Double, Ea As Double
    With Retos.Salas(Sala)
        ' Promedio y no suma: el 400 del Elo mide diferencia de rating, no de plantel. La
        ' ventaja numerica ya se cobra al repartir el delta entre mas jugadores.
        Exponente = (CDbl(.EloPromedioDer) - CDbl(.EloPromedioIzq)) / 400#
    End With
    ' Con ratings absurdos (dato corrupto) 10 ^ Exponente desborda el Double y tira error
    ' adentro de FinalizarReto. Saturamos.
    If Exponente > 40# Then Exponente = 40#
    If Exponente < -40# Then Exponente = -40#
    Ea = 1# / (1# + 10# ^ Exponente)
    DeltaEloEquipoIzq = CLng(CDbl(Retos.FactorK) * (Sa - Ea))
End Function

' Reparte Total entre Cantidad miembros sin perder ni inventar puntos: el resto se da de a
' uno a los primeros miembros, asi la suma de las cuotas es EXACTAMENTE Total. Funciona
' igual con Total negativo porque en VB6 "\" trunca hacia cero y Mod conserva el signo del
' dividendo (-15 \ 2 = -7 y -15 Mod 2 = -1, o sea -8 y -7, que suman -15).
Private Function CuotaElo(ByVal Total As Long, ByVal Cantidad As Integer, ByVal Orden As Integer) As Long
    If Cantidad <= 0 Then Exit Function
    Dim Base As Long, Resto As Long
    Base = Total \ Cantidad
    Resto = Total Mod Cantidad
    CuotaElo = Base
    If Orden < Abs(Resto) Then CuotaElo = Base + Sgn(Resto)
End Function

' Unico punto del modulo que escribe Stats.ELO.
Private Sub AjustarElo(ByVal UserIndex As Integer, ByVal Delta As Long)
    If Delta = 0 Then Exit Sub
    With UserList(UserIndex).Stats
        ' El rating no puede quedar negativo. Es el UNICO lugar donde el sistema deja de
        ' sumar cero (al recortar la perdida se inyectan hasta FactorK - 1 puntos), y solo
        ' le pasa a alguien con menos de FactorK.
        If .ELO + Delta < 0 Then Delta = -.ELO
        .ELO = .ELO + Delta
    End With
    If Delta > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MSG_HAS_GANADO_PUNTOS_ELO, CStr(Delta), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1695=Has ganado ¬1 puntos de ELO!
    ElseIf Delta < 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MSG_HAS_PERDIDO_PUNTOS_ELO, CStr(-Delta), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1696=Has perdido ¬1 puntos de ELO!
    End If
End Sub

' Liquida el ELO de todos los que siguen en la sala y todavia no saldaron.
Private Sub AplicarEloReto(ByVal Sala As Integer, ByVal Sa As Double)
    On Error GoTo ErrHandler
    Dim i As Integer, DeltaIzq As Long, DeltaDer As Long
    With Retos.Salas(Sala)
        If Not .EloHabilitado Then
            For i = 0 To UBound(.Jugadores)
                If IsValidUserRef(.Jugadores(i)) Then
                    Call SendData(SendTarget.ToIndex, .Jugadores(i).ArrayIndex, PrepareMessageLocaleMsg(MSG_PARTICIPANTE_RETO_TIENE_NIVEL_MENOR_ELO_PERMANECE, CStr(Retos.NivelMinimoELO), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
                End If
            Next i
            Exit Sub
        End If
        ' El divisor es el PLANTEL ORIGINAL, no los que quedan: con los presentes, el que se
        ' QUEDA pagaba la parte del que se fue (en un 5v5 con 4 desconectados, el unico leal
        ' comia el delta completo). El ordinal es estable (i \ 2, da 0,1,2 tanto para pares
        ' como impares): el que se fue y el que se quedo usan ordinales DISJUNTOS y las
        ' cuotas suman exactamente el delta del equipo.
        DeltaIzq = DeltaEloEquipoIzq(Sala, Sa)
        DeltaDer = -DeltaIzq
        For i = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(i)) And Not .EloSaldado(i) Then
                .EloSaldado(i) = True
                If i Mod 2 = 0 Then
                    Call AjustarElo(.Jugadores(i).ArrayIndex, CuotaElo(DeltaIzq, CInt(.TamañoInicialIzq), i \ 2))
                Else
                    Call AjustarElo(.Jugadores(i).ArrayIndex, CuotaElo(DeltaDer, CInt(.TamañoInicialDer), i \ 2))
                End If
            End If
        Next i
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AplicarEloReto", Erl)
End Sub

' Piso del castigo por abandono. El Elo puro devuelve 0 cuando la derrota estaba
' completamente descontada (a partir de unos 669 puntos de diferencia a favor del rival),
' y ahi abandonar seria gratis. Un punto de deflacion por abandono es barato al lado de eso.
Private Function PisoAbandono(ByVal Cuota As Long) As Long
    PisoAbandono = Cuota
    If Cuota = 0 Then PisoAbandono = -1
End Function

' Le cobra la derrota al que abandona o se desconecta: el delta que le habria tocado si su
' equipo perdia, dividido por el Tamaño ORIGINAL de su equipo. A proposito NO se le acredita
' a nadie (dos cuentas amigas farmearian rating entrando y saliendo): en equipos el abandono
' destruye puntos. En 1v1 el que se queda cobra esa misma cantidad en FinalizarReto.
Private Sub SaldarEloAbandono(ByVal Sala As Integer, ByVal UserIndex As Integer, ByVal Equipo As e_EquipoReto)
    On Error GoTo ErrHandler
    Dim i As Integer, Total As Long
    With Retos.Salas(Sala)
        If Not .EloHabilitado Then Exit Sub
        If .Finalizado Then Exit Sub
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex = UserIndex Then
                If .EloSaldado(i) Then Exit Sub
                .EloSaldado(i) = True
                ' Mismo ordinal (i \ 2) que usara AplicarEloReto para los que se quedan.
                If Equipo = e_EquipoReto.Izquierda Then
                    Total = DeltaEloEquipoIzq(Sala, 0)
                    Call AjustarElo(UserIndex, PisoAbandono(CuotaElo(Total, CInt(.TamañoInicialIzq), i \ 2)))
                Else
                    Total = -DeltaEloEquipoIzq(Sala, 1)
                    Call AjustarElo(UserIndex, PisoAbandono(CuotaElo(Total, CInt(.TamañoInicialDer), i \ 2)))
                End If
                Exit Sub
            End If
        Next i
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.SaldarEloAbandono", Erl)
End Sub

' ---------------------------------------------------------------------------------------
' Enfriamiento entre los mismos rivales (plan 27.001 ola 6 / 05.002)
' ---------------------------------------------------------------------------------------

' Clave canonica del enfrentamiento: nombres en mayusculas y ordenados, para que A vs B y
' B vs A sean el MISMO enfrentamiento. Si cambia un solo participante la clave cambia.
Private Function ClaveEnfrentamiento(ByVal Oferente As Integer) As String
    Dim Nombres() As String
    Dim i As Integer, j As Integer, Tmp As String
    With UserList(Oferente).flags.SolicitudReto
        ReDim Nombres(0 To UBound(.Jugadores) + 1)
        Nombres(0) = UCase$(UserList(Oferente).name)
        For i = 0 To UBound(.Jugadores)
            Nombres(i + 1) = UCase$(.Jugadores(i).nombre)
        Next i
    End With
    ' Insercion directa: son 10 nombres como maximo.
    For i = 1 To UBound(Nombres)
        Tmp = Nombres(i)
        j = i - 1
        Do While j >= 0
            If StrComp(Nombres(j), Tmp, vbBinaryCompare) <= 0 Then Exit Do
            Nombres(j + 1) = Nombres(j)
            j = j - 1
        Loop
        Nombres(j + 1) = Tmp
    Next i
    ClaveEnfrentamiento = Join(Nombres, "|")
End Function

' Minutos que faltan para que este mismo grupo pueda volver a retarse. 0 = puede.
Private Function MinutosParaRepetir(ByVal Clave As String) As Long
    Dim i As Integer, Transcurrido As Double, Ventana As Double
    If Retos.MaximoRetosSeguidos <= 0 Then Exit Function
    Ventana = CDbl(Retos.MinutosEnfriamiento) * 60000#
    If Ventana <= 0 Then Exit Function
    For i = 0 To MAX_HISTORIAL_RETOS - 1
        If HistorialRetos(i).Clave = Clave Then
            ' TicksElapsed es a prueba del wrap de 32 bits de timeGetTime (49.7 dias).
            Transcurrido = TicksElapsed(HistorialRetos(i).UltimoTick, GetTickCountRaw())
            If Transcurrido >= Ventana Then
                ' Paso el enfriamiento: el contador vuelve a cero.
                HistorialRetos(i).Clave = vbNullString
                HistorialRetos(i).Repeticiones = 0
                Exit Function
            End If
            If HistorialRetos(i).Repeticiones < Retos.MaximoRetosSeguidos Then Exit Function
            ' Redondeamos para arriba: "faltan 0 minutos" y rebotarlo igual es peor.
            MinutosParaRepetir = CLng(Fix((Ventana - Transcurrido) / 60000#)) + 1
            Exit Function
        End If
    Next i
End Function

Private Sub RegistrarEnfrentamiento(ByVal Clave As String)
    Dim i As Integer, Libre As Integer, MasViejo As Integer
    Dim Ahora As Long, Ventana As Double
    If LenB(Clave) = 0 Then Exit Sub
    Ahora = GetTickCountRaw()
    Ventana = CDbl(Retos.MinutosEnfriamiento) * 60000#
    Libre = -1
    MasViejo = 0
    For i = 0 To MAX_HISTORIAL_RETOS - 1
        If HistorialRetos(i).Clave = Clave Then
            If TicksElapsed(HistorialRetos(i).UltimoTick, Ahora) >= Ventana Then
                ' La ventana ya vencio (reto encolado o pedidos simultaneos): cuenta de nuevo desde 1.
                HistorialRetos(i).Repeticiones = 1
            Else
                HistorialRetos(i).Repeticiones = HistorialRetos(i).Repeticiones + 1
            End If
            HistorialRetos(i).UltimoTick = Ahora
            Exit Sub
        End If
        If Libre < 0 Then
            If LenB(HistorialRetos(i).Clave) = 0 Then
                Libre = i
            ElseIf TicksElapsed(HistorialRetos(i).UltimoTick, Ahora) >= Ventana Then
                Libre = i
            End If
        End If
        If TicksElapsed(HistorialRetos(i).UltimoTick, Ahora) > TicksElapsed(HistorialRetos(MasViejo).UltimoTick, Ahora) Then MasViejo = i
    Next i
    ' Si la tabla se lleno con enfrentamientos vivos pisamos el mas viejo (hacen falta 256
    ' grupos distintos en la ventana para llegar aca).
    If Libre < 0 Then Libre = MasViejo
    HistorialRetos(Libre).Clave = Clave
    HistorialRetos(Libre).Repeticiones = 1
    HistorialRetos(Libre).UltimoTick = Ahora
End Sub

' Solo tiene sentido frenar el win-trading donde hay rating en juego. Si alguno no llega al
' nivel minimo el reto no mueve ELO y no hay razon para restringirlo.
Private Function RetoOtorgaElo(ByVal Oferente As Integer) As Boolean
    Dim i As Integer
    If UserList(Oferente).Stats.ELV < Retos.NivelMinimoELO Then Exit Function
    With UserList(Oferente).flags.SolicitudReto
        For i = 0 To UBound(.Jugadores)
            If Not IsValidUserRef(.Jugadores(i).CurIndex) Then Exit Function
            If UserList(.Jugadores(i).CurIndex.ArrayIndex).Stats.ELV < Retos.NivelMinimoELO Then Exit Function
        Next i
    End With
    RetoOtorgaElo = True
End Function
