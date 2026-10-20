Attribute VB_Name = "modSafeRestart"
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
'    Copyright (C) 2002 Mrquez Pablo Ignacio
Option Explicit

' modSafeRestart
' Reinicio y apagado seguros del server sin intervencion manual (plan 09.001, Ola 1).
' ReinicioSeguro guarda a todos los conectados, avisa por consola, deja constancia
' en Logs\safe-restart.log, lanza restart_server.bat (que espera a que este proceso
' termine y relanza server.exe heredando la elevacion del padre) y recien entonces
' cierra por el mismo camino que "Guardar y cerrar" (Unload frmMain -> Form_Unload
' -> CerrarServidor). ApagadoSeguro es identico pero sin el Shell: el server queda
' apagado (pensado para el ciclo compilar->relanzar). MaybeRunSafeRestart sondea
' cada ~2s (patron de modUserAutoSave.MaybeRunUserAutoSave) los archivos de
' solicitud restart.request/shutdown.request. El comando GM por protocolo queda
' fuera de esta ola (ver plan 09.001, seccion "Fuera de alcance").

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const SAFE_RESTART_LOG_HEADER As String = "fecha_hora;evento;motivo;modo;pid"
Private Const SAFE_RESTART_POLL_INTERVAL_MS As Long = 2000

' True mientras un reinicio o apagado seguro esta en curso. Evita que un doble
' disparo (boton + archivo, o dos archivos seguidos) dispare el flujo dos veces.
Private m_ReinicioEnCurso As Boolean
Private m_LastPollTick As Long

''
' Sondea cada ~2s si hay un archivo de solicitud de reinicio/apagado en App.Path.
' Costo con nada que hacer: un Dir$() (via FileExist) cada 2 segundos, calcado del
' patron de MaybeRunUserAutoSave. Llamar desde el loop principal (General.bas).
Public Sub MaybeRunSafeRestart()
    On Error GoTo MaybeRunSafeRestart_Err

    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()

    If m_LastPollTick <> 0 Then
        If TicksElapsed(m_LastPollTick, nowRaw) < SAFE_RESTART_POLL_INTERVAL_MS Then Exit Sub
    End If
    m_LastPollTick = nowRaw

    Dim restartPath As String
    Dim shutdownPath As String
    restartPath = App.Path & "\restart.request"
    shutdownPath = App.Path & "\shutdown.request"

    If FileExist(restartPath) Then
        Dim motivoRestart As String
        Dim modoRestart As String
        Call LeerSolicitud(restartPath, motivoRestart, modoRestart)
        Kill restartPath
        Call ReinicioSeguro(motivoRestart, modoRestart)
        Exit Sub
    End If

    If FileExist(shutdownPath) Then
        Dim motivoShutdown As String
        Dim modoShutdown As String
        Call LeerSolicitud(shutdownPath, motivoShutdown, modoShutdown)
        Kill shutdownPath
        Call ApagadoSeguro(motivoShutdown)
        Exit Sub
    End If

    Exit Sub
MaybeRunSafeRestart_Err:
    Call TraceError(Err.Number, Err.Description, "modSafeRestart.MaybeRunSafeRestart", Erl)
End Sub

''
' Reinicio seguro: guarda a todos los conectados, avisa por consola, loguea,
' lanza restart_server.bat y recien entonces cierra el proceso actual (mismo
' camino que Command4_Click, sin duplicar CerrarServidor). Si el Shell falla,
' NO cierra el server: limpia el flag, loguea shell_fallo y sale.
Public Sub ReinicioSeguro(Optional ByVal Motivo As String = vbNullString, Optional ByVal Modo As String = "normal")
    On Error GoTo ReinicioSeguro_Err

    If m_ReinicioEnCurso Then
        Call LogSafeRestart("reinicio_ignorado", Motivo, Modo)
        Exit Sub
    End If
    m_ReinicioEnCurso = True

    Call LogSafeRestart("reinicio", Motivo, Modo)
    Call MensajeGlobal("Servidor > El servidor se va a reiniciar. Guardando personajes...", e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
    Call GuardarUsuarios
    Call EcharPjsNoPrivilegiados

    Dim shellResult As Double
    Dim shellErrNum As Long
    Dim shellErrDesc As String
    On Error Resume Next
    shellResult = Shell("cmd.exe /c """ & App.Path & "\restart_server.bat"" " & GetCurrentProcessId() & " " & Modo, vbHide)
    shellErrNum = Err.Number
    shellErrDesc = Err.Description
    On Error GoTo ReinicioSeguro_Err

    If shellErrNum <> 0 Or shellResult = 0 Then
        m_ReinicioEnCurso = False
        Call LogSafeRestart("shell_fallo", Motivo, Modo)
        Call TraceError(shellErrNum, shellErrDesc, "modSafeRestart.ReinicioSeguro (Shell)", Erl)
        Exit Sub
    End If

    frmMain.GuardarYCerrar = True
    Unload frmMain
    Exit Sub
ReinicioSeguro_Err:
    m_ReinicioEnCurso = False
    Call TraceError(Err.Number, Err.Description, "modSafeRestart.ReinicioSeguro", Erl)
End Sub

''
' Apagado seguro: identico a ReinicioSeguro pero sin lanzar el relanzador (el
' server queda apagado). Pensado para el ciclo compilar->relanzar (Ola 2/3).
Public Sub ApagadoSeguro(Optional ByVal Motivo As String = vbNullString)
    On Error GoTo ApagadoSeguro_Err

    If m_ReinicioEnCurso Then
        Call LogSafeRestart("apagado_ignorado", Motivo, "apagado")
        Exit Sub
    End If
    m_ReinicioEnCurso = True

    Call LogSafeRestart("apagado", Motivo, "apagado")
    Call MensajeGlobal("Servidor > El servidor se va a apagar. Guardando personajes...", e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
    Call GuardarUsuarios
    Call EcharPjsNoPrivilegiados

    frmMain.GuardarYCerrar = True
    Unload frmMain
    Exit Sub
ApagadoSeguro_Err:
    m_ReinicioEnCurso = False
    Call TraceError(Err.Number, Err.Description, "modSafeRestart.ApagadoSeguro", Erl)
End Sub

''
' Parsea el contenido opcional de un archivo de solicitud: lineas "motivo=..."
' y/o "modo=..." (cualquier orden). Vacio o sin esas claves = motivo vacio,
' modo "normal". No falla si el archivo esta vacio.
Private Sub LeerSolicitud(ByVal Ruta As String, ByRef Motivo As String, ByRef Modo As String)
    On Error GoTo LeerSolicitud_Err
    Motivo = vbNullString
    Modo = "normal"

    Dim fnum As Integer
    Dim linea As String
    Dim contenido As String
    fnum = FreeFile
    Open Ruta For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, linea
        contenido = contenido & linea & ";"
    Loop
    Close #fnum

    Dim partes() As String
    partes = Split(contenido, ";")
    Dim i As Long
    Dim parte As String
    For i = 0 To UBound(partes)
        parte = Trim$(partes(i))
        If LenB(parte) > 0 Then
            If LCase$(Left$(parte, 7)) = "motivo=" Then
                Motivo = Mid$(parte, 8)
            ElseIf LCase$(Left$(parte, 5)) = "modo=" Then
                Modo = Mid$(parte, 6)
            End If
        End If
    Next i
    Exit Sub
LeerSolicitud_Err:
    On Error Resume Next
    Close #fnum
    Call TraceError(Err.Number, Err.Description, "modSafeRestart.LeerSolicitud", Erl)
End Sub

''
' Append a Logs\safe-restart.log: fecha_hora;evento;motivo;modo;pid. Crea el
' archivo (y el header) si no existe. Mismo patron que
' modPoisonLogging.LogPoisonEvent: si falla, cierra en silencio sin recursar
' en TraceError (que tambien loguea a archivo).
Private Sub LogSafeRestart(ByVal Evento As String, ByVal Motivo As String, ByVal Modo As String)
    On Error GoTo LogSafeRestart_Err

    Dim fname As String
    fname = App.Path & "\Logs\safe-restart.log"

    Dim fnum As Integer
    fnum = FreeFile
    Dim writeHeader As Boolean
    writeHeader = (LenB(Dir$(fname)) = 0)

    Open fname For Append As #fnum
    If writeHeader Then Print #fnum, SAFE_RESTART_LOG_HEADER
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & ";" & _
                 Evento & ";" & Motivo & ";" & Modo & ";" & GetCurrentProcessId()
    Close #fnum
    Exit Sub
LogSafeRestart_Err:
    On Error Resume Next
    Close #fnum
End Sub
