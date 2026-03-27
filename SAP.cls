VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SapGui, SapGuiAuto, App, Connection, Session, WScript, WSHShell As Object
Public PeriodoInicial, PeriodoFinal As String

Public Sub SAP_Logon()
    'Abre o SAP
    Shell "C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe", vbNormalFocus
    
    ' Aguarda um pouco para o SAP abrir
    Application.Wait (Now + TimeValue("00:00:02"))
    
    
    '#######################
    'ABRE O SISTEMA
    Set WSHShell = CreateObject("WScript.Shell")
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.OpenConnection(Conexao, True)
    Set Session = Connection.Children(0)
    
    '#################
    'DADOS DO LOGIN NO SISTEMA
    Session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = Mandante 'Identificação do mandante
    Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = Worksheets("fbl3n").myUser 'Identificação do usuário
    Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = Worksheets("fbl3n").myPassword 'Identificação da senha do usuário
    Session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = Idioma   'Identificação do idioma
    Session.FindById("wnd[0]").SendVKey 0
End Sub

Public Sub Gerar_razao()
    Dim w As Worksheet
    Set w = Sheets("fbl3n")
    
    Session.FindById("wnd[0]").Maximize
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "fbl3n"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/rad[2]").Select
    Session.FindById("wnd[0]/usr/ctxt[2]").Text = w.Range("C8")
    Session.FindById("wnd[0]/usr/ctxt[11]").Text = w.Range("C9")
    Session.FindById("wnd[0]/usr/ctxt[12]").Text = w.Range("C10")
    Session.FindById("wnd[0]/usr/ctxt[12]").SetFocus
    Session.FindById("wnd[0]/usr/ctxt[12]").CaretPosition = 10
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    
    'aqui começa o clique no menu sistema
Session.FindById("wnd[0]").SendVKey 9
Session.FindById("wnd[1]/usr/sub/2/sub/2/1/rad[4,0]").Select
Session.FindById("wnd[1]/usr/sub/2/sub/2/1/rad[4,0]").SetFocus
Session.FindById("wnd[1]/tbar[0]/btn[0]").Press

Session.FindById("wnd[0]").Maximize
Session.FindById("wnd[0]").Close
Session.FindById("wnd[1]/usr/btn[0]").Press

End Sub

Public Sub Connect_sap()
    If Not IsObject(App) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set App = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = App.Children(0)
    End If
    If Not IsObject(Session) Then
       Set Session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject Session, "on"
       WScript.ConnectObject App, "on"
    End If
End Sub
