Attribute VB_Name = "main"
Sub Login_SAP()
    Set sap_auto = New SAP 'intancia um objeto da classe SAP
    sap_auto.SAP_Logon 'Faz login no SAP
    MsgBox "Login realizado com sucesso", vbInformation
End Sub

Sub Gerar_razao()
    'apago as linhas
    Rows("11:400").Select
    Selection.ClearContents
    
    Set sap_auto = New SAP
    sap_auto.SAP_Logon
    sap_auto.Connect_sap
    sap_auto.Gerar_razao
    
    'colo o conteúdo da ára de transferencia
    Range("B11").Select
    ActiveSheet.Paste
    
    Application.SendKeys "%{TAB}", True
    Application.SendKeys "{NUMLOCK}"
    
End Sub

Sub FocoForcado()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.AppActivate "Microsoft Excel"
    SendKeys "%{TAB}", True ' Envia Alt+Tab para alternar janelas
End Sub
