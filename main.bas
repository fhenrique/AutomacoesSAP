Attribute VB_Name = "main"
Sub ConectarSAP()
    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim Session As Object

    ' Abre o SAP Logon (ajuste o caminho se necessário)
    Shell "C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe", vbNormalFocus

    ' Aguarda um pouco para o SAP abrir
    Application.Wait (Now + TimeValue("00:00:02"))

    ' Conecta ao SAP GUI
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine

    ' Abre a conexão específica (substitua pelo nome no seu SAP Logon)
    Set Connection = App.OpenConnection("Stefanini - ECC", True)
    Set Session = Connection.Children(0)

    ' Login (substitua pelos seus dados)
    Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = "fhenrique"
    Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Senha@001"
    Session.FindById("wnd[0]").SendVKey 0 ' Pressionar Enter
End Sub


Sub me51n()
    
    Dim w As Worksheet
    Dim nRows As Long
    Dim contador As Integer
    Dim cel As Byte
    Dim mensagem As String
    Dim docNumero As String
    Dim celulaDoc As String
    Dim tempo1 As Date
    Dim tempo2 As Date
    Dim duracao As Date
    
    tempo1 = Time

    ConectarSAP

    Set w = Sheets("Lancto")

    nRows = w.Cells(w.Rows.Count, 2).End(xlUp).Row

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

    Session.FindById("wnd[0]").Maximize

    For contador = 7 To nRows
        Session.FindById("wnd[0]/tbar[0]/okcd").Text = "f-01"
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/usr/ctxt[2]").Text = w.Range("D" & contador) 'empresa
        Session.FindById("wnd[0]/usr/ctxt[4]").Text = "BRL"
        Session.FindById("wnd[0]/usr/txt[2]").Text = w.Range("E" & contador) 'referencia
        Session.FindById("wnd[0]/usr/ctxt[8]").Text = "40"
        Session.FindById("wnd[0]/usr/ctxt[9]").Text = w.Range("F" & contador) 'conta de débito
        Session.FindById("wnd[0]/usr/ctxt[9]").SetFocus
        Session.FindById("wnd[0]/usr/ctxt[9]").CaretPosition = 6
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/usr/txt[2]").Text = w.Range("H" & contador) 'montante
        Session.FindById("wnd[0]/usr/ctxt[4]").Text = w.Range("E" & contador) 'referencia
        Session.FindById("wnd[0]/usr/ctxt[5]").Text = "50"
        Session.FindById("wnd[0]/usr/ctxt[6]").Text = w.Range("G" & contador) 'conta de crédito
        Session.FindById("wnd[0]/usr/ctxt[6]").SetFocus
        Session.FindById("wnd[0]/usr/ctxt[6]").CaretPosition = 6
        Session.FindById("wnd[0]/tbar[0]/btn[11]").Press
        Session.FindById("wnd[0]/usr/txt[2]").Text = w.Range("H" & contador) 'montante
        Session.FindById("wnd[0]/usr/txt[2]").CaretPosition = 3
        Session.FindById("wnd[0]/tbar[0]/btn[11]").Press
        
        ' --- Retornar o número do documento ---
        ' Captura a mensagem na parte inferior da tela
        mensagem = Session.FindById("wnd[0]/sbar").Text
        
        docNumero = Mid(mensagem, 10, 11) ' Ajuste o número de caracteres se necessário
        
        celulaDoc = "I" & LTrim(RTrim(Str(contador)))
        
        Range(celulaDoc) = docNumero
        
        
        Session.FindById("wnd[0]/tbar[0]/btn[15]").Press
        Session.FindById("wnd[1]/usr/btn[0]").Press

        cel = cel + 1
    
    Next contador
    
    AppActivate "Automacao_f-02.xlsm - Excel"
    
    tempo2 = Time
    
    duracao = tempo2 - tempo1
    
    ' Retorna para o Excel na célula A1
    Sheets(1).Range("A1").Select
        
    MsgBox (nRows - 6) & " lançamentos realizados em: " & Format(duracao, "HH:mm:ss")
    
End Sub
