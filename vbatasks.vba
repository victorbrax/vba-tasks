Private Sub Worksheet_Change(ByVal Target As Range)

    Dim OutApp As Object
    Dim OutMail As Object
    Dim texto As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    linha = ActiveCell.Row - 1
    If Target.Address = "$J$" & linha Then
    
        If Sheet1.Cells(linha, 10) = "Em Progresso" And Sheet1.Cells(linha, 11) <> "SIM" Then
            texto = "@Callcenter Team," & vbCrLf & vbCrLf & _
            "Por favor, poderiam abrir um novo chamado?" & vbCrLf & vbCrLf & _
            "Cliente: " & Sheet1.Cells(linha, 1) & vbCrLf & _
            "Site: " & Sheet1.Cells(linha, 8) & vbCrLf & _
            "Serviço: " & Sheet1.Cells(linha, 6) & vbCrLf & _
            "Item de Configuração: " & Sheet1.Cells(linha, 7) & vbCrLf & _
            "Descrição: " & Sheet1.Cells(linha, 9) & vbCrLf & _
            "Tipo: " & "CHANGE" & vbCrLf & _
            "Data Inicial Estimada: " & Sheet1.Cells(linha, 2) & vbCrLf & _
            "Data Final Estimada: " & Sheet1.Cells(linha, 3) & vbCrLf & _
            "Prioridade: " & "4" & vbCrLf & _
            "Time: " & "PDCS_CM_SAP" & vbCrLf & vbCrLf & _
            "Obrigado!"

            With OutMail
                .To = "callcenter.brasil@sencinet.com" 'Caso nao funcione: .To = "" & Sheet1.Cells(linha, 1) & ""'
                .CC = ""
                .BCC = ""
                .Subject = Sheet1.Cells(linha, 9) 'Caso nao funcione: .To = "" & Sheet1.Cells(linha, 1) & ""'
                .Body = texto
                .Display
            End With
            Sheet1.Cells(linha, 11) = "SIM"
        End If
        
        On Error GoTo 0
        
        Set OutMail = Nothing
        Set OutApp = Nothing
    
    End If
End Sub