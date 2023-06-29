# Stanislav Stepanov

# Contacts
* Location: Russia, Khabarovsk
* E-mail: stanislav.msb@yandex.ru
* GitHub: https://github.com/Stdes

# Skills

* Assembler
* Borland Delphi
* MS Access
* MS Visual Basic
* SQL

# Code Examples
```
Private Sub ConfirmDocumentButton_Click()

If IsNull(Me.ClearanceID) Then
Exit Sub
End If

CurrentDocumentID = Me.ClearanceID

'спросим подтверждение
If MyMsg(TMsg("Are you sure about confirm document?"), vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
Exit Sub
End If

'On Error GoTo ErrorHandlingCode
Set rst = CurrentDb.OpenRecordset _
("SELECT ClearanceDetail.* FROM ClearanceDetail WHERE ClearanceID=" & GetCurrentDocumentID)
With rst
Do While Not .EOF

'изменим таможенный статус заблокированных на складе позиций 'ClearanceStatusID=2 (Passed)
'и разблокируем их RegisterStorageStatusID=1 (Ok), обнулим дату окончания нерастамож. хранения.
DoCmd.RunSQL "UPDATE RegisterStorage SET RegisterStorageStatusID = 1, " & _
"ClearanceStatusID = 2, ClearanceExpired = Null " & _
"WHERE RegisterStorageID=" & !RegisterStorageID

'зарегистрируем движение
Call RegisterTrafficInsert(3, 0, 0, 7, GetCurrentDocumentID(), 1, !ArrivalDetailID, 2, !Qty, !Weight, _
!ClearancePrice, !CurrencyID)

.MoveNext
Loop
.Close
End With

DoCmd.Close acForm, "ClearanceDetail"
'изменим статус документа DocumentStatusID=1 (Closed)
DoCmd.RunSQL "UPDATE Clearance SET DocumentStatusID=1 WHERE ClearanceID=" & GetCurrentDocumentID
'Me.DocumentStatusID = 1
Form_ClearanceDatasheet.Requery
DoCmd.OpenForm "ClearanceDetail", , , "ClearanceID=" & GetCurrentDocumentID, acFormReadOnly
MyMsg TMsg("Document has been confirmed!"), vbInformation

Exit Sub
End Sub
```

# Education
* Far Eastern State Transport University, Khabarovsk
    * Faculty of automation, telemechanics and communication, specialization of microprocessor's systems.

