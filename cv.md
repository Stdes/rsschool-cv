# Stanislav Stepanov

# Contacts
* Location: Russia, Khabarovsk
* E-mail: stanislav.msb@yandex.ru
* GitHub: https://github.com/Stdes
* Discord: Stanislav

# About myself
I am 41 years old electrical engineer. I had met programming in my school years starting with the Basic on Spectrum that my father crafted by blueprints. At university's internship I participated in a couple of project with programmable logic controllers and frequency converters. It's pretty interesting working on joint of of technologies: programming and electric drivers. Also I met with Assembler and studied Borland Delphi more than curriculum required just because I found its interesting. Never thought someone will pay me for this skill. 

However after I finished university I had rather found work of programmer than electrician engineer. So I got some experience developing visualisation programs of thermal power station's technological processes in Borland Delphi.

Few years I worked as system administrator of commercial electricity metering. That work was not much about coding more about servers, communication and electric equipment perhaps some shell scripts.

Next couple of years I spent solo-working as freelancer working on order to develop an ERP-program that allowed small  international company switch from excel's tables to automatization of warehouses accounting,  financial accounting and trading. I used MS Access, VBA and SQL developing this program (code example below).

Therefore... to summarize all of my programming experience is about desktop applications. So I would like to study web-developing as frontend as backend.


# Skills

* Assembler
* Borland Delphi
* MS Access
* MS Visual Basic
* SQL

 # Education
* Far Eastern State Transport University, Khabarovsk
    * Faculty of automation, telemechanics and communication, specialization of microprocessor's systems.

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