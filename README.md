<div align="center">

## Text Box Auto Fill


</div>

### Description

This simple code allows to make an auto-filled textbox (like an adress box in IE). This example uses an DataEnvironment connection, but it can be easy used in any other cases.
 
### More Info
 
Just place this code into the form module and correct the sub title (e.g. the textbox, which uses this code in my app is txtSTREET)

no side effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timur](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timur.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timur-text-box-auto-fill__1-2624/archive/master.zip)

### API Declarations

private firstcome as boolean


### Source Code

```
Private Sub txtSTREET_KeyUp(KeyCode As Integer, Shift As Integer)
Dim PrevLength  As Integer, PrevStart As Integer
If Not KeyCode >= 65 Then Exit Sub
If firstcome Then firstcome = False: Exit Sub
  With DataEnvironment.Connection1.Execute("SELECT ADDRESS from tblFlats WHERE UCASE(ADDRESS) like '" & UCase(Me.txtSTREET) & "%'")
    If Not .EOF Then
      If Not Me.txtSTREET = "" Then
        PrevStart = Len(Me.txtSTREET) + 1
        PrevLength = -Len(Me.txtSTREET) + Len(!ADDRESS)
        Me.txtSTREET.SelStart = PrevStart
        Me.txtSTREET.SelLength = PrevLength
        Me.txtSTREET.SelText = Mid$(!ADDRESS, Len(Me.txtSTREET) + 1)
        Me.txtSTREET.SelStart = PrevStart - 1
        Me.txtSTREET.SelLength = PrevLength
      End If
    'Else
      'MsgBox "The entered fragment is not found in the list!"
      'Me.STREET = ""
    End If
  End With
End Sub
```

