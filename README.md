<div align="center">

## DataGrid Sort


</div>

### Description

Sorts the records in a datagrid form when clicking on the column header. Toggles between ascending/descending sort order.
 
### More Info
 
I assume that the connection is established and that the recordset has been opened.

Besides that I would really appreciate your feedback. Let me know if you have other ideas on how to accomplish this. I'd specially like to know if there is a "nicer" alternative to the nested if-statement.

Not that I could see ;)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Gaudin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-gaudin.md)
**Level**          |Beginner
**User Rating**    |4.6 (69 globes from 15 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-gaudin-datagrid-sort__1-7345/archive/master.zip)

### API Declarations

Dim WithEvents adoPrimaryRS As Recordset


### Source Code

```
Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
 Dim strColName As String
 Static bSortAsc As Boolean
 Static strPrevCol As String
 strColName = grdDataGrid.Columns(ColIndex).DataField
' Did the user click again on the same column ? If so, check
' the previous state, in order to toggle between sorting ascending
' or descending. If this is the first time the user clicks on a column
' or if he/she clicks on another column, then sort ascending.
 If strColName = strPrevCol Then
  If bSortAsc Then
   adoPrimaryRS.Sort = strColName & " DESC"
   bSortAsc = False
  Else
   adoPrimaryRS.Sort = strColName
   bSortAsc = True
  End If
 Else
  adoPrimaryRS.Sort = strColName
  bSortAsc = True
 End If
 strPrevCol = strColName
End Sub
```

