<div align="center">

## Search Drive


</div>

### Description

Search Drive for files with a few lines o'code, Recursivly- Please VOTE!!!!
 
### More Info
 
Add: Commandbutton, dir list box, file list box, text1.text, label1, label3, listbox


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joseph Ronie](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joseph-ronie.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joseph-ronie-search-drive__1-23020/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
 File1.Pattern = "*.exe"
 Label3.Caption = "Searching " & File1.Pattern
 Call search("C:\")
End Sub
Public Sub search(dr As String)
Dim lst(5000) As String
Dim lstcnt As Integer
Dir1.Path = dr
listcnt = 0
 Do While (Dir1.ListIndex < Dir1.ListCount - 1)
  Dir1.ListIndex = Dir1.ListIndex + 1
  listcnt = listcnt + 1
  lst(listcnt) = Dir1.List(Dir1.ListIndex)
 Loop
 For i = 1 To listcnt
  search (lst(i))
 Next i
  File1.Path = Dir1.List(Dir1.ListIndex)
  DoEvents
  Do While (File1.ListIndex < _   File1.ListCount - 1)
   File1.ListIndex = File1.ListIndex + 1
   List1.AddItem (Dir1.List _(Dir1.ListIndex) & "\" & File1.FileName)
   DoEvents
  Loop
End Sub
```

