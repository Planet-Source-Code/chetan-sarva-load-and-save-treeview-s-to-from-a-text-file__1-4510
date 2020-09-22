<div align="center">

## Load and Save TreeView's to/from a text file


</div>

### Description

These are two functions I wrote to save and load a treeview's nodes (saves the .text, .tag, and .key properties) to and from a text file.

<BR><BR>

This is a very simple code and should be very easy to incorporate to any project.
 
### More Info
 
No inputs...

None needed...

Doesn't return anything...

No side effects...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chetan Sarva](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chetan-sarva.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chetan-sarva-load-and-save-treeview-s-to-from-a-text-file__1-4510/archive/master.zip)

### API Declarations

No API's used here...


### Source Code

```
Public Sub LoadTree(ByVal tvTree As TreeView, ByVal sFileName As String)
' Function by Chetan Sarva (November 17, 1999)
' Please include this comment if you use this code.
Dim curNode As Node
Dim sDelimiter As String
Dim freef As Integer
Dim buf As String
Dim nodeparts As Variant
sDelimiter = "" ' We want something extremely unique to delimit
        ' each of the pices of our treeview
 On Error Resume Next
 ' Get a free file and open our file for output
 freef = FreeFile()
 Open sFileName For Input As #freef
  Do
  DoEvents
   ' Read in the current line
   Line Input #freef, buf
   ' Split the line into pieces on our delimiter
   nodeparts = Split(buf, sDelimiter)
   ' See if it's a root or child node and add accordingly
   If nodeparts(3) = "parent" Then
    curNode = tvTree.Nodes.Add(, , nodeparts(1), nodeparts(0))
    curNode.Tag = nodeparts(2)
   Else
    curNode = tvTree.Nodes.Add(nodeparts(3), tvwChild, nodeparts(1), nodeparts(0))
    curNode.Tag = nodeparts(2)
   End If
  Loop Until EOF(freef)
 Close #freef
End Sub
Public Sub SaveTree(ByVal tvTree As TreeView, ByVal sFileName As String)
' Function by Chetan Sarva (November 17, 1999)
' Please include this comment if you use this code.
Dim curNode As Node
Dim sDelimiter As String
Dim freef As Integer
sDelimiter = "" ' We want something extremely unique to delimit
        ' each of the pices of our treeview
 On Error Resume Next
 ' Get a free file and open our file for output
 freef = FreeFile()
 Open sFileName For Output As #freef
  ' Loop through all the nodes and save all the
  ' important information
  For Each curNode In tvTree.Nodes
   If curNode.FullPath = curNode.Text Then
    Print #freef, curNode.Text; sDelimiter; curNode.Key; sDelimiter; curNode.Tag; sDelimiter; "parent"
   Else
    Print #freef, curNode.Text; sDelimiter; curNode.Key; sDelimiter; curNode.Tag; sDelimiter; curNode.Parent.Key
   End If
  Next curNode
 Close #freef
End Sub
```

