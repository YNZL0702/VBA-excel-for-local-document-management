   
Option Explicit

Sub Main_updating_List_Name()

' 文件 默认从Bx开始, x= 

Dim folderPath As String   

folderPath = ThisWorkbook.path & "\"

Dim newList
newList = getDocumentName(folderPath)

Dim oldList
oldList = getExist_ListName()

Dim re_oldList
re_oldList = Sum_two_Arr(oldList, newList)

Dim updateList
updateList = remove_same_item(re_oldList)

addHyperLink (folderPath)

End Sub

'-----------------------------------------------------

Function getDocumentName(folder_Path As String)

'only pdf


Dim file_name As String
Dim new_Name() As String
Dim n As Integer
n = 1

file_name = Dir(folder_Path & "*.pdf")     '

Do While file_name <> ""
    
    ReDim Preserve new_Name(1 To n)
    
    new_Name(n) = VBA.Split(file_name, ".")(0)      '
    
    n = n + 1
    
    file_name = Dir
    
Loop

getDocumentName = new_Name

End Function

'------------------------------------------------------------------------------------

Function getExist_ListName()

'start Excel B(Bx) 


Dim i As Integer
Dim old_Name()

Dim rn As Integer
rn = Range("B65536").End(xlUp).Row    

If rn < x Then
    
    old_Name = Array()
    
    Else:
    
    ReDim old_Name(1 To rn - 2)
    
    For i = 1 To rn - 2
    
        old_Name(i) = Cells(i + 2, 2)   '从 Bx单元格开始赋值
    Next

End If

    getExist_ListName = old_Name
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Sum_two_Arr(old_Name, new_Name)

Dim num_new_Name, num_old_Name As Integer
Dim i As Integer, j As Integer

j = 1

Dim rn As Integer
rn = Range("B65536").End(xlUp).Row                                   

num_new_Name = UBound(new_Name) - LBound(new_Name) + 1                
num_old_Name = UBound(old_Name) - LBound(old_Name) + 1                

ReDim Preserve old_Name(1 To num_new_Name + num_old_Name)

For i = rn - 1 To num_new_Name + num_old_Name                         '
    old_Name(i) = new_Name(j)
    j = j + 1
Next

Sum_two_Arr = old_Name

End Function

'---------------------------------------------------------------------------

Function remove_same_item(arr)


Dim dic As Object
Set dic = CreateObject("Scripting.Dictionary")

Dim ii As Integer

For ii = LBound(arr) To UBound(arr)
    dic(arr(ii)) = ""
Next

Range("B"&x).Resize(dic.Count, 1) = Application.WorksheetFunction.Transpose(dic.keys)


End Function

'-----------------------------------------------------------------------------------------

Function addHyperLink(folder_Path As String)


Dim k As Integer
Dim LinkPath

Dim rnNew As Integer

rnNew = Range("B65536").End(xlUp).Row

For k = x To rnNew  '从Bx单元格到最后

    LinkPath = folder_Path & Range("B" & k).Text & ".pdf"

    Range("B" & k).Hyperlinks.Add Anchor:=Range("B" & k), Address:=LinkPath
    
Next

End Function

