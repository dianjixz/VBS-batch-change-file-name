files_path = "D:\work\pythonyuan\text"
file_txt_path = "D:\work\pythonyuan\text\123.txt"
set fs = createobject("scripting.filesystemobject")



set log_file=fs.opentextfile(files_path + "\nihao.txt",2,True)
log_file.writeline("Welcome using lgx script!")

' file.writeline("line 1")

Dim find_file_name()
find_file_num = 0


Dim file_name()
Dim file_path()
file_num = 0






' ReDim file_name(1)
' ReDim fold_path(1)

' msgbox(file_num)
' Dim MyArray() '首先定义一个一维动态数组
' ReDim MyArray(3) '重新定义该数组的大小
' MyArray(0) = "我" '分别为数组赋值
' MyArray(1) = "要"
' MyArray(2) = "学"
' MyArray(3) = "习"
' ReDim Preserve MyArray(5) '重新定义该数组的大小
' MyArray(4) = "测" '继续为数组赋值
' MyArray(5) = "试"
' For i=0 To UBound(MyArray)
'   MsgBox MyArray(i) '循环遍历数组，并输出数组值
' Next












'寻找子文件并添加到数组中
function findfile(fie)
set folder = fs.getfolder(fie)      '设置子文件
for each file in folder.files
file_num = file_num + 1
ReDim Preserve file_name(file_num) '重新定义文件名数组的大小
file_name(file_num-1) = file.Name
if fie = files_path Then
ReDim Preserve file_path(file_num) '重新定义文件名数组的大小
file_path(file_num-1) = files_path + "\" + file.Name
Else
ReDim Preserve file_path(file_num) '重新定义文件名数组的大小
file_path(file_num-1) = file.Path
End If
next
end function

' call findfile(files_path)

' WScript.Echo("----------------------------------------------------")
' WScript.Echo(file_name(0))
' WScript.Echo(file_name(1))

' WScript.Echo("----------------------------------------------------")
' WScript.Echo(file_path(0))
' WScript.Echo(file_path(1))

Function file_lie(fie)
Call findfile(fie)
set folder = fs.getfolder(fie)
for each subfolder in folder.subfolders
Call file_lie(subfolder.Path)
next
end Function


' Set fso = CreateObject("Scripting.FileSystemObject")
' set f=fso.getfile("c:\Progra~1\Microsoft Office\Office10\WINWORD.EXE")
' f.name="WINWORD.txt" 

Function move_name(fie_path,nam)
set f=fs.getfile(fie_path)
f.Name = nam + ".pdf"
End Function






' Dim find_file_name()
' find_file_num = 0

Function read_txt(fie)
set f=fs.opentextfile(fie,1,False)
do while f.atendofline<>true

file_name_k =f.ReadLine
' WScript.Echo (file_name_k)

find_file_num = find_file_num + 1
ReDim Preserve find_file_name(find_file_num) '重新定义文件名数组的大小
find_file_name(find_file_num-1) = file_name_k
loop
end Function



' Function name_find(name)
' i = 0
' for each na in file_name
' if na = name Then
' name_find = i 
' Else
' i = i + 1
' Next

' name_find = -1 

' End Function



Function name_find(name)
i = 0
name_find = -1
for each na in file_name
if na = name Then
name_find = i
Exit For
End If
i = i + 1
next
if name_find = -1 Then

log_file.writeline("EERO:can't find name:" + name)

' WScript.Echo ("can't find name:" + name)
End If

End Function







Function main_run()
Call read_txt(file_txt_path)   '读取文件名字信息
call file_lie(files_path)       '读取目录下所有文件信息

i = 0
for each na in find_file_name
if na = "" Then
log_file.writeline("LOG:find null file name!")
Else
id = name_find(na)
if id <> -1 Then
log_file.writeline("success_LOG: success change file name:" + file_name(id))
call move_name(file_path(id),file_name(id))
i = i + 1
End If
End if

Next




log_file.writeline("--------------------------------------")
log_file.writeline("LOG: success change file number:")
log_file.writeline(Cstr(i))
end Function

call main_run()

log_file.close()



' Call file_lie(files_path)
' for Each name in file_name
' WScript.Echo(name)
' Next
' WScript.Echo(file_num)
' WScript.Echo("----------------------------------------")
' WScript.Echo(file_name(0))
' WScript.Echo(file_path(0))













' set drive = fs.getdrive(file)



' function findfold(fie)
' '寻找子文件夹
' set fs = createobject("scripting.filesystemobject")
' set folder = fs.getfolder(fie)      '设置子文件
' for each subfolder in folder.subfolders
' msgbox(subfolder)
' next
' end function

' call findfold(files_path)
