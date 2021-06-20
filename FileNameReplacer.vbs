Set objFso = CreateObject("Scripting.FileSystemObject")
Set Folder = objFso.GetFolder("D:\IDM\Compressed\Udemy")

For Each File In Folder.Files

    sNewFile = File.Name

    sNewFile = Replace(sNewFile,"old-word","new-word")

    if (sNewFile<>File.Name) then

        File.Move(File.ParentFolder + "\" + sNewFile)

    end if

Next
