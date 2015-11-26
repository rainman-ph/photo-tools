 'Todo 
 ' Convert current rename (jhead) to rename using exif 
 ' NOTE: Exif does not contain Seconds.  CHECK for file before creating.  Add suffix if needed. 
 ' Build in support for files that do not have exif info 
 ' Clean up script big time 
 '  
 
 option Explicit  
 Dim strFolder 
 Dim IgnoredExtensions 
 Dim ie 
 Dim ret  
 Dim scriptFolder  
 Dim f, f1, fc, myArray,fileExtension,newFolder,strNewFile 
 Dim picCount : picCount = 0 
 Dim aviCount : aviCount = 0 
 Dim folderCount : folderCount = 0 
 Dim skipCount : skipCount = 0 
 Dim objFolder 
 Dim strFilename,strMediaCreatedDate 
 Dim fso : Set fso = CreateObject("Scripting.FileSystemObject") 
 Dim wshShell : Set WshShell = WScript.CreateObject("WScript.Shell") 
 Dim oShell : Set oShell = CreateObject("Shell.Application") 
 Dim sNewFileName
 Dim iTemp
 
 scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName) 
 
 'Check if ignoredExtensions file is present.  This will be used to list extensions that do not need to be processed 
 If fso.FileExists(scriptFolder & "\IgnoredExtensions.txt") Then 
     Set ie = fso.OpenTextFile(scriptFolder & "\IgnoredExtensions.txt", 1) 
     IgnoredExtensions = ie.ReadAll 
     ie.Close 
 Else 
     IgnoredExtensions = "" 
 End If  
 
 'Path is Parent of Tools folder i.e.  
 ' ScriptFolder = d:\NewPictures\Tools 
 ' strFolder default = d:\newPictures 
 strFolder = InputBox("enter in path to folder","Source",fso.GetParentFolderName(scriptFolder)) 
 Set objFolder = oShell.Namespace(strFolder) 
 If strFolder = "" Then  
     WScript.Echo "Cancel" 
     WScript.Quit 
 End If  
 
 'Set current folder to inputbox as this is where the files will be processed.  needed for jhead to function correctly as files need to be in same folder as exe 
 WshShell.CurrentDirectory = strFolder 
 
 'While in testing, create backup of all files. 
 ret = MsgBox("This script is still in test.  Do you want to create a backup?",vbYesNo) 
 If ret = vbYes Then 
     Call doBackup  
 End If  
 
 'Rename all AVI's and THM's to MM_DD_YYY_HHMMSS format.  All THM's will be renamed to JPG 
 'set objFolder = wshShell.Namespace(strFolder) 

 
 Set f = fso.GetFolder(strFolder) 
 Set fc = f.Files 
 For Each f1 in fc 
     strFileName = f1.Name 
     If fso.GetExtensionName(ucase(strFilename)) = "AVI" Then      
         'Set strmediacreateddate based on info taken from properties of the file, 190 is the media Created Date 
         strMediaCreatedDate = objFolder.GetDetailsOf(objFolder.Parsename(strFileName),190) 
         If Len(strMediaCreatedDate) = 0 Then'If data is not present, skip file 
             Wscript.Echo strFilename & " - was Skipped, Unknown Media Created Date" 
             skipCount = skipCount + 1'increment skipCount(only for show) 
         Else 
		    iTemp = 1
		    sNewFileName = strFolder & "\" & CleanDate(strMediaCreatedDate) & ".avi"
		    do while fso.FileExists(sNewFileName)
			   sNewFileName = strFolder & "\" & CleanDate(strMediaCreatedDate) & "("&iTemp &").avi"
			   iTemp = iTemp +1
			Loop
			   
             'Rename file and leave in current folder 
             fso.MoveFile f1.Path, sNewFileName 'File name comes back as Month_DD_YYY_HHMM Seconds not needed as 2 videos will not be taken in the same second 
         End If  
     End if  
 Next 
 
 'Rename all JPG's to MM_DD_YYY_HHMMSS format (_ can be changed as needed) 
 'WshShell.Run "cmd /c " & scriptFolder & "\jhead -nf%m_%d_%Y_%H%M%S *.jpg -v >" & scriptFolder & "\jpg.txt",0,True   
 'msgbox "Pause" 'Uncomment out pause to see the the files name before they are moved. 
 
 REM Set f = fso.GetFolder(strFolder) 
 REM Set fc = f.Files 
 REM For Each f1 in fc 
     REM If fso.GetExtensionName(ucase(f1.Name)) = "AVI" Then  
         REM strNewFile = f1.Name'If AVI file, does not need to go through fndate function again doing so would turn October_31_2009 into _ober_31_2009 
      REM Else 
         REM strNewFile = fnDate(f1.name)'set filename to be Month_DD_YYYY_HHMMSS format 
      REM End if  
      
     REM fileExtension = fso.GetExtensionName(lcase(strNewFile)) 
       
     REM If fileExtension = "jpg" Or fileExtension = "bmp" then 
         REM picCount = picCount + 1'increment picCount(only for show) 
         REM Call CopyFiles(f1.Path,strNewFile) 
     REM ElseIf Lcase(fileExtension) = "avi" Then  
         REM aviCount = aviCount + 1'increment aviCount(only for show) 
         REM Call CopyFiles(f1.Path,strNewFile) 
     REM ElseIf lcase(strNewFile) = "thumbs.db" Then 'Delete thumbs.db if present 
         REM f1.attributes = f1.attributes - 6 
         REM fso.DeleteFile(f1.path) 
     REM Else 
         REM If InStr(IgnoredExtensions,fileExtension) = 0 Then 
             REM ret = Msgbox(f1.Name & " is an unknown extension.  Do you want to add it to the Ignored Extensions file?",vbYesNo) 
             REM If ret = vbYes Then 
                 REM Call UpdateIgnoredExtensions(fileExtension) 
             REM End If  
         REM End If  
     REM End If  
 REM Next 
 
 MsgBox picCount & " pictures have been processed" & vbcr _ 
 & aviCount & " movies have been processed" & vbCr _  
 & folderCount & " folders have been created" & vbCr _  
 & skipCount & " files have been skipped"   
   
 Sub DoBackup() 
     Dim backupCount : backupCount = 0 
     Set f = fso.GetFolder(strFolder) 
     Set fc = f.Files 
      
     For Each f1 in fc 
          strNewFile = f1.name 
          fileExtension = fso.GetExtensionName(lcase(strNewFile)) 
          If Not (fileExtension = "jpg" Or fileExtension = "bmp" Or fileExtension = "avi") Then  
                 'DoNothing 
         Else 
             backupCount = backupCount + 1     
             fso.CopyFile f1.Path,strFolder & "\backup\" & f1.Name  
         End If  
     Next 
     MsgBox backupCount & " files have been backuped" 
 End Sub 
   
 Sub CopyFiles(originalFile,theFile) 
     Dim cf 
     myArray = Split(theFile,"_") 
     'Set newfolder to equal first 3 parts of the File which will be 
     ' Month(spelled out) DD and YYYY 
     'ex January_01_2008 
     'newFolder = strFolder & "\" & myArray(0) & "_" & myArray(1) & "_" & myArray(2) 
    '  If Not fso.FolderExists(newFolder) Then  
    '     folderCount = folderCount + 1'increment folderCount(only for show) 
    '     Set cf = fso.CreateFolder(newFolder) ' Create the folder as needed 
    ' End If 
     If Not fso.FileExists(newFolder & "\" & theFile) Then  
         'Move file to new folder renaming it in the process 
         fso.MoveFile originalFile,newFolder & "\" & theFile 
     Else 
         Wscript.Echo theFile & " - was Skipped" 
         skipCount = skipCount + 1'increment skipCount(only for show) 
     End If  
 End Sub 
 
 Function CleanDate(strDate) 
     Dim RegEx : Set RegEx = New RegExp 
     RegEx.Pattern = "[^\w\s\/:]" ' not a word (A-Z, a-z, 0-9, _) or colon 
     RegEx.IgnoreCase = True 
     RegEx.Global = True 
     If fso.GetExtensionName(ucase(f1.Name)) = "AVI" then  
         CleanDate = FnDateAVI(RegEx.Replace(strDate, "")) 
     Else 
         CleanDate = FnDate(RegEx.Replace(strDate, ""))     
     End if  
 End Function 
 
 Function FnDateAVI(Dt) 
	 FnDateAVI =  FnN(Year(Dt), 4) &"-" &FnN(Month(Dt), 2) &"-" &FnN(Day(Dt), 2)& " " & Replace(FormatDateTime(dt,vbShortTime),":",".")
'     Dim tFnDate,Rnm 
'     tFnDate = FnN(Month(Dt), 2) 
 
'     Select Case tFnDate 
'         Case "01" Rnm = "January" 
'         Case "02" Rnm = "February" 
'         Case "03" Rnm = "March" 
'         Case "04" Rnm = "April" 
'         Case "05" Rnm = "May" 
'         Case "06" Rnm = "June" 
'         Case "07" Rnm = "July" 
'         Case "08" Rnm = "August" 
'         Case "09" Rnm = "September" 
'         Case "10" Rnm = "October" 
'         Case "11" Rnm = "November" 
'         Case "12" Rnm = "December" 
'     End Select 
     'FnDateAVI =  rNm & "_" & FnN(Day(Dt), 2) & "_" & FnN(Year(Dt), 4) & "_" & Replace(FormatDateTime(dt,vbShortTime),":","") 
End Function 
 
 Function FnN(V, N) 
  FnN = Right(String(N,"0") & V, N) 
 End Function 
 
 Function FnDate(Dt) 
 'Replace MM with Month 
 ' i.e. 01 becomes January 
     Dim tFnDate,Rnm 
     tFnDate = Left(dt,2) 
 
     Select Case tFnDate 
         Case "01" Rnm = "January" 
         Case "02" Rnm = "February" 
         Case "03" Rnm = "March" 
         Case "04" Rnm = "April" 
         Case "05" Rnm = "May" 
         Case "06" Rnm = "June" 
         Case "07" Rnm = "July" 
         Case "08" Rnm = "August" 
         Case "09" Rnm = "September" 
         Case "10" Rnm = "October" 
         Case "11" Rnm = "November" 
         Case "12" Rnm = "December" 
     End Select 
      
     FnDate =  Rnm & "_" & Mid(dt,4,len(dt)) 
 End Function 
 
 Sub UpdateIgnoredExtensions(theExtension) 
 'Update IgnoredExtensions 
     Dim s 
     Set s = fso.OpenTextFile(scriptFolder & "\IgnoredExtensions.txt", 8,True)'ForAppending 
     s.WriteLine(theExtension) 
     s.Close 
 End Sub  
 
