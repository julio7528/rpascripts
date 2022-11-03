'Global Variable: Once time declared here -  works in all functions
dim filePathName
dim aDate
dim strName

'call functions by steps
call wsfilename
call wshshell1
call FileExists(filePathName)
call wshShell2
call SheetImport
call KillTasks

'Object WScript.Shell: run windows commands
Function KillTasks
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.Run "taskkill /f /im excel.exe"
oShell.Run "taskkill /f /im notepad.exe"
End function

'This variable gets the username machine
Function wsfilename
Set wsfilename = WScript.CreateObject("WScript.Shell")
strName = wsfilename.ExpandEnvironmentStrings("%USERNAME%")
end function

'Function that write a text file in background
Function wshshell1
    Set wshshell1 = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\"&strName&"\Desktop\Info.txt",2,true)
    aDate = Date()    
    filePathName = "C:\Users\"&strName&"\Desktop\Info.txt"
    wshshell1.WriteLine(aDate)
    wshshell1.write(strName)
    wshshell1.WriteLine()
    wshshell1.WriteLine()
    wshshell1.WriteLine(aDate)    
    wshshell1.WriteLine()
    wshshell1.WriteLine(aDate)    
    wshshell1.WriteLine()
    wshshell1.WriteLine(aDate)    
    wshshell1.WriteLine()
    wshshell1.WriteLine(aDate&strName)    
    wshshell1.WriteLine(strName)
    wshshell1.WriteLine(aDate)
    wshshell1.Close    
End Function

'Function check if a file exists
Function FileExists(filePathName)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(filePathName) Then
    FileExists=CBool(1)    
  Else
    FileExists=CBool(0)    
  End If
End Function

'Run and write a text file in screen
Function wshShell2
    If (FileExists(filePathName)) = false then
    msgbox("Arquivo Não Existe")
    Else    
        Set WshShell2 = WScript.CreateObject("WScript.Shell")
        Call WshShell2.Run(filePathName)
        wscript.sleep 2000
        WshShell2.SendKeys ("^{END}")
        WshShell2.SendKeys ("^{ENTER}")
        WshShell2.sendkeys (aDate &" "& strName)
        wscript.sleep 1000
        WshShell2.SendKeys ("^{s}")
        WshShell2.SendKeys ("%{F4}")
    End If
End Function 

Function SheetImport
  Const xlDelimited  = 1
  Const xlWorkbookNormal = -4143

      Set objExcel1 = CreateObject("Excel.Application")      
      objExcel1.Visible = True
      Set objWorkbook = objExcel1.Workbooks.Open("C:\Users\"&strName&"\Desktop\Info.xlsx")
      objExcel1.Application.DisplayAlerts = False



  Set objExcel2 = CreateObject("Excel.Application")
      objExcel2.Visible = True
            objExcel2.Workbooks.OpenText filePathName, _
          , , xlDelimited, , , , , , , True, "~"
              ' .Range("A1"), _                ' Destination
              ' xlDelimited, _                 ' Data Type
              ' xlDoubleQuote, _               ' Text Qualifier
              ' True, _                        ' Consecutive Delimiters?
              ' , _                            ' Use Tab for Delimiter?
              ' , _                            ' Use Semicolon for Delimiter?
              ' , _                            ' Use Comma for Delimiter?
              ' True                           ' Use Space for Delimiter?
     
      objExcel2.Application.DisplayAlerts = False
      objExcel2.Range("A1").Select
      objExcel2.Cells.Select
      objExcel2.Selection.Copy
      
      objExcel1.Visible = True
      objExcel1.Range("A1").Select
      objExcel1.Cells.Select
      objExcel1.ActiveSheet.Paste

      objExcel1.Visible = True
      objExcel1.Range("A1").Select
      objExcel1.Selection.EntireRow.Insert
      objExcel1.Selection.FormulaR1C1 = "Tabela"
      objExcel1.Columns("A:A").Select
      objExcel1.Selection.AutoFilter
      objExcel1.Range("A1").Select
      ExemploAspasDuplas = Chr(34)   
      dateSignal = ("=*"&strName&"*")
      'msgbox(ExemploAspasDuplas&"=*"&strName&"*"&ExemploAspasDuplas)
      objExcel1.Range("A1").AutoFilter 1, dateSignal,,,True 

      objWorkbook.save
      objExcel1.Workbooks.Close
      objExcel2.Quit

End Function

