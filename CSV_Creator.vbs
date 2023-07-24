Function GetValuesToCSV()
    'fso = FileSystemObject
    'MyFile = TextStream (File)
    Dim fso, MyFile

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set MyFile = fso.CreateTextFile("C:\Users\julio\Desktop\TesteIBGE\Dados.csv", True, True)

    Dim xlApp, xlBook, xlSht

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open("C:\Users\julio\Desktop\TesteIBGE\Cadastro.xlsx")
    xlApp.Visible = True

    'Sheet Cadastro
    Dim vListCadastro
    vListCadastro = Array("B1", "B2", "B3", "B4", "B5", "B6", "B7", "M12", "M13", "M14")
    Set xlSht = xlBook.Worksheets("Cadastro")
    Call PopulateCSVFile(xlSht, vListCadastro, MyFile)

    'Sheet Validacao
    Dim vListValidacao
    vListValidacao = Array("B1", "J8", "J9")
    Set xlSht = xlBook.Worksheets("Validacao")
    Call PopulateCSVFile(xlSht, vListValidacao, MyFile)

    'Leitura do Banco de Dados
    vDataBase = ReadDatabase()
    for each vItem in vDataBase  
        WScript.Echo vItem
        Call PopulateDirectCSVFile(vItem, MyFile)     
    next
    MyFile.Close
End Function

Sub PopulateDirectCSVFile(vListParameter, MyFile)
        MyFile.Write(vListParameter & ";")
End Sub

Sub PopulateCSVFile(xlSht, vListParameter, MyFile)
    Dim vItem, vRegister
    For Each vItem In vListParameter
        vRegister = xlSht.Range(vItem).Value
        MyFile.Write(vRegister & ";")
    Next
End Sub

Function ReadDatabase()
    connectionString = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost:3306;Database=bd_devmedia;Uid=root;Pwd=;"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COLUNA1, COLUNA2 FROM tabela1 WHERE COLUNA1 = 'BANANA'", connectionString
    If Not rs.EOF And Not rs.BOF Then
        Dim vList
        vList = rs.GetRows
        rs.Close
        ReadDatabase = vList
    Else
        WScript.Echo "No data found in the recordset."
        ReadDatabase = Null
        rs.Close
    End If
End Function

Call GetValuesToCSV()

' Automation Anywhere Calling Code:

' Start
    ' Error handler: Try
        ' VBScript: OpenVBScript manual script of 59 lines
        ' VBScript: Run function“GetValuesToCSV”
        ' VBScript: CloseVBScript “Default”
    ' Error handler: CatchAllErrors
        ' Message box“$ErrorMessage$ $ErrorLineNumber.Number:toString$”
' End
