Sub UploadToIO()

    Const PATH = "C:\Users\Jeswin Annish\Pictures\Camera Roll\"
    Const FILENAME = "WIN_20200507_22_02_27_Pro.jpg"
    Const CONTENT = "image/jpg"
    'Define URL Components
    base_url = "https://api.github.com/repos/"
    repo_name = "Test_jes/"
    username = "jeswin04/"
    file_name = "Module1.bas"
    access_token = "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"
    
    'Build the full URL.
    Url = base_url + username + repo_name + "contents/" + FILENAME + "?ref=master"
    'Const URL = "https://file.io"

    ' generate boundary
    Dim BOUNDARY, s As String, n As Integer
    For n = 1 To 16: s = s & Chr(65 + Int(Rnd * 25)): Next
    BOUNDARY = s & CDbl(Now)

    Dim part As String, ado As Object
    part = "--" & BOUNDARY & vbCrLf
    part = part & "Content-Disposition: form-data; name=""file""; filename=""" & FILENAME & """" & vbCrLf
    part = part & "Content-Type: " & CONTENT & vbCrLf & vbCrLf

    ' read file into image
    Dim image
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1 'binary
    ado.Open
    ado.LoadFromFile PATH & FILENAME
    ado.Position = 0
    image = ado.read
    ado.Close

    ' combine part, image , end
    ado.Open
    ado.Position = 0
    ado.Type = 1 ' binary
    ado.Write ToBytes(part)
    ado.Write image
    ado.Write ToBytes(vbCrLf & "--" & BOUNDARY & "---")
    ado.Position = 0
    'ado.savetofile "c:\tmp\debug.bin", 2 ' overwrite

    ' send request
    
    With CreateObject("MSXML2.ServerXMLHTTP")
        .Open "PUT", Url, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
        
    'Set the headers.
.setRequestHeader "Accept", "application/vnd.github.v3+json"
.setRequestHeader "Authorization", "token " + access_token
        .send ado.read
        Debug.Print .responseText
    End With

    MsgBox "File: " & PATH & FILENAME & vbCrLf & _
           "Boundary: " & BOUNDARY, vbInformation, "Uploaded to " & Url

End Sub

Function ToBytes(str As String) As Variant

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Open
    ado.Type = 2 ' text
    ado.Charset = "_autodetect"
    ado.WriteText str
    ado.Position = 0
    ado.Type = 1
    ToBytes = ado.read
    ado.Close

End Function



