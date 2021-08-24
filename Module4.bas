Public Sub UploadFile()
'Dim sFormData As String
Dim sFormData, bFormData
Dim d As String, DestURL As String, FILENAME As String, filePath As String, FieldName As String
FieldName = "File"
DestURL = "https://localhost:44327/api/fileUpload"
'FileName = "testfile.txt"
'CONTENT = "text/plain"
FILENAME = "Virtusa_Application.pdf"
CONTENT = "application/pdf"
filePath = "C:\Users\Jeswin Annish\Desktop\" & FILENAME

'Define URL Components
    base_url = "https://api.github.com/repos/"
    repo_name = "Test_jes/"
    username = "jeswin04/"
    'file_name = "Module1.bas"
    access_token = "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"
    
    'Build the full URL.
    DestURL = base_url + username + repo_name + "contents/" + FILENAME + "?ref=master"
    'Const URL = "https://file.io"


'Boundary of fields.
'Be sure this string is Not In the source file

Const BOUNDARY As String = "---------------------------0123456789012"

Dim File, FILESIZE
Set ado = CreateObject("ADODB.Stream")
ado.Type = 1 'binary
ado.Open
ado.LoadFromFile processBase64Binary(PATH & FILENAME)
ado.Position = 0
FILESIZE = ado.Size
File = ado.read
ado.Close


Set ado = CreateObject("ADODB.Stream")
d = "--" + BOUNDARY + vbCrLf
d = d + "Content-Disposition: form-data; name=""" + FieldName + """;"
d = d + " filename=""" + FILENAME + """" + vbCrLf
d = d + "Content-Type: " & CONTENT + vbCrLf + vbCrLf
ado.Type = 1 'binary
ado.Open
ado.Write ToBytes(d)
ado.Write File
ado.Write ToBytes(vbCrLf + "--" + BOUNDARY + "--" + vbCrLf)
ado.Position = 0


  With CreateObject("MSXML2.ServerXMLHTTP")
    .Open "PUT", DestURL, False
    .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
    
    .setRequestHeader "Accept", "application/vnd.github.v3+json"
.setRequestHeader "Authorization", "token " + access_token
    .send ado.read
    Debug.Print .responseText
End With
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


Function processBase64Binary(ByVal filePath As String) As String

'define objects used
Dim mstream As Stream
Dim objXML As MSXML2.DOMDocument
Dim objNode As MSXML2.IXMLDOMElement

'load the data from the filepath provided
Set mstream = New Stream
mstream.Type = adTypeBinary
mstream.Open
mstream.LoadFromFile filePath

'convert to base 64 binary and convert to a string to be used in XML
Set objXML = New MSXML2.DOMDocument
Set objNode = objXML.createElement("Base64Data")
objNode.DataType = "bin.base64"
objNode.nodeTypedValue = mstream.read()
myText = objNode.text
testDecode (myText)
processBase64Binary = myText

End Function