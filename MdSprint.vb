Imports System.Xml
Imports System.Web
Imports System.Net
Imports System.Text
Imports System.IO
'Imports System.Web.Services



Module MdSprint

    'Function Sprint_request() As Boolean
    '    Dim req As New Chilkat.HttpRequest()
    '    Dim http As New Chilkat.Http()

    '    Dim success As Boolean
    '    Dim result As String = ""

    '    '  Any string unlocks the component for the 1st 30-days.
    '    success = http.UnlockComponent("Anything for 30-day trial")
    '    If (success <> True) Then
    '        MsgBox(http.LastErrorText)
    '        Exit Function
    '    End If


    '    '  Build this XML SOAP request:
    '    '  <?xml version="1.0" encoding="utf-8"?>
    '    '  <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    '    '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    '    '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    '    '    <soap:Body>
    '    '      <GetQuote xmlns="http://www.webserviceX.NET/">
    '    '        <symbol>string</symbol>
    '    '     </GetQuote>
    '    '    </soap:Body>
    '    '  </soap:Envelope>
    '    Dim soapReq As New Chilkat.Xml()
    '    soapReq.Encoding = "utf-8"
    '    soapReq.Tag = "soapenv:Envelope"
    '    'soapReq.AddAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    '    'soapReq.AddAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
    '    soapReq.AddAttribute("xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
    '    soapReq.AddAttribute("xmlns:v2", "http://integration.sprint.com/common/header/WSMessageHeader/v2")
    '    soapReq.AddAttribute("xmlns:quer", "http://integration.sprint.com/interfaces/queryCdmaDeviceInfo/v2/queryCdmaDeviceInfoV2.xsd")

    '    soapReq.NewChild2("soapenv:Header", "")
    '    soapReq.FirstChild2()
    '    soapReq.NewChild2("v2:wsMessageHeader", "")
    '    soapReq.FirstChild2()
    '    soapReq.NewChild2("v2:trackingMessageHeader", "")
    '    soapReq.FirstChild2()
    '    soapReq.NewChild2("v2:applicationId", "EXODUS")
    '    soapReq.NewChild2("v2:applicationUserId", "EXODUS")
    '    soapReq.NewChild2("v2:consumerId", "EXODUS")
    '    soapReq.NewChild2("v2:messageId", "EXODUS")
    '    soapReq.NewChild2("v2:timeToLive", "30")
    '    soapReq.NewChild2("v2:messageDateTimeStamp", Now)

    '    soapReq.GetParent2()
    '    soapReq.GetParent.GetParent2()

    '    soapReq.GetRoot2()

    '    soapReq.NewChild2("soapenv:Body", "")
    '    soapReq.LastChild2()
    '    soapReq.NewChild2("quer:queryCdmaDeviceInfoV2", "")
    '    soapReq.FirstChild2()
    '    soapReq.NewChild2("quer:serialNumber", "")
    '    soapReq.FirstChild2()
    '    soapReq.NewChild2("quer:deviceSerialNumberDecimal", "270113177604822159")

    '    soapReq.GetRoot2()

    '    result = result & soapReq.GetXml() & vbCrLf

    '    '  Build an SOAP request.
    '    req.UseXmlHttp(soapReq.GetXml())
    '    req.Path = ""

    '    req.AddHeader("SOAPAction", "https://webservicesgatewaytest.sprint.com:444/rtb1/services/wireless/account/QueryDeviceInfoService/v1?wsdl")

    '    '  Send the HTTP POST and get the response.  Note: This is a blocking call.
    '    '  The method does not return until the full HTTP response is received.
    '    Dim domain As String
    '    Dim port As Long
    '    Dim ssl As Boolean
    '    domain = "https://webservicesgatewaytest.sprint.com:444/rtb1/services/wireless/account/QueryDeviceInfoService/v1?wsdl"
    '    port = 444
    '    ssl = False
    '    Dim resp As Chilkat.HttpResponse
    '    resp = http.SynchronousRequest(domain, port, ssl, req)
    '    If (resp Is Nothing) Then
    '        result = result & http.LastErrorText & vbCrLf
    '    Else
    '        '  The XML response is in the BodyStr property of the response object:
    '        Dim soapResp As New Chilkat.Xml()
    '        soapResp.LoadXml(resp.BodyStr)

    '        result = ""
    '        '  The response will look like this:
    '        '  <?xml version="1.0" encoding="utf-8"?>
    '        '  <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    '        '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    '        '  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    '        '    <soap:Body>
    '        '      <GetQuoteResponse xmlns="http://www.webserviceX.NET/">
    '        '        <GetQuoteResult>string</GetQuoteResult>
    '        '      </GetQuoteResponse>
    '        '    </soap:Body>
    '        '  </soap:Envelope>
    '        '  Navigate to soap:Body
    '        ''            soapResp.FirstChild2()
    '        '  Navigate to GetQuoteResponse
    '        '          soapResp.FirstChild2()
    '        '  Navigate to GetQuoteResult
    '        '         soapResp.FirstChild2()

    '        '  The actual XML response is the data within GetQuoteResult:
    '        Dim xmlResp As New Chilkat.Xml()
    '        xmlResp.LoadXml(soapResp.Content)

    '        '  Display the XML response:
    '        result = result & xmlResp.GetXml() & vbCrLf

    '    End If

    'End Function

    'Function Sprint_request1() As Boolean

    '    'Dim bline As Byte()
    '    'Dim str As String
    '    Dim xml As String = ""

    '    'CHANGE HERE use the real path of the pdf physical path here. 
    '    'bline = System.IO.File.ReadAllBytes("C:\Sample.pdf")
    '    'str = Convert.ToBase64String(bline)

    '    xml = xml & "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbNewLine
    '    xml = xml & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v2=""http://integration.sprint.com/common/header/WSMessageHeader/v2"" xmlns:quer=""http://integration.sprint.com/interfaces/queryCdmaDeviceInfo/v2/queryCdmaDeviceInfoV2.xsd"">" & vbNewLine
    '    xml = xml & "  <soapenv:Header>" & vbNewLine
    '    xml = xml & "    <v2:wsMessageHeader>" & vbNewLine
    '    xml = xml & "      <v2:trackingMessageHeader>" & vbNewLine
    '    xml = xml & "        <v2:applicationId>EXODUS</v2:applicationId>" & vbNewLine
    '    xml = xml & "        <v2:applicationUserId>EXODUS</v2:applicationUserId>" & vbNewLine
    '    xml = xml & "        <v2:consumerId>EXODUS</v2:consumerId>" & vbNewLine
    '    xml = xml & "        <v2:messageId>EXODUS</v2:messageId>" & vbNewLine
    '    xml = xml & "        <v2:timeToLive>30</v2:timeToLive>" & vbNewLine
    '    xml = xml & "        <v2:messageDateTimeStamp>2013-11-06 오후 4:02:52</v2:messageDateTimeStamp>" & vbNewLine
    '    xml = xml & "      </v2:trackingMessageHeader>" & vbNewLine
    '    xml = xml & "    </v2:wsMessageHeader>" & vbNewLine
    '    xml = xml & "  </soapenv:Header>" & vbNewLine
    '    xml = xml & "  <soapenv:Body>" & vbNewLine
    '    xml = xml & "    <quer:queryCdmaDeviceInfoV2>" & vbNewLine
    '    xml = xml & "      <quer:serialNumber>" & vbNewLine
    '    xml = xml & "        <quer:deviceSerialNumberDecimal>270113177604822159</quer:deviceSerialNumberDecimal>" & vbNewLine
    '    xml = xml & "      </quer:serialNumber>" & vbNewLine
    '    xml = xml & "    </quer:queryCdmaDeviceInfoV2>" & vbNewLine
    '    xml = xml & "  </soapenv:Body>" & vbNewLine
    '    xml = xml & "</soapenv:Envelope>" & vbNewLine

    '    'Dim aa As WebRequest = CreateObject("MSSOAP.SoapClient30")
    '    'Dim xmldoc = CreateObject("MSXML2.DOMDocument.4.0")



    '    Dim data As String = xml
    '    Dim url As String = "http://integration.sprint.com/interfaces/queryCdmaDeviceInfo/v2/queryCdmaDeviceInfoV2.xsd"
    '    Dim responsestring As String = ""



    '    Dim myReq As WebRequest = WebRequest.Create(url)
    '    Dim proxy As IWebProxy = CType(myReq.Proxy, IWebProxy)
    '    Dim proxyaddress As String
    '    Dim myProxy As New WebProxy()
    '    Dim encoding As New ASCIIEncoding
    '    Dim buffer() As Byte = encoding.GetBytes(xml)
    '    Dim response As String


    '    'myReq.AllowWriteStreamBuffering = False
    '    myReq.Method = "POST"
    '    myReq.ContentType = "text/xml; charset=UTF-8"
    '    myReq.ContentLength = buffer.Length
    '    myReq.Headers.Add("SOAPAction", "http://integration.sprint.com/interfaces/queryCdmaDeviceInfo/v2/queryCdmaDeviceInfoV2.xsd")
    '    'myReq.Credentials = New NetworkCredential("abc", "123")
    '    myReq.PreAuthenticate = True
    '    proxyaddress = proxy.GetProxy(myReq.RequestUri).ToString



    '    Dim newUri As New Uri(proxyaddress)
    '    myProxy.Address = newUri
    '    myReq.Proxy = myProxy
    '    Dim post As Stream = myReq.GetRequestStream

    '    post.Write(buffer, 0, buffer.Length)
    '    post.Close()

    '    Dim myResponse As HttpWebResponse = myReq.GetResponse
    '    Dim responsedata As Stream = myResponse.GetResponseStream
    '    Dim responsereader As New StreamReader(responsedata)

    '    response = responsereader.ReadToEnd
    'End Function


End Module
