Attribute VB_Name = "Module3"
Function ServiceNow(table As String, columns, Query As String, limit As String) As String()
 
		Dim objHTTP As New WinHttp.WinHttpRequest
        Dim ColumnsArray As Variant
        Dim resp As New DOMDocument60
        Dim Result As IXMLDOMNode
        Dim count As Integer
        Dim list() As String
        Dim resperror(1, 1) As String
		Dim InstanceURL As String
		Dim AuthorizationCode As String
        ' Replace with your Service Now Inctance URL
        InstanceURL = "https://xxxx.service-now.com"
        ' Replace with your Authorization code
        AuthorizationCode = "Basic "
        ' Add more tables ascomma seperated with no spaces
               
        
        
        
                        
        ColumnsArray = Split(columns, ",")
        OtherSysParam = "&sysparm_limit=" & limit
        SysQuery = "&sysparm_query=" & Query
                      
          
        URL = InstanceURL & "/api/now/table/"
        sysParam = "sysparm_display_value=true&sysparm_exclude_reference_link=true" & OtherSysParam & SysQuery & "&sysparm_fields=" & columns
        URL = URL & table & "?" & sysParam
        objHTTP.Open "get", URL, False
        objHTTP.SetRequestHeader "Accept", "application/xml"
        objHTTP.SetRequestHeader "Content-Type", "application/xml"
        
        ' Authorization Code
        objHTTP.SetRequestHeader "Authorization", AuthorizationCode
        objHTTP.Send
        
        If objHTTP.Status <> "200" Then
        resperror(0, 0) = "error"
        ServiceNow = resperror
        Debug.Print resperror(0, 0)
        Exit Function
        End If
        
        Debug.Print objHTTP.Status
        
        resp.LoadXML objHTTP.ResponseText
        
                
        count = 0
        
        ReDim list(limit, UBound(ColumnsArray))
        
        For Each Result In resp.getElementsByTagName("result")
            For x = 0 To UBound(ColumnsArray)
              
              list(count, x) = Result.SelectSingleNode(ColumnsArray(x)).Text
              Debug.Print Result.SelectSingleNode(ColumnsArray(x)).Text
                  
            Next x
            count = count + 1
            
        Next Result
        
        ServiceNow = list
        
End Function

