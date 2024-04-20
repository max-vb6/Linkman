Attribute VB_Name = "modLink"
Option Explicit

Public Function LinkConfig(sIP As String, sSubnetMask As String, sGateway As String, sMainDNS As String, Optional sSecondDNS As String) As Long
    On Error GoTo LCErr:
    
    Dim strComputer As String, objWMIService As Object, colNetAdapters As Object
    Dim strIPAddress As Variant, strSubnetMask As Variant, strGateway As Variant, strGatewaymetric As Variant, strDNS As Variant
    Dim objNetAdapter As Object, errEnable As Long, errGateways As Long, errDNS As Long
    
    strComputer = "."
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    
    strIPAddress = Array(sIP) 'IPµØÖ·
    strSubnetMask = Array(sSubnetMask) '×ÓÍø
    strGateway = Array(sGateway) 'Gateways
    strDNS = Array(sMainDNS, sSecondDNS) 'MAIN DNS AND SECOND DNS
    strGatewaymetric = Array(1)
    
    For Each objNetAdapter In colNetAdapters
        errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
        errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
        errDNS = objNetAdapter.SetDNSServerSearchOrder(strDNS)
        
        LinkConfig = Not (errEnable = 0 And errGateways = 0 And errDNS = 0)
    Next
    
    Exit Function
LCErr:
    LinkConfig = 1
End Function

Public Function PingIP(sIP As String) As Boolean
    On Error GoTo PIErr
    
    Dim objWMIService As Object
    Dim colItems      As Object
    Dim objItem       As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PingStatus Where Address='" & sIP & "'")
    For Each objItem In colItems
        PingIP = (objItem.StatusCode = 0)
    Next
    
    Exit Function
PIErr:
    PingIP = False
End Function
