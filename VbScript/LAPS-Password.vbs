' Script to query Active Directory for a machine's LAPS password and related details
' Usage: cscript /nologo LAPS-Password.vbs

Option Explicit
Dim strLDAPDomain, strComputer
Dim objConnection, objCommand, objRecordSet, objShell, lngBiasKey, lngBias

strLDAPDomain = "LDAP://DC=Example,DC=com"
strComputer = InputBox ("Machine name", "Computer Name")

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

' Search for Specific Computer Accounts
Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
"Select Name, operatingSystem, operatingSystemVersion, Description, ms-Mcs-AdmPwd, ms-Mcs-AdmPwdExpirationTime, pwdLastSet from '" & _
strLDAPDomain &"' where objectClass='computer' and Name = '" & strComputer & "'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

' Obtain local time zone bias from machine registry.
' This bias changes with Daylight Savings Time.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngBias = lngBias + (lngBiasKey(k) * 256^k)
    Next
End If

Do Until objRecordSet.EOF
	Wscript.Echo "Computer Name: " & objRecordSet.Fields("Name").Value
	Wscript.Echo "OS Version: " & objRecordSet.Fields("operatingSystem").Value & " - " & objRecordSet.Fields("operatingSystemVersion").Value
	Wscript.Echo "LAPS Password: " & objRecordSet.Fields("ms-Mcs-AdmPwd").Value
	Dim objDatePwdLastSet,dtmPwdLastSet
	If (TypeName(objRecordSet.Fields("pwdLastSet").Value) = "Object") Then
		Set objDatePwdLastSet = objRecordSet.Fields("pwdLastSet").value
		dtmPwdLastSet = Integer8Date(objDatePwdLastSet, lngBias)
		Wscript.echo "Password Last Set: " & dtmPwdLastSet
	End If
	Dim objDatePwdExpire,dtmPwdExpire
	If (TypeName(objRecordSet.Fields("ms-Mcs-AdmPwdExpirationTime").Value) = "Object") Then
		Set objDatePwdExpire = objRecordSet.Fields("ms-Mcs-AdmPwdExpirationTime").value
		dtmPwdExpire = Integer8Date(objDatePwdExpire, lngBias)
		Wscript.echo "LAPS Password Expires: " & dtmPwdExpire
	End If
	objRecordSet.MoveNext
Loop

Function Integer8Date(ByVal objDate, ByVal lngBias)
    ' Function to convert Integer8 (64-bit) value to a date, adjusted for
    ' local time zone bias.
    Dim lngAdjust, lngDate, lngHigh, lngLow
    lngAdjust = lngBias
    lngHigh = objDate.HighPart
    lngLow = objdate.LowPart
    ' Account for error in IADsLargeInteger property methods.
    If (lngLow < 0) Then
        lngHigh = lngHigh + 1
    End If
    If (lngHigh = 0) And (lngLow = 0) Then
        lngAdjust = 0
    End If
    lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
        + lngLow) / 600000000 - lngAdjust) / 1440
    ' Trap error if lngDate is ridiculously huge.
    On Error Resume Next
    Integer8Date = CDate(lngDate)
    If (Err.Number <> 0) Then
        On Error GoTo 0
        Integer8Date = #1/1/1601#
    End If
    On Error GoTo 0
End Function
