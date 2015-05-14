'#########################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution	##
'##																						##
'## September 18, 2013																	##
'##																						##
'## Version 1.0																			##
'#########################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 6 Then 
	ShowError(3) 
End If
'Set Culture - en-us
SetLocale(1033)

Dim Host, MetricState, TargetIDList, SiteList, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
TargetIDList = WScript.Arguments(2)
Username = WScript.Arguments(3)
Password = WScript.Arguments(4)
Domain = WScript.Arguments(5)

Dim arrTargetsIDs, arrMetrics
arrTargetsIDs = Split(TargetIDList,",")
arrMetrics = Split(MetricState,",")

Dim objSWbemLocator, objSWbemServices, colItems

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	
Dim Counter, FullUserName,objItem
Counter = 0
	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If

	If Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, strComputer)
	End If
	If Err Then CALL ShowError(1, Host)
	
	if Err.Number = 0 Then
		objSWbemServices.Security_.ImpersonationLevel = 3

	'Status IIS Services ( IISADMIN , W3SVC )
		
		Dim colStatus, objStatus, Status
		Set colStatus = objSWbemServices.ExecQuery("SELECT State from Win32_Service Where Name ='W3SVC'",,16)
			If colStatus.Count <> 0 Then
				For Each objStatus In colStatus
					If objStatus.State = "Running" Then
						Status = 1
						If arrMetrics(18)=1 Then _
							Call Output("133:9",arrTargetsIDs(Counter),"1","")
					Else
						Status = 0
						If arrMetrics(18)=1 Then _
							Call Output("133:9",arrTargetsIDs(Counter),"0","")	
					End If
				Next
			Else
				'If there is no response in WMI query
				CALL ShowError(5, Host)
			End If
	
				Set colItems = objSWbemServices.ExecQuery("Select AnonymousUsersPersec,BytesReceivedPerSec,BytesSentPerSec,BytesTotalPerSec,ConnectionAttemptsPerSec,CurrentAnonymousUsers,CurrentConnections,CurrentNonAnonymousUsers,LogonAttemptsPersec,Name,ServiceUptime,TotalbytesReceived,TotalBytesSent,TotalBytesTransferred,TotalFilesreceived,TotalFilesSent,TotalFilesTransferred,TotalLogonAttempts,TotalMethodRequests,TotalMethodRequestsPerSec from Win32_PerfFormattedData_W3SVC_WebService where Name='_total'",,16) 
				If colItems.Count <> 0 Then
					For Each objItem in colItems
							'AnonymousUsersPersec
							If (arrMetrics(0)=1 and Status = 1) Then _
							Call Output("121:4",arrTargetsIDs(Counter),objItem.AnonymousUsersPersec,"")
							'kBytesReceivedPerSec
							If (arrMetrics(1)=1 and Status = 1) Then _
							Call Output("197:4",arrTargetsIDs(Counter),FormatNumber(objItem.BytesReceivedPerSec/1024),"")
							'kBytesSentPerSec
							If (arrMetrics(2)=1 and Status = 1) Then _
							CALL Output("166:4",arrTargetsIDs(Counter),FormatNumber(objItem.BytesSentPerSec/1024),"")
							'kBytesTotalPerSec
							If (arrMetrics(3)=1 and Status = 1) Then _
							CALL Output("111:4",arrTargetsIDs(Counter),FormatNumber(objItem.BytesTotalPerSec/1024),"")
							'ConnectionAttemptsPerSec
							If (arrMetrics(4)=1 and Status = 1) Then _
							CALL Output("127:4",arrTargetsIDs(Counter),objItem.ConnectionAttemptsPerSec,"")
							'CurrentAnonymousUsers
							If (arrMetrics(5)=1 and Status = 1) Then _
							Call Output("58:4",arrTargetsIDs(Counter),objItem.CurrentAnonymousUsers,"")
							'CurrentConnections
							If (arrMetrics(6)=1 and Status = 1) Then _
							Call Output("74:4",arrTargetsIDs(Counter),objItem.CurrentConnections,"")
							'CurrentNonAnonymousUsers
							If (arrMetrics(7)=1 and Status = 1) Then _
							CALL Output("171:4",arrTargetsIDs(Counter),objItem.CurrentNonAnonymousUsers,"")
							'LogonAttemptsPersec
							If (arrMetrics(8)=1 and Status = 1) Then _
							CALL Output("32:4",arrTargetsIDs(Counter),objItem.LogonAttemptsPersec,"")
							'ServiceUptime
							If (arrMetrics(9)=1 and Status = 1) Then _
							CALL Output("89:4",arrTargetsIDs(Counter),objItem.ServiceUptime,"")
							'TotalkbytesReceived
							If (arrMetrics(10)=1 and Status = 1) Then _
							CALL Output("135:4",arrTargetsIDs(Counter),FormatNumber(objItem.TotalbytesReceived/1024),"")
							'TotalkBytesSent
							If (arrMetrics(11)=1 and Status = 1) Then _
							CALL Output("101:4",arrTargetsIDs(Counter),FormatNumber(objItem.TotalBytesSent/1024),"")
							'TotalkBytestransferred
							If (arrMetrics(12)=1 and Status = 1) Then _
							CALL Output("42:4",arrTargetsIDs(Counter),FormatNumber(objItem.TotalBytestransferred/1024),"")
							'TotalFilesreceived
							If (arrMetrics(13)=1 and Status = 1) Then _
							CALL Output("120:4",arrTargetsIDs(Counter),objItem.TotalFilesreceived,"")
							'TotalFilesSent
							If (arrMetrics(14)=1 and Status = 1) Then _
							CALL Output("138:4",arrTargetsIDs(Counter),objItem.TotalFilesSent,"")
							'TotalFilesTransferred
							If (arrMetrics(15)=1 and Status = 1) Then _
							CALL Output("38:4",arrTargetsIDs(Counter),objItem.TotalFilesTransferred,"")
							'TotalLogonAttempts
							If (arrMetrics(16)=1 and Status = 1) Then _
							CALL Output("119:4",arrTargetsIDs(Counter),objItem.TotalLogonAttempts,"")
							'TotalMethodRequests
							If (arrMetrics(17)=1 and Status = 1) Then _
							CALL Output("211:4",arrTargetsIDs(Counter),objItem.TotalMethodRequests,"")
							'TotalMethodRequestsPerSec
							If (arrMetrics(19)=1 and Status = 1) Then _
							CALL Output("150:4",arrTargetsIDs(Counter),objItem.TotalMethodRequestsPerSec,"")
					Next
				Else
					'If there is no response in WMI query
					CALL ShowError(5, Host)
				End If
			If Err.number <> 0 Then
				CALL ShowError(5, Host)
				Err.Clear
			End If
	Counter = Counter + 1
End if		

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Sub Output(SourceUUID, TargetUUID, SourceValue, SourceObject)
	If SourceObject <> "" Then
		If SourceValue <> "" Then
			WScript.Echo ToUTC() & "|" & SourceUUID & "|" & TargetUUID & "|" & SourceValue & "|" & SourceObject & vbCr 
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If SourceValue <> "" Then
			WScript.Echo ToUTC() & "|" & SourceUUID & "|" & TargetUUID & "|" & SourceValue & vbCr 
		Else
			CALL ShowError(5, Host) 
		End If
	End If
End Sub

Function ToUTC()
	Dim dtmDateValue, dtmAdjusted
	Dim objShell, lngBiasKey, lngBias, k, UTC
	dtmDateValue = Now()
	'Obtain local Time Zone bias from machine registry.
	Set objShell = CreateObject("Wscript.Shell")
	lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
	If (UCase(TypeName(lngBiasKey)) = "LONG") Then
		lngBias = lngBiasKey
		ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
			lngBias = 0
		For k = 0 To UBound(lngBiasKey)
			lngBias = lngBias + (lngBiasKey(k) * 256^k)
		Next
	End If
	'Convert datetime value to UTC.
	UTC = DateAdd("n", lngBias, dtmDateValue)
	ToUTC =  FormatDateTime(UTC,2) & " " & FormatDateTime(UTC,3)
End Function
