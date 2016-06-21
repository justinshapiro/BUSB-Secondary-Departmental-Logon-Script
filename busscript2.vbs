'===============================================================================
' NAME: 	busscript2.vbs
' AUTHOR: 	OIT, CU Denver
' REVISED:	REVISED BY Justin Shapiro
' DATE  : 	06/20/16
' COMMENT:  	Secondary departmental login script for UCDENVER.PVT scripting
'				environment; called from the AD user objects in the department OU
'===============================================================================

'comment out the below line for debugging
on error resume next 

' declare variables
Dim objNetwork, objDrives, objShell
Dim strSubst, strSubstVal, strSubstName, strEnumDrive

'dimension 1 never resizes, dimension 2 dynamically sizes to number of current mapped drives
ReDim existingDrives(1, 0) 

' initialize network objects
Set fso = CreateObject("Scripting.FileSystemObject")
Set objNetwork = WScript.CreateObject("WScript.Network")
Set wshshell = wscript.createobject("WScript.Shell")
Set wshnetwork  = wscript.createobject("WScript.Network")
Set objShell = CreateObject("Shell.Application")
Set objDrives = objNetwork.EnumNetworkDrives

'determine user name
StrUser = wshnetwork.username

'use WINNT provider get list of groups the user belongs To
ADSpath = "WinNT://UNIVERSITY/" & StrUser

'fill a two-dimensional array with the current mapped drives on the computer
GetCurrentMappedDrives

'loop through all the groups and execute commands
'to avoid ambiguity, all group names must be capitalized!

Set ADSObj = GetObject(ADSpath)
UserMembership = ""
For Each Group In ADSObj.groups
	UserMembership = UserMembership & VbCrLf & UCase(Group.name)
	Select Case Ucase(Group.name)
		Case "TBS BUSIT"
			driveLetter = "A:"
			drivePath = "\\tbs-fs4\DeptShare\BUS IT"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS ANNUAL REPORTS" 
			driveLetter = "B:"
			drivePath = "\\tbs-fs4\DeptShare\Annual Reports"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS ARCHIVE" 
			driveLetter = "G:"
			drivePath = "\\tbs-fs4\DeptShare\Archive"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS JPMCC-RMI"
			driveLetter = "H:"
			drivePath = "\\tbs-fs4\DeptShare\JPMCC-RMI"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS AWARDS"
			driveLetter = "I:"
			drivePath = "\\tbs-fs4\DeptShare\awards"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS BCC"
			driveLetter = "J:"
			drivePath = "\\tbs-fs4\DeptShare\BCC"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS BUDGET"
			driveLetter = "K:"
			drivePath = "\\tbs-fs4\DeptShare\Budget"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS BUSINT"
			driveLetter = "L:"
			drivePath = "\\tbs-fs4\DeptShare\BusInt"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS DEANS OFFICE"
			driveLetter = "M:"
			drivePath = "\\tbs-fs4\DeptShare\Dean's"
			Call ProcessDriveMap(driveLetter, drivePath)
			
		' Drive letter N is available
		
		Case "TBS GRAD"
			driveLetter = "O:"
			drivePath = "\\tbs-fs4\DeptShare\Grad"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS GEM"
			driveLetter = "O:"
			drivePath = "\\Wilson\cob$\GEM"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS HEALTHADMIN"
			driveLetter = "P:"
			drivePath = "\\tbs-fs4\DeptShare\HealthAdmin"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS JPMCC"
			driveLetter = "Q:"
			drivePath = "\\tbs-fs4\DeptShare\JPMCC"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS RMI"
			driveLetter = "R:"
			drivePath = "\\tbs-fs4\DeptShare\RMI"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS MGMT"
			driveLetter = "S:"
			drivePath = "\\tbs-fs4\DeptShare\MGMT"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS OFFICE SUPPORT"
			driveLetter = "T:"
			drivePath = "\\tbs-fs4\DeptShare\OfficeSupport"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS OSSEVENT"
			driveLetter = "U:"
			drivePath = "\\tbs-fs4\DeptShare\OSSEVENT"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS PROGRAMS OFFICE"
			driveLetter = "V:"
			drivePath = "\\tbs-fs4\DeptShare\Programs"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS FACULTY" 
			driveLetter = "W:"
			drivePath = "\\tbs-fs4\FacultyShare\" & StrUser
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS STAFF" 
			driveLetter = "W:"
			drivePath = "\\tbs-fs4\StaffShare\" & StrUser
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS UNDERGRAD"
			driveLetter = "X:"
			drivePath = "\\tbs-fs4\DeptShare\Undergrad"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS PRISK" 
			driveLetter = "Y:"
			drivePath = "\\tbs-fs4\DeptShare\prisk"
			Call ProcessDriveMap(driveLetter, drivePath)
		Case "TBS COMC" 
			driveLetter = "Z:"
			drivePath = "\\tbs-fs4\DeptShare\Comc"
			Call ProcessDriveMap(driveLetter, drivePath)
	End Select
Next

'==========================================================================================='
' HELPER FUNCTION: "GetCurrentMappedDrives"
' 		  Purpose: To dynamically fill a two-dimensional array with the drive path and letter 
'                  of current mapped drives on the user profile.
'==========================================================================================='
Function GetCurrentMappedDrives
	For i = 0 to objDrives.Count - 1 Step 2
		strSubst = objShell.NameSpace(objDrives.Item(i) & Chr(92)).Self.Name 
		strSubstVal = inStr(1,strSubst, Chr(40)) - 2
		strSubstName = Mid(strSubst, 1, strSubstVal)
		
		existingDriveLetter = objDrives.Item(i)
		existingDrivePath = objDrives.Item(i + 1)
				
		ReDim Preserve existingDrives(1, i)
		existingDrives(0, i) = existingDriveLetter
		existingDrives(1, i) = existingDrivePath		
	Next
End Function


'==========================================================================================='
' HELPER FUNCTION: "ProcessDriveMap"
' 		  Purpose: Using the array of mapped drives to the user profile, remove each currently
'                  mapped drive. Then, map the current Case's associated network share.         
'==========================================================================================='
Function ProcessDriveMap(driveLetter, drivePath)
	For i = 0 to UBound(existingDrives, 2)
		If (existingDrives(1, i) = drivePath) Then
			wshnetwork.RemoveNetworkDrive existingDrives(0, i), True, True 
		End If
	Next
	
	objNetwork.MapNetworkDrive driveLetter, drivePath, True
End Function

'Garbage Collection
Set wshshell = Nothing
Set wshnetwork = Nothing
Set objNetwork = Nothing
Set ADSObj = Nothing
Set objNetwork = Nothing
Set objShell = Nothing
Set objDrives = Nothing
Set strSubst = Nothing
Set strSubstVal = Nothing
Set strSubstName = Nothing
Set strEnumDrive = Nothing

'Quit Script
WScript.quit
