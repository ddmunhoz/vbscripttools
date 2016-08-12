'###############################################################################################################################################
'# Author: Diego Munhóz - munhozdiego@gmail.com - ca.linkedin.com/in/munhozdiego#
'# GitHub: https://github.com/munhozdiego/				
'# Date: 29/09/2014
'###############################################################################################################################################
'# This script will install zabbix agent and create zabbix.conf file as well.
'# It can be executed by group policy or using psexec.
'# First it copies all files to servers at root of system drive specified by parameter /s 
'# 
'# Example:
'#    Install Zabbix Agent on server XPTO in folder ZABBIX_AGENT and configure it to use zabbix server ZABBIX.
'#    Intall files are located on networkshare: \\ZabbixDistributionXPTO
'#    Zabbix config file must be named: Zabbix_CONFIG_FILE.conf
'#
'#      Command line: cmd /c cscript.exe ZBSX-mng.vbs /h:ZABBIX /c:Zabbix_Agent\Zabbix_CONFIG_FILE.conf /s:\\ZabbixDistributionXPTO /a:i
'# 
'#      Parameters
'#			/h - Zabbix Server
'#			/c - Folder where zabbix agent will be installed followed by desired config file name
'#			/s - Location of zabbix agent install files
'#			/a - Action that will be executed: Use /a:i to install or /a:u to uninstall
'#
'###############################################################################################################################################
'#----------------------Definitions--------------------------
Dim strHost, strConfName,strSourcePath, arg,confF, shell, rootFolder, strAction
Set Arg = WScript.Arguments.Named
'#-----------------------------------------------------------

'#----------------------Parameters---------------------------
strHost = Arg.Item("h")
strConfName = Arg.Item("c")
strSourcePath = Arg.Item("s")
strAction = Arg.Item("a")

set shell = WScript.CreateObject("WScript.Shell")
rootFolder =left(shell.ExpandEnvironmentStrings("%windir%"), 3)
'#-----------------------------------------------------------

'#----------------------ZabbixConfFile-----------------------
confF ="LogFile=" & rootFolder & left(strConfName,InStrRev(strConfName,"\")) & loadHostInfo("localhost")(1) & ".log" & vbCrLf &_
"Hostname=" & loadHostInfo("localhost")(1) & "." & loadHostInfo("localhost")(0) & vbCrLf & "Server=" & strHost & vbCrLf &_
"ServerActive="& strHost & vbCrLf 
'#-----------------------------------------------------------


'#-----------------------Main--------------------------------
If strAction = "i" then 
	if RegistryRead("SYSTEM\ControlSet001\Services\Zabbix Agent\","ImagePath") <> "0" then
		RunProgramWait "cmd /c " & Replace(RegistryRead("SYSTEM\ControlSet001\Services\Zabbix Agent\","ImagePath"),chr(34),"") & " --stop","n","","n","","",true,0
		Wscript.Sleep (5000)
		RunProgramWait "cmd /c " & Replace(RegistryRead("SYSTEM\ControlSet001\Services\Zabbix Agent\","ImagePath"),chr(34),"") & " --uninstall","n","","n","","",true,0 
		Wscript.Sleep (5000) 
		MngExp confF,strConfName,1
		Wscript.Sleep (1000)
		CopyF "exe",CreateObject("Scripting.FilesystemObject"), strSourcePath, rootFolder & left(strConfName,InStrRev(strConfName,"\"))
		Wscript.Sleep (5000) 
		RunProgramWait "cmd /c " & rootFolder & left(strConfName,InStrRev(strConfName,"\")) & "zabbix_agentd --conf " & rootFolder & strConfName & " --install","n","","n","","",true,0 
		Wscript.Sleep (5000)
		RunProgramWait "cmd /c " & rootFolder & left(strConfName,InStrRev(strConfName,"\")) & "zabbix_agentd --conf " & rootFolder & strConfName & " --start","n","","n","","",true,0 
	else
		MngExp confF,strConfName,1
		Wscript.Sleep (1000)
		CopyF "exe",CreateObject("Scripting.FilesystemObject"), strSourcePath, rootFolder & left(strConfName,InStrRev(strConfName,"\"))
		Wscript.Sleep (5000) 
		RunProgramWait "cmd /c " & rootFolder & left(strConfName,InStrRev(strConfName,"\")) & "zabbix_agentd --conf " & rootFolder & strConfName & " --install","n","","n","","",true,0 
		Wscript.Sleep (5000)
		RunProgramWait "cmd /c " & rootFolder & left(strConfName,InStrRev(strConfName,"\")) & "zabbix_agentd --conf " & rootFolder & strConfName & " --start","n","","n","","",true,0 
	End If
else
	If strAction ="u" then 
		RunProgramWait "cmd /c " & Replace(RegistryRead("SYSTEM\ControlSet001\Services\Zabbix Agent\","ImagePath"),chr(34),"") & " --stop","n","localhost","n","","",true,0
		Wscript.Sleep (5000)
		RunProgramWait "cmd /c " & Replace(RegistryRead("SYSTEM\ControlSet001\Services\Zabbix Agent\","ImagePath"),chr(34),"") & " --uninstall","n","localhost","n","","",true,0 
	End If
End If 
'#------------------------------------------------------------


'#-----------------------ZabbixUninstall----------------------
Function RegistryRead(iPath,iValue)
	Dim strComputer, strKeyPath, strValueName
	Const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."
	Set objRegistry = GetObject("winmgmts:\\" & _ 
	strComputer & "\root\default:StdRegProv")

	strKeyPath = iPath
	strValueName = iValue
	objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

	If IsNull(strValue) Then 
		RegistryRead = 0
	Else
		RegistryRead = strValue
	End If
End Function
'#------------------------------------------------------------
'#-----------------------Host Info----------------------------
Function loadHostInfo(iHost)
	Dim objWMIService, colComputer, objComputer
	Dim hostDomain,hostName,hostIP
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & iHost & "\root\cimv2")
	Set colComputer = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each objComputer in colComputer 
		hostDomain=objcomputer.domain 
		hostName=objcomputer.name 
	Next
	loadHostInfo=array(hostDomain,hostName,hostIp)
End Function
'#------------------------------------------------------------
'#-----------------------Buid Path----------------------------
Sub BuildPath(ByVal Path)
	Dim objFSO
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FolderExists(Path) Then
		BuildPath objFSO.GetParentFolderName(Path)
		objFSO.CreateFolder Path
	End If
End Sub
'#------------------------------------------------------------
'#-----------------------EXP_CONF-----------------------------
Function MngExp(iVar,Ifile,ifn)
	Dim File,objFSO,FileC
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	File=rootFolder & Ifile 
	' wscript.echo File 
	' wscript.echo right(file,len(file) - InStrRev(File,"\"))
	' wscript.echo left(file,InStrRev(File,"\"))
	If Not objFSO.FolderExists(File) Then
		BuildPath left(file,InStrRev(File,"\"))
	End if 
	If (ifn=1) then 
		Set objFile = objFSO.CreateTextFile(File,True)
		objFile.Write iVar
		objFile.Close
	else
		If (ifn=0) then 
			Set objFile = objFSO.OpenTextFile(File)
			Do Until objFile.AtEndOfStream
				'FileC=objFile.ReadLine
				FileC=objFile.ReadAll
				iVar = FileC
			Loop
			objFile.Close
		End If
	End If 
End Function
'#------------------------------------------------------------

'#-----------------------Copy LNK-----------------------------
Function CopyF(sTyp, oFSO, sStart, sEnd) 'As Integer
	Dim d, f, i, ext
	If Right(sEnd, 1) <> "\" Then _
		sEnd = sEnd & "\"

	If Not oFSO.FolderExists(sEnd) Then _
		oFSO.CreateFolder sEnd
		With oFSO.GetFolder(sStart)
			For Each f In .Files
				ext = oFSO.GetExtensionName(f)
				If Len(ext) And InStr(1, sTyp, ext, 1) And _
					ShouldCopy(oFSO, sEnd & f.Name, f.DateLastModified) Then
					f.Copy sEnd
					i = i + 1
				End If
			Next 
		End With
CopyLnkFolder = i
End Function
'#------------------------------------------------------------


Function ShouldCopy(oFSO, sFile, dCutoff) 'As Boolean
	With oFSO
		If Not .FileExists(sFile) Then
			ShouldCopy = True
			Exit Function
		End If
		ShouldCopy = .GetFile(sFile).DateLastModified < dCutoff
	End With
End Function
'#------------------------------------------------------------
'#-----------------------Check Service------------------------
Function Check(Service)
	Dim ctrlS
	Set objComputer = GetObject("WinNT://localhost,computer")
	objComputer.Filter = Array("Service")
	strServiceName = Service
	For Each aService In objComputer
		If LCase(strServiceName) = LCase(aService.Name) Then
		ctrlS = aService.Status 
	End If
	Next 
		if ctrls = 4 then
			check = 4
		else
			check = 0
		End if
End Function 
'#------------------------------------------------------------
'#--------------Run Program Subs------------------------------
Sub RunProgramWait(iApp,iVrb,iName,iPrm,iPrmS,iPrmE,iPrmW,iPrmShw)
	Dim oShell 
	If IPrmShw <> "" then
		If IPrmShw=1 then
			IPrmShw=1
		else
			IPrmShw=0
		End If 
	else
		IPrmShw=0
	End If 

	If IPrmw <> "" then
		If (Iprmw="true" or Iprmw="TRUE") then
			Iprmw="true"
		else
			Iprmw="false"
		End If 
	else
		Iprmw="true"
	End If


	Set oShell = WScript.CreateObject("WSCript.shell")
	If (iVrb = "y" or iVrb = "Y") then 
		If (iPrm = "y" or iPrm = "Y") then 
			icmd = iApp &" "& iPrms & iName & iPrme &" > " & loadHostInfo("localhost")(1) & "_log.txt"
			oShell.run iCmd, iPrmShw, iPrmw
		else
			oShell.run iApp &" > " & loadHostInfo("localhost")(1) & "_log.txt", iPrmShw, iPrmw
		end If 
	else
		iCmd = iApp &" "& iPrms & iName & iPrme
		oShell.run iCmd , iPrmShw, iPrmw
	End If 
	Set oShell = Nothing
End Sub 
'#------------------------------------------------------------