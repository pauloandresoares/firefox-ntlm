'Name: firefox_ntlm_preference_conf.vbs
'Author: Matt Martz <mdmartz@sivel.net>
'Date: 2007-05-29
'
'This script is designed to allow an organization to manage a list of sites that Firefox will send 
'NTLM authentication to.  This list is similar to the Local Intranet zone located in the Internet 
'Explorer configuration options.  With this preference enabled users will be automatically 
'authenticated to sites which Internet Explorer would automatically be authenticated to.
'
'Note from the author:
'	1) I don't claim that this (script)  is well written.  It started as a script to just append the 
'	proposed site list to the end of the file and grew into a beast that would add the proposed 
'	site list and and sites that were configured by the user that didn't exist in the proposed 
'	site list
'	2) VBScript sucks at parsing and editing files.  It wasn't meant for this purpose but with
'	some innovation can be made to do what you want.  Once again this is sloppy code and it 
'	would be more efficient written in a different language.  But VBScript is natively supported
'	on Windows which makes it easy for group policy deployment
'	3) I'm sure something else sucks about this script but I just haven't found it yet...So if you 
'	find something else that sucks you can let me know at the above email address and I'll see 
'	if I can make it suck less.
'
'Define Variables
option explicit
Dim fsObj, WshShell, fullpath, SubFolders, prefsfile, strLine
Dim strSiteList, strAppData, strFireProf, strProfFolder, strSearchString
Dim Folder, i, arrSiteFirst, arrSiteLast, strNewContents
Dim intMatchNTLM, strLineResult, intMatchLine, strAddExistSiteList
Dim arrCurSiteList, intCurSiteArrayLength, strExistSiteList
Dim arrSiteList, intSiteArrayLength, i_addsite, intExistSiteListLen

'Configure Environment
Set fsObj = CreateObject("Scripting.FileSystemObject") 
Set WshShell = CreateObject( "WScript.Shell" )

'Define the path to the profiles directory for Firefox
strAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
strFireProf = "\Mozilla\Firefox\Profiles\"

'Check if firefox even has a profiles directory, if not than quit
If fsObj.FolderExists(strAppData & strFireProf) = False Then
	wscript.quit
End If

'Define the proposed site list, split it into an array and determine length of array
'strSiteList should be a list of sites beginning with http:// or https:// 
'Only the DNS portion of the site should follow and there should be no trailing slash
'THIS (strSiteList) SHOULD BE THE ONLY VARIABLE THAT NEEDS CONFIGURATION
strSiteList = "sso.biosfera.mma.gov.br, biosferasso.florestal.gov.br"
arrSiteList = Split(strSiteList, ",", -1, 1)
intSiteArrayLength = UBound(arrSiteList)

'Create an array of the subfolders in the profiles directory (there should only be one)
Set fullpath = fsObj.GetFolder(strAppData & strFireProf)
Set Subfolders = fullpath.subfolders

'Set the full path of the prefs.js file
For Each Folder in Subfolders
	strProfFolder=Folder
	Set prefsfile = fsObj.OpenTextFile(Folder & "\prefs.js", 1, True) 
Next

'Lets load the prefsfile as a string for easy searching and then search for the preference we are looking for
'If found intMatchNTLM will be greater than 0
strSearchString = prefsfile.ReadAll
intMatchNTLM = InStr(strSearchString, "network.automatic-ntlm-auth.trusted-uris")
prefsfile.close

'Lets reopen the prefs.js file so we are at the beginning
Set prefsfile = fsObj.OpenTextFile(strProfFolder & "\prefs.js", 1, True) 

'Now we do a while loop to parse the file and look for the ntlm preference
do while prefsfile.AtEndOfStream = 0 
	'Read in the current line and check for presence of the preference name
	'If found intMatchLine will be greater than 0
	strLineResult = prefsfile.ReadLine
	intMatchLine = InStr(strLineResult, "network.automatic-ntlm-auth.trusted-uris")
	'Check if the preference exists and if so...
	if intMatchLine>0 Then
		'Format the string, removing unwanted characters so delimiting is easier and convert to an array
		strLineResult = Replace(strLineResult, chr(34), "")
		strLineResult = Replace(strLineResult, ")", "")
		strLineResult = Replace(strLineResult, ";", "")
		arrCurSiteList = Split(strLineResult, ",", -1, 1)
		'Get the length of the array and loop it to create a list of sites that are in the existing list but not in the proposed list
		intCurSiteArrayLength = UBound(arrCurSiteList)
		for i=1 to intCurSiteArrayLength
			strExistSiteList = strExistSiteList & "," & arrCurSiteList(i)
		Next
		intExistSiteListLen = Len(strExistSiteList)
		strExistSiteList = Right(strExistSiteList, intExistSiteListLen-2)
		strAddExistSiteList = strExistSiteList
		for i_addsite=0 to intSiteArrayLength
			strAddExistSiteList = Replace(strAddExistSiteList,arrSiteList(i_addsite), "")
		next
		'Continuing with the above comment we need to check for formatting issues caused by removing records and fix them
		strAddExistSiteList = Replace(strAddExistSiteList,",,", ",")
		if InStr(strAddExistSiteList, ",") = 1 then
			strAddExistSiteList = Right(strAddExistSiteList, Len(strAddExistSiteList)-1)
		end if
		if InStrRev(strAddExistSiteList, ",") = Len(strAddExistSiteList) and InStrRev(strAddExistSiteList, ",") <> 0 then
			strAddExistSiteList = Left(strAddExistSiteList, Len(strAddExistSiteList)-1)
		end if
		if Len(strAddExistSiteList) = 0 Then
			strSiteList = strSiteList
			strSiteList = Replace(strSiteList,",,", ",")
		else
			strSiteList = strSiteList & "," & strAddExistSiteList
			strSiteList = Replace(strSiteList,",,", ",")
		end if
		
	end If
Loop

'Now that we have the list of sites to match the proposed list plus the sites that didn't overlap between existing and proposed
'If the preference was not set at all we add just the proposed list
if intMatchNTLM<1 Then
	'Open the file for appending as we aren't editing the contents
	Set prefsfile = fsObj.OpenTextFile(strProfFolder & "\prefs.js", 8, True)
	prefsfile.writeline("user_pref(" & chr(34) & "network.automatic-ntlm-auth.trusted-uris" & chr(34) & ", " & chr(34) & strSiteList & chr(34) & ");")
	prefsfile.close
else 
	'If the existing list is longer than the proposed list and the existing list contains sites then quit
	if Len(strExistSiteList) >= Len(strSiteList) and Len(strAddExistSiteList) = 0 Then
		wscript.quit
	'else lets check to see how we add the list
	else
		'Open the file for reading
		Set prefsfile = fsObj.OpenTextFile(strProfFolder & "\prefs.js", 1, True)
		'Loop the file writing all lines excluding the preference we are writing to a string for easy manipulation
		Do Until prefsfile.AtEndOfStream
			strLine = prefsfile.Readline
			strLine = Trim(strLine)
			If strLine <> "user_pref(" & chr(34) & "network.automatic-ntlm-auth.trusted-uris" & chr(34) & ", " & chr(34) & strExistSiteList & chr(34) & ");" Then
				strNewContents = strNewContents & strLine & vbCrLf
			End If
		Loop
		prefsfile.close
		'Lets reopen the file in editing mode and write the contents of the file minus the preference we are setting
		Set prefsfile = fsObj.OpenTextFile(strProfFolder & "\prefs.js", 2, True)
		prefsfile.write strNewContents
		prefsfile.close
		'Now lets reopen the file in append mode and append the preference we are setting to the end
		Set prefsfile = fsObj.OpenTextFile(strProfFolder & "\prefs.js", 8, True)
		prefsfile.writeline("user_pref(" & chr(34) & "network.automatic-ntlm-auth.trusted-uris" & chr(34) & ", " & chr(34) & strSiteList & chr(34) & ");")
		prefsfile.close
	end if
end If