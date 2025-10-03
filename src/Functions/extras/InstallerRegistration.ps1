<#
.SYNOPSIS
    Forcefully removes Windows Installer (MSI) product registrations when normal
    uninstall or repair is not possible.

.DESCRIPTION
    This function scrubs Windows Installer registration for a given product when
    cached MSI/MSP files are missing and neither repair nor uninstall can complete.

    > WARNING: This is a destructive recovery tool. It does not remove files
    > from disk, only the MSI/MSP registration data in the registry and Windows
    > Installer metadata. Use it only when standard uninstall/repair paths fail.

    This function wraps and adapts routines originally authored by Microsoft as
    part of their Program Install and Uninstall Troubleshooter. Specifically,
    the underlying helper functions are extracted and adapted from:
        MicrosoftProgram_Install_and_Uninstall.meta.diagcab\MSIMATSFN.ps1

    Original Microsoft tool and details:
    https://support.microsoft.com/en-us/topic/fix-problems-that-block-programs-from-being-installed-or-removed-cca7d1b6-65a9-3d98-426b-e9f927e1eb4d

.PARAMETER Filter
    A scriptblock filter used to target one or more registered MSI installations
    from the registry. The scriptblock is evaluated against uninstall registry keys.

    Example values:
        { $_.DisplayName -like "*SQL*" }
        { $_.Publisher -eq "Microsoft Corporation" -and $_.DisplayName -like "*Agent*" }

.EXAMPLE
    PS> Remove-InstallerRegistration -Filter { $_.DisplayName -like "Azure Connected Machine Agent*" -and $_.DisplayVersion -eq "1.56.03167" }

    Removes registration data for all MSI products with DisplayName matching "Azure Connected Machine Agent*" and DisplayVersion equal to "1.56.03167"

.EXAMPLE
    PS> Remove-InstallerRegistration -Filter { $_.Publisher -eq "Microsoft Corporation" -and $_.DisplayName -like "*SQL*" }

    Removes registration for Microsoft SQL Server-related products.

.NOTES
    Author: Joey Eckelbarger
    Derived from Microsoft’s Program Install/Uninstall Troubleshooter files (MSIMATSFN.ps1).

    License Note:
        The embedded helper functions are Microsoft-authored and used under the
        terms of their published diagnostic tool conditions. They are included here
        solely to enable advanced recovery where cached MSI/MSP files cannot be sourced
        and attempts to repair the installation via MSI fail. 

    Limitations:
        - Removes registry/metadata only. Program files may remain on disk.
        - Should be considered a last resort to enable reinstallation.
        - Always test in a lab environment prior to production use.
#>
function Remove-InstallerRegistration {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Filter
    )
    #################################################################################
    # Copyright (c) 2011, Microsoft Corporation. All rights reserved.
    #
    # You may use this code and information and create derivative works of it,
    # provided that the following conditions are met:
    # 1. This code and information and any derivative works may only be used for
    # troubleshooting a) Windows and b) products for Windows, in either case using
    # the Windows Troubleshooting Platform
    # 2. Any copies of this code and information
    # and any derivative works must retain the above copyright notice, this list of
    # conditions and the following disclaimer.
    # 3. THIS CODE AND INFORMATION IS PROVIDED `AS IS'' WITHOUT WARRANTY OF ANY 
    # KIND, WHETHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
    # WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. IF THIS
    # CODE AND INFORMATION IS USED OR MODIFIED, THE ENTIRE RISK OF USE OR RESULTS IN
    # CONNECTION WITH THE USE OF THIS CODE AND INFORMATION REMAINS WITH THE USER.
    #################################################################################

    # .\MicrosoftProgram_Install_and_Uninstall.meta.diagcab\MSIMATSFN.ps1
    function MATSFingerPrint {
        Param (
            $ProductCode,
            $Value,
            $Type,
            $DateTimeRun
        )
        if (!(test-path "HKLM:\Software\Microsoft\MATS\WindowsInstaller\$ProductCode\$DateTimeRun")){
            New-Item -Path hklm:\software\Microsoft\MATS\ -ErrorAction SilentlyContinue
            New-Item -Path hklm:\software\Microsoft\MATS\WindowsInstaller -ErrorAction SilentlyContinue
            New-Item -Path hklm:\software\Microsoft\MATS\WindowsInstaller\$ProductCode -ErrorAction SilentlyContinue
            New-Item -Path hklm:\software\Microsoft\MATS\WindowsInstaller\$ProductCode\$DateTimeRun -ErrorAction SilentlyContinue
        }
        New-ItemProperty hklm:\software\Microsoft\MATS\WindowsInstaller\$ProductCode\$DateTimeRun -name $Value -value $Type -propertyType String
    }

    function UninstallProduct {
$ProductUninstall=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
public class UninstallProduct2
{
    public enum INSTALLUILEVEL
    {
        INSTALLUILEVEL_NOCHANGE = 0,     // UI level is unchanged
        INSTALLUILEVEL_DEFAULT = 1,     // default UI is used
        INSTALLUILEVEL_NONE = 2,     // completely silent installation
        INSTALLUILEVEL_BASIC = 3,     // simple progress and error handling
        INSTALLUILEVEL_REDUCED = 4,     // authored UI, wizard dialogs suppressed
        INSTALLUILEVEL_FULL = 5,     // authored UI with wizards, progress, errors
        INSTALLUILEVEL_ENDDIALOG = 0x80, // display success/failure dialog at end of install
        INSTALLUILEVEL_PROGRESSONLY = 0x40, // display only progress dialog
        INSTALLUILEVEL_HIDECANCEL = 0x20, // do not display the cancel button in basic UI
        INSTALLUILEVEL_SOURCERESONLY = 0x100, // force display of source resolution even if quiet
    }
        [DllImport("msi.dll")]
                public static extern int MsiConfigureProduct(
                string szProduct,                // product code
                int iInstallLevel,                // install level
                int eInstallState     // install state
                );
                [DllImport("msi.dll", SetLastError = true)]
        public static extern int MsiSetInternalUI(INSTALLUILEVEL dwUILevel, ref int phWnd);
                public static int LogandUninstallProduct(string szProductCode)
        {
                            int phWnd=0;
                            MsiSetInternalUI(INSTALLUILEVEL.INSTALLUILEVEL_NONE, ref phWnd);
                int iErrorCode = MsiConfigureProduct(szProductCode, 1, 2);
                return (iErrorCode);
        }
}
"@
                    
        if ([UninstallProduct2] -eq $null) {
            #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
        } else {
            #type is loaded, do nothing
        }

        Trap [Exception] {
            #type doesn't exist.  Load it
            $type = Add-Type -TypeDefinition $ProductUninstall -PassThru
            continue
        }
    }

    function CreateRegistryFileRecoveryFile {
        Param($ProductCode)
        $root=$Env:SystemDrive
        $DirectoryPath= $root+"\MATS\$ProductCode"
        if (!(Test-Path -path $DirectoryPath)) { 
            [void]( New-Item  "$DirectoryPath" -type directory -force -ErrorAction SilentlyContinue) #Create the backup folder
        }

$WriteFile=
@"
Function RestoreFilesFromXML
{
    [void](New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT) # Allows access to registry HKEY_CLASSES_ROOT
    if ((Test-Path -path "$DirectoryPath\\FileBackupTemplate.xml")) 
    { 
        `$xml =[xml] (get-content "$DirectoryPath\\FileBackupTemplate.xml")
        `$root =`$xml.FileBackup.File
        foreach (`$XMLItems in `$xml.FileBackup.File)
        {     
                if (!(Test-Path `$XMLItems.FileBackupLocation))
                {
                    if (!(Test-Path -path `$XMLItems.FileBackupLocation.substring(0,(`$XMLItems.FileBackupLocation.LastIndexOf("\"))))) 
                        { 
                                [void](New-Item `$XMLItems.FileBackupLocation.substring(0,(`$XMLItems.FileBackupLocation.LastIndexOf("\"))) -type directory) #Create the backup directories ready to copy the files into then
                        }
                    Copy-Item `$XMLItems.FileDestination -Destination (`$XMLItems.FileBackupLocation) -ErrorAction SilentlyContinue #copy files to the backup folder
                }
        }
    }
}

Function RestoreRegistryFiles
{
    [void](New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT) # Allows access to registry HKEY_CLASSES_ROOT
    if ((Test-Path -path "$DirectoryPath\registryBackupTemplate.xml")) 
    { 
        `$xml =[xml] (get-content "$DirectoryPath\registryBackupTemplate.xml")
        `$root =`$xml.RegistryBackup.Registry
        foreach (`$XMLItems in `$xml.RegistryBackup.Registry)
        {     
                if (!(CheckRegistryValueExists -RegRoot `$XMLItems.RegistryHive -RegValue `$XMLItems.RegistryName))
                {
                if (`$XMLItems.RegistryType -eq "MultiString")
                {
                [string[]]`$MultiString=`$XMLItems.RegistryValue.Split(";")
                new-item -path `$XMLItems.RegistryHive -ErrorAction SilentlyContinue | Out-Null 
                New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$MultiString -propertyType `$XMLItems.RegistryType
                }
                elseif(`$XMLItems.RegistryType -eq "Binary")
                {
                new-item -path `$XMLItems.RegistryHive                         
                [Byte[]]`$ByteArray = @()
                foreach(`$byte in `$XMLItems.RegistryValue.Split(",")){`$ByteArray += `$byte}
                New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$ByteArray -propertyType `$XMLItems.RegistryType
                }
                elseif(`$XMLItems.RegistryType -eq "Dword")
                {
                `$DwordValue =Hex2Dec `$XMLItems.RegistryValue
                new-item -path `$XMLItems.RegistryHive -ErrorAction SilentlyContinue | Out-Null  
                New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$DwordValue  -propertyType `$XMLItems.RegistryType
                }
                else
                {
                ##Main Registry key missing so create here
                new-item -path `$XMLItems.RegistryHive -ErrorAction SilentlyContinue | Out-Null  
                New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$XMLItems.RegistryValue -propertyType `$XMLItems.RegistryType
                }
                }
            }
    }
}
Function CheckRegistryValueExists
{
Param(`$RegRoot,`$RegValue)
Get-ItemProperty `$RegRoot `$RegValue -ErrorAction SilentlyContinue | Out-Null  
`$?
}
function Hex2Dec
{
param(`$HEX)
ForEach (`$value in `$HEX)
{
    [Convert]::ToInt32(`$value,16)
}
}

function Test-Administrator
{
`$user = [Security.Principal.WindowsIdentity]::GetCurrent() 
(New-Object Security.Principal.WindowsPrincipal `$user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
`$UsersTemp= `$Env:temp
Copy-Item  "$DirectoryPath\\RestoreYourFilesAndRegistry.ps1" -Destination (`$UsersTemp+"\RestoreYourFilesAndRegistry.ps1") -ErrorAction SilentlyContinue
If(Test-Administrator)
{
RestoreRegistryFiles
RestoreFilesFromXML
}
else
{
cls
`$ElevatedProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell"     
`$ElevatedProcess.Arguments =  (`$UsersTemp+"\RestoreYourFilesAndRegistry.ps1")
`$ElevatedProcess.Verb = "runas"      
[System.Diagnostics.Process]::Start(`$ElevatedProcess)      
exit
}
"@

    $writefile | Out-File $DirectoryPath\RestoreYourFilesAndRegistry.ps1 -ErrorAction SilentlyContinue | Out-Null  
    
    }
    function Compressed {
$sourceDelete = @"
        using System;
        using System.Collections.Generic;
        using System.Collections;
        using System.Text;
        using Microsoft.Win32;
public class CleanUpRegistry
{
        public static string ReverseString(string szGUID)
        {
                char[] arr = szGUID.ToCharArray();
                Array.Reverse(arr);
                return new string(arr);
        }
        public static string CompressGUID(string szStandardGUID)
        {
                return (ReverseString(szStandardGUID.Substring(1, 8)) +
                    ReverseString(szStandardGUID.Substring(10, 4)) +
                    ReverseString(szStandardGUID.Substring(15, 4)) +
                    ReverseString(szStandardGUID.Substring(20, 2)) +
                    ReverseString(szStandardGUID.Substring(22, 2)) +
                    ReverseString(szStandardGUID.Substring(25, 2)) +
                    ReverseString(szStandardGUID.Substring(27, 2)) +
                    ReverseString(szStandardGUID.Substring(29, 2)) +
                    ReverseString(szStandardGUID.Substring(31, 2)) +
                    ReverseString(szStandardGUID.Substring(33, 2)) +
                    ReverseString(szStandardGUID.Substring(35, 2)));
        }
}
"@
                
        if ([CleanUpRegistry] -eq $null) {
            #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
        } else {
            #type is loaded, do nothing
        }

        Trap [Exception] {
            #type doesn't exist.  Load it
            $type=Add-Type -TypeDefinition $sourceDelete -PassThru
            continue
        }
    }

    function RegistryHiveReplace {
        Param(
            [string]$RegistryKey
        )

        $RegistryKey=$RegistryKey.toupper()
        $RegistryKey=$RegistryKey -Replace("HKEY_LOCAL_MACHINE","HKLM:") 
        $RegistryKey=$RegistryKey -Replace("HKEY_CLASSES_ROOT","HKCR:")      
        $RegistryKey=$RegistryKey -Replace("HKEY_CURRENT_USER","HKCU:")  
        $RegistryKey=$RegistryKey -Replace("HKEY_USERS","HKU:")  
        $RegistryKey
    }

    function GetRegistryType {
    $GetRegistryTypeCode=
@"
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
public class Registry
{
        [DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA")]
        public static extern int RegQueryValueEx(int hKey, string lpValueName, int lpReserved, out int lpType, StringBuilder lpData, ref int lpcbData);
        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int RegOpenKeyEx(Microsoft.Win32.RegistryHive hKey, string subKey, int ulOptions, int samDesired, out int phkResult);
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int RegCloseKey(int hKey);
enum RegistryRights
        {
            ReadKey = 131097,
            WriteKey = 131078
        }
    static Microsoft.Win32.RegistryHive RegHive(string szRegHive)
        {
            Microsoft.Win32.RegistryHive rhRegHiveOut;
            
            switch (szRegHive)
            {
                case "HKLM":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.LocalMachine;
                    break;
                case "HKCR":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.ClassesRoot ;
                    break;
                case "HKCU":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.CurrentUser;
                    break;
                case "HKU":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.Users;
                    break;
                default:
                    rhRegHiveOut = Microsoft.Win32.RegistryHive .LocalMachine;
                    break;
            }
            return rhRegHiveOut;
        }
        static string RegValueType(int iRegType)
        {
            string szRegistryType = "REG_NONE";
            switch (iRegType)
            {
                case 0:
                    szRegistryType = "REG_NONE";
                    break;
                case 1:
                    szRegistryType = "String";
                    break;
                case 2:
                    szRegistryType = "ExpandString";
                    break;
                case 3:
                    szRegistryType = "Binary";//Binary
                    break;
                case 4:
                    szRegistryType = "DWord";
                    break;
                case 5:
                    szRegistryType = "DWord"; //"REG_DWORD_BIG_ENDIAN";
                    break;
                case 6:
                    szRegistryType = "REG_LINK";
                    break;
                case 7:
                    szRegistryType = "MultiString";//multistring
                    break;
            }
            
            return szRegistryType;
        }
public static string GetType(string szHive,string szRegRoot, string szRegValue,int iWow)
{

            //iwow 64bit =256 and 32 = 512
            int hKeyVal = 0, type = 0, hwndOpenKey=0;
            string szType = "NA", szValue="NA" ;
            int valueRet = RegOpenKeyEx(RegHive(szHive), @szRegRoot, 0, (int)RegistryRights.ReadKey | iWow, out hKeyVal);
            if (valueRet == 0)
            {
                int iBuffSize = 10,iError=0;
                bool blSize = false;
                while (!blSize)
                {
                    StringBuilder sb = new StringBuilder(iBuffSize);
                    hwndOpenKey = iBuffSize;
                    try
                    {
                        iError = RegQueryValueEx(hKeyVal, szRegValue, 0, out type, sb, ref hwndOpenKey);
                    }
                    catch { }
                    if (iError == 234)
                    {
                        iBuffSize = iBuffSize + 1;
                    }
                    else if (iError == 0)
                    {
                        blSize = true;
                        szValue = sb.ToString();
                        szType = RegValueType(type);
                    }
                    else
                    {
                        blSize = true;
                    }
                    sb.Remove(0, sb.Length);
                }
            }
            return szType;
            //RegCloseKey(hKeyVal);
}
}
"@
                
        if ([Registry] -eq $null)
        {
        #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
        } else {
        #type is loaded, do nothing
        }

        Trap [Exception]{
            #type doesn't exist.  Load it
            $type = Add-Type -TypeDefinition $GetRegistryTypeCode -PassThru
            continue
        }
    }

    Function CheckRegistryValueExists {
        param(
            [string]$RegRoot,
            [string]$RegValue
        )
        Get-ItemProperty -path $RegRoot -name $RegValue -ErrorAction SilentlyContinue | Out-Null  
        
        if (!$?){
            #check64bit
            $RegRoot=$RegRoot.tolower()
            $RegValue=$RegValue.tolower()
            $Regroot=$Regroot.Replace("software\","software\wow6432node\")
            $Regroot=$Regroot.Replace("hkcr:\","hkcr:\wow6432node\")
            Get-ItemProperty $RegRoot $RegValue -ErrorAction SilentlyContinue | Out-Null  
        }
        $?
    }

    function ProductListing {
$ProductListing=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
public class MakeStringTest
        {
            [DllImport("msi.dll")]
            public static extern Int32 MsiEnumProducts(int iProductIndex, StringBuilder lpProductBuf);
            [DllImport("msi.dll")]
            public static extern Int32 MsiGetProductInfo(string szProduct, string szProperty, StringBuilder lpValueBuf, ref int pcchValueBuf);
            public static ArrayList AllProductCodes = new ArrayList();
            public static ArrayList HashArrary = new ArrayList();
            public static ArrayList loop()
            {
                AllProductCodes = MsiEnumProducts();
                HashArrary = MsiGetProductInfoExtended(AllProductCodes);
                return (HashArrary);
            }
            public static string GetMSIProductInformation(string szProductCode, string szInformationRequested)
            {
                string szReturnValue = "";
                StringBuilder sb = new StringBuilder(1024);
                int iret = 1024, iErrorReturn = 0;
                iErrorReturn = MsiGetProductInfo(szProductCode, szInformationRequested, sb, ref iret);
                if (iErrorReturn == 0)
                {
                    szReturnValue = sb.ToString();
                }
                else
                {
                    szReturnValue = "Unknown";
                }
                return szReturnValue;
            }
            public static ArrayList MsiGetProductInfoExtended(ArrayList alAllProductCodes)
            {
                ArrayList alProdList = new ArrayList();
                string szProductName = "";
                foreach (string szCheck in alAllProductCodes)
                {
                    szProductName = GetMSIProductInformation(szCheck, "ProductName");
                    if (szProductName == "")
                    {
                        szProductName = "Name not available";
                    }
                    Hashtable Hash = new Hashtable();
                    Hash.Add("Name", szProductName);
                    Hash.Add("Value", szCheck);
                    Hash.Add("Description", szCheck);
                    Hash.Add("ExtensionPoint", "<Default/><Icon>@resource.dll,-104</Icon>");
                    alProdList.Add(Hash);
                }
                return alProdList;
            }
            public static ArrayList MsiEnumProducts()
            {
                int iErrorReturn = 0, iProductIndex = 0;
                StringBuilder szProdCode = new StringBuilder(39);
                string szProductCode = "";
                ArrayList alProdCodes = new ArrayList();
                while (iErrorReturn != 259)
                {
                    iErrorReturn = MsiEnumProducts(iProductIndex, szProdCode);
                    if (iErrorReturn == 0)
                    {
                        szProductCode = szProdCode.ToString();
                        alProdCodes.Add(szProductCode);
                    }

                    iProductIndex++;
                }
                return (alProdCodes);
            }
        }
"@
                
        if ([MakeString] -eq $null){
            #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
        } else {
            #type is loaded, do nothing
        }

        Trap [Exception] {
            #type doesn't exist.  Load it
            $type = Add-Type -TypeDefinition $ProductListing -PassThru
            continue
        }
    }

    function BackUpRegistry {
        Param(
            $Item,
            $DeleteParent
        )

        switch ($Item.substring(0,2)) { 
            "-1"{
                $QueryAssignmenttype=[MakeString]::GetProductInfo($ProductCode,"AssignmentType") # cant seem to find where this is defined 
                switch($QueryAssignmenttype){
                    "0"     { $Item=$Item.Replace("21:","HKCU:")}
                    "1"     { $Item=$Item.Replace("02:","HKLM:")}
                    default{ $Item=$Item.Replace("02:","HKLM:")}
                }
            }    
            "00"{ $Item=$Item.Replace("00:","HKCR:")} 
            "01"{ $Item=$Item.Replace("01:","HKCU:")}
            "02"{ $Item=$Item.Replace("02:","HKLM:")} 
            "03"{ $Item=$Item.Replace("03:","HKU:")}
            "20"{ $Item=$Item.Replace("20:","HKCR:")} 
            "21"{ $Item=$Item.Replace("21:","HKCU:")}
            "22"{ $Item=$Item.Replace("22:","HKLM:")} 
            "23"{ $Item=$Item.Replace("23:","HKU:")}
        }      
                
        $RegRoot = $Item.substring(0,$Item.LastIndexOf("\")) #Get Registry root value
        $RegValue =$Item.substring($Item.LastIndexOf("\")+1,$Item.length-$Item.LastIndexof("\")-1) #Get registry value 
        $Exista=$false
        $Existsa=CheckRegistryValueExists -RegRoot $Regroot -RegValue $RegValue #Does this key still exist in the registry? MSI thinks it does
        
        if (!$Existsa) { #Maybe a path on the reg info so search value for known good
            $InLoop=$true
            $iCounter=0
            $Slash =$Item.IndexOf("\",$iCounter)+1

            while($InLoop){
            $TempRoot=$Item.substring(0,$Slash)
            $TempValue=$Item.substring($Slash,($Item.length-$Slash))
                $Existsa=CheckRegistryValueExists  -RegRoot $TempRoot  -RegValue $TempValue
                if (!$Existsa){
                    $iCounter=$Slash
                    $Slash =$Item.IndexOf("\",$iCounter)+1
                    If ($Slash -eq 0 -or $TempRoot -eq "") { #$Slash =-1 before DEBUG
                        $InLoop=$False
                    }         
                } else {
            #FOUND IT return here
                    $InLoop=$False
                    $RegRoot=$TempRoot
                    $RegValue=$TempValue
                    $RegistryType=[Registry]::GetType($item.substring(0,($item.IndexOf(":"))),$RegRoot.substring($RegRoot.indexOf(":")+2),$RegValue,256)
                }
            }
        }

        if ($Existsa){ #If we have found the registry key lets continue
            $RegistryType="NA"
            $RegistryType= [Registry]::GetType($item.substring(0,($item.IndexOf(":"))),$RegRoot.substring($RegRoot.indexOf(":")+2),$RegValue,256)  #Wow32 to false for registry redirection
            
            If($RegistryType -ne "NA"){
                WriteXMLRegistry $RegistryType $RegRoot $RegValue $DeleteParent
            }

            $RegistryType="NA"
            $regroot=$regroot.tolower()
            $regroot=$regroot.replace("software\","software\wow6432node\")
            $regroot=$regroot.replace("HKCR:","HKCR:\WOW6432node\")
            $regroot=$regroot.replace("hkcr:","HKCR:\WOW6432node\")
            $RegistryType = [Registry]::GetType($item.substring(0,($item.IndexOf(":"))),$RegRoot.substring($RegRoot.indexOf(":")+2),$RegValue,256) #Wow32 to false for registry redirection
            If ($RegistryType -ne "NA") {
                WriteXMLRegistry $RegistryType $RegRoot $RegValue $DeleteParent
            }
        }
        $RegistryCountFailures
    }

    function WriteXMLRegistry
    {
        param(
            $RegistryType,
            $RegRoot,
            $RegValue,
            $DeleteParent
        )

        $root=$Env:SystemDrive
        $DirectoryPath= $root+"\MATS\$ProductCode"
        if (!($RegContent=Get-ItemProperty $RegRoot $RegValue)){
            #If the registry key is missing or has issues fail out
            return $True
        }
            
        if ($Regcontent -eq ""){
            $Failure=$true
        }

        $OutcomeContent=""
        [string]$RegName=""
            
        switch ($RegistryType){
            "String" {
                $RegName=$RegContent.$RegValue
            }
                
            "MultiString" {
                foreach ($Value in $RegContent.$RegValue) {
                    $RegName=$RegName+$value+";"
                }

                $RegName=$RegName.substring(0,$RegName.length-1)
            }
            
            "Binary" {
                foreach ($Value in $RegContent.$RegValue){
                    $RegName=$RegName+$value+","
                }
                $RegName=$RegName.substring(0,$RegName.length-1)
            }
                
            default{
                Foreach ($Value in $RegContent.$RegValue){
                    $Outcome= "{0:x}" -f $Value
                    $OutcomeContent=$OutcomeContent+$Outcome
                }
                $RegContent.$RegValue=$OutcomeContent
                $RegName=$RegContent.$RegValue
            }
        }
            
        $xml = New-Object xml
        If (!(Test-Path -path $DirectoryPath\registryBackupTemplate.xml)) { 
            $root = $xml.CreateElement("RegistryBackup")
            [void]$xml.AppendChild($root)
            $Record = $xml.CreateElement("Registry")
            #$Record.PSBase.InnerText = $RegRoot
            [void]$root.AppendChild($Record)
            $RegistryHive= $xml.CreateElement("RegistryHive")
            $RegistryHive.PSBase.InnerText = ($RegRoot)
            [void]$Record.AppendChild($RegistryHive)
            $XMLDeleteParent = $xml.CreateElement("RegistryDeleteParent")
            $XMLDeleteParent.PSBase.InnerText = [string]$DeleteParent
            [void]$Record.AppendChild($XMLDeleteParent)
            $Name = $xml.CreateElement("RegistryType")
            $Name.PSBase.InnerText = $RegistryType
            [void]$Record.AppendChild($name)
            $Value = $xml.CreateElement("RegistryName")
            $Value.PSBase.InnerText = ($RegValue)
            [void]$Record.AppendChild($Value)
            $Date = $xml.CreateElement("RegistryValue")
            $Date.PSBase.InnerText = ($RegName)
            [void]$Record.AppendChild($date)
            if (!(Test-Path -path $DirectoryPath)) { 
                [void]( New-Item  "$DirectoryPath" -type directory -force )#Create the backup folder for the files we will delete
            }

            [void]$xml.Save("$DirectoryPath\registryBackupTemplate.xml")
        } else {
            $xml =[xml] (get-content "$DirectoryPath\registryBackupTemplate.xml")
            $root = $xml.RegistryBackup
            $Record = $xml.CreateElement("Registry")
            [void]$root.AppendChild($Record)
            $RegistryHive= $xml.CreateElement("RegistryHive")
            $RegistryHive.PSBase.InnerText = ($RegRoot)
            [void]$Record.AppendChild($RegistryHive)
            $XMLDeleteParent = $xml.CreateElement("RegistryDeleteParent")
            $XMLDeleteParent.PSBase.InnerText = [string]$DeleteParent
            [void]$Record.AppendChild($XMLDeleteParent)
            $Name = $xml.CreateElement("RegistryType")
            $Name.PSBase.InnerText = $RegistryType
            [void]$Record.AppendChild($name)
            $Value = $xml.CreateElement("RegistryName")
            $Value.PSBase.InnerText = ($RegValue)
            [void]$Record.AppendChild($Value)
            $Date = $xml.CreateElement("RegistryValue")
            $Date.PSBase.InnerText = ($RegName)
            [void]$Record.AppendChild($date)
            [void]$xml.Save("$DirectoryPath\registryBackupTemplate.xml")
        }       
    }

    function RapidProductRegistryRemove {
        Param(
            $ProductCode
        )
        CreateRegistryFileRecoveryFile($ProductCode)
        New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT # Allows access to registry HKEY_CLASSES_ROOT
        
        $CompressedGUID=[CleanupRegistry]::CompressGUID([string]$ProductCode)

        if ($CompressedGUID) {
            ##################HKLM
            $RPRHives=Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Userdata\ #IF need to delete component then add back recurse HKL,HKCR code here
            foreach($SubKey in $RPRHives) { 
                $RegistryKey=RegistryHiveReplace $SubKey.name
        
                if (Test-Path "$RegistryKey\Products\$CompressedGUID") {        
                    New-ItemProperty "$RegistryKey\Products\$CompressedGUID" -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null  
                    BackupRegistry ("$RegistryKey\Products\$CompressedGUID"+"\RPRTAG") $False
                    $RegistryBackup=dir "$RegistryKey\Products\$CompressedGUID" -recurse
                    if ($RegistryBackup) {
                        foreach ($a in $RegistryBackup) {
                            foreach($Property in $a.Property) {
                                $a=RegistryHiveReplace $a
                                BackupRegistry ($a+"\"+$Property) $False
                            }
                        }
                    }
                
                    del "$RegistryKey\Products\$CompressedGUID" -recurse #Special delete for parent
                }
            }

            ############HKCR
            if (Test-Path "HKCR:\Installer\Products\$CompressedGUID" ) {
                $RegistryBackup=dir "HKCR:\Installer\Products\$CompressedGUID" -recurse
                if ($RegistryBackup) {
                        $ParentValues= Get-ItemProperty -Path "HKCR:\Installer\Products\$CompressedGUID" 
                        #[string]$rr=get-itemproperty $Parentvalues
                        [string]$temp=[string]$ParentValues
                        $temp=$temp.Replace("@{","")
                        $temp=$temp.Replace("}","")
                        foreach ($ParentItem in $temp.Split(";")){
                            #FN to backup the root elements
                            $BackupSplit= $ParentItem.substring(0,$ParentItem.indexof("=")).trim()
                            BackupRegistry("HKCR:\Installer\Products\$CompressedGUID\$BackupSplit") $False
                        }
            
                        foreach ($a in $RegistryBackup){
                            if ($a.Property -ne ""){
                                foreach($Property in $a.Property){
                                        $a=RegistryHiveReplace $a
                                        BackupRegistry($a+"\"+$Property) $False
                                }
                            } else {
                                New-ItemProperty (RegistryHiveReplace $a.name) -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null 
                                BackupRegistry((RegistryHiveReplace $a.name) +"\RPRTAG") $False
                            }
                        }
                        del "HKCR:\Installer\Products\$CompressedGUID" -recurse #Special delete for parent
                }
            }

            ############HKCU
            if (Test-Path "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID" ){
                $RegistryBackup=dir "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID" -recurse
                if ($RegistryBackup){
                    $ParentValues= Get-ItemProperty -Path "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID"
                    
                    [string]$SplitString=[string]$ParentValues
                    $SplitString=$SplitString.Replace("@{","")
                    $SplitString=$SplitString.Replace("}","")
                    
                    foreach ($ParentItem in $SplitString.Split(";")) {
                        #FN to backup the root elements
                        $BackupSplit= $ParentItem.substring(0,$ParentItem.indexof("=")).trim()
                        BackupRegistry("HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID\$BackupSplit") $False
                    }
        
                    foreach ($a in $RegistryBackup){
                        if ($a.Property -ne ""){
                            foreach($Property in $a.Property){
                                    $a=RegistryHiveReplace $a
                                    BackupRegistry($a+"\"+$Property) $False
                            }
                        } else {
                            New-ItemProperty (RegistryHiveReplace $a.name) -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null 
                            BackupRegistry((RegistryHiveReplace $a.name) +"\RPRTAG") $False
                        }
                    }
                    del "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID" -recurse #Special delete for parent
                }
            }
        }
    }

    #.NET Functions
    #========================
    Compressed      | Out-Null
    ProductListing  | Out-Null
    GetRegistryType | Out-Null

    $ErrorActionPreference = "Continue"

    $registeredInstallPaths = if([Environment]::Is64BitOperatingSystem){
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    } else {
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    [array]$productToScrub = Get-ChildItem $registeredInstallPaths |
        Where-Object { $_.PsPath -like "*{*-*-*-*}"}  | # fixes errors if host has unexpected (non-guid) uninstall entries in registry 
        Foreach-Object { Get-ItemProperty $_.PSPath } |
        Select-Object *,@{Name='ProductCode';E={$_.PSChildName}} | # add ProductCode (which is parent reg folder name) as property to the objs
        Where-Object -FilterScript $Filter

    if($productToScrub.Count -eq 0){
        Write-Warning "No products matching filter: $Filter"
        return
    }

    $DateTimeRun="{0:yyyy.MM.dd.HH.mm.ss}" -f (get-date) #Run Date\Time

    foreach($product in $productToScrub){
        MATSFingerPrint $product.ProductCode "Version_RS_RapidProductRemoval" "1.3" $DateTimeRun
        MATSFingerPrint $product.ProductCode "ProductName" $product.DisplayName $DateTimeRun
        $root=$Env:SystemDrive
        $DirectoryPath= $root+"\MATS\$($product.ProductCode)"
        
        "Scrubbing Product Code $($product.ProductCode)"
        
        $value=[UninstallProduct2]::LogandUninstallProduct($product.ProductCode) # try Uninstall with the Windows Installer -x switch first
        MATSFingerPrint $product.ProductCode "msiexec -x" $value $DateTimeRun

        RapidProductRegistryRemove $product.ProductCode
        MATSFingerPrint $product.ProductCode "RPR" $true $DateTimeRun
        Get-Item $($product.PSPath) | Remove-Item -Recurse -Force -Confirm:$false
    }
}

<#
.SYNOPSIS
    Lists MSI-based registered installations from the Windows Installer registry.

.DESCRIPTION
    Queries the standard registry locations for MSI-based products and returns
    their registered properties along with the ProductCode (GUID).

    This is a helper function intended to let you explore what products are
    registered before using Remove-InstallerRegistration to scrub them.

.PARAMETER Filter
    A scriptblock filter used to target results. This filter is evaluated
    against each registry entry object.

    Example:
        { $_.DisplayName -like "*SQL*" }
        { $_.Publisher -eq "Microsoft Corporation" }

.EXAMPLE
    PS> Get-InstallerRegistration

    Returns all MSI-registered products.

.EXAMPLE
    PS> Get-InstallerRegistration -Filter { $_.DisplayName -like "Azure Connected Machine Agent*" }

    Lists all registered Azure Connected Machine Agent installs.

.OUTPUTS
    PSCustomObject with properties from the registry entry, plus:
    - ProductCode (GUID from registry key name)

.NOTES
    Author: Joey Eckelbarger
#>
function Get-InstallerRegistration {
    $registeredInstallPaths = if([Environment]::Is64BitOperatingSystem){
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    } else {
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    [array]$allRegisteredProducts = Get-ChildItem $registeredInstallPaths |
        Where-Object { $_.PsPath -like "*{*-*-*-*}"}  | # fixes errors if host has unexpected (non-guid) uninstall entries in registry 
        Foreach-Object { Get-ItemProperty $_.PSPath } |
        Select-Object *,@{Name='ProductCode';E={$_.PSChildName}} 

    if($Filter){
        $matchingFilter = $allRegisteredProducts | Where-Object -FilterScript $Filter
        if($matchingFilter.Count -eq 0){
            Write-Warning "No products matching filter: $Filter"
            return
        } else {
            return $matchingFilter
        }
    } else {
        return $allRegisteredProducts
    }
}