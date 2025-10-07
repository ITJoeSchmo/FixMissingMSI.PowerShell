#region Microsoft Derived Functions from MicrosoftProgram_Install_and_Uninstall.meta.diagcab\MSIMATSFN.ps1

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

#region funcs needed for both RapidProductRegistryRemove and LPR 
# MATSFingerPrint
# BackupRegistry
# -> CheckRegistryValueExists
# -> GetRegistryType
# -> ProductListing
# -> WriteXMLRegistry
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
    $Existsa=$false
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

#region funcs needed for BackUpRegistry
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

function WriteXMLRegistry {
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
    
#endregion

#endregion 

#region funcs for RapidProductRegistryRemove process
# UninstallProduct
# CreateRegistryFileRecoveryFile
# Compressed
# RegistryHiveReplace
# RapidProductRegistryRemove
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
#endregion

#region funcs for LPR process

function ProductLPR {
    $ProductLPR=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
public class ProductLPRMain
{
[DllImport("msi.dll")]
public static extern Int32 MsiEnumComponents(
                int iComponentIndex,
                StringBuilder lpComponentBuf
                );
[DllImport("msi.dll")]
public static extern int MsiGetComponentPath(
    string szProduct,   // product code for client product
    string szComponent, // component ID
    StringBuilder lpPathBuf,    // returned path
    ref int pcchBuf       // buffer character count
    );
[DllImport("msi.dll")]
public static extern int MsiEnumClients(
    string szComponent, // component code, string GUID
    int iProductIndex, // 0-based index into client products
    StringBuilder lpProductBuf  // buffer to receive GUID
    );
    
    public static ArrayList GetComponentList()
        {
            StringBuilder sb = new StringBuilder(40);
            ArrayList alComponents = new ArrayList();
            int iError = 0,iCounter=1;
            iError= MsiEnumComponents(0,sb);
            while (iError == 0)
            {
                iError = MsiEnumComponents(iCounter , sb);
                if (iError == 0)
                {
                    alComponents.Add(sb.ToString());
                    iCounter = iCounter + 1;
                }
            }
            return alComponents;
        }
        public static bool SharedComponents(string szComponent,string szProductCode)
        {
            
            bool blShared=false;
            StringBuilder sb = new StringBuilder(1024);
            int iError = 0,iCounter=0;
            while (blShared != true && iError !=259)
            {
                sb.Remove(0, sb.Length);
                iError = MsiEnumClients(szComponent, iCounter, sb);
                if (iError == 0  && (szProductCode !=sb.ToString ()))
                {
                    blShared = true;                    
                }
                iCounter = iCounter + 1;
            }
            return blShared;
        }
        public static ArrayList GetComponentPath(string szProductCode)
        {           
            int pathBuf = 1024,iError=0;
            StringBuilder sbPath = new StringBuilder(1024);
            ArrayList alItems = new ArrayList();
            ArrayList alComponents= new ArrayList();
            alComponents=GetComponentList();
            
            foreach (string szComponent in alComponents)
            {
                pathBuf = 1024;
                iError=  MsiGetComponentPath(szProductCode, szComponent, sbPath, ref pathBuf);
                string szCompPath = sbPath.ToString().ToLower ();
                if (szCompPath=="")szCompPath ="123NA123";
                
                //if ((iError != -1 && szCompPath != "123NA123") && (szCompPath.IndexOf("\\", szCompPath.Length - 1, 1)==-1) && !(szCompPath.Contains(Environment.GetEnvironmentVariable("WinDir").ToLower ())))
        if ((iError != -1 && szCompPath != "123NA123") && (szCompPath.IndexOf("\\", szCompPath.Length - 1, 1)==-1))
                {
                    if ( !(SharedComponents (szComponent,szProductCode)))
                    {
        alItems.Add(sbPath.ToString());
                    }                    
                }  
        }
            return alItems;
        }
    }
"@    
    if ([ProductLPRMain] -eq $null){
        #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
    } else {
        #type is loaded, do nothing
    }

    Trap [Exception] {
        #type doesn't exist.  Load it
        $type = Add-Type -TypeDefinition $ProductLPR -PassThru
        continue
    }
}
Function ShimXML{
    param ($ProductCode)
    # Change by Joey Eckelbarger -- embedded shim.xml 
    $shimXML = @"
<Shim>
  <Product>
    <ProductCode>{CE2CDD62-0124-36CA-84D3-9F4DCF5C5BD9}</ProductCode>
    <ProductName>Microsoft .NET Framework 3.5 SP1</ProductName>
    <OS Build="5.1.2600">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft .NET Framework 3.5 SP1</RegistryValue>
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727</RegistryValue>
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
      </Language>
    </OS>
    <OS Build="6.0.6001">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft .NET Framework 3.5 SP1</RegistryValue>
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
      </Language>
    </OS>
    <OS Build="6.0.6002">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft .NET Framework 3.5 SP1</RegistryValue>
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
      </Language>
    </OS>
    <OS Build="6.0.6000">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft .NET Framework 3.5 SP1</RegistryValue>
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
      </Language>
    </OS>
    <ProductCode>{774088D4-0777-4D78-904D-E435B318F5D2}</ProductCode>
    <ProductName>Microsoft Antimalware</ProductName>
    <OS Build="ALL">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft Antimalware</RegistryValue>
            <RegistryValue>HKEY_CLASSES_ROOT\Installer\UpgradeCodes\1F69ACF0D1CF2B7418F292F0E05EC20B</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
        <Services>
          <Service>Microsoft Antimalware Service</Service>
        </Services>
      </Language>
    </OS>
    <ProductCode>{50779A29-834E-4E36-BBEB-B7CABC67A825}</ProductCode>
    <ProductName>Microsoft Security Client JA-JP Language Pack</ProductName>
    <OS Build="ALL">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_CLASSES_ROOT\Installer\UpgradeCodes\18EAFAA8B504E9C4781184E19FC7A0F8</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
        <Services>
          <Service>Microsoft Antimalware Service</Service>
        </Services>
      </Language>
    </OS>
    <ProductCode>{05BFB060-4F22-4710-B0A2-2801A1B606C5}</ProductCode>
    <ProductName>Microsoft Antimalware</ProductName>
    <OS Build="ALL">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_CLASSES_ROOT\Installer\UpgradeCodes\1F69ACF0D1CF2B7418F292F0E05EC20B</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
        <Services>
          <Service>Microsoft Antimalware Service</Service>
        </Services>
      </Language>
    </OS>
    <ProductCode>{3B1E1F4C-031D-410F-A93A-1220236608C8}</ProductCode>
    <ProductName>Microsoft Antimalware Service JA-JP Language Pack</ProductName>
    <OS Build="ALL">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_CLASSES_ROOT\Installer\UpgradeCodes\307BC1C5473CE634DB482731070ADD5C</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
        <Services>
          <Service>Microsoft Antimalware Service</Service>
        </Services>
      </Language>
    </OS>
    <ProductCode>{54B6DC7D-8C5B-4DFB-BC15-C010A3326B2B}</ProductCode>
    <ProductName>Microsoft Security Client</ProductName>
    <OS Build="ALL">
      <Language Code="ALL">
        <Registry>
          <RegistryHive Action="*">
            <RegistryValue>HKEY_CLASSES_ROOT\Installer\UpgradeCodes\11BB99F8B7FD53D4398442FBBAEF050F</RegistryValue>
          </RegistryHive>
          <RegistryHive Action="-">
            <RegistryValue></RegistryValue>
          </RegistryHive>
        </Registry>
        <Services>
          <Service>Microsoft Antimalware Service</Service>
        </Services>
      </Language>
    </OS>

  </Product>

</Shim>
"@

    $Win32_OS=Get-WmiObject Win32_OperatingSystem | select BuildNumber,OSLanguage,version
    $Build=$Win32_OS.Version
    $Language=$Win32_OS.OSLanguage
    # Valid directory/file tokens are the following:
    # %5% - Root drive
    # %10% - Windows directory; example – c:\windows
    # %11% - system directory; example – c:\windows\system32
    # %16422% - Program Files directory – c:\Program Files
    # %16427% - Common Files directory – c:\Program Files\Common Files
    # %16527% - USERPROFILE C:\Documents and Settings\someone
    # %16627% - APPDATA     C:\Documents and Settings\kmyer\Application Data
    $ShimItems= New-Object System.Collections.ArrayList($null)
        
    $xml = [xml]$shimXML
    $XMLProductCode=$xml.shim.Product | Where-Object {$_.ProductCode -match $ProductCode} 
    if ($XMLProductCode -ne $null) {
        $XMLBuild= $XMLProductCode.os | Where-Object {($_.Build -match $Build) -or ($_.Build -match "ALL")}
        if ($XMLBuild -ne $null){
            Foreach ($Build in $XMLBuild) {        
                $XMLLanguage=$Build.Language |Where-Object {($_.Code -match $Language) -or ($_.Code -match "ALL") }
                if ($XMLLanguage -ne $null) { 
                    Foreach ($Language in $XMLLanguage) { 
                        foreach ($XMLRegistryKey in $Language.Registry) { #Registry shims
                            if ($XMLRegistryKey-ne $null) {
                                $RegistryContainer=$XMLRegistryKey.RegistryHive
                                if ($RegistryContainer.length  -gt 0) {
                                    foreach ($Action in $RegistryContainer) {  #Registry shims 
                                        switch ($Action.Action){
                                            "*"{
                                                $EnumShim=ShimRegistryCollection $Action.registryvalue
                                                if ($EnumShim.length -gt 0) {
                                                    $ShimItems=$ShimItems+$EnumShim
                                                }
                                                break;
                                            }
                                            "-" {
                                                $ShimItems+=((RegistryHiveReplaceShim($Action.registryvalue))) #Registry value Additions
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        foreach ($XMLFileKey in $Language.File) { #File Shims
                            if (($XMLFileKey.FileName).length -gt 0){       
                                $FilePathType = $XMLFileKey.FileName.substring($XMLFileKey.FileName.indexof("%"),$XMLFileKey.FileName.lastindexof("%")+1)
                                
                                switch ($FilePathType) {
                                    "%5%" {
                                        $FilePath=$XMLFileKey.FileName
                                        $root=$Env:SystemDrive
                                        $FilePath=$FilePath.replace("%5%",$root)
                                        break;
                                    }
                                    "%10%" {
                                        $FilePath=$XMLFileKey.FileName
                                        $root=$Env:WinDir
                                        $FilePath=$FilePath.replace("%10%",$root)
                                        break;
                                    }
                                    "%16627%" {
                                        $FilePath=$XMLFileKey.FileName
                                        $root=$Env:AppData
                                        $FilePath=$FilePath.replace("%16627%",$root)
                                        break;
                                    }
                                }
                                $ShimItems+=($FilePath) #File Additions
                            }                
                        }
            
                        foreach ($XMLService in $Language.Services) { #Services
                            if (($XMLService.Service).length -gt 0) {
                                if ($XMLService.Service -ne $null){ #Stop Services
                                    stop-service $XMLService.Service -force -ErrorAction SilentlyContinue
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    $ShimItems
}

Function ShimRegistryCollection {
    param($Registrykey)
    
    New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT |out-null # Allows access to registry HKEY_CLASSES_ROOT
    $ShimArray=@()
    foreach ($RegistryRoot in $RegistryKey){
        $RegistryRoot=RegistryHiveReplace($RegistryRoot)
        If (ShimLoop ($RegistryRoot)) {
            $ShimArray= $ShimArray+(ShimLoop ($RegistryRoot))
        }

        #IS WOW ALSO
        $RegistryRoot=$RegistryRoot.tolower()
        $RegistryRoot=$RegistryRoot.replace("software\","software\wow6432node\")
        $RegistryRoot=$RegistryRoot.replace("HKCR:","HKCR:\WOW6432node\")
        $RegistryRoot=$RegistryRoot.replace("hkcr:","HKCR:\WOW6432node\")
        
        $RegistryRoot=RegistryHiveReplace($RegistryRoot)
        If (ShimLoop ($RegistryRoot)){
            $ShimArray= $ShimArray+(ShimLoop ($RegistryRoot))
        }
    }
    return $ShimArray 
}
Function ShimLoop {
    Param($RegistryRoot)
    $Itemstoadd= @()
    if((Test-Path $RegistryRoot)){
        $RegistryRoot=RegistryHiveReplace $RegistryRoot

        [string]$ParentValues= get-itemproperty -path $RegistryRoot
        [string]$SplitString=[string]$ParentValues
        $SplitString=$SplitString.Replace("@{","")
        $SplitString=$SplitString.Replace("}","")
        
        foreach ($ParentItem in $SplitString.Split(";")){
            if ($ParentItem.length -gt 0){
                $Itemstoadd+=((RegistryHiveReplaceShim($RegistryRoot))+"\"+($ParentItem.substring(0,$ParentItem.indexof("=")).trim()))
            }
        }
                    
        $RegistryPatchList = dir $RegistryRoot -recurse
        if ($RegistryPatchList){
            foreach ($Patch in $RegistryPatchList){
                foreach ($Property in $Patch.property){
                    if ($Property.length  -gt 0){
                            $Itemstoadd+=((RegistryHiveReplaceShim($Patch.name))+"\"+ $Property)
                    }
                }
            }
        }
        
        New-ItemProperty "$RegistryRoot" -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null  
        BackupRegistry("$RegistryRoot"+"\RPRTAG") $True
    }
    $Itemstoadd
}

Function RegistryHiveReplaceShim {
    Param ([string]$RegistryKey)
    $RegistryKey=$RegistryKey.toupper()
    $RegistryKey=$RegistryKey.Replace("HKEY_CLASSES_ROOT","00:")
    $RegistryKey=$RegistryKey.Replace("HKEY_CURRENT_USER","01:")
    $RegistryKey=$RegistryKey.Replace("HKEY_LOCAL_MACHINE","02:")
    $RegistryKey=$RegistryKey.Replace("HKEY_USERS","03:")
    $RegistryKey=$RegistryKey.Replace("HKCR:","00:")
    $RegistryKey=$RegistryKey.Replace("HKCU:","01:")
    $RegistryKey=$RegistryKey.Replace("HKLM:","02:")
    $RegistryKey=$RegistryKey.Replace("HKU:","03:")

    $RegistryKey  
}

Function ARPEntries {
    param($ProductCode,$ProductName)

    $Itemstoadd= @()
    $ARPUninstall="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

    if($ProductName.length -gt 0){
        $ARPRegistryKeyName=$ARPUninstall+$ProductName
        if (!(Test-Path -path $ARPRegistryKeyName)){
            $ReturnedARP=(ARPFindDisplayName $ARPUninstall $ProductName)

            if ($ReturnedARP){
                $Itemstoadd+=($ReturnedARP)
            }
        } else {
            $Itemstoadd+=(ShimRegistryCollection($ARPRegistryKeyName))
        }
            
        if (Test-Path -path ($ARPUninstall+$ProductCode)){
            $Itemstoadd+=(ShimRegistryCollection($ARPUninstall+$ProductCode))
        }
    }
    $Itemstoadd
}
Function ARPFindDisplayName {
    param ($ARPUninstall, $ProductName)
    #Enum for DisplayName
    $ARPUninstallKey = dir $ARPUninstall -recurse
       
    if ($ARPUninstallKey){
        foreach ($ARPEntry in $ARPUninstallKey){
            foreach ($Property in $ARPEntry.property){
                if ($Property -eq "DisplayName"){
                    $ARPDisplayName =(Get-ItemProperty -Path (RegistryHiveReplace($ARPEntry).Name))
                    if ($ARPDisplayName.DisplayName -eq $ProductName){
                        return ShimRegistryCollection($ARPEntry.Name)
                    } 
                }
            }
        }
    }
    $false
}

Function BackupFiles {
    Param ($Items,$ProductCode)

    $root=$Env:SystemDrive
    $DirectoryPath= $root+"\MATS\$ProductCode"
    [void](New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT ) # Allows access to registry HKEY_CLASSES_ROOT
    if (!(Test-Path -path $DirectoryPath)) { 
        [void]( New-Item  "$DirectoryPath\FileBackup" -type directory -force )#Create the backup folder for the files we will delete
    }

    foreach ($Item in $Items){
        [void]$Item
        if (($Item -ne $Null) -and ($Item.length -gt 2)){
            if(!($Item.substring(0,1)-match '^\d+$')) { #not a registry
                $DuplicateDirector="$DirectoryPath\FileBackup\"+$item.Replace(":","")
                if (!(Test-Path -path $DuplicateDirector.substring(0,($DuplicateDirector.LastIndexOf("\"))))) { 
                    [void](New-Item $DuplicateDirector.substring(0,($DuplicateDirector.LastIndexOf("\"))) -type directory) #Create the backup directories ready to copy the files into then
                }
            
                if(Test-Path $Item){
                    $iError=Copy-Item $Item -Destination ($DuplicateDirector) -ErrorAction SilentlyContinue #copy files to the backup folder
                    WriteXMLFiles ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($Item).InternalName) ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($Item).FileVersion) $Item  $DuplicateDirector
                }
            } else {
                #We are dealing with the registry   
                BackUpRegistry $Item $False
            }
        }
    }
}

Function WriteXMLFiles {
    PARAM (
        $FileNamePath,
        $FileVersion,
        $FileBackupLocation, 
        $FileDestination
    )

    $xml = New-Object xml
    if (!(Test-Path -path "$DirectoryPath\FileBackupTemplate.xml")) { 
        $root = $xml.CreateElement("FileBackup")
        [void]$xml.AppendChild($root)
        $Record = $xml.CreateElement("File")
        #$Record.PSBase.InnerText = $RegRoot
        [void]$root.AppendChild($Record)
        $RecordName = $xml.CreateElement("FileName")
        $RecordName.PSBase.InnerText = ($FileNamePath)
        [void]$Record.AppendChild($RecordName)
        $RecordVersion= $xml.CreateElement("FileVersion")
        $RecordVersion.PSBase.InnerText = ($FileVersion)
        [void]$Record.AppendChild($RecordVersion)
        $RecordType = $xml.CreateElement("FileBackupLocation")
        $RecordType.PSBase.InnerText = $FileBackupLocation
        [void]$Record.AppendChild($RecordType)
        $RecordBackup = $xml.CreateElement("FileDestination")
        $RecordBackup.PSBase.InnerText = $FileDestination
        [void]$Record.AppendChild($RecordBackup)
        [void]$xml.Save("$DirectoryPath\FileBackupTemplate.xml")
    } else {
        $xml =[xml] (get-content "$DirectoryPath\FileBackupTemplate.xml")
        $root = $xml.FileBackup
        $Record = $xml.CreateElement("File")
        [void]$root.AppendChild($Record)
        $RecordName = $xml.CreateElement("FileName")
        $RecordName.PSBase.InnerText = ($FileNamePath)
        [void]$Record.AppendChild($RecordName)
        $RecordVersion= $xml.CreateElement("FileVersion")
        $RecordVersion.PSBase.InnerText = ($FileVersion)
        [void]$Record.AppendChild($RecordVersion)
        $RecordType = $xml.CreateElement("FileBackupLocation")
        $RecordType.PSBase.InnerText = $FileBackupLocation
        [void]$Record.AppendChild($RecordType)
        $RecordBackup = $xml.CreateElement("FileDestination")
        $RecordBackup.PSBase.InnerText = $FileDestination
        [void]$Record.AppendChild($RecordBackup)
        [void]$xml.Save("$DirectoryPath\FileBackupTemplate.xml")
    }       
}

Function DeleteFilesFromXML {
    Param($ProductCode)
    $root=$Env:SystemDrive
    $DirectoryPath= $root+"\MATS\$ProductCode"
    if ((Test-Path -path $DirectoryPath\FileBackupTemplate.xml)) { 
        $xml =[xml] (get-content $DirectoryPath\FileBackupTemplate.xml)
        $root =$xml.FileBackup.File
        foreach ($XMLItems in $xml.FileBackup.File){    
            remove-item $XMLItems.FileBackupLocation -ErrorAction SilentlyContinue
        }
    }   
}
Function DeleteRegistryKeysFromXML{
    Param($ProductCode)
    $root=$Env:SystemDrive
    $DirectoryPath= $root+"\MATS\$ProductCode"
    New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT # Allows access to registry HKEY_CLASSES_ROOT

    if ((Test-Path -path $DirectoryPath\registryBackupTemplate.xml)){ 
        $xml =[xml] (get-content $DirectoryPath\registryBackupTemplate.xml)
        #$root =$xml.RegistryBackup.Registry
        foreach ($XMLItems in $xml.RegistryBackup.Registry){
            if ((($XMLItems.RegistryDeleteParent).toupper()) -eq "TRUE"){
                del $XMLItems.RegistryHive -recurse -ErrorAction SilentlyContinue
            } else {
                remove-itemproperty -path $XMLItems.RegistryHive -name $XMLItems.RegistryName -ErrorAction SilentlyContinue
            }
        }
    } else {
        #Failed to find registry backup so exit
        Return $False
    }
}

function LPR{
    Param(
        $ProductCode,
        $ProductName,
        $DateTimeRun
    )

    MATSFingerPrint $ProductCode "LPR" $true $DateTimeRun

    $alItemsToDelete = New-Object System.Collections.ArrayList($null)
    $alItemsToDelete = [ProductLPRMain]::GetComponentPath($ProductCode) #generate list of installed items
    $alItemsToDelete = $alItemsToDelete+(ShimXML($ProductCode)) #Do we have any shims If so add
    $alItemsToDelete = $alItemsToDelete+(ARPEntries $ProductCode $ProductName ) #Do we have any extra ARP entry

    if ($alItemsToDelete) { #If we found files and registry keys proceed 
        BackupFiles $alItemsToDelete $ProductCode #Attempting to backup all files and registry keys for selected product
        MATSFingerPrint $ProductCode "BackupFiles" $true $DateTimeRun

        DeleteFilesFromXML($ProductCode)
        MATSFingerPrint $ProductCode "DeleteFiles" $True $DateTimeRun
        DeleteRegistryKeysFromXML($ProductCode)
    }
}

#endregion

<#
.SYNOPSIS
    Forcefully removes Windows Installer (MSI) product registrations when normal
    uninstall or repair is not possible.

.DESCRIPTION
    This function scrubs Windows Installer registration for a given product when
    cached MSI/MSP files are missing and neither repair nor uninstall can complete.

    > WARNING: This is a destructive recovery tool. It removes Windows Installer
    > metadata and, when -DeepCleanup is specified, may also remove product files
    > and related registry data. Use it only when standard uninstall/repair paths fail.

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
        { $_.ProductCode -eq "{B9E5C55D-B7B3-4A36-BBFE-18543570333D}"}

.PARAMETER DeepClean
    When specified, performs a full "Locate–Purge–Remove" (LPR) phase in addition
    to the standard Rapid Product Removal (RPR) phase.

    The DeepClean phase attempts to:
      - Enumerate all MSI components and related client registrations.
      - Identify associated product files, registry keys, and ARP entries.
      - Create backups of detected items.
      - Remove residual files, shims, and orphaned registry data.

    Use this switch only if the standard cleanup does not allow reinstallation
    or if remnants continue to block setup operations.

.EXAMPLE
    PS> Remove-InstallerRegistration -Filter { $_.DisplayName -like "Azure Connected Machine Agent*" -and $_.DisplayVersion -eq "1.56.03167" }

    Removes registration data for all MSI products with DisplayName matching
    "Azure Connected Machine Agent*" and DisplayVersion equal to "1.56.03167".

.EXAMPLE
    PS> Remove-InstallerRegistration -Filter { $_.Publisher -eq "Microsoft Corporation" -and $_.DisplayName -like "*SQL*" } -DeepClean

    Performs both registry scrubbing and full residual cleanup for Microsoft SQL Server-related products.

.NOTES
    Author: Joey Eckelbarger
    Derived from Microsoft’s Program Install/Uninstall Troubleshooter files (MSIMATSFN.ps1).

    License Note:
        The embedded helper functions are Microsoft-authored and used under the
        terms of their published diagnostic tool conditions. They are included here
        solely to enable advanced recovery where cached MSI/MSP files cannot be sourced
        and attempts to repair the installation via MSI fail. 

    Limitations:
        - Without -DeepCleanup, removes registry/metadata only (RPR phase).
        - With -DeepCleanup, also removes discovered product files, registry keys,
          and other installation remnants (LPR phase) with a backup in C:\MATS\{productcode}
        - Should be considered a last resort to enable reinstallation.
        - Always test in a lab environment prior to production use.
#>
function Remove-InstallerRegistration {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Filter,
        [Switch]$DeepClean
    )

    $ErrorActionPreference = "Continue"

    # 'import' .NET Functions from MSFT funcs
    Compressed       | Out-Null
    ProductListing   | Out-Null
    GetRegistryType  | Out-Null
    UninstallProduct | Out-Null
    ProductLPR       | Out-Null

    [array]$productToScrub = Get-InstallerRegistration -Filter $Filter 

    if($productToScrub.Count -eq 0){
        Write-Warning "No products matching filter: $Filter"
        return
    }

    $DateTimeRun="{0:yyyy.MM.dd.HH.mm.ss}" -f (get-date) #Run Date\Time

    $confirmMsg = if($DeepClean){
        "DEEPCLEAN ENABLED: forcibly remove associated registry keys AND PROGRAM FILES"
    } else {
        "Forcibly remove registry keys (installer/uninstall metadata)"
    }

    foreach($product in $productToScrub){
        if ($PSCmdlet.ShouldProcess(
            "", 
            "$($product.DisplayName)-$($Product.DisplayVersion)-$($product.ProductCode)",
            $confirmMsg
        )) {
            MATSFingerPrint $product.ProductCode "Version_RS_RapidProductRemoval" "1.3" $DateTimeRun | Out-Null 
            MATSFingerPrint $product.ProductCode "ProductName" $product.DisplayName $DateTimeRun     | Out-Null 
            $root=$Env:SystemDrive
            $DirectoryPath= $root+"\MATS\$($product.ProductCode)"
            
            $exitCode=[UninstallProduct2]::LogandUninstallProduct($product.ProductCode) # try Uninstall with the Windows Installer -x switch first
            MATSFingerPrint $product.ProductCode "msiexec -x" $exitCode $DateTimeRun | Out-Null

            RapidProductRegistryRemove $product.ProductCode | Out-Null
            MATSFingerPrint $product.ProductCode "RPR" $true $DateTimeRun | Out-Null

            if($DeepClean){
                LPR $product.ProductCode $product.DisplayName $DateTimeRun | Out-Null
            }

            [PSCustomObject]@{
                DisplayName        = $product.DisplayName
                ProductCode        = $product.ProductCode
                "MsiExec ExitCode" = $exitCode
                DeepClean          = $DeepClean
                BackupPath         = $DirectoryPath
            }
        }
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
    param(
        [ScriptBlock]$Filter
    )

    $registeredInstallPaths = if([Environment]::Is64BitOperatingSystem){
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    } else {
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    [array]$allRegisteredProducts = Get-ChildItem $registeredInstallPaths |
        Where-Object { $_.PsPath -like "*{*-*-*-*}"}  | # fixes errors if host has unexpected (non-guid) uninstall entries in registry 
        Foreach-Object { Get-ItemProperty $_.PSPath } |
        Select-Object *,@{Name='ProductCode';E={$_.PSChildName}}  # add ProductCode (which is parent reg folder name) as property to the objs

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