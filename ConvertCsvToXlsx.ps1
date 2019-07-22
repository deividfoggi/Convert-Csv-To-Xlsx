#    Get-O365AcceptedDomainByProxyAddresses.ps1
#
#    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
#    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,        
#    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#    We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute
#    the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks
#    to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on
#    Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us
#    and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or resultfrom the 
#    use or distribution of the Sample Code.
#    Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained 
#    within the Premier Customer Services Description.
#
#

########################################################################################################################
# MICROSOFT - PFE Team Brazil
#
# File : ConvertCsvToXlsx.ps1
# Version : 1.0
# Creation date : Jul 22nd, 2019
# Modification date : Jul 22nd, 2019
#
# Author: Deivid de Foggi - Office 365 PFE
#
# 
#########################################################################################################################


# Parameters
Param(
    [Parameter(Mandatory=$true)][String]$Source,
    [String]$TableName
)

If(!$TableName){
    $TableName = "Table"
}

$source = (Get-Item $source).FullName
$destination = $source.Replace(".csv",".xlsx")

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$workbook = $Excel.Workbooks.Open($source)
$worksheet = $workbook.ActiveSheet

$ListObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.Name = $TableName

$workbook.SaveAs($destination,51)
$workbook.Close()
$Excel.quit()