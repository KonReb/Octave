#----!!!! If using Access 32bit use Powershell (x86)-------------------


Param (
[string]$DBLocation
)

$IsAvailable=0
$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$PrimaryKeyRecordset = New-Object -comobject ADODB.Recordset

#$DBLocation="C:\\Users\\MyName\\DBConnection\\TestDB1.accdb"


$OpenString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + $DBLocation + ";Persist Security Info = False"

try
    {
      
      #Write-Host $OpenString
      $objConnection.Open($OpenString)

      $objConnection.Close();
      $IsAvailable=1;
    }
catch
    {
     #Write-Warning $Error[0]
     $IsAvailable=0;
    }


return $IsAvailable;