

Param (
[string]$DBLocation,
[string]$Table,
[string]$NewRecordSetText
)

$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$InsertRecordset = New-Object -comobject ADODB.Recordset


#$objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\MyName\\TestDB1.accdb;Persist Security Info = False")
$OpenString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + $DBLocation + ";Persist Security Info = False"
$objConnection.Open($OpenString)
$query = "SELECT * FROM " + $Table
$InsertRecordset.open($query,$objConnection,$adOpenStatic,$adLockOptimistic)
#write-Host $objConnection.State

$Alllines=@()
$AllItems=@()
$AllItemNames=@()
$Alllines=$NewRecordSetText.Split("#");
#Write-Host $NewRecordSetText
#Write-Host $Alllines.Count
$AllItemNames=$Alllines[0].Split(";")  #Get the names of items from first line


for ($i=1;$i -le $Alllines.Count-1; $i++)
{
#Write-Host $Alllines[$i]
$InsertRecordset.Addnew()
$AllItems=$Alllines[$i].Split(";")
  for ($j=0;$j -le $Allitems.Count-1;$j++)
  {
    #Write-Host $AllItemNames[$j],$AllItems[$j]
    $InsertRecordset.Fields.Item($AllItemNames[$j]).value = $AllItems[$j]
    #Write-Host $AllItems[$j]
  }
  $InsertRecordset.Update()
}

$InsertRecordset.close() 

$objConnection.Close()

return 1