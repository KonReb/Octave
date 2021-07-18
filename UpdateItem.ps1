Param (
[string]$DBLocation,
[string]$Table,
[string]$TablePK,
[string]$ItemID,
[string]$Field,
[string]$NewItemValue
)



$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$UpdateyRecordset = New-Object -comobject ADODB.Recordset


#$objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\MyName\\TestDB1.accdb;Persist Security Info = False")
$OpenString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + $DBLocation + ";Persist Security Info = False"
$objConnection.Open($OpenString)

#write-host $objConnection.State

$SQLname="SELECT * FROM " + $Table+ " WHERE "+ $TablePK +"=" + $ItemID
#write-host $SQLname
#write-host $SQLname

$UpdateyRecordset.Open($SQLname,$objConnection,$adOpenStatic,$adLockOptimistic)


$UpdateyRecordset.Fields.Item($Field).Value=$NewItemValue
$UpdateyRecordset.Update()


$UpdateyRecordset.close() 

$objConnection.Close()


return $NewID


