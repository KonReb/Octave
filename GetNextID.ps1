Param (
[string]$DBLocation,
[string]$Table,
[string]$IDField
)

$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$PrimaryKeyRecordset = New-Object -comobject ADODB.Recordset

#$objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\MyName\\TestDB1.accdb;Persist Security Info = False")
$OpenString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + $DBLocation + ";Persist Security Info = False"
$objConnection.Open($OpenString)

$query = "SELECT " + $IDField + " FROM " + $Table

$NewID=0

$PrimaryKeyRecordset.open($query,$objConnection,$adOpenStatic,$adLockOptimistic)
#write-host $PrimaryKeyRecordset.EOF
#$PrimaryKeyRecordset.Sort #= "ID ASC"
if (($PrimaryKeyRecordset.EOF -eq $True) -or ($PrimaryKeyRecordset.BOF -eq $True))
    {
      $NewID=1
    }
else
    {
    $PrimaryKeyRecordset.MoveLast()
    $NewID=$PrimaryKeyRecordset.Fields.Item($IDField).Value

    
    [int]$NewIDNum = [convert]::ToInt32($NewID, 10)

    }


$PrimaryKeyRecordset.close() 

$objConnection.Close()



return $NewIDNum