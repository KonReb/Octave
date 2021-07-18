
#if you use  Access 32 bit use  ps(x86) version. Execute Set-ExecutionPolicy RemoteSigned first


Param (
[string]$DBLocation,
[string]$SQLDELname
)



$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$objRecordset = New-Object -comobject ADODB.Recordset

$OpenString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + $DBLocation + ";Persist Security Info = False"
#Write-Host $OpenString

#$objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\MyName\\TestDB1.accdb;Persist Security Info = False")
$objConnection.Open($OpenString)
#write-host $SQLDELname
#write-Host $objConnection.State
$success=$objRecordset.Open($SQLDELname,$objConnection,$adOpenStatic,$adLockOptimistic)


write-host $success
$objConnection.Close()


