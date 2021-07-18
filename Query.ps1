
#If you use Access 32 bit use the ps(x86) version. Execute Set-ExecutionPolicy RemoteSigned first


Param (
[string]$DBLocation,
[string]$SQLname
)

$Result=@()

$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$objRecordset = New-Object -comobject ADODB.Recordset

$OpenString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + $DBLocation + ";Persist Security Info = False"
#Write-Host $OpenString

#$objConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\MyName\\TestDB1.accdb;Persist Security Info = False")
$objConnection.Open($OpenString)

#write-Host $objConnection.State
$objRecordset.Open($SQLname,$objConnection,$adOpenStatic,$adLockOptimistic)
#write-Host $objRecordset.EOF


if (($objRecordset.State) -and -not ($objRecordset.EOF))
{

    $objRecordset.MoveFirst()
    $AllItemNames=@()
    $j=0
    foreach($field in $objRecordset.Fields)
        {
            $AllItemNames+= $objRecordset.Fields.item($j).name 
            $j=$j+1
        }
    
    
    
    $objRecordset.MoveFirst()
   
    do 
    {
        $NewDBitem = New-Object  PSObject
        $i=0
        foreach($field in $objRecordset.Fields)
        {
            $NewDBitem | Add-Member -type NoteProperty -name $objRecordset.Fields.item($i).name -Value $objRecordset.Fields.Item($i).value
            $i=$i+1 
        }
    $Result+=$NewDBitem
    $objRecordset.MoveNext()
    } 
    until ($objRecordset.EOF -eq $True)


    

    # metainformation in header
    $RSSize=@()                    
    $RSSize+=$Result.length
    $RSSize+=$i
    #$RSSize+="X"  # put this to define end of header

    return 0,$RSSize,$AllItemNames,$Result | Format-List | Out-String -width 2048    

}

else
{

    Write-Host 0;0;0   #"Connection to DB is not open."


}
$objRecordset.Close()
$objConnection.Close()

