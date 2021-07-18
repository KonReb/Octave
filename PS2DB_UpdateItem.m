function PS2DB_UpdateItem(PSlocation,DBLocation,Table,TablePK,ItemID, Field, NewItemValue)
  
% Update an item identified with ID in PK for a single attribute
  
  Insertstr=cstrcat(PSlocation,"powershell.exe ",PSlocation,"\UpdateItem.ps1 -DBLocation '", DBLocation,"' -Table '",Table,"' -TablePK '",TablePK,"' -ItemID '",ItemID,"' -Field '",Field,"' -NewItemValue '",NewItemValue,"'"); 
  [sts text] =system (Insertstr);
  
endfunction
