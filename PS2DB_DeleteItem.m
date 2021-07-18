function [text]=PS2DB_DeleteItem(PSlocation,DBLocation,Table,TablePK,ItemID)

% Delete one item in specified table of Database. Item identified by ID of Primary Key
% PSloation : Powershell home
% DBLocation: For Location of access-file
% Table for deletion
% TablePK PrimaryKey of table needs to be included
% ItemID  ID of item to be deleted
 

SQLDel=cstrcat("DELETE FROM ",Table," WHERE ",TablePK," = ",ItemID);
 
PSstring=cstrcat(PSlocation,"powershell.exe ",PSlocation,"\DeleteItemDB.ps1 -DBLocation '",DBLocation,"' -SQLDelname '",SQLDel,"'"); 

[sts text] =system (PSstring); 

endfunction
