function IsAvailable=PS2DB_IsAvailable(PSlocation,DBLocation)
  
% Checks if a database is available. Connection can be established. Is not established a priori.

% IsConnected=1  DBConnection is available
% IsConnected=0  DBConnection is not available
  
  
  GetIDstr=cstrcat(PSlocation,"powershell.exe ",PSlocation,"\IsAvailable.ps1 -DBLocation '",DBLocation,"'"); 
  [sts IsAvailable] =system (GetIDstr);
  
  IsAvailable=str2num(deblank(IsAvailable));
  
 endfunction 
  