function [NewID]=PS2DB_InsertItem(PSlocation,DBLocation,Table,TablePK,NewRecordset)

% insert a single item with all fields into the specified table
% PSloation : Powershell home
% DBLocation 
% Table for insert
% TablePK PrimaryKey of table needs to be included
% NewRecordset  The recordset to be inserted including fieldnames as a header, use any value for PK, will overwrite
% Format of Recordset in Octave is cell array with {n+1,m} cells, with n lines (items), first line is header with m column names
% RecordSetRuns cell  4x2 Dimension, RunID is PK
% RecordSetRuns =
%  {
%  [1,1] = RunID
%  [2,1] = 25
%  [3,1] = 26
%  [4,1] = 27
%  [1,2] = RunName
%  [2,2] = PullP2-1
%  [3,2] = PullP2-2
%  [4,2] = PullP2-3
%  }
%
%


Numlines=size(NewRecordset,1);
NumFields=size(NewRecordset,2);
NewRSText="";


%First get new ID for PK of table
GetIDstr=cstrcat(PSlocation,"powershell.exe ",PSlocation,"\GetNextID.ps1 -DBLocation '",DBLocation,"' -Table '",Table,"' -IDField '",TablePK,"'"); 
[sts NewIDraw] =system (GetIDstr);
%NewID=deblank(NewID);
idxnum=isdigit(NewIDraw);
NewID=NewIDraw(idxnum);
 
PKCol="";
for col=1:NumFields
  if strcmp(NewRecordset{1,col},TablePK)
    PKCol=col;
    break;
  endif
endfor

if (PKCol=="") 
  disp("TablePK not found in list of fields")
else
  for iline=1:Numlines-1
   NewRecordset{iline+1,PKCol}=num2str(str2num(NewID)+iline-1);
  end
endif

for line=1:Numlines
   for Field=1:NumFields-1
     NewRSText=cstrcat(NewRSText,strtrim(num2str(NewRecordset{line,Field})),";");
     
   endfor
     NewRSText=cstrcat(NewRSText,strtrim(num2str(NewRecordset{line,NumFields})));
     if (line ~=Numlines)
       NewRSText=strcat(NewRSText,"#");  %\n
     endif  
endfor
  

if length(NewRSText)<6000
  PSstring=cstrcat(PSlocation,"powershell.exe ",PSlocation,"\InsertDB.ps1 -DBLocation '",DBLocation,"' -Table '",Table,"' -NewRecordSetText '",NewRSText,"'"); 
  [sts text] =system (PSstring); 
else
  timenow=time();
  timestamp = strrep(num2str(timenow), '.', '_');
  filename=strcat(pwd,"\\",timestamp,".txt");
  %filename="C:\\temp\\keks.txt";
  
  fid = fopen (filename, "w");
  fputs (fid, NewRSText);
  fclose (fid); 
  
  %save(filename,'NewRSText');
  PSstring=cstrcat(PSlocation,"powershell.exe ",PSlocation,"\InsertDBfromFile.ps1 -DBLocation '",DBLocation,"' -Table '",Table,"' -filename '",filename,"'"); 
  [sts text] =system (PSstring);
  system(cstrcat("rm ",filename));
endif

  
endfunction
