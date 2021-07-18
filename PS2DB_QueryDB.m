function [RecordSet,numitems,numfields]=PS2DB_QueryDB(PSHome,DBLocation,SQLCommand)

% Run SQL Query on Database
% Returns RecordSet as Cellarray with first line as field names
% Call like:
% PSHome="C:\\Windows\\SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe";
% DBLocation="C:\\Users\\MyName\\TestDB1.accdb";
% SQLCommand="SELECT * FROM Runs WHERE RunName=\"run1\"";  


qrystr=cstrcat(PSHome,"powershell.exe ",PSHome,"\Query.ps1 -DBLocation '",DBLocation,"' -SQLname "," '",SQLCommand,"'");
qrystr = strrep (qrystr, '"', '\"');
fid = popen (qrystr, "r");

fskipl (fid, 1);

    s = fgetl (fid);
    numitems=str2num(s);
    s = fgetl (fid);
    numfields=str2num(s);
    RecordSet=cell(numitems+1,numfields);
  
  for fieldj=1:numfields
      s = fgetl (fid);
      RecordSet{1,fieldj}=s;
  endfor



  fskipl (fid, 2);
    
  for item=1:numitems
    fieldj=1;
    while (not(isempty(s = fgetl (fid))))   
      
      startval=strfind(s,":");
      if not(isempty(startval))
        RecordSet{item+1,fieldj}=s(startval+2:length(s));
      else  %add to last line
        RecordSet{item+1,fieldj-1}=strcat(RecordSet{item+1,fieldj-1},s);    
        fieldj=fieldj-1; 
      end
    fieldj=fieldj+1; 
    endwhile
  endfor
  

  RecordSet=strtrim(RecordSet);
  fclose(fid);

endfunction