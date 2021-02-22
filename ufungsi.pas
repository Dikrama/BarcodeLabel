unit uFungsi;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,ZConnection,IniFiles,forms,ZDataset,dialogs,Graphics,grids,
  fpspreadsheet,fpspreadsheetctrls,db,ExtCtrls,  comobj,fpsTypes;

  function Connection(Conn:TZConnection;Pass,host,database,user:string):boolean;
  function LoadConfig(Value:string):string;
  function SaveConfig(Field,Value:string):Boolean;
  function LoadSQL(conn:TZConnection;id:string):string;
  function LoadGrid(DBSource:TZQuery;grid:TStringGrid):boolean;
  //function insertImgToGrid(Acol,ARow:integer;Grid:TstringGrid;Dia:TOpenDialog):Boolean;
  function CreateQuery(ParamList:tstringlist;ID:String;Conn:TZConnection):TZQuery;
  function AddTotal(aGrid : TStringGrid; New : Boolean):boolean;
  //function LoadComment(Q : TZQuery; Table:string; sg : TStringGrid):boolean;
  function PrintEx(Afile : variant) : boolean;
  function FindField (DB : TZQuery; sKol:string) : boolean;
  function ExcelLoop(DBExcel, DBMaster : TZQuery; FileExcel : tsWorkSheet) : integer;
  procedure LoadGridx(DBSource:TZQuery;grid:TStringGrid);
  procedure LoadGridx2(DBSource:TZQuery;grid:TStringGrid;RowStart:Integer);
  function StrMySQLDate(ss:string):string;
  function AdjustGrid(grid:TStringGrid;ColHide:Integer):integer;
  function SaveToExcel(AFile:String;grid:TStringGrid):boolean;
  function GridColHide(grid:TStringGrid;ColHide:Integer): integer;
  
 // function LoadGridK(DBSource:TDataset;grid:TKGrid):boolean;
 
//  function AddTotalK(aGrid : TKGrid; New : Boolean):boolean;

implementation

var
  PSQL, LastError : string;

function SaveToExcel(AFile:String;grid:TStringGrid):boolean;
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  i, j : Integer;
  io : real;
begin
  try
      begin
      // Create the spreadsheet
          MyWorkbook := TsWorkbook.Create;
          MyWorksheet := MyWorkbook.AddWorksheet('Worksheet1');
          MYWORKSHEET.DefaultColWidth:=11;
     with grid do
          begin
      // Write all cells to the worksheet
            for j:=0 to rowcount - 1 do
            begin
              for i := 0 to ColCount - 1 do
              begin
                if (trystrtofloat(stringreplace(cells[i,j],',','',[rfreplaceall]),io)) and (i>1) and (j>1) then
                MyWorksheet.WriteNumber(j + 1, i, io) else
                MyWorksheet.WriteText(j + 1, i, cells[i,j]);
                MyWorksheet.WriteBackgroundColor(1,i,clyellow);
                MyWorksheet.WriteBorders(j+1,i,[cbNorth,cbsouth,cbWest,cbEast]);
              end;
            end;

      // Save the spreadsheet to a file
            MyWorkbook.WriteToFile(AFile+'.'+'xlsx');
            MyWorkbook.Free;
          end;
      end;
  except
     on E: Exception do LastError:= 'An exception was raised: ' + E.Message;
  end;
result := true;
end;

function GridColHide(grid: TStringGrid; ColHide: Integer): integer;
begin
  grid.ColWidths[ColHide]:=0;
  result := ColHide;
end;

function AdjustGrid(grid:TStringGrid;ColHide:Integer):integer;
var
  i : integer;
begin
  for i := ColHide to grid.ColCount-1 do
      begin
        grid.ColWidths[i]:=round(grid.Width/(grid.ColCount-ColHide));
      end;
  result :=round(grid.Width/(grid.ColCount-ColHide));
end;

function StrMySQLDate(ss:string):string;
var
  D : TDateTime;
begin
if trystrtodate(ss,D) then
 result := formatdatetime('yyyy-mm-dd',D) else result := '';
end;

procedure LoadGridx2(DBSource:TZQuery;grid:TStringGrid;RowStart:Integer);
var
  i,j:integer;
  ii : real;
  date : TDateTime;
begin
if DBSource.Active then
begin
  with Grid do begin
      RowCount      := DBSource.RecordCount+1;
      ColCount      := DBSource.FieldCount;
      try
          for i := 0 to DBSource.Fields.Count - 1 do
              cells[i, 0] := DBSource.Fields[i].FieldName;

              DBSource.First;
              j := RowStart;
              while not DBSource.EOF do
              begin
                for i := 0 to DBSource.Fields.Count - 1 do
                    cells[i, j + 1]:= DBSource.Fields[i].AsString;
                DBSource.Next;
                Inc(j);
              end;

       except
          on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
       end;
  end;
end
else
grid.Cells[0,0]:= 'Dataset not activated';
end;

procedure LoadGridx(DBSource:TZQuery;grid:TStringGrid);
var
  i,j:integer;
  ii : real;
begin
  with Grid do begin
      RowCount      := DBSource.RecordCount+1;
      ColCount      := DBSource.FieldCount;
      try
          for i := 0 to DBSource.Fields.Count - 1 do
              cells[i, 0] := DBSource.Fields[i].FieldName;

              DBSource.First;
              j := 0;
              while not DBSource.EOF do
              begin
                for i := 0 to DBSource.Fields.Count - 1 do
                    cells[i, j + 1]:= DBSource.Fields[i].AsString;
                DBSource.Next;
                Inc(j);
              end;

       except
          on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
       end;

  end;

end;

function ExcelLoop(DBExcel, DBMaster : TZQuery; FileExcel : tsWorkSheet) : integer;
var 
	i,ii : integer;
begin
	DBExcel.first;
	ii :=0;
	for i :=1 to DBExcel.RecordCount do
	    begin
		if FindField(DBMaster,DBExcel.FieldByName('kolom').AsString) then
		   begin
		        FileExcel.WriteText(DBExcel.FieldByName('row').AsInteger, DBExcel.FieldByName('col').AsInteger,
				DBMaster.FieldByName(DBExcel.FieldByName('kolom').AsString).AsString);
				inc(ii);					
				end;
		DBExcel.Next;
		end;
		result := ii;
end;

function FindField (DB : TZQuery; sKol:string) : boolean;
var 
	sss : TStringlist;
begin
	try
		sss := TStringlist.Create;
		DB.getfieldnames(sss);
		
		if pos(sKol,sss.text)>0 then 
				result := true 
			else 
				result := false;
	finally
		sss.free;
	end;
end;

function PrintEx(Afile : variant) : boolean;
var
  obj : OleVariant;
begin
  try
    obj := CreateOleObject('Excel.Application');
    obj.Workbooks.Open(afile);
    obj.ActiveSheet.PrintOut;    
  finally
    obj.Quit;
    obj := Unassigned;
  end;
  result := true;
end;



{function LoadGridK(DBSource:TDataset;grid:TKGrid):boolean;
var
  i,j:integer;
  ii : real;
begin
  grid.RowCount:=DBSource.RecordCount+1;
  grid.ColCount:=DBSource.FieldCount;
  try
      for i := 0 to DBSource.Fields.Count - 1 do
          grid.cells[i, 0]:= DBSource.Fields[i].FieldName;
      DBSource.First;
      j := 0;
      while not DBSource.EOF do
      begin
        for i := 0 to DBSource.Fields.Count - 1 do
         if tryStrToFloat(DBSource.Fields[i].AsString,ii)=true then
         grid.cells[i, j + 1]:= formatfloat('#,###',ii) else
           grid.cells[i, j + 1]:= DBSource.Fields[i].AsString;
        DBSource.Next;
        Inc(j);
      end;
        //grid.AutoSizeColumns;
   except
      on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
   end;
  result:=true;
end;    }
{function AddTotalK(aGrid : TKGrid; New : Boolean):boolean;
var i,ii : integer; iii : real;
begin
iii := 0;
with aGrid do
    begin
    if new=true then ColCount:=ColCount+1;
    for ii :=1 to RowCount - 1 do
      begin
         for i := 1 to colCount-2 do
             begin
                  iii  := iii + strtofloatdef(trim(stringreplace(Cells[i,ii], ',','',[rfreplaceall])),0);
             end;
          Cells[ColCount-1,ii]  := formatfloat('#,###',iii);
          iii:=0;
          Cells[colcount-1,0] := 'Total';
      end;
    result := true;
    end;
end; }

function AddTotal(aGrid : TStringGrid; New : Boolean):boolean;
var i,ii : integer; iii : real;
begin
iii := 0;
with aGrid do
    begin
    if new=true then ColCount:=ColCount+1;
    for ii :=1 to RowCount - 1 do
      begin
         for i := 1 to colCount-2 do
             begin
                  iii  := iii + strtofloatdef(trim(stringreplace(Cells[i,ii], ',','',[rfreplaceall])),0);
             end;
          Cells[ColCount-1,ii]  := formatfloat('#,###',iii);
          iii:=0;
          Cells[colcount-1,0] := 'Total';
      end;
    result := true;
    end;
end;

function CreateQuery(ParamList:tstringlist;ID:String;Conn:TZConnection):TZQuery;
var QProses:TZQuery; i:integer;
begin
  try
    QProses:=TZQuery.Create(nil);
    QProses.Connection:=conn;

    with QProses do
         begin
                SQL.Text:=loadsql(Conn,ID);
                for i:=0 to ParamList.Count -1 do
                Params[i].Value:= ParamList[i];
                if pos('select',sql.Text)<>0 then
                Open else execSQL;
         end;
         result:=QProses;
  except
    on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
  end;
  //QProses.Free;
end;


{function insertImgToGrid(Acol,ARow:integer;Grid:TstringGrid;Dia:TOpenDialog):Boolean;
var
  Image:TImage;

  Rect:TRect;
begin
  if dia.Execute then
     begin
         Image:=TImage.Create(nil);
         Rect:=TRect.Create(Image.Left,Image.Top,image.Height,image.Width);
         Image.Picture.LoadFromFile(Dia.FileName);

     if (Grid.Col=Acol) and (Grid.Row=ARow) then
        begin;
              grid.Canvas.StretchDraw(Rect,Image.Picture.Bitmap);
        end;
    image.Free;
  end;
 result :=true;
end;     }

{
function SaveToExcel(DBSource:TDataset;FileName:String):Boolean;
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  i, j: Integer;
begin
  try
  begin
  DBSource.Open;

// Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet1');

  // Write the field names
  for i := 0 to DBSource.Fields.Count - 1 do
    MyWorksheet.WriteText(0, i, DBSource.Fields[i].FieldName);

  // Write all cells to the worksheet
  DBSource.First;
  j := 0;
  while not DBSource.EOF do
  begin
    for i := 0 to DBSource.Fields.Count - 1 do
    MyWorksheet.WriteText(j + 1, i, DBSource.Fields[i].AsString);
    DBSource.Next;
    Inc(j);
  end;

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(FileName+'.'+'xlsx');
  MyWorkbook.Free;
  showmessage('Data sudah disimpan');
  Result:=True;
  end;
  except
       on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
  end;

end;
}
function LoadGrid(DBSource:TZQuery;grid:TStringGrid):boolean;
var
  i,j:integer;
  ii : real;
begin
  grid.RowCount:=DBSource.RecordCount+1;
  //grid.ColCount:=DBSource.FieldCount;
  try
      //for i := 0 to DBSource.Fields.Count - 1 do
      //    grid.cells[i, 0]:= DBSource.Fields[i].FieldName;
      DBSource.First;
      j := 0;
      while not DBSource.EOF do
      begin
        for i := 0 to DBSource.Fields.Count - 1 do
         if tryStrToFloat(DBSource.Fields[i].AsString,ii)=true then
         grid.cells[i, j + 1]:= formatfloat('#,###',ii) else
           grid.cells[i, j + 1]:= DBSource.Fields[i].AsString;
        DBSource.Next;
        Inc(j);
      end;
        //grid.AutoSizeColumns;
   except
      on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
   end;
  result:=true;
end;


function LoadSQL(conn:TZConnection;id:string):string;
var Query:TZQuery;
begin
  ///load sql from DB, result : SQL Text
  try
    Query:=TZQuery.Create(nil);
    Query.Connection:=conn;
    Query.SQL.Text:='select*from tb_storedsql where id= :id' ;
    Query.ParamByName('id').Value:=id;
    Query.Open;

    result:=Query.FieldByName('sql_sql').AsString;

  except
    on E: Exception do ShowMessage('An exception was raised: ' + E.Message);
  end;
  Query.free;
end;

function Connection(Conn:TZConnection;Pass,host,database,user:string):boolean;
begin
  //setting config connection
  try
    Conn.Disconnect;
    Conn.Database:=database;
    Conn.User:=User;
    Conn.HostName:=Host;
    Conn.Password:=Pass;
    Conn.Connect;
    result:=true;
  except
    result:=false;
  end;

end;

function LoadConfig(Value:string):string;
var
  config:TInifile;
begin
     //create file ini
  config:=TInifile.Create(extractfilepath(application.ExeName)+'config.ini');

  result := config.ReadString('Database',Value,'');
      //Connection(loadconfig(database,ss),loadconfig(host),loadconfig(....
end;

function SaveConfig(Field,Value:string):boolean;
var
  config:TInifile;
begin
     //create file ini
  config:=TInifile.Create(extractfilepath(application.ExeName)+'config.ini');

  config.WriteString('Database',Field,Value);

  result:=true;
end;

end.

