unit uData;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, fpspreadsheetctrls, fpspreadsheetgrid, ZDataset,
  Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Panel1: TPanel;
    sWorkbookSource1: TsWorkbookSource;
    grid: TsWorksheetGrid;
    Q: TZQuery;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

uses uMain;

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button2Click(Sender: TObject);
begin

end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  with Q do
       begin
         SQL.Text:= 'insert into barcode values(:code,:ket1,:ket2:,:ket3,:ket4)';
         parambyname('code').Value:=grid.;
       end;
end;

end.

