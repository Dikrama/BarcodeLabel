unit uMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, db, FileUtil, ZConnection, ZDataset, SpkToolbar, spkt_Tab,
  spkt_Pane, spkt_Buttons, Forms, Controls, Graphics, Dialogs, DBGrids, DbCtrls,
  ComCtrls, Grids, ExtCtrls, Buttons, UFUNGSI, LR_BarC, LR_Class, LR_Desgn,
  LR_DBSet, lr_e_pdf;

type

  { TFMain }

  TFMain = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    frBarCodeObject1: TfrBarCodeObject;
    frDBDataSet1: TfrDBDataSet;
    frDesigner1: TfrDesigner;
    Panel2: TPanel;
    R: TfrReport;
    frTNPDFExport1: TfrTNPDFExport;
    PageControl1: TPageControl;
    Panel1: TPanel;
    SpkLargeButton1: TSpkLargeButton;
    SpkLargeButton2: TSpkLargeButton;
    SpkLargeButton3: TSpkLargeButton;
    SpkLargeButton4: TSpkLargeButton;
    SpkPane1: TSpkPane;
    SpkTab1: TSpkTab;
    SpkToolbar1: TSpkToolbar;
    Con: TZConnection;
    Q: TZQuery;
    grid: TStringGrid;
    TabSheet1: TTabSheet;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpkLargeButton2Click(Sender: TObject);
    procedure SpkLargeButton3Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  FMain: TFMain;

implementation

uses uData;

{$R *.lfm}

{ TFMain }

procedure TFMain.FormShow(Sender: TObject);
begin
  try
    Con.Connect;
    Q.Open;
    LoadGridx(Q,grid);
  finally

  end;

end;

procedure TFMain.BitBtn1Click(Sender: TObject);
begin
  udata.Form1.Show;
end;

procedure TFMain.SpkLargeButton2Click(Sender: TObject);
begin
 R.DesignReport;
end;

procedure TFMain.SpkLargeButton3Click(Sender: TObject);
begin

end;

end.

