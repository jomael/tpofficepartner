(* ***** BEGIN LICENSE BLOCK *****
 * Version: MPL 1.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is TurboPower OfficePartner
 *
 * The Initial Developer of the Original Code is
 * TurboPower Software
 *
 * Portions created by the Initial Developer are Copyright (C) 2000-2002
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *
 * ***** END LICENSE BLOCK ***** *)

unit XLDemo1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, Grids, DBGrids, DBCtrls, StdCtrls, ComCtrls, OpExcel, activeX,
  OpMSO, OpShared, OpXLXP, OleCtrls;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    DBGrid2: TDBGrid;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    btnMasterDetail: TButton;
    btnCustomer: TButton;
    btnChart: TButton;
    stsbarDoubleClick: TStatusBar;
    OpExcel: TOpExcel;
    LaunchExcelBtn: TButton;
    procedure btnMasterDetailClick(Sender: TObject);
    procedure btnCustomerClick(Sender: TObject);
    procedure OpExcelBeforeSheetDoubleClick(Sender: TObject;
      const Sh: IDispatch; const Target: ExcelRange; var Cancel: WordBool);
    procedure btnChartClick(Sender: TObject);
    procedure LaunchExcelBtnClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses XLDemo2;

{$R *.DFM}

procedure TForm1.btnMasterDetailClick(Sender: TObject);
begin
  if OpExcel.Connected = False then
    LaunchExcelBtnClick(Self);
  {Properties hooked up at design-time, just go!}
  with DataModule1 do begin
    mdlCustomers.DetailModel := mdlOrders;
    {Accessing Ranges through collection hierarchy}
    OpExcel.RangeByName['MasterDetail'].Populate;
    OpExcel.Workbooks[0].Worksheets[1].Activate;
    mdlCustomers.DetailModel := nil;
  end;
end;

procedure TForm1.btnCustomerClick(Sender: TObject);
var
 Rng : TOpExcelRange;
begin
  if OpExcel.Connected = False then
    LaunchExcelBtnClick(Self);
  {the following 5 properties can be set at design-time.
  They are set here in code for demonstration purposes
  only}
  {add a range to the first worksheet}
  Rng := OpExcel.Workbooks[0].Worksheets[0].Ranges.Add;
  {Give the name a range (simplifies access)}
  Rng.Name := 'CustomerRange';
  {We'll only set the anchor cell since we are
    populating the range through a dataset (unknown
    columns & rows)}
  Rng.Address := 'A1';
  {associate the DataSeodel with the range}
  Rng.OfficeModel := DataModule1.mdlCustomers;
  {associate the DataSeodel with the DataSet}
  DataModule1.mdlCustomers.Dataset :=
    DataModule1.tblCustomers;
  {Fill the first worksheet (starting with the anchor
  cell) with every data column and row in the DataSet}
  OpExcel.RangeByName['CustomerRange'].Populate;
  {Activate the first worksheet}
  OpExcel.Workbooks[0].Worksheets[0].Activate;
  Rng.AsRange.Columns.AutoFit;
end;


procedure TForm1.OpExcelBeforeSheetDoubleClick(Sender: TObject;
  const Sh: IDispatch; const Target: ExcelRange; var Cancel: WordBool);
begin
  stsbarDoubleClick.SimpleText := 'You just double clicked ' +
    Target.Address[True,True,xlA1,EmptyParam,0];
end;

procedure TForm1.btnChartClick(Sender: TObject);
var
  Chart : TOpExcelChart;
begin
  if OpExcel.Connected = false then
    LaunchExcelBtnClick(Self);
  {Run Query & add data to worksheet[2]}
  DataModule1.qryChartData.Active := True;
  OpExcel.RangeByName['ChartData'].Populate;

  {Add and position a chart to the worksheet}
  Chart := OpExcel.Workbooks[0].Worksheets[2].Charts.Add;
  Chart.Left := 350;
  Chart.Width := 550;
  Chart.Height := 400;
  {We'll create a pie chart..}
  Chart.ChartType := xlct3DPie;
  {Tell the chart where to find the data}
  Chart.DataRange := 'A1:C100';

  {Close the query}
  DataModule1.qryChartData.Active := False;
  OpExcel.Workbooks[0].Worksheets[2].Activate;
end;

procedure TForm1.LaunchExcelBtnClick(Sender: TObject);
begin
  {launch instance of Excel}
  OpExcel.Connected := True;
  {Office applications start up hidden by default. To show Excel, all
  we need to do is set its visible property to True.}
  OpExcel.Visible := True;
end;

procedure TForm1.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  {make sure we close Excel}
  if OpExcel.Connected = True then
    OpExcel.Connected := false;
end;

end.
 