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

unit ExTbl2u;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, OpShared, OpWrdXP, OpWord;

type
  TForm1 = class(TForm)
    btnUpdate: TButton;
    gridTableData: TStringGrid;
    Label1: TLabel;
    Label2: TLabel;
    edtCol1Header: TEdit;
    edtCol2Header: TEdit;
    edtCol3Header: TEdit;
    OpWord1: TOpWord;
    btnRefresh: TButton;
    procedure btnUpdateClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnRefreshClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

const
  ProjectCaption = 'Populating a Word Table';

var
  oDocument : TOpWordDocument;
  iTable    : OpWrdXP.Table;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Caption := ProjectCaption;
  { Turn on Word and create new document }
  OpWord1.Connected := True;
  oDocument := OpWord1.NewDocument;
  oDocument.Insert(Caption + #13#10#10#10, True);

  { Create a table with a header row }
  iTable := oDocument.AsDocument.Tables.AddOld(oDocument.AsDocument.Characters.Last,
                                               gridTableData.RowCount + 1,
                                               gridTableData.ColCount);
  iTable.AllowAutoFit := True;

  { Populate the table }
  btnUpdateClick(nil);

  { Make the document visible }
  OpWord1.Visible := True;
end;

procedure TForm1.btnUpdateClick(Sender: TObject);
var
  i, j : Integer;
begin

  { Populate header row }
  iTable.Cell(1, 1).Range.Text := edtCol1Header.Text;
  iTable.Cell(1, 1).Range.Bold := 1;
  iTable.Cell(1, 2).Range.Text := edtCol2Header.Text;
  iTable.Cell(1, 2).Range.Bold := 1;
  iTable.Cell(1, 3).Range.Text := edtCol3Header.Text;
  iTable.Cell(1, 3).Range.Bold := 1;

  { Populate table data }
  for j := 1 to gridTableData.RowCount do
    for i := 1 to gridTableData.ColCount do
      iTable.Cell(j+1, i).Range.Text := gridTableData.Cells[i-1, j-1];
  iTable.Columns.AutoFit;
end;

procedure TForm1.btnRefreshClick(Sender: TObject);
var
  i, j : Integer;
begin
  { Refresh header row }
  edtCol1Header.Text := Trim(iTable.Cell(1, 1).Range.Text);
  edtCol2Header.Text := Trim(iTable.Cell(1, 2).Range.Text);
  edtCol3Header.Text := Trim(iTable.Cell(1, 3).Range.Text);

  { Refresh table data }
  for j := 1 to gridTableData.RowCount do
    for i := 1 to gridTableData.ColCount do
      gridTableData.Cells[i-1, j-1] := Trim(iTable.Cell(j+1, i).Range.Text);
  iTable.Columns.AutoFit;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(oDocument) then
    oDocument.Free;
  OpWord1.Connected := False;
end;

end.
