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

unit ExMerge1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DBCtrls, ComCtrls, ExtCtrls, Grids, DBGrids, OpWord,
  Db, DBTables, OpModels, OpDbOfc, OpShared, OpWrdXP;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    StatusBar1: TStatusBar;
    cbEvent: TCheckBox;
    DBNavigator1: TDBNavigator;
    btnPopulateTable: TButton;
    btnPopulateMailMerge: TButton;
    btnExecuteMailMerge: TButton;
    DBGrid1: TDBGrid;
    OpWord1: TOpWord;
    OpDataSetModel1: TOpDataSetModel;
    OpDataSetModel2: TOpDataSetModel;
    Table1: TTable;
    Table1NAME: TStringField;
    Table1SIZE: TSmallintField;
    Table1WEIGHT: TSmallintField;
    Table1AREA: TStringField;
    Table1BMP: TBlobField;
    MergeQuery: TQuery;
    DataSource1: TDataSource;
    procedure btnPopulateTableClick(Sender: TObject);
    procedure btnPopulateMailMergeClick(Sender: TObject);
    procedure btnExecuteMailMergeClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

uses
  ExMerge2;

var
  MergeDoc: TOpWordDocument;

procedure TForm1.btnPopulateTableClick(Sender: TObject);
begin
  with OpWord1.Documents[0].Tables.Add do begin
    OfficeModel := OpDatasetModel1;
    PopulateDocTable;
  end;
end;

procedure TForm1.btnPopulateMailMergeClick(Sender: TObject);
begin
  if (Form2.ShowModal = mrOK) then
    if (Form2.edtAlias.Text <> '') and (Form2.mmoSQL.Text <> '') then begin
      MergeQuery.Close;
      MergeQuery.DatabaseName := Form2.edtAlias.Text;
      MergeQuery.SQL.Clear;
      MergeQuery.SQL.Add(Form2.mmoSQL.Text);
      MergeQuery.Open;
      MergeDoc.MailMerge.OfficeModel := OpDatasetModel2;
      MergeDoc.PopulateMailMerge;
      btnExecuteMailMerge.Enabled := True;
    end;
end;

procedure TForm1.btnExecuteMailMergeClick(Sender: TObject);
begin
  MergeDoc.MailMerge.AsMailMerge.Destination := wdSendToNewDocument;
  MergeDoc.ExecuteMailMerge;
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
  OpWord1.Visible := True;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  OpWord1.Connected := True;
  OpWord1.Visible := False;
  MergeDoc := OpWord1.NewDocument;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(MergeDoc) then
    MergeDoc.Free;
  OpWord1.Connected := False;
end;

end.
