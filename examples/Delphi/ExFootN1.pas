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

unit ExFootN1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OpShared, OpWrdXP, OpWord, OleCtrls, Mask, Buttons;

type
  TForm1 = class(TForm)
    OpWord1: TOpWord;
    btnOpen: TButton;
    OpenDialog1: TOpenDialog;
    btnClose: TButton;
    edtFileName: TEdit;
    btnOpenDialog: TSpeedButton;
    GroupBox1: TGroupBox;
    mmoFootnote: TMemo;
    btnAddFootnote: TButton;
    btnGetFootnote: TButton;
    Label2: TLabel;
    Label1: TLabel;
    btnDeleteFootnote: TButton;
    edtFootnoteRef: TEdit;
    procedure btnOpenClick(Sender: TObject);
    procedure btnAddFootnoteClick(Sender: TObject);
    procedure btnOpenDialogClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnGetFootnoteClick(Sender: TObject);
    procedure btnDeleteFootnoteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

var
  iFootNotes : FootNotes;

procedure TForm1.FormCreate(Sender: TObject);
begin
  OpWord1.Connected := True;
  OpWord1.Visible := True;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  btnCloseClick(nil);
end;

procedure TForm1.btnOpenClick(Sender: TObject);
begin
  OpWord1.OpenDocument(edtFileName.Text);
  iFootNotes := OpWord1.Documents[0].AsDocument.Footnotes;
end;

procedure TForm1.btnCloseClick(Sender: TObject);
begin
  if OpWord1.Connected then
    OpWord1.Connected := False;
end;

procedure TForm1.btnAddFootnoteClick(Sender: TObject);
var
  varText : OleVariant;
  iRange : Range;
begin
  varText := mmoFootnote.Lines.Text;
  iRange := OpWord1.Server.Selection.Range;  // current cursor location
  iFootNotes.Add(iRange, emptyParam, varText);
end;

procedure TForm1.btnGetFootnoteClick(Sender: TObject);
var
  varText : OleVariant;
  Index : Integer;
begin
  mmoFootnote.Clear;
  Index := StrToIntDef(edtFootnoteRef.Text, 1);
  if (Index > 0) and (Index <= iFootNotes.Count) then begin
    varText := iFootnotes.Item(Index).Range.Text;
    mmoFootnote.Text := varText;
  end;
end;

procedure TForm1.btnDeleteFootnoteClick(Sender: TObject);
var
  Index : Integer;
begin
  mmoFootnote.Clear;
  Index := StrToIntDef(edtFootnoteRef.Text, 1);
  if (Index > 0) and (Index <= iFootNotes.Count) then
    iFootnotes.Item(Index).Delete;
end;

procedure TForm1.btnOpenDialogClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
    edtFileName.Text := OpenDialog1.FileName;
end;

end.
