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

unit ExField1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OpShared, OpWrdXP, OpWord, OleCtrls;

type
  TForm1 = class(TForm)
    btnNewDocument: TButton;
    btnAddField: TButton;
    btnShowCode: TButton;
    OpWord1: TOpWord;
    procedure btnNewDocumentClick(Sender: TObject);
    procedure btnAddFieldClick(Sender: TObject);
    procedure btnShowCodeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Doc: TOpWordDocument;
  Flds: OpWrdXP.Fields;
  Sel: OpWrdXP.Selection;
  Fld: OpWrdXP.Field;

implementation

{$R *.DFM}

procedure TForm1.FormCreate(Sender: TObject);
begin
  OpWord1.Connected:= True;
  OpWord1.Visible:= False;
end;

procedure TForm1.btnNewDocumentClick(Sender: TObject);
begin
  Doc := OpWord1.NewDocument;
  OpWord1.Visible:= True;
end;

procedure TForm1.btnAddFieldClick(Sender: TObject);
var
  direction, FieldType, Text: OleVariant;
begin
  direction:= wdCollapseStart;
  // You could add any type of field.
  // There are 91 different types available.
  // See OpWrd2K.pas
  FieldType:= wdFieldHyperlink;
  Text:= 'http://www.turbopower.com';
  Sel:= OpWord1.Server.Selection;
  Sel.Collapse(direction);
  Flds:= Doc.AsDocument.Fields;
  Fld:= Flds.Add(sel.Range, FieldType, Text, emptyParam);
end;

procedure TForm1.btnShowCodeClick(Sender: TObject);
begin
  ShowMessage(Fld.Result.Text);
  ShowMessage(Fld.Code.Text);
  Fld.ShowCodes:= True;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(Doc) then
    Doc.Free;
  OpWord1.Connected := False;
end;

end.
