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

unit ExAppen1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OpShared, OpWord, OpWrdXP, OleCtrls;

type
  TForm1 = class(TForm)
    btnNewDocument: TButton;
    btnAppendDocument: TButton;
    OpenDialog1: TOpenDialog;
    OpWord1: TOpWord;
    procedure btnNewDocumentClick(Sender: TObject);
    procedure btnAppendDocumentClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  MainDoc: TOpWordDocument;
implementation

{$R *.DFM}

procedure TForm1.btnNewDocumentClick(Sender: TObject);
begin
  OpWord1.Visible:= True;
  MainDoc:=  OpWord1.Documents.Add;
  MainDoc.Insert('New Document' + #10#13#10#13, True);
  btnNewDocument.Enabled := False;
end;

procedure TForm1.btnAppendDocumentClick(Sender: TObject);
var
  EndOfDocVariant: OleVariant;
  EndOfDocRange: Range;
begin
  if OpenDialog1.Execute then
  begin
    
    MainDoc.Insert(OpenDialog1.FileName + #10#13#10#13, True);
    EndOfDocVariant:= (MainDoc.AsDocument).Characters.Count - 1;
    EndOfDocRange:= (MainDoc.AsDocument).Range(EndOfDocVariant, EndOfDocVariant);
    EndOfDocRange.InsertFile(OpenDialog1.FileName, emptyParam, emptyParam, emptyParam, emptyParam);
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(MainDoc) then
    MainDoc.Free;
  OpWord1.Connected := False;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  OpWord1.Connected := True;
  OpWord1.Visible := False;
end;

end.
