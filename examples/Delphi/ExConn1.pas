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

unit ExConn1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, OleCtnrs, StdCtrls, OpShared, OpWrdXP, OpWord;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    btnConnect: TButton;
    OpWord1: TOpWord;
    procedure btnConnectClick(Sender: TObject);
    procedure OpWord1GetInstance(Sender: TObject; var Instance: IDispatch;
      const CLSID, IID: TGUID);
    procedure OpWord1OpConnect(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  end;

var
  Form1: TForm1;

implementation

uses ComObj, ActiveX;

{$R *.DFM}

procedure TForm1.btnConnectClick(Sender: TObject);
begin
  OpWord1.Connected := True;
end;

procedure TForm1.OpWord1GetInstance(Sender: TObject;
  var Instance: IDispatch; const CLSID, IID: TGUID);
var
  Unk: IUnknown;
begin
  if IsEqualGuid(CLASS_Application_, CLSID) and
    IsEqualGuid(_Application, IID) then
  begin
    // Get Active Instance of Word and connect to it
    OleCheck(GetActiveObject(CLASS_Application_, nil, Unk));
    Instance := Unk as _Application;
  end;
end;

procedure TForm1.OpWord1OpConnect(Sender: TObject);
begin
  OpWord1.Caption := 'Hello from OfficePartner';
  OpWord1.Visible := True;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  OpWord1.Connected := False;
end;

end.
