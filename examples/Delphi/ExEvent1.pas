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

unit ExEvent1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OpShared, OpOlkXP, OpOutlk;

type
  TFrmMain = class(TForm)
    BtnSend: TButton;
    GroupBox1: TGroupBox;
    EdtEventTo: TEdit;
    MemEventMsg: TMemo;
    GroupBox2: TGroupBox;
    EdtTo: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    MemMsg: TMemo;
    Label3: TLabel;
    Label4: TLabel;
    BtnExit: TButton;
    OpOutlook1: TOpOutlook;
    procedure BtnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BtnSendClick(Sender: TObject);
    procedure OpOutlook1ItemSend(Item: IDispatch; var Cancel: WordBool);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmMain: TFrmMain;

implementation

{$R *.DFM}

procedure TFrmMain.BtnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmMain.FormCreate(Sender: TObject);
begin
  OpOutlook1.Connected := True;
end;

procedure TFrmMain.BtnSendClick(Sender: TObject);
begin
  with OpOutlook1.CreateMailItem do
    try
      MsgTo := EdtTo.Text;
      Subject := 'OpOutlook Event Test';
      HTMLBody := MemMsg.Lines.Text;
      Send;
    finally
      Free;
    end;
end;

procedure TFrmMain.OpOutlook1ItemSend(Item: IDispatch;
  var Cancel: WordBool);
var
  Mail: TOpMailItem;
begin
  if Item <> nil then
  begin
    Mail := TOpMailItem.Create(Item as _MailItem);
    try
      EdtEventTo.Text := Mail.MsgTo;
      MemEventMsg.Lines.Text := Mail.Body;
    finally
      Mail.Free;
    end;
  end;
end;

end.
