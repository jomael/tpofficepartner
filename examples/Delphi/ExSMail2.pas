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

unit ExSMail2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, OpShared, OpOlkXP, OpOutlk;

type
  TForm2 = class(TForm)
    Panel1: TPanel;
    edtTo: TEdit;
    btnTo: TButton;
    edtCC: TEdit;
    btnCC: TButton;
    edtBcc: TEdit;
    btnBCC: TButton;
    edtSubject: TEdit;
    Label1: TLabel;
    Panel2: TPanel;
    mmoBody: TMemo;
    btnSend: TButton;
    btnCancel: TButton;
    procedure btnToClick(Sender: TObject);
    procedure edtToChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSendClick(Sender: TObject);
  private
    fOutlook: TOpOutlook;
  public
    property Outlook: TOpOutlook read fOutlook write fOutlook;
  end;

var
  Form2: TForm2;

implementation

{$R *.DFM}

uses
  ExSMail3;

procedure TForm2.btnToClick(Sender: TObject);
var
  Dlg : TForm3;
begin
  Dlg := TForm3.Create( Self );
  try
    Dlg.Outlook := Outlook;
    if Dlg.ShowModal = mrOK then
    begin
      if edtTo.Text = '' then
        edtTo.Text := Dlg.RecipientTo
      else
        edtTo.Text := edtTo.Text + ';' + Dlg.RecipientTo;

      if edtCC.Text = '' then
        edtCC.Text := Dlg.RecipientCC
      else
        edtCC.Text := edtCC.Text + ';' + Dlg.RecipientCC;

      if edtBCC.Text = '' then
        edtBCC.Text := Dlg.RecipientBCC
      else
        edtBCC.Text := edtBCC.Text + ';' + Dlg.RecipientBCC;
    end;
  finally
    Dlg.Free
  end;
end;

procedure TForm2.edtToChange(Sender: TObject);
begin
  btnSend.Enabled := edtTo.Text <> '';
end;

procedure TForm2.FormShow(Sender: TObject);
begin
  if ( not Assigned(Outlook) ) or ( not(Outlook.Connected) ) then
    Close;
end;

procedure TForm2.btnSendClick(Sender: TObject);
var
  MailItem: TOpMailItem;
begin
  MailItem := Outlook.CreateMailItem;
  if assigned( MailItem ) then
  with MailItem do
  begin
    MsgTo := edtTo.Text;
    if edtCC.Text <> '' then
      MsgCC := edtCC.Text;
    if edtBCC.Text <> '' then
      MsgBCC := edtBCC.Text;
    if edtSubject.Text <> '' then
      Subject := edtSubject.Text;
    if mmoBody.Text <> '' then
      Body := mmoBody.Text;
    try
      Send;
    except
      MessageDlg( 'Mail send failure', mtWarning, [mbOK], 0 )
    end;
  end;
  Close
end;

end.
