// ***** BEGIN LICENSE BLOCK *****
// * Version: MPL 1.1
// *
// * The contents of this file are subject to the Mozilla Public License Version
// * 1.1 (the "License"); you may not use this file except in compliance with
// * the License. You may obtain a copy of the License at
// * http://www.mozilla.org/MPL/
// *
// * Software distributed under the License is distributed on an "AS IS" basis,
// * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
// * for the specific language governing rights and limitations under the
// * License.
// *
// * The Original Code is TurboPower OfficePartner
// *
// * The Initial Developer of the Original Code is
// * TurboPower Software
// *
// * Portions created by the Initial Developer are Copyright (C) 2000-2002
// * the Initial Developer. All Rights Reserved.
// *
// * Contributor(s):
// *
// * ***** END LICENSE BLOCK *****
//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "ExSMail1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpOlkXP"
#pragma link "OpOutlk"
#pragma link "OpShared"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnSendClick(TObject *Sender)
{
  if (!OpOutlook1->Connected)
    OpOutlook1->Connected = true;
  TOpMailItem* MailItem = OpOutlook1->CreateMailItem();
  if (MailItem) {
    MailItem->MsgTo = edtTo->Text;
    MailItem->MsgCC = edtCC->Text;
    MailItem->MsgBCC = edtBcc->Text;
    MailItem->Subject = edtSubject->Text;
    MailItem->Body = mmoBody->Text;
    MailItem->Send();
  }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
  OpOutlook1->Connected = false;    
}
//---------------------------------------------------------------------------
