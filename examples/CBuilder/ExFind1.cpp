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

#include "ExFind1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpShared"
#pragma link "OpWord"
#pragma link "OpWrdXP"
#pragma resource "*.dfm"
TForm1 *Form1;
TOpWordDocument* Doc1;

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnFindClick(TObject *Sender)
{
  Doc1->Find(edtFindText->Text, true);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnFindNextClick(TObject *Sender)
{
  Doc1->FindNext();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::OpWord1OpDisconnect(TObject *Sender)
{
  Caption = "Word Find example";
  OpWord1->Visible = false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnOpenDocClick(TObject *Sender)
{
  btnCloseDocClick(0);
  if (OpenDialog1->Execute()) {
    Screen->Cursor = crHourGlass;
    Doc1 = OpWord1->OpenDocument(OpenDialog1->FileName);
  }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnCloseDocClick(TObject *Sender)
{
  if (Doc1)
    delete Doc1;
  Doc1 = 0;
  OpWord1->Connected = false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
  btnCloseDocClick(0);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::OpWord1DocumentOpen(TObject *Sender,
      _Document *Doc)
{
  WideString WS;
  Doc->Get_FullName(WS);
  Caption = WS;
  OpWord1->Visible = true;
  Screen->Cursor = crDefault;
}
//---------------------------------------------------------------------------

