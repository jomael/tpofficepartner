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

#include "ExAppen1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpShared"
#pragma link "OpWord"
#pragma link "OpWrdXP"
#pragma resource "*.dfm"

#define spIDocument Opwrdxp::_di__Document  //safe pointer
#define spIRange Opwrdxp::_di_Range         //safe pointer

TForm1 *Form1;
TOpWordDocument *MainDoc;

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnNewDocumentClick(TObject *Sender)
{
  OpWord1->Visible = true;
  MainDoc = OpWord1->Documents->Add();
  MainDoc->Insert("New Document\n\n", true);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnAppendDocumentClick(TObject *Sender)
{
  if (OpenDialog1->Execute()) {
    MainDoc->Insert(OpenDialog1->FileName, true);
    spIDocument IDocument = MainDoc->AsDocument;
    OleVariant EndOfDoc = IDocument->Characters->Count - 1;
    spIRange IRange = IDocument->Content;
    IRange->InsertFile(OpenDialog1->FileName, EmptyParam,
                       EmptyParam, EmptyParam, EmptyParam);
  }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormActivate(TObject *Sender)
{
  OpWord1->Connected = true;
  OpWord1->Visible = false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
  if (MainDoc)
    delete MainDoc;
  OpWord1->Connected = false;
}
//---------------------------------------------------------------------------
