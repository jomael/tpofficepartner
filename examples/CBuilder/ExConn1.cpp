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

#include "ExConn1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpShared"
#pragma link "OpWord"
#pragma link "OpWrdXP"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnConnectClick(TObject *Sender)
{
  OpWord1->Connected = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::OpWord1GetInstance(TObject *Sender,
      IDispatch *&Instance, const TGUID &CLSID, const TGUID &IID)
{
  IUnknown* pUnk;
  if (IsEqualGUID(Opwrdxp::CLASS_Application_, CLSID) &
      IsEqualGUID(Opwrdxp::CLASS_Application_, IID)) {

    // Get Active Instance of Word and connect to it
    OleCheck(GetActiveObject(Opwrdxp::CLASS_Application_, 0, &pUnk));
    Instance = (IDispatch*) pUnk;
  }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::OpWord1OpConnect(TObject *Sender)
{
  OpWord1->Caption = "Hello from OfficePartner";
  OpWord1->Visible = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
  OpWord1->Connected = false;
}
//---------------------------------------------------------------------------
