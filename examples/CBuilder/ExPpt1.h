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
#ifndef ExPpt1H
#define ExPpt1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "OpPower.hpp"
#include "OpPptXP.hpp"
#include "OpShared.hpp"
#include <Dialogs.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
    TButton *btnOpenPresentation;
    TButton *btnRunSlideShow;
    TButton *btnClosePresentation;
    TGroupBox *gbOptions;
    TLabel *Label1;
    TLabel *Label2;
    TLabel *Label3;
    TLabel *Label4;
    TComboBox *cbxTransitionEffect;
    TComboBox *cbxLayout;
    TEdit *edtAdvanceTime;
    TComboBox *cbxTransitionSpeed;
    TOpPowerPoint *OpPowerPoint1;
    TOpenDialog *OpenDialog1;
    TCheckBox *chkAdvanceOnClick;
    TCheckBox *chkAdvanceOnTime;
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall btnOpenPresentationClick(TObject *Sender);
    void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
    void __fastcall btnRunSlideShowClick(TObject *Sender);
    void __fastcall btnClosePresentationClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
