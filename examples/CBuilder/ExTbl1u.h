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
#ifndef ExTbl1uH
#define ExTbl1uH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Grids.hpp>
#include <Forms.hpp>
#include "OpModels.hpp"
#include "OpShared.hpp"
#include "OpWord.hpp"
#include "OpWrdXP.hpp"
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
    TLabel *Label2;
    TButton *btnUpdate;
    TStringGrid *StringGrid1;
    TOpWord *OpWord1;
    TOpEventModel *OpEventModel1;
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall btnUpdateClick(TObject *Sender);
    void __fastcall OpEventModel1GetColCount(TObject *Sender,
          int &ColCount);
    void __fastcall OpEventModel1GetColHeaders(TObject *Sender,
          Variant &ColHeaders);
    void __fastcall OpEventModel1GetData(TObject *Sender, int Index,
          int Row, TOpRetrievalMode Mode, int &Size, Variant &Data);
    void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
private:	// User declarations
public:		// User declarations
    __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
