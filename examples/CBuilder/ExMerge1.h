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
#ifndef ExMerge1H
#define ExMerge1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <Db.hpp>
#include <DBCtrls.hpp>
#include <DBGrids.hpp>
#include <DBTables.hpp>
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include "OpDbOfc.hpp"
#include "OpModels.hpp"
#include "OpShared.hpp"
#include "OpWord.hpp"
#include "OpWrdXP.hpp"
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
    TPanel *Panel1;
    TCheckBox *cbEvent;
    TDBNavigator *DBNavigator1;
    TPanel *Panel2;
    TButton *btnPopulateTable;
    TButton *btnPopulateMailMerge;
    TButton *btnExecuteMailMerge;
    TPanel *Panel3;
    TDBGrid *DBGrid1;
    TStatusBar *StatusBar1;
    TOpWord *OpWord1;
    TOpDataSetModel *OpDataSetModel1;
    TOpDataSetModel *OpDataSetModel2;
    TTable *Table1;
    TStringField *Table1NAME;
    TSmallintField *Table1SIZE;
    TSmallintField *Table1WEIGHT;
    TStringField *Table1AREA;
    TBlobField *Table1BMP;
    TQuery *MergeQuery;
    TDataSource *DataSource1;
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall FormActivate(TObject *Sender);
    void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
    void __fastcall btnPopulateTableClick(TObject *Sender);
    void __fastcall btnPopulateMailMergeClick(TObject *Sender);
    void __fastcall btnExecuteMailMergeClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
