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

#include "ExO2XL1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpDbOfc"
#pragma link "OpDbOlk"
#pragma link "OpExcel"
#pragma link "OpModels"
#pragma link "OpOlkXP"
#pragma link "OpOutlk"
#pragma link "OpShared"
#pragma link "OpXLXP"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnExportClick(TObject *Sender)
{
  TCursor oCursor = Screen->Cursor;
  Screen->Cursor = crHourGlass;

  // check all connections
  if (!OpOutlook1->Connected)
    OpOutlook1->Connected = true;
  if (!OpExcel1->Connected)
    OpExcel1->Connected = true;
  if (!OpContactsDataSet1->Active)
    OpContactsDataSet1->Active = true;

  TOpExcelWorkbook* oWorkBook = OpExcel1->Workbooks->Add(); // create a workbook
  TOpExcelWorksheet* oSheet = oWorkBook->Worksheets->Add(); // add a worksheet
  TOpExcelRange* oRange = oSheet->Ranges->Add();            // create range for output
  oRange->Address = "A1";                                   // locate range
  oRange->OfficeModel = OpDataSetModel1;                    // assign the model to the range
  oRange->Populate();                                       // fill the range from the model
  OpExcel1->Visible = true;                                 // show it
  OpExcel1->Interactive = true;                             // Let the user access it
  Screen->Cursor = oCursor;
}
//---------------------------------------------------------------------------
