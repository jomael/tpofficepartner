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

#include "XlDemo1.h"
#include "XLDemo2.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpExcel"
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
void __fastcall TForm1::LaunchExcelBtnClick(TObject *Sender)
{
  OpExcel->Connected = true;
  OpExcel->Visible = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnMasterDetailClick(TObject *Sender)
{
  if (!OpExcel->Connected)
    LaunchExcelBtnClick(0);

  // Properties hooked up at design-time, just go!
  DataModule1->mdlCustomers->DetailModel = DataModule1->mdlOrders;
  // Accessing Ranges through collection hierarchy
  OpExcel->RangeByName["MasterDetail"]->Populate();
  OpExcel->Workbooks->Items[0]->Worksheets->Items[1]->Activate();
  DataModule1->mdlCustomers->DetailModel = 0;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnCustomerClick(TObject *Sender)
{
  if (!OpExcel->Connected)
    LaunchExcelBtnClick(0);

  TOpExcelRange* Rng = OpExcel->Workbooks->Items[0]->Worksheets->Items[0]->Ranges->Add();
  Rng->Name = "CustomerRange";
  Rng->Address = "A1";
  Rng->OfficeModel = DataModule1->mdlCustomers;
  DataModule1->mdlCustomers->Dataset = DataModule1->tblCustomers;
  OpExcel->RangeByName["CustomerRange"]->Populate();
  OpExcel->Workbooks->Items[0]->Worksheets->Items[0]->Activate();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnChartClick(TObject *Sender)
{
  if (!OpExcel->Connected)
    LaunchExcelBtnClick(0);

  //Run Query & add data to worksheet[2]
  DataModule1->qryChartData->Active = true;
  OpExcel->RangeByName["ChartData"]->Populate();

  // Add and position a chart to the worksheet
  TOpExcelChart* Chart = OpExcel->Workbooks->Items[0]->Worksheets->Items[2]->Charts->Add();
  Chart->Left = 350;
  Chart->Width = 550;
  Chart->Height = 400;
  // create a pie chart
  Chart->ChartType = xlct3DPie;
  // Tell the chart where to find the data
  Chart->DataRange = "A1:C100";

  // Close the query
  DataModule1->qryChartData->Active = false;
  OpExcel->Workbooks->Items[0]->Worksheets->Items[2]->Activate();
}
//---------------------------------------------------------------------------
