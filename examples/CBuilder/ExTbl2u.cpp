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

#include "ExTbl2u.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "OpShared"
#pragma link "OpWord"
#pragma link "OpWrdXP"
#pragma resource "*.dfm"
TForm1 *Form1;
AnsiString ProjectCaption = "Populating a Word Table";
TOpWordDocument* Doc;
Opwrdxp::_di_Table ITable;
Opwrdxp::_di__Document IDoc;

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormCreate(TObject *Sender)
{
  Caption = ProjectCaption;
  // Turn on Word and create new document
  OpWord1->Connected = true;
  Doc = OpWord1->NewDocument();
  Doc->Insert(Caption, true);

  // Create a table with a header row
  IDoc = Doc->AsDocument;
  IDoc->Tables->AddOld(IDoc->Characters->Last,
                       gridTableData->RowCount + 1,
                       gridTableData->ColCount,
                       ITable);
  ITable->AllowAutoFit = true;

  // Populate the table
  btnUpdateClick(0);

  // Make the document visible
  OpWord1->Visible = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnUpdateClick(TObject *Sender)
{
  // Populate header row
  Opwrdxp::_di_Cell ICell;
  ITable->Cell(1, 1, ICell);
  ICell->Range->Text = edtCol1Header->Text;
  ICell->Range->Bold = 1;
  ITable->Cell(1, 2, ICell);
  ICell->Range->Text = edtCol2Header->Text;
  ICell->Range->Bold = 1;
  ITable->Cell(1, 3, ICell);
  ICell->Range->Text = edtCol3Header->Text;
  ICell->Range->Bold = 1;

  // Populate table data
  for (int j = 1; j <= gridTableData->RowCount; j++)
    for (int i = 1; i <= gridTableData->ColCount; i++) {
      ITable->Cell(j+1, i, ICell);
      ICell->Range->Text = gridTableData->Cells[i-1][j-1];
    }
  ITable->Columns->AutoFit();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
  if (Doc)
    delete Doc;
  OpWord1->Connected = false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnRefreshClick(TObject *Sender)
{
  // Refresh header row
  Opwrdxp::_di_Cell ICell;
  ITable->Cell(1, 1, ICell);
  edtCol1Header->Text = Trim(ICell->Range->Text);
  ITable->Cell(1, 2, ICell);
  edtCol2Header->Text = Trim(ICell->Range->Text);
  ITable->Cell(1, 3, ICell);
  edtCol3Header->Text = Trim(ICell->Range->Text);

  // Refresh table data
  for (int j = 1; j <= gridTableData->RowCount; j++)
    for (int i = 1; i <= gridTableData->ColCount; i++)
    {
      ITable->Cell(j+1, i, ICell);
      gridTableData->Cells[i-1][j-1] = Trim(ICell->Range->Text);
    }
}
//---------------------------------------------------------------------------

