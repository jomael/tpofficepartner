(* ***** BEGIN LICENSE BLOCK *****
 * Version: MPL 1.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is TurboPower OfficePartner
 *
 * The Initial Developer of the Original Code is
 * TurboPower Software
 *
 * Portions created by the Initial Developer are Copyright (C) 2000-2002
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *
 * ***** END LICENSE BLOCK ***** *)

unit ExO2XL1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OpXLXP, OpExcel, OpModels, OpDbOfc, Db,
  OpDbOlk, OpShared, OpOlkXP, OpOutlk, Buttons;

type
  TForm1 = class(TForm)
    btnExport: TButton;
    OpExcel1: TOpExcel;
    OpOutlook1: TOpOutlook;
    OpDataSetModel1: TOpDataSetModel;
    OpContactsDataSet1: TOpContactsDataSet;
    procedure btnExportClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.btnExportClick(Sender: TObject);
var
  oCursor: TCursor;
  oWorkBook: TOpExcelWorkbook;
  oSheet: TOpExcelWorksheet;
  oRange: TOpExcelRange;
begin
  oCursor := Screen.Cursor;
  Screen.Cursor := crHourglass;
  try
    // check all connections
    if not OpOutlook1.Connected then
      OpOutlook1.Connected := True;
    if not OpExcel1.Connected then
      OpExcel1.Connected := True;
    if not OpContactsDataSet1.Active then
      OpContactsDataSet1.Active := True;

    with OpExcel1 do
    begin
      oWorkBook := Workbooks.Add;            // create a workbook
      oSheet := oWorkBook.Worksheets.Add;    // add a worksheet
      oRange := oSheet.Ranges.Add;           // create range for output
      oRange.Address := 'A1';                // locate range
      oRange.OfficeModel := OpDatasetModel1; // assign the model to the range
      oRange.Populate;                       // fill the range from the model
      Visible := True;                       // show it
      Interactive := True                    // Let the user access it
    end;
  finally
    Screen.Cursor := oCursor;
  end;
end;

end.
