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

unit XlRange1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, OpShared, OpExcel, OpXLXP;

type
  TExRgPro1 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    BorderWeightCB: TComboBox;
    GroupBox1: TGroupBox;
    BorderLeftcb: TCheckBox;
    BorderRightcb: TCheckBox;
    BorderTopcb: TCheckBox;
    BorderBottomcb: TCheckBox;
    BorderStyleCB: TComboBox;
    ColorBtn: TButton;
    GroupBox2: TGroupBox;
    Boldcb: TCheckBox;
    Italiccb: TCheckBox;
    Underlinecb: TCheckBox;
    Strikethroughcb: TCheckBox;
    FontColorBtn: TButton;
    FontNameEdit: TEdit;
    FontSizeEdit: TEdit;
    hAligncb: TComboBox;
    vAligncb: TComboBox;
    IndentEdit: TEdit;
    Orientationcb: TComboBox;
    Patterncb: TComboBox;
    PatternColorBtn: TButton;
    RotateEdit: TEdit;
    ShrinkToFitcb: TCheckBox;
    SimpleTextEdit: TEdit;
    AddressEdit: TEdit;
    ClearAllBtn: TButton;
    ColWidthTB: TTrackBar;
    RowHeightTB: TTrackBar;
    WrapTextcb: TCheckBox;
    AutoFitColBtn: TButton;
    AutoFitRowBtn: TButton;
    SetAddressBtn: TButton;
    OpExcel1: TOpExcel;
    ColorDialog1: TColorDialog;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BorderWeightCBChange(Sender: TObject);
    procedure BorderStyleCBChange(Sender: TObject);
    procedure BorderLeftcbClick(Sender: TObject);
    procedure BorderRightcbClick(Sender: TObject);
    procedure BorderTopcbClick(Sender: TObject);
    procedure BorderBottomcbClick(Sender: TObject);
    procedure ColorBtnClick(Sender: TObject);
    procedure FontColorBtnClick(Sender: TObject);
    procedure PatternColorBtnClick(Sender: TObject);
    procedure FontNameEditExit(Sender: TObject);
    procedure FontSizeEditExit(Sender: TObject);
    procedure IndentEditExit(Sender: TObject);
    procedure RotateEditExit(Sender: TObject);
    procedure ShrinkToFitcbClick(Sender: TObject);
    procedure BoldcbClick(Sender: TObject);
    procedure ItaliccbClick(Sender: TObject);
    procedure UnderlinecbClick(Sender: TObject);
    procedure StrikethroughcbClick(Sender: TObject);
    procedure PatterncbChange(Sender: TObject);
    procedure hAligncbChange(Sender: TObject);
    procedure vAligncbChange(Sender: TObject);
    procedure OrientationcbChange(Sender: TObject);
    procedure ClearAllBtnClick(Sender: TObject);
    procedure ColWidthTBChange(Sender: TObject);
    procedure RowHeightTbChange(Sender: TObject);
    procedure AutoFitColBtnClick(Sender: TObject);
    procedure AutoFitRowBtnClick(Sender: TObject);
    procedure WrapTextcbClick(Sender: TObject);
    procedure SimpleTextEditChange(Sender: TObject);
    procedure SetAddressBtnClick(Sender: TObject);
  private
    FColor        : TColor;
    FFontColor    : TColor;
    FPatternColor : TColor;
    procedure UpdateAll;
  end;

var
  ExRgPro1: TExRgPro1;

implementation

{$R *.DFM}

procedure TExRgPro1.FormShow(Sender: TObject);
var
  i : Integer;
begin
  OpExcel1.Connected := True;
  OpExcel1.Visible := True;
  for i := 0 to self.ComponentCount - 1 do
    begin
      if self.Components[i] is TComboBox then
        (Components[i] as TComboBox).ItemIndex := 0;
    end;
  FColor := clWindow;
  FFontColor := clBlack;
  FPatternColor := clWindow;
  UpdateAll;
end;

procedure TExRgPro1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  OpExcel1.Connected := False;
end;

procedure TExRgPro1.BorderWeightCBChange(Sender: TObject);
begin
  case BorderWeightCB.ItemIndex of
    0 : OpExcel1.RangeByName['r'].BorderLineWeight := xlbwHairline;
    1 : OpExcel1.RangeByName['r'].BorderLineWeight := xlbwThin;
    2 : OpExcel1.RangeByName['r'].BorderLineWeight := xlbwMedium;
    3 : OpExcel1.RangeByName['r'].BorderLineWeight := xlbwThick;
  end;
end;

procedure TExRgPro1.BorderStyleCBChange(Sender: TObject);
begin
  case BorderStyleCB.ItemIndex of
    0 : OpExcel1.RangeByName['r'].BorderStyle := xlblsContinuous;
    1 : OpExcel1.RangeByName['r'].BorderStyle := xlblsDash;
    2 : OpExcel1.RangeByName['r'].BorderStyle := xlblsDashDot;
    3 : OpExcel1.RangeByName['r'].BorderStyle := xlblsDashDotDot;
    4 : OpExcel1.RangeByName['r'].BorderStyle := xlblsDot;
    5 : OpExcel1.RangeByName['r'].BorderStyle := xlblsDouble;
    6 : OpExcel1.RangeByName['r'].BorderStyle := xlblsLineStyleNone;
    7 : OpExcel1.RangeByName['r'].BorderStyle := xlblsSlantDashDot;
  end;
end;

procedure TExRgPro1.BorderLeftcbClick(Sender: TObject);
begin
  if BorderLeftcb.Checked then
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders + [xlbLeft]
  else
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders - [xlbLeft]
end;

procedure TExRgPro1.BorderRightcbClick(Sender: TObject);
begin
  if BorderRightcb.Checked then
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders + [xlbRight]
  else
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders - [xlbRight]
end;

procedure TExRgPro1.BorderTopcbClick(Sender: TObject);
begin
  if BorderTopcb.Checked then
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders + [xlbTop]
  else
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders - [xlbTop]
end;

procedure TExRgPro1.BorderBottomcbClick(Sender: TObject);
begin
  if BorderBottomcb.Checked then
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders + [xlbBottom]
  else
    OpExcel1.RangeByName['r'].Borders := OpExcel1.RangeByName['r'].Borders - [xlbBottom]
end;

procedure TExRgPro1.ColorBtnClick(Sender: TObject);
begin
  if ColorDialog1.Execute then begin
    FColor := ColorDialog1.Color;
    OpExcel1.RangeByName['r'].Color := FColor;
  end;
end;

procedure TExRgPro1.FontColorBtnClick(Sender: TObject);
begin
  if ColorDialog1.Execute then begin
    FFontColor := ColorDialog1.Color;
    OpExcel1.RangeByName['r'].FontColor := FFontColor;
  end;
end;

procedure TExRgPro1.PatternColorBtnClick(Sender: TObject);
begin
  if ColorDialog1.Execute then begin
    FPatternColor := ColorDialog1.Color;
    OpExcel1.RangeByName['r'].PatternColor := FPatternColor;
  end;
end;

procedure TExRgPro1.FontNameEditExit(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].FontName := FontNameEdit.Text;
end;

procedure TExRgPro1.FontSizeEditExit(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].FontSize := StrToInt(FontSizeEdit.Text);
end;

procedure TExRgPro1.IndentEditExit(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].IndentLevel:= StrToInt(IndentEdit.Text);
end;

procedure TExRgPro1.RotateEditExit(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].RotateDegrees:= StrToInt(RotateEdit.Text);
end;

procedure TExRgPro1.ShrinkToFitcbClick(Sender: TObject);
begin
  if ShrinkToFitCB.Checked then
    OpExcel1.RangeByName['r'].ShrinkToFit := True
  else
    OpExcel1.RangeByName['r'].ShrinkToFit := false;
end;

procedure TExRgPro1.BoldcbClick(Sender: TObject);
begin
  if Boldcb.Checked then
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes + [xlfaBold]
  else
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes - [xlfaBold]
end;

procedure TExRgPro1.ItaliccbClick(Sender: TObject);
begin
  if Italiccb.Checked then
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes + [xlfaItalic]
  else
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes - [xlfaItalic]
end;

procedure TExRgPro1.UnderlinecbClick(Sender: TObject);
begin
  if Underlinecb.Checked then
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes + [xlfaUnderline]
  else
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes - [xlfaUnderline]
end;

procedure TExRgPro1.StrikethroughcbClick(Sender: TObject);
begin
  if Strikethroughcb.Checked then
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes + [xlfaStrikethrough]
  else
    OpExcel1.RangeByName['r'].FontAttributes := OpExcel1.RangeByName['r'].FontAttributes - [xlfaStrikethrough]
end;

procedure TExRgPro1.PatterncbChange(Sender: TObject);
begin
  case Patterncb.ItemIndex of
    0 : OpExcel1.RangeByName['r'].Pattern := xlipPatternAutomatic;
    1 : OpExcel1.RangeByName['r'].Pattern := xlipPatternChecker;
    2 : OpExcel1.RangeByName['r'].Pattern := xlipPatternCrissCross;
    3 : OpExcel1.RangeByName['r'].Pattern := xlipPatternDown;
    4 : OpExcel1.RangeByName['r'].Pattern := xlipPatternGray16;
    5 : OpExcel1.RangeByName['r'].Pattern := xlipPatternGray25;
    6 : OpExcel1.RangeByName['r'].Pattern := xlipPatternGray50;
    7 : OpExcel1.RangeByName['r'].Pattern := xlipPatternGray75;
    8 : OpExcel1.RangeByName['r'].Pattern := xlipPatternGray8;
    9 : OpExcel1.RangeByName['r'].Pattern := xlipPatternGrid;
    10 : OpExcel1.RangeByName['r'].Pattern := xlipPatternHorizontal;
    11 : OpExcel1.RangeByName['r'].Pattern := xlipPatternLightDown;
    12 : OpExcel1.RangeByName['r'].Pattern := xlipPatternLightHorizontal;
    13 : OpExcel1.RangeByName['r'].Pattern := xlipPatternLightUp;
    14 : OpExcel1.RangeByName['r'].Pattern := xlipPatternLightVertical;
    15 : OpExcel1.RangeByName['r'].Pattern := xlipPatternNone;
    16 : OpExcel1.RangeByName['r'].Pattern := xlipPatternSemiGray75;
    17 : OpExcel1.RangeByName['r'].Pattern := xlipPatternSolid;
    18 : OpExcel1.RangeByName['r'].Pattern := xlipPatternUp;
    19 : OpExcel1.RangeByName['r'].Pattern := xlipPatternVertical;
  end;
end;

procedure TExRgPro1.hAligncbChange(Sender: TObject);
begin
  case hAligncb.ItemIndex of
    0 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaGeneral;
    1 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaLeft;
    2 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaCenter;
    3 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaRight;
    4 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaFill;
    5 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaJustify;
    6 : OpExcel1.RangeByName['r'].HorizontalAlignment := xlchaCenterAcrossSelection;
  end;
end;

procedure TExRgPro1.vAligncbChange(Sender: TObject);
begin
  case vAligncb.ItemIndex of
    0 : OpExcel1.RangeByName['r'].VerticalAlignment := xlcvaTop;
    1 : OpExcel1.RangeByName['r'].VerticalAlignment := xlcvaCenter;
    2 : OpExcel1.RangeByName['r'].VerticalAlignment := xlcvaBottom;
    3 : OpExcel1.RangeByName['r'].VerticalAlignment := xlcvaJustify;
    4 : OpExcel1.RangeByName['r'].VerticalAlignment := xlcvaDistributed;
  end;
end;

procedure TExRgPro1.OrientationcbChange(Sender: TObject);
begin
  case Orientationcb.ItemIndex of
    0 : OpExcel1.RangeByName['r'].Orientation := xlcoDownward;
    1 : OpExcel1.RangeByName['r'].Orientation := xlcoHorizontal;
    2 : OpExcel1.RangeByName['r'].Orientation := xlcoUpward;
    3 : OpExcel1.RangeByName['r'].Orientation := xlcoVertical;
    4 : OpExcel1.RangeByName['r'].Orientation := xlcoRotated;
  end;
end;

procedure TExRgPro1.ClearAllBtnClick(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].Clear;
end;

procedure TExRgPro1.ColWidthTBChange(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].ColumnWidth := ColWidthTB.Position;
end;

procedure TExRgPro1.RowHeightTbChange(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].RowHeight := RowHeightTB.Position;
end;

procedure TExRgPro1.AutoFitColBtnClick(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].AutoFitColumns;
end;

procedure TExRgPro1.AutoFitRowBtnClick(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].AutoFitRows;
end;

procedure TExRgPro1.WrapTextcbClick(Sender: TObject);
begin
  if WrapTextcb.Checked then
    OpExcel1.RangeByName['r'].WrapText := True
  else
    OpExcel1.RangeByName['r'].WrapText := False;
end;

procedure TExRgPro1.SimpleTextEditChange(Sender: TObject);
begin
  OpExcel1.RangeByName['r'].SimpleText := SimpleTextEdit.Text;
end;

procedure TExRgPro1.SetAddressBtnClick(Sender: TObject);
begin
  ClearAllBtnClick(self);
  OpExcel1.RangeByName['r'].Address := AddressEdit.Text;
  UpdateAll;
end;

procedure TExRgPro1.UpdateAll;
begin
  SimpleTextEditChange(nil);
  WrapTextcbClick(nil);
  RowHeightTbChange(nil);
  ColWidthTBChange(nil);
  OrientationcbChange(nil);
  vAligncbChange(nil);
  hAligncbChange(nil);
  PatterncbChange(nil);
  StrikethroughcbClick(nil);
  UnderlinecbClick(nil);
  ItaliccbClick(nil);
  BoldcbClick(nil);
  ShrinkToFitcbClick(nil);
  RotateEditExit(nil);
  IndentEditExit(nil);
  FontSizeEditExit(nil);
  FontNameEditExit(nil);
  BorderBottomcbClick(nil);
  BorderTopcbClick(nil);
  BorderRightcbClick(nil);
  BorderLeftcbClick(nil);
  BorderStyleCBChange(nil);
  BorderWeightCBChange(nil);
  OpExcel1.RangeByName['r'].Color := FColor;
  OpExcel1.RangeByName['r'].FontColor := FFontColor;
  OpExcel1.RangeByName['r'].PatternColor := FPatternColor;
  AutoFitColBtnClick(nil);
  AutoFitRowBtnClick(nil);
end;

end.
