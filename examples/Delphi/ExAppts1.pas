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

unit ExAppts1;
{
 OutlookAppoints sample application
   Connects to Outlook and pulls all appointments from the calendar.
   Fills a list box with the subject of each appointment
   As the user navigates the list box it displays the Start date/time,
     End date/time, and body for the selected appointment
}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OpShared, OpOlkXP, OpOutlk, ExtCtrls, Menus;

type
  TForm1 = class(TForm)
    ListBox1: TListBox;
    Label1: TLabel;
    Panel1: TPanel;
    Label2: TLabel;
    Memo1: TMemo;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    Label3: TLabel;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    Label4: TLabel;
    Label5: TLabel;
    OpOutlook1: TOpOutlook;
    btnPopulate: TButton;
    procedure ListBox1Click(Sender: TObject);
    procedure btnPopulateClick(Sender: TObject);
  private
    { Private declarations }
  protected
    { Protected declarations }
  public
    { Public declarations }

  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.ListBox1Click(Sender: TObject);
var
  oApptItem: TOpAppointmentItem;
  oDateTime: TDateTime;
begin
  if ListBox1.ItemIndex >= 0 then
    if assigned( ListBox1.Items.Objects[ListBox1.ItemIndex] ) then
    begin
      oApptItem := ( TObject(ListBox1.Items.Objects[ListBox1.ItemIndex]) as TOpAppointmentItem);
      Memo1.Text := oApptItem.Body;                // appointment body
      oDateTime := oApptItem.Start;                // get start date/time
      StaticText1.Caption := DateToStr(oDateTime); // show start date
      StaticText2.Caption := TimeToStr(oDateTime); // show start time
      oDateTime := oApptItem.End_;                 // get end date/time
      StaticText3.Caption := DateToStr(oDateTime); // show end date
      StaticText4.Caption := TimeToStr(oDateTime)  // show end time
    end
end;

procedure TForm1.btnPopulateClick(Sender: TObject);
var
  i: Integer;
  oMAPIFolder: MAPIFolder;
  oItems: Items;
  oAppt: _AppointmentItem;      // calendar folder contains these
  oApptItem: TOpAppointmentItem;// Delphi wrapper object for _AppointmentItem
begin
  ListBox1.Clear;
  Memo1.Clear;
  with OpOutlook1 do
  begin
    if not Connected then
      Connected := True;

    oMAPIFolder := OpOutlook1.MAPINameSpace.GetDefaultFolder(olFolderCalendar);
    if assigned( oMAPIFolder ) then
    begin
      oItems := oMAPIFolder.Items;
      if oItems.Count > 0 then
        for i := 1 to oItems.Count do // Items uses a 1 based index
          if oItems.Item(i).QueryInterface( _AppointmentItem, oAppt ) = s_OK then
          begin
            oApptItem := TOpAppointmentItem.Create( oAppt );
            ListBox1.Items.AddObject(oApptItem.Subject, oApptItem )
          end
    end
  end
end;

end.
