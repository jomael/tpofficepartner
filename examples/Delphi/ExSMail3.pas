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

unit ExSMail3;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, OpShared, OpOlkXP, OpOutlk;

type
  {: Enum for recipient routing: To. CC, Bcc. Used internal to this file only }
  TRecipientType = (rtTo, rtCC, rtBCC);
  {: Simple object wrap an AddressEntry interface since interfaces can not be
     reliably stored in pointer lists without manually handling _Addref and _Release }
  TAddressEntryWrapper = class( TObject )
  private
    fAddress: AddressEntry;
  public
    {: Return a duplicate of self with Address assigned }
    function Clone: TAddressEntryWrapper;
    property Address: AddressEntry read fAddress write fAddress;
  end;

  TForm3 = class(TForm)
    Panel1: TPanel;
    Bevel1: TBevel;
    Panel2: TPanel;
    Bevel2: TBevel;
    Panel3: TPanel;
    Label1: TLabel;
    cbAddressLists: TComboBox;
    Label2: TLabel;
    lbNames: TListBox;
    btnTo: TButton;
    edtName: TEdit;
    btnCC: TButton;
    btnBcc: TButton;
    Label3: TLabel;
    btnOK: TButton;
    btnCancel: TButton;
    lbTo: TListBox;
    lbCC: TListBox;
    lbBcc: TListBox;
    procedure FormShow(Sender: TObject);
    procedure cbAddressListsChange(Sender: TObject);
    procedure edtNameChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnToClick(Sender: TObject);
    procedure lbToDblClick(Sender: TObject);
    procedure lbNamesDblClick(Sender: TObject);
    procedure lbToKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    fOutlook: TOpOutlook;
    procedure AddCurrentToRecipients( Which: TRecipientType );
    {: clear all items and free all objects in the passed list box }
    procedure FlushListBox( aListBox: TListBox );
    {: clear all items and free all objects from the Names ListBox }
    procedure FlushNamesListBox;
    {: clear all items and free all objects from the To, CC, & BCC ListBoxes }
    procedure FlushRecipientListBoxes;
    {: Accessor for RecipientTo, RecipientCC, & RecipientBCC }
    function GetRecipient( Index: integer ): string;
    {: Set up initial state of all controls }
    procedure InitControls( bEnabled: Boolean );
  public
    property Outlook: TOpOutlook read fOutlook write fOutlook;
    //: semicolon delimited string containing the "BCC" recipients' email addresses
    property RecipientBCC: string index 2 read GetRecipient;
    //: semicolon delimited string containing the "CC" recipients' email addresses
    property RecipientCC: string index 1 read GetRecipient;
    //: semicolon delimited string containing the "To" recipients' email addresses
    property RecipientTo: string index 0 read GetRecipient;
  end;

var
  Form3: TForm3;

implementation

{$R *.DFM}

{ TAddressEntryWrapper }

function TAddressEntryWrapper.Clone: TAddressEntryWrapper;
begin
  result := TAddressEntryWrapper.Create;
  result.Address := Address;
end;

{ TdlgSelectNames }

procedure TForm3.AddCurrentToRecipients(Which: TRecipientType);
var
  oListBox: TListBox;
begin
  case Which of
    rtTo  : oListBox := lbTo;
    rtCC  : oListBox := lbCC;
    rtBCC : oListBox := lbBcc;
    else
      oListBox := lbTo;
  end;
  if lbNames.ItemIndex >= 0 then
    oListBox.Items.AddObject( lbNames.Items[lbNames.ItemIndex],
       TAddressEntryWrapper(lbNames.Items.Objects[lbNames.ItemIndex]).Clone );
end;

procedure TForm3.btnToClick(Sender: TObject);
begin
  if Sender is TButton then
    AddCurrentToRecipients( TRecipientType(TButton(Sender).Tag ))
end;

procedure TForm3.FormShow(Sender: TObject);
begin
  if assigned( fOutlook ) and ( fOutlook.Connected )then
    InitControls( True )
  else
    InitControls( False );
end;

procedure TForm3.cbAddressListsChange(Sender: TObject);
var
  idx: integer;
  oAddressList: AddressList;
  i: Integer;
  oAddress: TAddressEntryWrapper;
begin
  cbAddressLists.Update;
  FlushNamesListBox;
  idx := cbAddressLists.Items.IndexOf( cbAddressLists.Text );
  if idx >= 0 then
  begin
    oAddressList := fOutlook.MapiNamespace.AddressLists.Item(idx + 1);
    if assigned( oAddressList ) then
      if oAddressList.AddressEntries.Count > 0 then
      begin
        lbNames.Items.BeginUpdate;
        try
          for i := 1 to oAddressList.AddressEntries.Count do
          begin
            oAddress := TAddressEntryWrapper.Create;
            oAddress.Address := oAddressList.AddressEntries.Item(i);
            lbNames.Items.AddObject( oAddress.Address.Name, oAddress )
          end
        finally
          lbNames.Items.EndUpdate;
        end;
      end
  end;
end;

procedure TForm3.edtNameChange(Sender: TObject);
var
  i: Integer;
begin
  if lbNames.Items.Count > 0 then
  begin
    for i := 0 to lbNames.Items.Count do    { Iterate }
      if CompareText( edtName.Text, Copy(lbNames.Items[i], 1, Length(edtName.Text)) ) = 0  then
      begin
        lbNames.ItemIndex := i;
        break;
      end;
  end;
end;

procedure TForm3.FlushListBox( aListBox: TListBox );
var
  i: Integer;
begin
  if aListBox.Items.Count > 0 then
  begin
    for i := aListBox.Items.Count - 1 downto 0 do
      if assigned( aListBox.Items.Objects[i] ) then
        aListBox.Items.Objects[i].Free;
    aListBox.Items.Clear;
  end;
end;

procedure TForm3.FlushNamesListBox;
begin
  FlushListBox( lbNames );
end;

procedure TForm3.FlushRecipientListBoxes;
begin
  FlushListBox(lbTo);
  FlushListBox(lbCC);
  FlushListBox(lbBCC);
end;

procedure TForm3.FormDestroy(Sender: TObject);
begin
  FlushNamesListBox;
  FlushRecipientListBoxes;
end;

function TForm3.GetRecipient(Index: integer): string;
var
  i: Integer;
  oListBox: TListBox;
begin
  result := ''; // this must be done because otherwise result acts like a static
  case Index of    { }
    0: oListBox := lbTo;
    1: oListBox := lbCC;
    2: oListBox := lbBCC;
    else
      oListBox := lbTo;
  end;    { case }
  if oListBox.Items.Count > 0 then
    for i := 0 to oListBox.Items.Count - 1 do
      if assigned( oListBox.Items.Objects[i] ) then
        result := result + TAddressEntryWrapper(oListBox.Items.Objects[i]).
                     Address.Address + ';';
  if Length(result) > 0 then
    Delete(result, Length(Result), 1 );
end;

procedure TForm3.InitControls(bEnabled: Boolean);
var
  i: Integer;
begin
  for i := 0 to ControlCount -1  do
    Controls[i].Enabled := bEnabled;
  cbAddressLists.Clear;
  edtName.Clear;
  if bEnabled then
    if fOutlook.MapiNamespace.AddressLists.Count > 0 then
      for i := 1 to fOutlook.MapiNamespace.AddressLists.Count do
        cbAddressLists.Items.Add(fOutlook.MapiNamespace.AddressLists.Item(i).Name );
  FlushNamesListBox;
  FlushRecipientListBoxes;
end;

procedure TForm3.lbNamesDblClick(Sender: TObject);
begin
  AddCurrentToRecipients(rtTo);
end;

procedure TForm3.lbToDblClick(Sender: TObject);
begin
  if Sender is TListBox then
    with TListBox(Sender) do
      if ItemIndex >= 0 then
      begin
        Items.Objects[ItemIndex].Free;
        Items.Delete(ItemIndex)
      end;
end;

//: type ahead search
procedure TForm3.lbToKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Sender is TListBox then
    if Key = VK_Delete then
      with TListBox(Sender) do
        if ItemIndex >= 0 then
        begin
          Items.Objects[ItemIndex].Free;
          Items.Delete(ItemIndex)
        end;
end;

end.
