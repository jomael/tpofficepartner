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

{$I OPDEFINE.INC}

{$IFDEF CBuilder}
  {$Warnings Off}
{$ENDIF}

unit OpWrd2k;

// ************************************************************************ //
// WARNING                                                                  //
// -------                                                                  //
// The types declared in this file were generated from data read from a     //
// Type Library. If this type library is explicitly or indirectly (via      //
// another type library referring to this type library) re-imported, or the //
// 'Refresh' command of the Type Library Editor activated while editing the //
// Type Library, the contents of this file will be regenerated and all      //
// manual modifications will be lost.                                       //
// ************************************************************************ //

// PASTLWTR : $Revision$
// File generated on 1/14/99 12:41:44 PM from Type Library described below.

// ************************************************************************ //
// Type Lib: C:\Program Files\Microsoft Office\Office\msword9.olb
// IID\LCID: {00020905-0000-0000-C000-000000000046}\0
// Helpfile: C:\Program Files\Microsoft Office\Office\VBAWRD9.CHM
// HelpString: Microsoft Word 9.0 Object Library
// Version:    9.0
// ************************************************************************ //

interface

uses Windows, ActiveX, Classes, Graphics, OleCtrls, StdVCL, 
  OpOfc2K, OpVBId2k, OpShared;

// Base class exists for automatic uses clause inclusion.

type
  TOpWordBase = class(TOpOfficeComponent);

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:      //
//   Type Libraries     : LIBID_xxxx                                    //
//   CoClasses          : CLASS_xxxx                                    //
//   DISPInterfaces     : DIID_xxxx                                     //
//   Non-DISP interfaces: IID_xxxx                                      //
// *********************************************************************//
const
  LIBID_Word: TGUID = '{00020905-0000-0000-C000-000000000046}';
  IID__Application: TGUID = '{00020970-0000-0000-C000-000000000046}';
//  IID__Global: TGUID = '{000209B9-0000-0000-C000-000000000046}';
  IID_FontNames: TGUID = '{0002096F-0000-0000-C000-000000000046}';
  IID_Languages: TGUID = '{0002096E-0000-0000-C000-000000000046}';
  IID_Language: TGUID = '{0002096D-0000-0000-C000-000000000046}';
  IID_Documents: TGUID = '{0002096C-0000-0000-C000-000000000046}';
  IID__Document: TGUID = '{0002096B-0000-0000-C000-000000000046}';
  IID_Template: TGUID = '{0002096A-0000-0000-C000-000000000046}';
  IID_Templates: TGUID = '{000209A2-0000-0000-C000-000000000046}';
  IID_RoutingSlip: TGUID = '{00020969-0000-0000-C000-000000000046}';
  IID_Bookmark: TGUID = '{00020968-0000-0000-C000-000000000046}';
  IID_Bookmarks: TGUID = '{00020967-0000-0000-C000-000000000046}';
  IID_Variable: TGUID = '{00020966-0000-0000-C000-000000000046}';
  IID_Variables: TGUID = '{00020965-0000-0000-C000-000000000046}';
  IID_RecentFile: TGUID = '{00020964-0000-0000-C000-000000000046}';
  IID_RecentFiles: TGUID = '{00020963-0000-0000-C000-000000000046}';
  IID_Window_: TGUID = '{00020962-0000-0000-C000-000000000046}';
  IID_Windows: TGUID = '{00020961-0000-0000-C000-000000000046}';
  IID_Pane: TGUID = '{00020960-0000-0000-C000-000000000046}';
  IID_Panes: TGUID = '{0002095F-0000-0000-C000-000000000046}';
  IID_Range: TGUID = '{0002095E-0000-0000-C000-000000000046}';
  IID_ListFormat: TGUID = '{000209C0-0000-0000-C000-000000000046}';
  IID_Find: TGUID = '{000209B0-0000-0000-C000-000000000046}';
  IID_Replacement: TGUID = '{000209B1-0000-0000-C000-000000000046}';
  IID_Characters: TGUID = '{0002095D-0000-0000-C000-000000000046}';
  IID_Words: TGUID = '{0002095C-0000-0000-C000-000000000046}';
  IID_Sentences: TGUID = '{0002095B-0000-0000-C000-000000000046}';
  IID_Sections: TGUID = '{0002095A-0000-0000-C000-000000000046}';
  IID_Section: TGUID = '{00020959-0000-0000-C000-000000000046}';
  IID_Paragraphs: TGUID = '{00020958-0000-0000-C000-000000000046}';
  IID_Paragraph: TGUID = '{00020957-0000-0000-C000-000000000046}';
  IID_DropCap: TGUID = '{00020956-0000-0000-C000-000000000046}';
  IID_TabStops: TGUID = '{00020955-0000-0000-C000-000000000046}';
  IID_TabStop: TGUID = '{00020954-0000-0000-C000-000000000046}';
  IID__ParagraphFormat: TGUID = '{00020953-0000-0000-C000-000000000046}';
  IID__Font: TGUID = '{00020952-0000-0000-C000-000000000046}';
  IID_Table: TGUID = '{00020951-0000-0000-C000-000000000046}';
  IID_Row: TGUID = '{00020950-0000-0000-C000-000000000046}';
  IID_Column: TGUID = '{0002094F-0000-0000-C000-000000000046}';
  IID_Cell: TGUID = '{0002094E-0000-0000-C000-000000000046}';
  IID_Tables: TGUID = '{0002094D-0000-0000-C000-000000000046}';
  IID_Rows: TGUID = '{0002094C-0000-0000-C000-000000000046}';
  IID_Columns: TGUID = '{0002094B-0000-0000-C000-000000000046}';
  IID_Cells: TGUID = '{0002094A-0000-0000-C000-000000000046}';
  IID_AutoCorrect: TGUID = '{00020949-0000-0000-C000-000000000046}';
  IID_AutoCorrectEntries: TGUID = '{00020948-0000-0000-C000-000000000046}';
  IID_AutoCorrectEntry: TGUID = '{00020947-0000-0000-C000-000000000046}';
  IID_FirstLetterExceptions: TGUID = '{00020946-0000-0000-C000-000000000046}';
  IID_FirstLetterException: TGUID = '{00020945-0000-0000-C000-000000000046}';
  IID_TwoInitialCapsExceptions: TGUID = '{00020944-0000-0000-C000-000000000046}';
  IID_TwoInitialCapsException: TGUID = '{00020943-0000-0000-C000-000000000046}';
  IID_Footnotes: TGUID = '{00020942-0000-0000-C000-000000000046}';
  IID_Endnotes: TGUID = '{00020941-0000-0000-C000-000000000046}';
  IID_Comments: TGUID = '{00020940-0000-0000-C000-000000000046}';
  IID_Footnote: TGUID = '{0002093F-0000-0000-C000-000000000046}';
  IID_Endnote: TGUID = '{0002093E-0000-0000-C000-000000000046}';
  IID_Comment: TGUID = '{0002093D-0000-0000-C000-000000000046}';
  IID_Borders: TGUID = '{0002093C-0000-0000-C000-000000000046}';
  IID_Border: TGUID = '{0002093B-0000-0000-C000-000000000046}';
  IID_Shading: TGUID = '{0002093A-0000-0000-C000-000000000046}';
  IID_TextRetrievalMode: TGUID = '{00020939-0000-0000-C000-000000000046}';
  IID_AutoTextEntries: TGUID = '{00020937-0000-0000-C000-000000000046}';
  IID_AutoTextEntry: TGUID = '{00020936-0000-0000-C000-000000000046}';
  IID_System_: TGUID = '{00020935-0000-0000-C000-000000000046}';
  IID_OLEFormat: TGUID = '{00020933-0000-0000-C000-000000000046}';
  IID_LinkFormat: TGUID = '{00020931-0000-0000-C000-000000000046}';
  IID__OLEControl: TGUID = '{000209A4-0000-0000-C000-000000000046}';
  IID_Fields: TGUID = '{00020930-0000-0000-C000-000000000046}';
  IID_Field: TGUID = '{0002092F-0000-0000-C000-000000000046}';
  IID_Browser: TGUID = '{0002092E-0000-0000-C000-000000000046}';
  IID_Styles: TGUID = '{0002092D-0000-0000-C000-000000000046}';
  IID_Style: TGUID = '{0002092C-0000-0000-C000-000000000046}';
  IID_Frames: TGUID = '{0002092B-0000-0000-C000-000000000046}';
  IID_Frame: TGUID = '{0002092A-0000-0000-C000-000000000046}';
  IID_FormFields: TGUID = '{00020929-0000-0000-C000-000000000046}';
  IID_FormField: TGUID = '{00020928-0000-0000-C000-000000000046}';
  IID_TextInput: TGUID = '{00020927-0000-0000-C000-000000000046}';
  IID_CheckBox: TGUID = '{00020926-0000-0000-C000-000000000046}';
  IID_DropDown: TGUID = '{00020925-0000-0000-C000-000000000046}';
  IID_ListEntries: TGUID = '{00020924-0000-0000-C000-000000000046}';
  IID_ListEntry: TGUID = '{00020923-0000-0000-C000-000000000046}';
  IID_TablesOfFigures: TGUID = '{00020922-0000-0000-C000-000000000046}';
  IID_TableOfFigures: TGUID = '{00020921-0000-0000-C000-000000000046}';
  IID_MailMerge: TGUID = '{00020920-0000-0000-C000-000000000046}';
  IID_MailMergeFields: TGUID = '{0002091F-0000-0000-C000-000000000046}';
  IID_MailMergeField: TGUID = '{0002091E-0000-0000-C000-000000000046}';
  IID_MailMergeDataSource: TGUID = '{0002091D-0000-0000-C000-000000000046}';
  IID_MailMergeFieldNames: TGUID = '{0002091C-0000-0000-C000-000000000046}';
  IID_MailMergeFieldName: TGUID = '{0002091B-0000-0000-C000-000000000046}';
  IID_MailMergeDataFields: TGUID = '{0002091A-0000-0000-C000-000000000046}';
  IID_MailMergeDataField: TGUID = '{00020919-0000-0000-C000-000000000046}';
  IID_Envelope: TGUID = '{00020918-0000-0000-C000-000000000046}';
  IID_MailingLabel: TGUID = '{00020917-0000-0000-C000-000000000046}';
  IID_CustomLabels: TGUID = '{00020916-0000-0000-C000-000000000046}';
  IID_CustomLabel: TGUID = '{00020915-0000-0000-C000-000000000046}';
  IID_TablesOfContents: TGUID = '{00020914-0000-0000-C000-000000000046}';
  IID_TableOfContents: TGUID = '{00020913-0000-0000-C000-000000000046}';
  IID_TablesOfAuthorities: TGUID = '{00020912-0000-0000-C000-000000000046}';
  IID_TableOfAuthorities: TGUID = '{00020911-0000-0000-C000-000000000046}';
  IID_Dialogs: TGUID = '{00020910-0000-0000-C000-000000000046}';
  IID_Dialog: TGUID = '{000209B8-0000-0000-C000-000000000046}';
  IID_PageSetup: TGUID = '{00020971-0000-0000-C000-000000000046}';
  IID_LineNumbering: TGUID = '{00020972-0000-0000-C000-000000000046}';
  IID_TextColumns: TGUID = '{00020973-0000-0000-C000-000000000046}';
  IID_TextColumn: TGUID = '{00020974-0000-0000-C000-000000000046}';
  IID_Selection: TGUID = '{00020975-0000-0000-C000-000000000046}';
  IID_TablesOfAuthoritiesCategories: TGUID = '{00020976-0000-0000-C000-000000000046}';
  IID_TableOfAuthoritiesCategory: TGUID = '{00020977-0000-0000-C000-000000000046}';
  IID_CaptionLabels: TGUID = '{00020978-0000-0000-C000-000000000046}';
  IID_CaptionLabel: TGUID = '{00020979-0000-0000-C000-000000000046}';
  IID_AutoCaptions: TGUID = '{0002097A-0000-0000-C000-000000000046}';
  IID_AutoCaption: TGUID = '{0002097B-0000-0000-C000-000000000046}';
  IID_Indexes: TGUID = '{0002097C-0000-0000-C000-000000000046}';
  IID_Index: TGUID = '{0002097D-0000-0000-C000-000000000046}';
  IID_AddIn: TGUID = '{0002097E-0000-0000-C000-000000000046}';
  IID_AddIns: TGUID = '{0002097F-0000-0000-C000-000000000046}';
  IID_Revisions: TGUID = '{00020980-0000-0000-C000-000000000046}';
  IID_Revision: TGUID = '{00020981-0000-0000-C000-000000000046}';
  IID_Task: TGUID = '{00020982-0000-0000-C000-000000000046}';
  IID_Tasks: TGUID = '{00020983-0000-0000-C000-000000000046}';
  IID_HeadersFooters: TGUID = '{00020984-0000-0000-C000-000000000046}';
  IID_HeaderFooter: TGUID = '{00020985-0000-0000-C000-000000000046}';
  IID_PageNumbers: TGUID = '{00020986-0000-0000-C000-000000000046}';
  IID_PageNumber: TGUID = '{00020987-0000-0000-C000-000000000046}';
  IID_Subdocuments: TGUID = '{00020988-0000-0000-C000-000000000046}';
  IID_Subdocument: TGUID = '{00020989-0000-0000-C000-000000000046}';
  IID_HeadingStyles: TGUID = '{0002098A-0000-0000-C000-000000000046}';
  IID_HeadingStyle: TGUID = '{0002098B-0000-0000-C000-000000000046}';
  IID_StoryRanges: TGUID = '{0002098C-0000-0000-C000-000000000046}';
  IID_ListLevel: TGUID = '{0002098D-0000-0000-C000-000000000046}';
  IID_ListLevels: TGUID = '{0002098E-0000-0000-C000-000000000046}';
  IID_ListTemplate: TGUID = '{0002098F-0000-0000-C000-000000000046}';
  IID_ListTemplates: TGUID = '{00020990-0000-0000-C000-000000000046}';
  IID_ListParagraphs: TGUID = '{00020991-0000-0000-C000-000000000046}';
  IID_List: TGUID = '{00020992-0000-0000-C000-000000000046}';
  IID_Lists: TGUID = '{00020993-0000-0000-C000-000000000046}';
  IID_ListGallery: TGUID = '{00020994-0000-0000-C000-000000000046}';
  IID_ListGalleries: TGUID = '{00020995-0000-0000-C000-000000000046}';
  IID_KeyBindings: TGUID = '{00020996-0000-0000-C000-000000000046}';
  IID_KeysBoundTo: TGUID = '{00020997-0000-0000-C000-000000000046}';
  IID_KeyBinding: TGUID = '{00020998-0000-0000-C000-000000000046}';
  IID_FileConverter: TGUID = '{00020999-0000-0000-C000-000000000046}';
  IID_FileConverters: TGUID = '{0002099A-0000-0000-C000-000000000046}';
  IID_SynonymInfo: TGUID = '{0002099B-0000-0000-C000-000000000046}';
  IID_Hyperlinks: TGUID = '{0002099C-0000-0000-C000-000000000046}';
  IID_Hyperlink: TGUID = '{0002099D-0000-0000-C000-000000000046}';
  IID_Shapes: TGUID = '{0002099F-0000-0000-C000-000000000046}';
  IID_ShapeRange: TGUID = '{000209B5-0000-0000-C000-000000000046}';
  IID_GroupShapes: TGUID = '{000209B6-0000-0000-C000-000000000046}';
  IID_Shape: TGUID = '{000209A0-0000-0000-C000-000000000046}';
  IID_TextFrame: TGUID = '{000209B2-0000-0000-C000-000000000046}';
  IID__LetterContent: TGUID = '{000209A1-0000-0000-C000-000000000046}';
  IID_View: TGUID = '{000209A5-0000-0000-C000-000000000046}';
  IID_Zoom: TGUID = '{000209A6-0000-0000-C000-000000000046}';
  IID_Zooms: TGUID = '{000209A7-0000-0000-C000-000000000046}';
  IID_InlineShape: TGUID = '{000209A8-0000-0000-C000-000000000046}';
  IID_InlineShapes: TGUID = '{000209A9-0000-0000-C000-000000000046}';
  IID_SpellingSuggestions: TGUID = '{000209AA-0000-0000-C000-000000000046}';
  IID_SpellingSuggestion: TGUID = '{000209AB-0000-0000-C000-000000000046}';
  IID_Dictionaries: TGUID = '{000209AC-0000-0000-C000-000000000046}';
  IID_HangulHanjaConversionDictionaries: TGUID = '{000209E0-0000-0000-C000-000000000046}';
  IID_Dictionary: TGUID = '{000209AD-0000-0000-C000-000000000046}';
  IID_ReadabilityStatistics: TGUID = '{000209AE-0000-0000-C000-000000000046}';
  IID_ReadabilityStatistic: TGUID = '{000209AF-0000-0000-C000-000000000046}';
  IID_Versions: TGUID = '{000209B3-0000-0000-C000-000000000046}';
  IID_Version: TGUID = '{000209B4-0000-0000-C000-000000000046}';
  IID_Options: TGUID = '{000209B7-0000-0000-C000-000000000046}';
  IID_MailMessage: TGUID = '{000209BA-0000-0000-C000-000000000046}';
  IID_ProofreadingErrors: TGUID = '{000209BB-0000-0000-C000-000000000046}';
  IID_Mailer: TGUID = '{000209BD-0000-0000-C000-000000000046}';
  IID_WrapFormat: TGUID = '{000209C3-0000-0000-C000-000000000046}';
  IID_HangulAndAlphabetExceptions: TGUID = '{000209D1-0000-0000-C000-000000000046}';
  IID_HangulAndAlphabetException: TGUID = '{000209D2-0000-0000-C000-000000000046}';
  IID_Adjustments: TGUID = '{000209C4-0000-0000-C000-000000000046}';
  IID_CalloutFormat: TGUID = '{000209C5-0000-0000-C000-000000000046}';
  IID_ColorFormat: TGUID = '{000209C6-0000-0000-C000-000000000046}';
  IID_ConnectorFormat: TGUID = '{000209C7-0000-0000-C000-000000000046}';
  IID_FillFormat: TGUID = '{000209C8-0000-0000-C000-000000000046}';
  IID_FreeformBuilder: TGUID = '{000209C9-0000-0000-C000-000000000046}';
  IID_LineFormat: TGUID = '{000209CA-0000-0000-C000-000000000046}';
  IID_PictureFormat: TGUID = '{000209CB-0000-0000-C000-000000000046}';
  IID_ShadowFormat: TGUID = '{000209CC-0000-0000-C000-000000000046}';
  IID_ShapeNode: TGUID = '{000209CD-0000-0000-C000-000000000046}';
  IID_ShapeNodes: TGUID = '{000209CE-0000-0000-C000-000000000046}';
  IID_TextEffectFormat: TGUID = '{000209CF-0000-0000-C000-000000000046}';
  IID_ThreeDFormat: TGUID = '{000209D0-0000-0000-C000-000000000046}';
  DIID_ApplicationEvents: TGUID = '{000209F7-0000-0000-C000-000000000046}';
//  CLASS_Global: TGUID = '{000209F0-0000-0000-C000-000000000046}';
  DIID_ApplicationEvents2: TGUID = '{000209FE-0000-0000-C000-000000000046}';
  DIID_DocumentEvents: TGUID = '{000209F6-0000-0000-C000-000000000046}';
  CLASS_Document: TGUID = '{00020906-0000-0000-C000-000000000046}';
  CLASS_Font: TGUID = '{000209F5-0000-0000-C000-000000000046}';
  CLASS_ParagraphFormat: TGUID = '{000209F4-0000-0000-C000-000000000046}';
  DIID_OCXEvents: TGUID = '{000209F3-0000-0000-C000-000000000046}';
  CLASS_OLEControl: TGUID = '{000209F2-0000-0000-C000-000000000046}';
  CLASS_LetterContent: TGUID = '{000209F1-0000-0000-C000-000000000046}';
  IID_IApplicationEvents: TGUID = '{000209F7-0001-0000-C000-000000000046}';
  IID_IApplicationEvents2: TGUID = '{000209FE-0001-0000-C000-000000000046}';
  CLASS_Application_: TGUID = '{000209FF-0000-0000-C000-000000000046}';
  IID_EmailAuthor: TGUID = '{000209D7-0000-0000-C000-000000000046}';
  IID_EmailOptions: TGUID = '{000209DB-0000-0000-C000-000000000046}';
  IID_EmailSignature: TGUID = '{000209DC-0000-0000-C000-000000000046}';
  IID_Email: TGUID = '{000209DD-0000-0000-C000-000000000046}';
  IID_HorizontalLineFormat: TGUID = '{000209DE-0000-0000-C000-000000000046}';
  IID_Frameset: TGUID = '{000209E2-0000-0000-C000-000000000046}';
  IID_DefaultWebOptions: TGUID = '{000209E3-0000-0000-C000-000000000046}';
  IID_WebOptions: TGUID = '{000209E4-0000-0000-C000-000000000046}';
  IID_OtherCorrectionsExceptions: TGUID = '{000209DF-0000-0000-C000-000000000046}';
  IID_OtherCorrectionsException: TGUID = '{000209E1-0000-0000-C000-000000000046}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                  //
// *********************************************************************//
// WdMailSystem constants
type
  WdMailSystem = TOleEnum;
const
  wdNoMailSystem = $00000000;
  wdMAPI = $00000001;
  wdPowerTalk = $00000002;
  wdMAPIandPowerTalk = $00000003;

// WdTemplateType constants
type
  WdTemplateType = TOleEnum;
const
  wdNormalTemplate = $00000000;
  wdGlobalTemplate = $00000001;
  wdAttachedTemplate = $00000002;

// WdContinue constants
type
  WdContinue = TOleEnum;
const
  wdContinueDisabled = $00000000;
  wdResetList = $00000001;
  wdContinueList = $00000002;

// WdIMEMode constants
type
  WdIMEMode = TOleEnum;
const
  wdIMEModeNoControl = $00000000;
  wdIMEModeOn = $00000001;
  wdIMEModeOff = $00000002;
  wdIMEModeHiragana = $00000004;
  wdIMEModeKatakana = $00000005;
  wdIMEModeKatakanaHalf = $00000006;
  wdIMEModeAlphaFull = $00000007;
  wdIMEModeAlpha = $00000008;
  wdIMEModeHangulFull = $00000009;
  wdIMEModeHangul = $0000000A;

// WdBaselineAlignment constants
type
  WdBaselineAlignment = TOleEnum;
const
  wdBaselineAlignTop = $00000000;
  wdBaselineAlignCenter = $00000001;
  wdBaselineAlignBaseline = $00000002;
  wdBaselineAlignFarEast50 = $00000003;
  wdBaselineAlignAuto = $00000004;

// WdIndexFilter constants
type
  WdIndexFilter = TOleEnum;
const
  wdIndexFilterNone = $00000000;
  wdIndexFilterAiueo = $00000001;
  wdIndexFilterAkasatana = $00000002;
  wdIndexFilterChosung = $00000003;
  wdIndexFilterLow = $00000004;
  wdIndexFilterMedium = $00000005;
  wdIndexFilterFull = $00000006;

// WdIndexSortBy constants
type
  WdIndexSortBy = TOleEnum;
const
  wdIndexSortByStroke = $00000000;
  wdIndexSortBySyllable = $00000001;

// WdJustificationMode constants
type
  WdJustificationMode = TOleEnum;
const
  wdJustificationModeExpand = $00000000;
  wdJustificationModeCompress = $00000001;
  wdJustificationModeCompressKana = $00000002;

// WdFarEastLineBreakLevel constants
type
  WdFarEastLineBreakLevel = TOleEnum;
const
  wdFarEastLineBreakLevelNormal = $00000000;
  wdFarEastLineBreakLevelStrict = $00000001;
  wdFarEastLineBreakLevelCustom = $00000002;

// WdMultipleWordConversionsMode constants
type
  WdMultipleWordConversionsMode = TOleEnum;
const
  wdHangulToHanja = $00000000;
  wdHanjaToHangul = $00000001;

// WdColorIndex constants
type
  WdColorIndex = TOleEnum;
const
  wdAuto = $00000000;
  wdBlack = $00000001;
  wdBlue = $00000002;
  wdTurquoise = $00000003;
  wdBrightGreen = $00000004;
  wdPink = $00000005;
  wdRed = $00000006;
  wdYellow = $00000007;
  wdWhite = $00000008;
  wdDarkBlue = $00000009;
  wdTeal = $0000000A;
  wdGreen = $0000000B;
  wdViolet = $0000000C;
  wdDarkRed = $0000000D;
  wdDarkYellow = $0000000E;
  wdGray50 = $0000000F;
  wdGray25 = $00000010;
  wdByAuthor = $FFFFFFFF;
  wdNoHighlight = $00000000;

// WdTextureIndex constants
type
  WdTextureIndex = TOleEnum;
const
  wdTextureNone = $00000000;
  wdTexture2Pt5Percent = $00000019;
  wdTexture5Percent = $00000032;
  wdTexture7Pt5Percent = $0000004B;
  wdTexture10Percent = $00000064;
  wdTexture12Pt5Percent = $0000007D;
  wdTexture15Percent = $00000096;
  wdTexture17Pt5Percent = $000000AF;
  wdTexture20Percent = $000000C8;
  wdTexture22Pt5Percent = $000000E1;
  wdTexture25Percent = $000000FA;
  wdTexture27Pt5Percent = $00000113;
  wdTexture30Percent = $0000012C;
  wdTexture32Pt5Percent = $00000145;
  wdTexture35Percent = $0000015E;
  wdTexture37Pt5Percent = $00000177;
  wdTexture40Percent = $00000190;
  wdTexture42Pt5Percent = $000001A9;
  wdTexture45Percent = $000001C2;
  wdTexture47Pt5Percent = $000001DB;
  wdTexture50Percent = $000001F4;
  wdTexture52Pt5Percent = $0000020D;
  wdTexture55Percent = $00000226;
  wdTexture57Pt5Percent = $0000023F;
  wdTexture60Percent = $00000258;
  wdTexture62Pt5Percent = $00000271;
  wdTexture65Percent = $0000028A;
  wdTexture67Pt5Percent = $000002A3;
  wdTexture70Percent = $000002BC;
  wdTexture72Pt5Percent = $000002D5;
  wdTexture75Percent = $000002EE;
  wdTexture77Pt5Percent = $00000307;
  wdTexture80Percent = $00000320;
  wdTexture82Pt5Percent = $00000339;
  wdTexture85Percent = $00000352;
  wdTexture87Pt5Percent = $0000036B;
  wdTexture90Percent = $00000384;
  wdTexture92Pt5Percent = $0000039D;
  wdTexture95Percent = $000003B6;
  wdTexture97Pt5Percent = $000003CF;
  wdTextureSolid = $000003E8;
  wdTextureDarkHorizontal = $FFFFFFFF;
  wdTextureDarkVertical = $FFFFFFFE;
  wdTextureDarkDiagonalDown = $FFFFFFFD;
  wdTextureDarkDiagonalUp = $FFFFFFFC;
  wdTextureDarkCross = $FFFFFFFB;
  wdTextureDarkDiagonalCross = $FFFFFFFA;
  wdTextureHorizontal = $FFFFFFF9;
  wdTextureVertical = $FFFFFFF8;
  wdTextureDiagonalDown = $FFFFFFF7;
  wdTextureDiagonalUp = $FFFFFFF6;
  wdTextureCross = $FFFFFFF5;
  wdTextureDiagonalCross = $FFFFFFF4;

// WdUnderline constants
type
  WdUnderline = TOleEnum;
const
  wdUnderlineNone = $00000000;
  wdUnderlineSingle = $00000001;
  wdUnderlineWords = $00000002;
  wdUnderlineDouble = $00000003;
  wdUnderlineDotted = $00000004;
  wdUnderlineThick = $00000006;
  wdUnderlineDash = $00000007;
  wdUnderlineDotDash = $00000009;
  wdUnderlineDotDotDash = $0000000A;
  wdUnderlineWavy = $0000000B;
  wdUnderlineWavyHeavy = $0000001B;
  wdUnderlineDottedHeavy = $00000014;
  wdUnderlineDashHeavy = $00000017;
  wdUnderlineDotDashHeavy = $00000019;
  wdUnderlineDotDotDashHeavy = $0000001A;
  wdUnderlineDashLong = $00000027;
  wdUnderlineDashLongHeavy = $00000037;
  wdUnderlineWavyDouble = $0000002B;

// WdEmphasisMark constants
type
  WdEmphasisMark = TOleEnum;
const
  wdEmphasisMarkNone = $00000000;
  wdEmphasisMarkOverSolidCircle = $00000001;
  wdEmphasisMarkOverComma = $00000002;
  wdEmphasisMarkOverWhiteCircle = $00000003;
  wdEmphasisMarkUnderSolidCircle = $00000004;

// WdInternationalIndex constants
type
  WdInternationalIndex = TOleEnum;
const
  wdListSeparator = $00000011;
  wdDecimalSeparator = $00000012;
  wdThousandsSeparator = $00000013;
  wdCurrencyCode = $00000014;
  wd24HourClock = $00000015;
  wdInternationalAM = $00000016;
  wdInternationalPM = $00000017;
  wdTimeSeparator = $00000018;
  wdDateSeparator = $00000019;
  wdProductLanguageID = $0000001A;

// WdAutoMacros constants
type
  WdAutoMacros = TOleEnum;
const
  wdAutoExec = $00000000;
  wdAutoNew = $00000001;
  wdAutoOpen = $00000002;
  wdAutoClose = $00000003;
  wdAutoExit = $00000004;

// WdCaptionPosition constants
type
  WdCaptionPosition = TOleEnum;
const
  wdCaptionPositionAbove = $00000000;
  wdCaptionPositionBelow = $00000001;

// WdCountry constants
type
  WdCountry = TOleEnum;
const
  wdUS = $00000001;
  wdCanada = $00000002;
  wdLatinAmerica = $00000003;
  wdNetherlands = $0000001F;
  wdFrance = $00000021;
  wdSpain = $00000022;
  wdItaly = $00000027;
  wdUK = $0000002C;
  wdDenmark = $0000002D;
  wdSweden = $0000002E;
  wdNorway = $0000002F;
  wdGermany = $00000031;
  wdPeru = $00000033;
  wdMexico = $00000034;
  wdArgentina = $00000036;
  wdBrazil = $00000037;
  wdChile = $00000038;
  wdVenezuela = $0000003A;
  wdJapan = $00000051;
  wdTaiwan = $00000376;
  wdChina = $00000056;
  wdKorea = $00000052;
  wdFinland = $00000166;
  wdIceland = $00000162;

// WdHeadingSeparator constants
type
  WdHeadingSeparator = TOleEnum;
const
  wdHeadingSeparatorNone = $00000000;
  wdHeadingSeparatorBlankLine = $00000001;
  wdHeadingSeparatorLetter = $00000002;
  wdHeadingSeparatorLetterLow = $00000003;
  wdHeadingSeparatorLetterFull = $00000004;

// WdSeparatorType constants
type
  WdSeparatorType = TOleEnum;
const
  wdSeparatorHyphen = $00000000;
  wdSeparatorPeriod = $00000001;
  wdSeparatorColon = $00000002;
  wdSeparatorEmDash = $00000003;
  wdSeparatorEnDash = $00000004;

// WdPageNumberAlignment constants
type
  WdPageNumberAlignment = TOleEnum;
const
  wdAlignPageNumberLeft = $00000000;
  wdAlignPageNumberCenter = $00000001;
  wdAlignPageNumberRight = $00000002;
  wdAlignPageNumberInside = $00000003;
  wdAlignPageNumberOutside = $00000004;

// WdBorderType constants
type
  WdBorderType = TOleEnum;
const
  wdBorderTop = $FFFFFFFF;
  wdBorderLeft = $FFFFFFFE;
  wdBorderBottom = $FFFFFFFD;
  wdBorderRight = $FFFFFFFC;
  wdBorderHorizontal = $FFFFFFFB;
  wdBorderVertical = $FFFFFFFA;
  wdBorderDiagonalDown = $FFFFFFF9;
  wdBorderDiagonalUp = $FFFFFFF8;

// WdBorderTypeHID constants
type
  WdBorderTypeHID = TOleEnum;
const
  emptyenum = $00000000;

// WdFramePosition constants
type
  WdFramePosition = TOleEnum;
const
  wdFrameTop = $FFF0BDC1;
  wdFrameLeft = $FFF0BDC2;
  wdFrameBottom = $FFF0BDC3;
  wdFrameRight = $FFF0BDC4;
  wdFrameCenter = $FFF0BDC5;
  wdFrameInside = $FFF0BDC6;
  wdFrameOutside = $FFF0BDC7;

// WdAnimation constants
type
  WdAnimation = TOleEnum;
const
  wdAnimationNone = $00000000;
  wdAnimationLasVegasLights = $00000001;
  wdAnimationBlinkingBackground = $00000002;
  wdAnimationSparkleText = $00000003;
  wdAnimationMarchingBlackAnts = $00000004;
  wdAnimationMarchingRedAnts = $00000005;
  wdAnimationShimmer = $00000006;

// WdCharacterCase constants
type
  WdCharacterCase = TOleEnum;
const
  wdNextCase = $FFFFFFFF;
  wdLowerCase = $00000000;
  wdUpperCase = $00000001;
  wdTitleWord = $00000002;
  wdTitleSentence = $00000004;
  wdToggleCase = $00000005;
  wdHalfWidth = $00000006;
  wdFullWidth = $00000007;
  wdKatakana = $00000008;
  wdHiragana = $00000009;

// WdCharacterCaseHID constants
type
  WdCharacterCaseHID = TOleEnum;
const
  emptyenum_ = $00000000;

// WdSummaryMode constants
type
  WdSummaryMode = TOleEnum;
const
  wdSummaryModeHighlight = $00000000;
  wdSummaryModeHideAllButSummary = $00000001;
  wdSummaryModeInsert = $00000002;
  wdSummaryModeCreateNew = $00000003;

// WdSummaryLength constants
type
  WdSummaryLength = TOleEnum;
const
  wd10Sentences = $FFFFFFFE;
  wd20Sentences = $FFFFFFFD;
  wd100Words = $FFFFFFFC;
  wd500Words = $FFFFFFFB;
  wd10Percent = $FFFFFFFA;
  wd25Percent = $FFFFFFF9;
  wd50Percent = $FFFFFFF8;
  wd75Percent = $FFFFFFF7;

// WdStyleType constants
type
  WdStyleType = TOleEnum;
const
  wdStyleTypeParagraph = $00000001;
  wdStyleTypeCharacter = $00000002;

// WdUnits constants
type
  WdUnits = TOleEnum;
const
  wdCharacter = $00000001;
  wdWord = $00000002;
  wdSentence = $00000003;
  wdParagraph = $00000004;
  wdLine = $00000005;
  wdStory = $00000006;
  wdScreen = $00000007;
  wdSection = $00000008;
  wdColumn = $00000009;
  wdRow = $0000000A;
  wdWindow = $0000000B;
  wdCell = $0000000C;
  wdCharacterFormatting = $0000000D;
  wdParagraphFormatting = $0000000E;
  wdTable = $0000000F;
  wdItem = $00000010;

// WdGoToItem constants
type
  WdGoToItem = TOleEnum;
const
  wdGoToBookmark = $FFFFFFFF;
  wdGoToSection = $00000000;
  wdGoToPage = $00000001;
  wdGoToTable = $00000002;
  wdGoToLine = $00000003;
  wdGoToFootnote = $00000004;
  wdGoToEndnote = $00000005;
  wdGoToComment = $00000006;
  wdGoToField = $00000007;
  wdGoToGraphic = $00000008;
  wdGoToObject = $00000009;
  wdGoToEquation = $0000000A;
  wdGoToHeading = $0000000B;
  wdGoToPercent = $0000000C;
  wdGoToSpellingError = $0000000D;
  wdGoToGrammaticalError = $0000000E;
  wdGoToProofreadingError = $0000000F;

// WdGoToDirection constants
type
  WdGoToDirection = TOleEnum;
const
  wdGoToFirst = $00000001;
  wdGoToLast = $FFFFFFFF;
  wdGoToNext = $00000002;
  wdGoToRelative = $00000002;
  wdGoToPrevious = $00000003;
  wdGoToAbsolute = $00000001;

// WdCollapseDirection constants
type
  WdCollapseDirection = TOleEnum;
const
  wdCollapseStart = $00000001;
  wdCollapseEnd = $00000000;

// WdRowHeightRule constants
type
  WdRowHeightRule = TOleEnum;
const
  wdRowHeightAuto = $00000000;
  wdRowHeightAtLeast = $00000001;
  wdRowHeightExactly = $00000002;

// WdFrameSizeRule constants
type
  WdFrameSizeRule = TOleEnum;
const
  wdFrameAuto = $00000000;
  wdFrameAtLeast = $00000001;
  wdFrameExact = $00000002;

// WdInsertCells constants
type
  WdInsertCells = TOleEnum;
const
  wdInsertCellsShiftRight = $00000000;
  wdInsertCellsShiftDown = $00000001;
  wdInsertCellsEntireRow = $00000002;
  wdInsertCellsEntireColumn = $00000003;

// WdDeleteCells constants
type
  WdDeleteCells = TOleEnum;
const
  wdDeleteCellsShiftLeft = $00000000;
  wdDeleteCellsShiftUp = $00000001;
  wdDeleteCellsEntireRow = $00000002;
  wdDeleteCellsEntireColumn = $00000003;

// WdListApplyTo constants
type
  WdListApplyTo = TOleEnum;
const
  wdListApplyToWholeList = $00000000;
  wdListApplyToThisPointForward = $00000001;
  wdListApplyToSelection = $00000002;

// WdAlertLevel constants
type
  WdAlertLevel = TOleEnum;
const
  wdAlertsNone = $00000000;
  wdAlertsMessageBox = $FFFFFFFE;
  wdAlertsAll = $FFFFFFFF;

// WdCursorType constants
type
  WdCursorType = TOleEnum;
const
  wdCursorWait = $00000000;
  wdCursorIBeam = $00000001;
  wdCursorNormal = $00000002;
  wdCursorNorthwestArrow = $00000003;

// WdEnableCancelKey constants
type
  WdEnableCancelKey = TOleEnum;
const
  wdCancelDisabled = $00000000;
  wdCancelInterrupt = $00000001;

// WdRulerStyle constants
type
  WdRulerStyle = TOleEnum;
const
  wdAdjustNone = $00000000;
  wdAdjustProportional = $00000001;
  wdAdjustFirstColumn = $00000002;
  wdAdjustSameWidth = $00000003;

// WdParagraphAlignment constants
type
  WdParagraphAlignment = TOleEnum;
const
  wdAlignParagraphLeft = $00000000;
  wdAlignParagraphCenter = $00000001;
  wdAlignParagraphRight = $00000002;
  wdAlignParagraphJustify = $00000003;
  wdAlignParagraphDistribute = $00000004;
  wdAlignParagraphJustifyMed = $00000005;
  wdAlignParagraphJustifyHi = $00000007;
  wdAlignParagraphJustifyLow = $00000008;

// WdParagraphAlignmentHID constants
type
  WdParagraphAlignmentHID = TOleEnum;
const
  emptyenum__ = $00000000;

// WdListLevelAlignment constants
type
  WdListLevelAlignment = TOleEnum;
const
  wdListLevelAlignLeft = $00000000;
  wdListLevelAlignCenter = $00000001;
  wdListLevelAlignRight = $00000002;

// WdRowAlignment constants
type
  WdRowAlignment = TOleEnum;
const
  wdAlignRowLeft = $00000000;
  wdAlignRowCenter = $00000001;
  wdAlignRowRight = $00000002;

// WdTabAlignment constants
type
  WdTabAlignment = TOleEnum;
const
  wdAlignTabLeft = $00000000;
  wdAlignTabCenter = $00000001;
  wdAlignTabRight = $00000002;
  wdAlignTabDecimal = $00000003;
  wdAlignTabBar = $00000004;
  wdAlignTabList = $00000006;

// WdVerticalAlignment constants
type
  WdVerticalAlignment = TOleEnum;
const
  wdAlignVerticalTop = $00000000;
  wdAlignVerticalCenter = $00000001;
  wdAlignVerticalJustify = $00000002;
  wdAlignVerticalBottom = $00000003;

// WdCellVerticalAlignment constants
type
  WdCellVerticalAlignment = TOleEnum;
const
  wdCellAlignVerticalTop = $00000000;
  wdCellAlignVerticalCenter = $00000001;
  wdCellAlignVerticalBottom = $00000003;

// WdTrailingCharacter constants
type
  WdTrailingCharacter = TOleEnum;
const
  wdTrailingTab = $00000000;
  wdTrailingSpace = $00000001;
  wdTrailingNone = $00000002;

// WdListGalleryType constants
type
  WdListGalleryType = TOleEnum;
const
  wdBulletGallery = $00000001;
  wdNumberGallery = $00000002;
  wdOutlineNumberGallery = $00000003;

// WdListNumberStyle constants
type
  WdListNumberStyle = TOleEnum;
const
  wdListNumberStyleArabic = $00000000;
  wdListNumberStyleUppercaseRoman = $00000001;
  wdListNumberStyleLowercaseRoman = $00000002;
  wdListNumberStyleUppercaseLetter = $00000003;
  wdListNumberStyleLowercaseLetter = $00000004;
  wdListNumberStyleOrdinal = $00000005;
  wdListNumberStyleCardinalText = $00000006;
  wdListNumberStyleOrdinalText = $00000007;
  wdListNumberStyleKanji = $0000000A;
  wdListNumberStyleKanjiDigit = $0000000B;
  wdListNumberStyleAiueoHalfWidth = $0000000C;
  wdListNumberStyleIrohaHalfWidth = $0000000D;
  wdListNumberStyleArabicFullWidth = $0000000E;
  wdListNumberStyleKanjiTraditional = $00000010;
  wdListNumberStyleKanjiTraditional2 = $00000011;
  wdListNumberStyleNumberInCircle = $00000012;
  wdListNumberStyleAiueo = $00000014;
  wdListNumberStyleIroha = $00000015;
  wdListNumberStyleArabicLZ = $00000016;
  wdListNumberStyleBullet = $00000017;
  wdListNumberStyleGanada = $00000018;
  wdListNumberStyleChosung = $00000019;
  wdListNumberStyleGBNum1 = $0000001A;
  wdListNumberStyleGBNum2 = $0000001B;
  wdListNumberStyleGBNum3 = $0000001C;
  wdListNumberStyleGBNum4 = $0000001D;
  wdListNumberStyleZodiac1 = $0000001E;
  wdListNumberStyleZodiac2 = $0000001F;
  wdListNumberStyleZodiac3 = $00000020;
  wdListNumberStyleTradChinNum1 = $00000021;
  wdListNumberStyleTradChinNum2 = $00000022;
  wdListNumberStyleTradChinNum3 = $00000023;
  wdListNumberStyleTradChinNum4 = $00000024;
  wdListNumberStyleSimpChinNum1 = $00000025;
  wdListNumberStyleSimpChinNum2 = $00000026;
  wdListNumberStyleSimpChinNum3 = $00000027;
  wdListNumberStyleSimpChinNum4 = $00000028;
  wdListNumberStyleHanjaRead = $00000029;
  wdListNumberStyleHanjaReadDigit = $0000002A;
  wdListNumberStyleHangul = $0000002B;
  wdListNumberStyleHanja = $0000002C;
  wdListNumberStyleHebrew1 = $0000002D;
  wdListNumberStyleArabic1 = $0000002E;
  wdListNumberStyleHebrew2 = $0000002F;
  wdListNumberStyleArabic2 = $00000030;
  wdListNumberStyleLegal = $000000FD;
  wdListNumberStyleLegalLZ = $000000FE;
  wdListNumberStyleNone = $000000FF;

// WdListNumberStyleHID constants
type
  WdListNumberStyleHID = TOleEnum;
const
  emptyenum___ = $00000000;

// WdNoteNumberStyle constants
type
  WdNoteNumberStyle = TOleEnum;
const
  wdNoteNumberStyleArabic = $00000000;
  wdNoteNumberStyleUppercaseRoman = $00000001;
  wdNoteNumberStyleLowercaseRoman = $00000002;
  wdNoteNumberStyleUppercaseLetter = $00000003;
  wdNoteNumberStyleLowercaseLetter = $00000004;
  wdNoteNumberStyleSymbol = $00000009;
  wdNoteNumberStyleArabicFullWidth = $0000000E;
  wdNoteNumberStyleKanji = $0000000A;
  wdNoteNumberStyleKanjiDigit = $0000000B;
  wdNoteNumberStyleKanjiTraditional = $00000010;
  wdNoteNumberStyleNumberInCircle = $00000012;
  wdNoteNumberStyleHanjaRead = $00000029;
  wdNoteNumberStyleHanjaReadDigit = $0000002A;
  wdNoteNumberStyleTradChinNum1 = $00000021;
  wdNoteNumberStyleTradChinNum2 = $00000022;
  wdNoteNumberStyleSimpChinNum1 = $00000025;
  wdNoteNumberStyleSimpChinNum2 = $00000026;
  wdNoteNumberStyleHebrewLetter1 = $0000002D;
  wdNoteNumberStyleArabicLetter1 = $0000002E;
  wdNoteNumberStyleHebrewLetter2 = $0000002F;
  wdNoteNumberStyleArabicLetter2 = $00000030;

// WdNoteNumberStyleHID constants
type
  WdNoteNumberStyleHID = TOleEnum;
const
  emptyenum____ = $00000000;

// WdCaptionNumberStyle constants
type
  WdCaptionNumberStyle = TOleEnum;
const
  wdCaptionNumberStyleArabic = $00000000;
  wdCaptionNumberStyleUppercaseRoman = $00000001;
  wdCaptionNumberStyleLowercaseRoman = $00000002;
  wdCaptionNumberStyleUppercaseLetter = $00000003;
  wdCaptionNumberStyleLowercaseLetter = $00000004;
  wdCaptionNumberStyleArabicFullWidth = $0000000E;
  wdCaptionNumberStyleKanji = $0000000A;
  wdCaptionNumberStyleKanjiDigit = $0000000B;
  wdCaptionNumberStyleKanjiTraditional = $00000010;
  wdCaptionNumberStyleNumberInCircle = $00000012;
  wdCaptionNumberStyleGanada = $00000018;
  wdCaptionNumberStyleChosung = $00000019;
  wdCaptionNumberStyleZodiac1 = $0000001E;
  wdCaptionNumberStyleZodiac2 = $0000001F;
  wdCaptionNumberStyleHanjaRead = $00000029;
  wdCaptionNumberStyleHanjaReadDigit = $0000002A;
  wdCaptionNumberStyleTradChinNum2 = $00000022;
  wdCaptionNumberStyleTradChinNum3 = $00000023;
  wdCaptionNumberStyleSimpChinNum2 = $00000026;
  wdCaptionNumberStyleSimpChinNum3 = $00000027;
  wdCaptionNumberStyleHebrewLetter1 = $0000002D;
  wdCaptionNumberStyleArabicLetter1 = $0000002E;
  wdCaptionNumberStyleHebrewLetter2 = $0000002F;
  wdCaptionNumberStyleArabicLetter2 = $00000030;

// WdCaptionNumberStyleHID constants
type
  WdCaptionNumberStyleHID = TOleEnum;
const
  emptyenum_____ = $00000000;

// WdPageNumberStyle constants
type
  WdPageNumberStyle = TOleEnum;
const
  wdPageNumberStyleArabic = $00000000;
  wdPageNumberStyleUppercaseRoman = $00000001;
  wdPageNumberStyleLowercaseRoman = $00000002;
  wdPageNumberStyleUppercaseLetter = $00000003;
  wdPageNumberStyleLowercaseLetter = $00000004;
  wdPageNumberStyleArabicFullWidth = $0000000E;
  wdPageNumberStyleKanji = $0000000A;
  wdPageNumberStyleKanjiDigit = $0000000B;
  wdPageNumberStyleKanjiTraditional = $00000010;
  wdPageNumberStyleNumberInCircle = $00000012;
  wdPageNumberStyleHanjaRead = $00000029;
  wdPageNumberStyleHanjaReadDigit = $0000002A;
  wdPageNumberStyleTradChinNum1 = $00000021;
  wdPageNumberStyleTradChinNum2 = $00000022;
  wdPageNumberStyleSimpChinNum1 = $00000025;
  wdPageNumberStyleSimpChinNum2 = $00000026;
  wdPageNumberStyleHebrewLetter1 = $0000002D;
  wdPageNumberStyleArabicLetter1 = $0000002E;
  wdPageNumberStyleHebrewLetter2 = $0000002F;
  wdPageNumberStyleArabicLetter2 = $00000030;

// WdPageNumberStyleHID constants
type
  WdPageNumberStyleHID = TOleEnum;
const
  emptyenum______ = $00000000;

// WdStatistic constants
type
  WdStatistic = TOleEnum;
const
  wdStatisticWords = $00000000;
  wdStatisticLines = $00000001;
  wdStatisticPages = $00000002;
  wdStatisticCharacters = $00000003;
  wdStatisticParagraphs = $00000004;
  wdStatisticCharactersWithSpaces = $00000005;
  wdStatisticFarEastCharacters = $00000006;

// WdStatisticHID constants
type
  WdStatisticHID = TOleEnum;
const
  emptyenum_______ = $00000000;

// WdBuiltInProperty constants
type
  WdBuiltInProperty = TOleEnum;
const
  wdPropertyTitle = $00000001;
  wdPropertySubject = $00000002;
  wdPropertyAuthor = $00000003;
  wdPropertyKeywords = $00000004;
  wdPropertyComments = $00000005;
  wdPropertyTemplate = $00000006;
  wdPropertyLastAuthor = $00000007;
  wdPropertyRevision = $00000008;
  wdPropertyAppName = $00000009;
  wdPropertyTimeLastPrinted = $0000000A;
  wdPropertyTimeCreated = $0000000B;
  wdPropertyTimeLastSaved = $0000000C;
  wdPropertyVBATotalEdit = $0000000D;
  wdPropertyPages = $0000000E;
  wdPropertyWords = $0000000F;
  wdPropertyCharacters = $00000010;
  wdPropertySecurity = $00000011;
  wdPropertyCategory = $00000012;
  wdPropertyFormat = $00000013;
  wdPropertyManager = $00000014;
  wdPropertyCompany = $00000015;
  wdPropertyBytes = $00000016;
  wdPropertyLines = $00000017;
  wdPropertyParas = $00000018;
  wdPropertySlides = $00000019;
  wdPropertyNotes = $0000001A;
  wdPropertyHiddenSlides = $0000001B;
  wdPropertyMMClips = $0000001C;
  wdPropertyHyperlinkBase = $0000001D;
  wdPropertyCharsWSpaces = $0000001E;

// WdLineSpacing constants
type
  WdLineSpacing = TOleEnum;
const
  wdLineSpaceSingle = $00000000;
  wdLineSpace1pt5 = $00000001;
  wdLineSpaceDouble = $00000002;
  wdLineSpaceAtLeast = $00000003;
  wdLineSpaceExactly = $00000004;
  wdLineSpaceMultiple = $00000005;

// WdNumberType constants
type
  WdNumberType = TOleEnum;
const
  wdNumberParagraph = $00000001;
  wdNumberListNum = $00000002;
  wdNumberAllNumbers = $00000003;

// WdListType constants
type
  WdListType = TOleEnum;
const
  wdListNoNumbering = $00000000;
  wdListListNumOnly = $00000001;
  wdListBullet = $00000002;
  wdListSimpleNumbering = $00000003;
  wdListOutlineNumbering = $00000004;
  wdListMixedNumbering = $00000005;

// WdStoryType constants
type
  WdStoryType = TOleEnum;
const
  wdMainTextStory = $00000001;
  wdFootnotesStory = $00000002;
  wdEndnotesStory = $00000003;
  wdCommentsStory = $00000004;
  wdTextFrameStory = $00000005;
  wdEvenPagesHeaderStory = $00000006;
  wdPrimaryHeaderStory = $00000007;
  wdEvenPagesFooterStory = $00000008;
  wdPrimaryFooterStory = $00000009;
  wdFirstPageHeaderStory = $0000000A;
  wdFirstPageFooterStory = $0000000B;

// WdSaveFormat constants
type
  WdSaveFormat = TOleEnum;
const
  wdFormatDocument = $00000000;
  wdFormatTemplate = $00000001;
  wdFormatText = $00000002;
  wdFormatTextLineBreaks = $00000003;
  wdFormatDOSText = $00000004;
  wdFormatDOSTextLineBreaks = $00000005;
  wdFormatRTF = $00000006;
  wdFormatUnicodeText = $00000007;
  wdFormatEncodedText = $00000007;
  wdFormatHTML = $00000008;

// WdOpenFormat constants
type
  WdOpenFormat = TOleEnum;
const
  wdOpenFormatAuto = $00000000;
  wdOpenFormatDocument = $00000001;
  wdOpenFormatTemplate = $00000002;
  wdOpenFormatRTF = $00000003;
  wdOpenFormatText = $00000004;
  wdOpenFormatUnicodeText = $00000005;
  wdOpenFormatEncodedText = $00000005;
  wdOpenFormatAllWord = $00000006;
  wdOpenFormatWebPages = $00000007;

// WdHeaderFooterIndex constants
type
  WdHeaderFooterIndex = TOleEnum;
const
  wdHeaderFooterPrimary = $00000001;
  wdHeaderFooterFirstPage = $00000002;
  wdHeaderFooterEvenPages = $00000003;

// WdTocFormat constants
type
  WdTocFormat = TOleEnum;
const
  wdTOCTemplate = $00000000;
  wdTOCClassic = $00000001;
  wdTOCDistinctive = $00000002;
  wdTOCFancy = $00000003;
  wdTOCModern = $00000004;
  wdTOCFormal = $00000005;
  wdTOCSimple = $00000006;

// WdTofFormat constants
type
  WdTofFormat = TOleEnum;
const
  wdTOFTemplate = $00000000;
  wdTOFClassic = $00000001;
  wdTOFDistinctive = $00000002;
  wdTOFCentered = $00000003;
  wdTOFFormal = $00000004;
  wdTOFSimple = $00000005;

// WdToaFormat constants
type
  WdToaFormat = TOleEnum;
const
  wdTOATemplate = $00000000;
  wdTOAClassic = $00000001;
  wdTOADistinctive = $00000002;
  wdTOAFormal = $00000003;
  wdTOASimple = $00000004;

// WdLineStyle constants
type
  WdLineStyle = TOleEnum;
const
  wdLineStyleNone = $00000000;
  wdLineStyleSingle = $00000001;
  wdLineStyleDot = $00000002;
  wdLineStyleDashSmallGap = $00000003;
  wdLineStyleDashLargeGap = $00000004;
  wdLineStyleDashDot = $00000005;
  wdLineStyleDashDotDot = $00000006;
  wdLineStyleDouble = $00000007;
  wdLineStyleTriple = $00000008;
  wdLineStyleThinThickSmallGap = $00000009;
  wdLineStyleThickThinSmallGap = $0000000A;
  wdLineStyleThinThickThinSmallGap = $0000000B;
  wdLineStyleThinThickMedGap = $0000000C;
  wdLineStyleThickThinMedGap = $0000000D;
  wdLineStyleThinThickThinMedGap = $0000000E;
  wdLineStyleThinThickLargeGap = $0000000F;
  wdLineStyleThickThinLargeGap = $00000010;
  wdLineStyleThinThickThinLargeGap = $00000011;
  wdLineStyleSingleWavy = $00000012;
  wdLineStyleDoubleWavy = $00000013;
  wdLineStyleDashDotStroked = $00000014;
  wdLineStyleEmboss3D = $00000015;
  wdLineStyleEngrave3D = $00000016;
  wdLineStyleOutset = $00000017;
  wdLineStyleInset = $00000018;

// WdLineWidth constants
type
  WdLineWidth = TOleEnum;
const
  wdLineWidth025pt = $00000002;
  wdLineWidth050pt = $00000004;
  wdLineWidth075pt = $00000006;
  wdLineWidth100pt = $00000008;
  wdLineWidth150pt = $0000000C;
  wdLineWidth225pt = $00000012;
  wdLineWidth300pt = $00000018;
  wdLineWidth450pt = $00000024;
  wdLineWidth600pt = $00000030;

// WdBreakType constants
type
  WdBreakType = TOleEnum;
const
  wdSectionBreakNextPage = $00000002;
  wdSectionBreakContinuous = $00000003;
  wdSectionBreakEvenPage = $00000004;
  wdSectionBreakOddPage = $00000005;
  wdLineBreak = $00000006;
  wdPageBreak = $00000007;
  wdColumnBreak = $00000008;
  wdLineBreakClearLeft = $00000009;
  wdLineBreakClearRight = $0000000A;
  wdTextWrappingBreak = $0000000B;

// WdTabLeader constants
type
  WdTabLeader = TOleEnum;
const
  wdTabLeaderSpaces = $00000000;
  wdTabLeaderDots = $00000001;
  wdTabLeaderDashes = $00000002;
  wdTabLeaderLines = $00000003;
  wdTabLeaderHeavy = $00000004;
  wdTabLeaderMiddleDot = $00000005;

// WdTabLeaderHID constants
type
  WdTabLeaderHID = TOleEnum;
const
  emptyenum________ = $00000000;

// WdMeasurementUnits constants
type
  WdMeasurementUnits = TOleEnum;
const
  wdInches = $00000000;
  wdCentimeters = $00000001;
  wdMillimeters = $00000002;
  wdPoints = $00000003;
  wdPicas = $00000004;

// WdMeasurementUnitsHID constants
type
  WdMeasurementUnitsHID = TOleEnum;
const
  emptyenum_________ = $00000000;

// WdDropPosition constants
type
  WdDropPosition = TOleEnum;
const
  wdDropNone = $00000000;
  wdDropNormal = $00000001;
  wdDropMargin = $00000002;

// WdNumberingRule constants
type
  WdNumberingRule = TOleEnum;
const
  wdRestartContinuous = $00000000;
  wdRestartSection = $00000001;
  wdRestartPage = $00000002;

// WdFootnoteLocation constants
type
  WdFootnoteLocation = TOleEnum;
const
  wdBottomOfPage = $00000000;
  wdBeneathText = $00000001;

// WdEndnoteLocation constants
type
  WdEndnoteLocation = TOleEnum;
const
  wdEndOfSection = $00000000;
  wdEndOfDocument = $00000001;

// WdSortSeparator constants
type
  WdSortSeparator = TOleEnum;
const
  wdSortSeparateByTabs = $00000000;
  wdSortSeparateByCommas = $00000001;
  wdSortSeparateByDefaultTableSeparator = $00000002;

// WdTableFieldSeparator constants
type
  WdTableFieldSeparator = TOleEnum;
const
  wdSeparateByParagraphs = $00000000;
  wdSeparateByTabs = $00000001;
  wdSeparateByCommas = $00000002;
  wdSeparateByDefaultListSeparator = $00000003;

// WdSortFieldType constants
type
  WdSortFieldType = TOleEnum;
const
  wdSortFieldAlphanumeric = $00000000;
  wdSortFieldNumeric = $00000001;
  wdSortFieldDate = $00000002;
  wdSortFieldSyllable = $00000003;
  wdSortFieldJapanJIS = $00000004;
  wdSortFieldStroke = $00000005;
  wdSortFieldKoreaKS = $00000006;

// WdSortFieldTypeHID constants
type
  WdSortFieldTypeHID = TOleEnum;
const
  emptyenum__________ = $00000000;

// WdSortOrder constants
type
  WdSortOrder = TOleEnum;
const
  wdSortOrderAscending = $00000000;
  wdSortOrderDescending = $00000001;

// WdTableFormat constants
type
  WdTableFormat = TOleEnum;
const
  wdTableFormatNone = $00000000;
  wdTableFormatSimple1 = $00000001;
  wdTableFormatSimple2 = $00000002;
  wdTableFormatSimple3 = $00000003;
  wdTableFormatClassic1 = $00000004;
  wdTableFormatClassic2 = $00000005;
  wdTableFormatClassic3 = $00000006;
  wdTableFormatClassic4 = $00000007;
  wdTableFormatColorful1 = $00000008;
  wdTableFormatColorful2 = $00000009;
  wdTableFormatColorful3 = $0000000A;
  wdTableFormatColumns1 = $0000000B;
  wdTableFormatColumns2 = $0000000C;
  wdTableFormatColumns3 = $0000000D;
  wdTableFormatColumns4 = $0000000E;
  wdTableFormatColumns5 = $0000000F;
  wdTableFormatGrid1 = $00000010;
  wdTableFormatGrid2 = $00000011;
  wdTableFormatGrid3 = $00000012;
  wdTableFormatGrid4 = $00000013;
  wdTableFormatGrid5 = $00000014;
  wdTableFormatGrid6 = $00000015;
  wdTableFormatGrid7 = $00000016;
  wdTableFormatGrid8 = $00000017;
  wdTableFormatList1 = $00000018;
  wdTableFormatList2 = $00000019;
  wdTableFormatList3 = $0000001A;
  wdTableFormatList4 = $0000001B;
  wdTableFormatList5 = $0000001C;
  wdTableFormatList6 = $0000001D;
  wdTableFormatList7 = $0000001E;
  wdTableFormatList8 = $0000001F;
  wdTableFormat3DEffects1 = $00000020;
  wdTableFormat3DEffects2 = $00000021;
  wdTableFormat3DEffects3 = $00000022;
  wdTableFormatContemporary = $00000023;
  wdTableFormatElegant = $00000024;
  wdTableFormatProfessional = $00000025;
  wdTableFormatSubtle1 = $00000026;
  wdTableFormatSubtle2 = $00000027;
  wdTableFormatWeb1 = $00000028;
  wdTableFormatWeb2 = $00000029;
  wdTableFormatWeb3 = $0000002A;

// WdTableFormatApply constants
type
  WdTableFormatApply = TOleEnum;
const
  wdTableFormatApplyBorders = $00000001;
  wdTableFormatApplyShading = $00000002;
  wdTableFormatApplyFont = $00000004;
  wdTableFormatApplyColor = $00000008;
  wdTableFormatApplyAutoFit = $00000010;
  wdTableFormatApplyHeadingRows = $00000020;
  wdTableFormatApplyLastRow = $00000040;
  wdTableFormatApplyFirstColumn = $00000080;
  wdTableFormatApplyLastColumn = $00000100;

// WdLanguageID constants
type
  WdLanguageID = TOleEnum;
const
  wdLanguageNone = $00000000;
  wdNoProofing = $00000400;
  wdAfrikaans = $00000436;
  wdAlbanian = $0000041C;
  wdArabicAlgeria = $00001401;
  wdArabicBahrain = $00003C01;
  wdArabicEgypt = $00000C01;
  wdArabicIraq = $00000801;
  wdArabicJordan = $00002C01;
  wdArabicKuwait = $00003401;
  wdArabicLebanon = $00003001;
  wdArabicLibya = $00001001;
  wdArabicMorocco = $00001801;
  wdArabicOman = $00002001;
  wdArabicQatar = $00004001;
  wdArabic = $00000401;
  wdArabicSyria = $00002801;
  wdArabicTunisia = $00001C01;
  wdArabicUAE = $00003801;
  wdArabicYemen = $00002401;
  wdArmenian = $0000042B;
  wdAssamese = $0000044D;
  wdAzeriCyrillic = $0000082C;
  wdAzeriLatin = $0000042C;
  wdBasque = $0000042D;
  wdByelorussian = $00000423;
  wdBengali = $00000445;
  wdBosniaHerzegovina = $0000101A;
  wdBulgarian = $00000402;
  wdBurmese = $00000455;
  wdCatalan = $00000403;
  wdChineseHongKong = $00000C04;
  wdChineseMacao = $00001404;
  wdSimplifiedChinese = $00000804;
  wdChineseSingapore = $00001004;
  wdTraditionalChinese = $00000404;
  wdCroatian = $0000041A;
  wdCzech = $00000405;
  wdDanish = $00000406;
  wdBelgianDutch = $00000813;
  wdDutch = $00000413;
  wdEnglishAUS = $00000C09;
  wdEnglishBelize = $00002809;
  wdEnglishCanadian = $00001009;
  wdEnglishCaribbean = $00002409;
  wdEnglishIreland = $00001809;
  wdEnglishJamaica = $00002009;
  wdEnglishNewZealand = $00001409;
  wdEnglishPhilippines = $00003409;
  wdEnglishSouthAfrica = $00001C09;
  wdEnglishTrinidad = $00002C09;
  wdEnglishUK = $00000809;
  wdEnglishUS = $00000409;
  wdEnglishZimbabwe = $00003009;
  wdEstonian = $00000425;
  wdFaeroese = $00000438;
  wdFarsi = $00000429;
  wdFinnish = $0000040B;
  wdBelgianFrench = $0000080C;
  wdFrenchCameroon = $00002C0C;
  wdFrenchCanadian = $00000C0C;
  wdFrenchCotedIvoire = $0000300C;
  wdFrench = $0000040C;
  wdFrenchLuxembourg = $0000140C;
  wdFrenchMali = $0000340C;
  wdFrenchMonaco = $0000180C;
  wdFrenchReunion = $0000200C;
  wdFrenchSenegal = $0000280C;
  wdSwissFrench = $0000100C;
  wdFrenchWestIndies = $00001C0C;
  wdFrenchZaire = $0000240C;
  wdFrisianNetherlands = $00000462;
  wdGaelicIreland = $0000083C;
  wdGaelicScotland = $0000043C;
  wdGalician = $00000456;
  wdGeorgian = $00000437;
  wdGermanAustria = $00000C07;
  wdGerman = $00000407;
  wdGermanLiechtenstein = $00001407;
  wdGermanLuxembourg = $00001007;
  wdSwissGerman = $00000807;
  wdGreek = $00000408;
  wdGujarati = $00000447;
  wdHebrew = $0000040D;
  wdHindi = $00000439;
  wdHungarian = $0000040E;
  wdIcelandic = $0000040F;
  wdIndonesian = $00000421;
  wdItalian = $00000410;
  wdSwissItalian = $00000810;
  wdJapanese = $00000411;
  wdKannada = $0000044B;
  wdKashmiri = $00000460;
  wdKazakh = $0000043F;
  wdKhmer = $00000453;
  wdKirghiz = $00000440;
  wdKonkani = $00000457;
  wdKorean = $00000412;
  wdLao = $00000454;
  wdLatvian = $00000426;
  wdLithuanian = $00000427;
  wdMacedonian = $0000042F;
  wdMalaysian = $0000043E;
  wdMalayBruneiDarussalam = $0000083E;
  wdMalayalam = $0000044C;
  wdMaltese = $0000043A;
  wdManipuri = $00000458;
  wdMarathi = $0000044E;
  wdMongolian = $00000450;
  wdNepali = $00000461;
  wdNorwegianBokmol = $00000414;
  wdNorwegianNynorsk = $00000814;
  wdOriya = $00000448;
  wdPolish = $00000415;
  wdBrazilianPortuguese = $00000416;
  wdPortuguese = $00000816;
  wdPunjabi = $00000446;
  wdRhaetoRomanic = $00000417;
  wdRomanianMoldova = $00000818;
  wdRomanian = $00000418;
  wdRussianMoldova = $00000819;
  wdRussian = $00000419;
  wdSamiLappish = $0000043B;
  wdSanskrit = $0000044F;
  wdSerbianCyrillic = $00000C1A;
  wdSerbianLatin = $0000081A;
  wdSindhi = $00000459;
  wdSlovak = $0000041B;
  wdSlovenian = $00000424;
  wdSorbian = $0000042E;
  wdSpanishArgentina = $00002C0A;
  wdSpanishBolivia = $0000400A;
  wdSpanishChile = $0000340A;
  wdSpanishColombia = $0000240A;
  wdSpanishCostaRica = $0000140A;
  wdSpanishDominicanRepublic = $00001C0A;
  wdSpanishEcuador = $0000300A;
  wdSpanishElSalvador = $0000440A;
  wdSpanishGuatemala = $0000100A;
  wdSpanishHonduras = $0000480A;
  wdMexicanSpanish = $0000080A;
  wdSpanishNicaragua = $00004C0A;
  wdSpanishPanama = $0000180A;
  wdSpanishParaguay = $00003C0A;
  wdSpanishPeru = $0000280A;
  wdSpanishPuertoRico = $0000500A;
  wdSpanishModernSort = $00000C0A;
  wdSpanish = $0000040A;
  wdSpanishUruguay = $0000380A;
  wdSpanishVenezuela = $0000200A;
  wdSesotho = $00000430;
  wdSutu = $00000430;
  wdSwahili = $00000441;
  wdSwedishFinland = $0000081D;
  wdSwedish = $0000041D;
  wdTajik = $00000428;
  wdTamil = $00000449;
  wdTatar = $00000444;
  wdTelugu = $0000044A;
  wdThai = $0000041E;
  wdTibetan = $00000451;
  wdTsonga = $00000431;
  wdTswana = $00000432;
  wdTurkish = $0000041F;
  wdTurkmen = $00000442;
  wdUkrainian = $00000422;
  wdUrdu = $00000420;
  wdUzbekCyrillic = $00000843;
  wdUzbekLatin = $00000443;
  wdVenda = $00000433;
  wdVietnamese = $0000042A;
  wdWelsh = $00000452;
  wdXhosa = $00000434;
  wdZulu = $00000435;

// WdFieldType constants
type
  WdFieldType = TOleEnum;
const
  wdFieldEmpty = $FFFFFFFF;
  wdFieldRef = $00000003;
  wdFieldIndexEntry = $00000004;
  wdFieldFootnoteRef = $00000005;
  wdFieldSet = $00000006;
  wdFieldIf = $00000007;
  wdFieldIndex = $00000008;
  wdFieldTOCEntry = $00000009;
  wdFieldStyleRef = $0000000A;
  wdFieldRefDoc = $0000000B;
  wdFieldSequence = $0000000C;
  wdFieldTOC = $0000000D;
  wdFieldInfo = $0000000E;
  wdFieldTitle = $0000000F;
  wdFieldSubject = $00000010;
  wdFieldAuthor = $00000011;
  wdFieldKeyWord = $00000012;
  wdFieldComments = $00000013;
  wdFieldLastSavedBy = $00000014;
  wdFieldCreateDate = $00000015;
  wdFieldSaveDate = $00000016;
  wdFieldPrintDate = $00000017;
  wdFieldRevisionNum = $00000018;
  wdFieldEditTime = $00000019;
  wdFieldNumPages = $0000001A;
  wdFieldNumWords = $0000001B;
  wdFieldNumChars = $0000001C;
  wdFieldFileName = $0000001D;
  wdFieldTemplate = $0000001E;
  wdFieldDate = $0000001F;
  wdFieldTime = $00000020;
  wdFieldPage = $00000021;
  wdFieldExpression = $00000022;
  wdFieldQuote = $00000023;
  wdFieldInclude = $00000024;
  wdFieldPageRef = $00000025;
  wdFieldAsk = $00000026;
  wdFieldFillIn = $00000027;
  wdFieldData = $00000028;
  wdFieldNext = $00000029;
  wdFieldNextIf = $0000002A;
  wdFieldSkipIf = $0000002B;
  wdFieldMergeRec = $0000002C;
  wdFieldDDE = $0000002D;
  wdFieldDDEAuto = $0000002E;
  wdFieldGlossary = $0000002F;
  wdFieldPrint = $00000030;
  wdFieldFormula = $00000031;
  wdFieldGoToButton = $00000032;
  wdFieldMacroButton = $00000033;
  wdFieldAutoNumOutline = $00000034;
  wdFieldAutoNumLegal = $00000035;
  wdFieldAutoNum = $00000036;
  wdFieldImport = $00000037;
  wdFieldLink = $00000038;
  wdFieldSymbol = $00000039;
  wdFieldEmbed = $0000003A;
  wdFieldMergeField = $0000003B;
  wdFieldUserName = $0000003C;
  wdFieldUserInitials = $0000003D;
  wdFieldUserAddress = $0000003E;
  wdFieldBarCode = $0000003F;
  wdFieldDocVariable = $00000040;
  wdFieldSection = $00000041;
  wdFieldSectionPages = $00000042;
  wdFieldIncludePicture = $00000043;
  wdFieldIncludeText = $00000044;
  wdFieldFileSize = $00000045;
  wdFieldFormTextInput = $00000046;
  wdFieldFormCheckBox = $00000047;
  wdFieldNoteRef = $00000048;
  wdFieldTOA = $00000049;
  wdFieldTOAEntry = $0000004A;
  wdFieldMergeSeq = $0000004B;
  wdFieldPrivate = $0000004D;
  wdFieldDatabase = $0000004E;
  wdFieldAutoText = $0000004F;
  wdFieldCompare = $00000050;
  wdFieldAddin = $00000051;
  wdFieldSubscriber = $00000052;
  wdFieldFormDropDown = $00000053;
  wdFieldAdvance = $00000054;
  wdFieldDocProperty = $00000055;
  wdFieldOCX = $00000057;
  wdFieldHyperlink = $00000058;
  wdFieldAutoTextList = $00000059;
  wdFieldListNum = $0000005A;
  wdFieldHTMLActiveX = $0000005B;

// WdBuiltinStyle constants
type
  WdBuiltinStyle = TOleEnum;
const
  wdStyleNormal = $FFFFFFFF;
  wdStyleEnvelopeAddress = $FFFFFFDB;
  wdStyleEnvelopeReturn = $FFFFFFDA;
  wdStyleBodyText = $FFFFFFBD;
  wdStyleHeading1 = $FFFFFFFE;
  wdStyleHeading2 = $FFFFFFFD;
  wdStyleHeading3 = $FFFFFFFC;
  wdStyleHeading4 = $FFFFFFFB;
  wdStyleHeading5 = $FFFFFFFA;
  wdStyleHeading6 = $FFFFFFF9;
  wdStyleHeading7 = $FFFFFFF8;
  wdStyleHeading8 = $FFFFFFF7;
  wdStyleHeading9 = $FFFFFFF6;
  wdStyleIndex1 = $FFFFFFF5;
  wdStyleIndex2 = $FFFFFFF4;
  wdStyleIndex3 = $FFFFFFF3;
  wdStyleIndex4 = $FFFFFFF2;
  wdStyleIndex5 = $FFFFFFF1;
  wdStyleIndex6 = $FFFFFFF0;
  wdStyleIndex7 = $FFFFFFEF;
  wdStyleIndex8 = $FFFFFFEE;
  wdStyleIndex9 = $FFFFFFED;
  wdStyleTOC1 = $FFFFFFEC;
  wdStyleTOC2 = $FFFFFFEB;
  wdStyleTOC3 = $FFFFFFEA;
  wdStyleTOC4 = $FFFFFFE9;
  wdStyleTOC5 = $FFFFFFE8;
  wdStyleTOC6 = $FFFFFFE7;
  wdStyleTOC7 = $FFFFFFE6;
  wdStyleTOC8 = $FFFFFFE5;
  wdStyleTOC9 = $FFFFFFE4;
  wdStyleNormalIndent = $FFFFFFE3;
  wdStyleFootnoteText = $FFFFFFE2;
  wdStyleCommentText = $FFFFFFE1;
  wdStyleHeader = $FFFFFFE0;
  wdStyleFooter = $FFFFFFDF;
  wdStyleIndexHeading = $FFFFFFDE;
  wdStyleCaption = $FFFFFFDD;
  wdStyleTableOfFigures = $FFFFFFDC;
  wdStyleFootnoteReference = $FFFFFFD9;
  wdStyleCommentReference = $FFFFFFD8;
  wdStyleLineNumber = $FFFFFFD7;
  wdStylePageNumber = $FFFFFFD6;
  wdStyleEndnoteReference = $FFFFFFD5;
  wdStyleEndnoteText = $FFFFFFD4;
  wdStyleTableOfAuthorities = $FFFFFFD3;
  wdStyleMacroText = $FFFFFFD2;
  wdStyleTOAHeading = $FFFFFFD1;
  wdStyleList = $FFFFFFD0;
  wdStyleListBullet = $FFFFFFCF;
  wdStyleListNumber = $FFFFFFCE;
  wdStyleList2 = $FFFFFFCD;
  wdStyleList3 = $FFFFFFCC;
  wdStyleList4 = $FFFFFFCB;
  wdStyleList5 = $FFFFFFCA;
  wdStyleListBullet2 = $FFFFFFC9;
  wdStyleListBullet3 = $FFFFFFC8;
  wdStyleListBullet4 = $FFFFFFC7;
  wdStyleListBullet5 = $FFFFFFC6;
  wdStyleListNumber2 = $FFFFFFC5;
  wdStyleListNumber3 = $FFFFFFC4;
  wdStyleListNumber4 = $FFFFFFC3;
  wdStyleListNumber5 = $FFFFFFC2;
  wdStyleTitle = $FFFFFFC1;
  wdStyleClosing = $FFFFFFC0;
  wdStyleSignature = $FFFFFFBF;
  wdStyleDefaultParagraphFont = $FFFFFFBE;
  wdStyleBodyTextIndent = $FFFFFFBC;
  wdStyleListContinue = $FFFFFFBB;
  wdStyleListContinue2 = $FFFFFFBA;
  wdStyleListContinue3 = $FFFFFFB9;
  wdStyleListContinue4 = $FFFFFFB8;
  wdStyleListContinue5 = $FFFFFFB7;
  wdStyleMessageHeader = $FFFFFFB6;
  wdStyleSubtitle = $FFFFFFB5;
  wdStyleSalutation = $FFFFFFB4;
  wdStyleDate = $FFFFFFB3;
  wdStyleBodyTextFirstIndent = $FFFFFFB2;
  wdStyleBodyTextFirstIndent2 = $FFFFFFB1;
  wdStyleNoteHeading = $FFFFFFB0;
  wdStyleBodyText2 = $FFFFFFAF;
  wdStyleBodyText3 = $FFFFFFAE;
  wdStyleBodyTextIndent2 = $FFFFFFAD;
  wdStyleBodyTextIndent3 = $FFFFFFAC;
  wdStyleBlockQuotation = $FFFFFFAB;
  wdStyleHyperlink = $FFFFFFAA;
  wdStyleHyperlinkFollowed = $FFFFFFA9;
  wdStyleStrong = $FFFFFFA8;
  wdStyleEmphasis = $FFFFFFA7;
  wdStyleNavPane = $FFFFFFA6;
  wdStylePlainText = $FFFFFFA5;
  wdStyleHtmlNormal = $FFFFFFA1;
  wdStyleHtmlAcronym = $FFFFFFA0;
  wdStyleHtmlAddress = $FFFFFF9F;
  wdStyleHtmlCite = $FFFFFF9E;
  wdStyleHtmlCode = $FFFFFF9D;
  wdStyleHtmlDfn = $FFFFFF9C;
  wdStyleHtmlKbd = $FFFFFF9B;
  wdStyleHtmlPre = $FFFFFF9A;
  wdStyleHtmlSamp = $FFFFFF99;
  wdStyleHtmlTt = $FFFFFF98;
  wdStyleHtmlVar = $FFFFFF97;

// WdWordDialogTab constants
type
  WdWordDialogTab = TOleEnum;
const
  wdDialogToolsOptionsTabView = $000000CC;
  wdDialogToolsOptionsTabGeneral = $000000CB;
  wdDialogToolsOptionsTabEdit = $000000E0;
  wdDialogToolsOptionsTabPrint = $000000D0;
  wdDialogToolsOptionsTabSave = $000000D1;
  wdDialogToolsOptionsTabProofread = $000000D3;
  wdDialogToolsOptionsTabTrackChanges = $00000182;
  wdDialogToolsOptionsTabUserInfo = $000000D5;
  wdDialogToolsOptionsTabCompatibility = $0000020D;
  wdDialogToolsOptionsTabTypography = $000002E3;
  wdDialogToolsOptionsTabFileLocations = $000000E1;
  wdDialogToolsOptionsTabFuzzy = $00000316;
  wdDialogToolsOptionsTabHangulHanjaConversion = $00000312;
  wdDialogToolsOptionsTabBidi = $00000405;
  wdDialogFilePageSetupTabMargins = $000249F0;
  wdDialogFilePageSetupTabPaperSize = $000249F1;
  wdDialogFilePageSetupTabPaperSource = $000249F2;
  wdDialogFilePageSetupTabLayout = $000249F3;
  wdDialogFilePageSetupTabCharsLines = $000249F4;
  wdDialogInsertSymbolTabSymbols = $00030D40;
  wdDialogInsertSymbolTabSpecialCharacters = $00030D41;
  wdDialogNoteOptionsTabAllFootnotes = $000493E0;
  wdDialogNoteOptionsTabAllEndnotes = $000493E1;
  wdDialogInsertIndexAndTablesTabIndex = $00061A80;
  wdDialogInsertIndexAndTablesTabTableOfContents = $00061A81;
  wdDialogInsertIndexAndTablesTabTableOfFigures = $00061A82;
  wdDialogInsertIndexAndTablesTabTableOfAuthorities = $00061A83;
  wdDialogOrganizerTabStyles = $0007A120;
  wdDialogOrganizerTabAutoText = $0007A121;
  wdDialogOrganizerTabCommandBars = $0007A122;
  wdDialogOrganizerTabMacros = $0007A123;
  wdDialogFormatFontTabFont = $000927C0;
  wdDialogFormatFontTabCharacterSpacing = $000927C1;
  wdDialogFormatFontTabAnimation = $000927C2;
  wdDialogFormatBordersAndShadingTabBorders = $000AAE60;
  wdDialogFormatBordersAndShadingTabPageBorder = $000AAE61;
  wdDialogFormatBordersAndShadingTabShading = $000AAE62;
  wdDialogToolsEnvelopesAndLabelsTabEnvelopes = $000C3500;
  wdDialogToolsEnvelopesAndLabelsTabLabels = $000C3501;
  wdDialogFormatParagraphTabIndentsAndSpacing = $000F4240;
  wdDialogFormatParagraphTabTextFlow = $000F4241;
  wdDialogFormatParagraphTabTeisai = $000F4242;
  wdDialogFormatDrawingObjectTabColorsAndLines = $00124F80;
  wdDialogFormatDrawingObjectTabSize = $00124F81;
  wdDialogFormatDrawingObjectTabPosition = $00124F82;
  wdDialogFormatDrawingObjectTabWrapping = $00124F83;
  wdDialogFormatDrawingObjectTabPicture = $00124F84;
  wdDialogFormatDrawingObjectTabTextbox = $00124F85;
  wdDialogFormatDrawingObjectTabWeb = $00124F86;
  wdDialogFormatDrawingObjectTabHR = $00124F87;
  wdDialogToolsAutoCorrectExceptionsTabFirstLetter = $00155CC0;
  wdDialogToolsAutoCorrectExceptionsTabInitialCaps = $00155CC1;
  wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = $00155CC2;
  wdDialogToolsAutoCorrectExceptionsTabIac = $00155CC3;
  wdDialogFormatBulletsAndNumberingTabBulleted = $0016E360;
  wdDialogFormatBulletsAndNumberingTabNumbered = $0016E361;
  wdDialogFormatBulletsAndNumberingTabOutlineNumbered = $0016E362;
  wdDialogLetterWizardTabLetterFormat = $00186A00;
  wdDialogLetterWizardTabRecipientInfo = $00186A01;
  wdDialogLetterWizardTabOtherElements = $00186A02;
  wdDialogLetterWizardTabSenderInfo = $00186A03;
  wdDialogToolsAutoManagerTabAutoCorrect = $0019F0A0;
  wdDialogToolsAutoManagerTabAutoFormatAsYouType = $0019F0A1;
  wdDialogToolsAutoManagerTabAutoText = $0019F0A2;
  wdDialogToolsAutoManagerTabAutoFormat = $0019F0A3;
  wdDialogEmailOptionsTabSignature = $001CFDE0;
  wdDialogEmailOptionsTabStationary = $001CFDE1;
  wdDialogEmailOptionsTabQuoting = $001CFDE2;
  wdDialogWebOptionsGeneral = $001E8480;
  wdDialogWebOptionsFiles = $001E8481;
  wdDialogWebOptionsPictures = $001E8482;
  wdDialogWebOptionsEncoding = $001E8483;
  wdDialogWebOptionsFonts = $001E8484;

// WdWordDialogTabHID constants
type
  WdWordDialogTabHID = TOleEnum;
const
  emptyenum___________ = $00000000;

// WdWordDialog constants
type
  WdWordDialog = TOleEnum;
const
  wdDialogHelpAbout = $00000009;
  wdDialogHelpWordPerfectHelp = $0000000A;
  wdDialogHelpWordPerfectHelpOptions = $000001FF;
  wdDialogFormatChangeCase = $00000142;
  wdDialogToolsOptionsFuzzy = $00000316;
  wdDialogToolsWordCount = $000000E4;
  wdDialogDocumentStatistics = $0000004E;
  wdDialogFileNew = $0000004F;
  wdDialogFileOpen = $00000050;
  wdDialogMailMergeOpenDataSource = $00000051;
  wdDialogMailMergeOpenHeaderSource = $00000052;
  wdDialogMailMergeUseAddressBook = $0000030B;
  wdDialogFileSaveAs = $00000054;
  wdDialogFileSummaryInfo = $00000056;
  wdDialogToolsTemplates = $00000057;
  wdDialogOrganizer = $000000DE;
  wdDialogFilePrint = $00000058;
  wdDialogMailMerge = $000002A4;
  wdDialogMailMergeCheck = $000002A5;
  wdDialogMailMergeQueryOptions = $000002A9;
  wdDialogMailMergeFindRecord = $00000239;
  wdDialogMailMergeInsertIf = $00000FD1;
  wdDialogMailMergeInsertNextIf = $00000FD5;
  wdDialogMailMergeInsertSkipIf = $00000FD7;
  wdDialogMailMergeInsertFillIn = $00000FD0;
  wdDialogMailMergeInsertAsk = $00000FCF;
  wdDialogMailMergeInsertSet = $00000FD6;
  wdDialogMailMergeHelper = $000002A8;
  wdDialogLetterWizard = $00000335;
  wdDialogFilePrintSetup = $00000061;
  wdDialogFileFind = $00000063;
  wdDialogMailMergeCreateDataSource = $00000282;
  wdDialogMailMergeCreateHeaderSource = $00000283;
  wdDialogEditPasteSpecial = $0000006F;
  wdDialogEditFind = $00000070;
  wdDialogEditReplace = $00000075;
  wdDialogEditGoToOld = $0000032B;
  wdDialogEditGoTo = $00000380;
  wdDialogCreateAutoText = $00000368;
  wdDialogEditAutoText = $000003D9;
  wdDialogEditLinks = $0000007C;
  wdDialogEditObject = $0000007D;
  wdDialogConvertObject = $00000188;
  wdDialogTableToText = $00000080;
  wdDialogTextToTable = $0000007F;
  wdDialogTableInsertTable = $00000081;
  wdDialogTableInsertCells = $00000082;
  wdDialogTableInsertRow = $00000083;
  wdDialogTableDeleteCells = $00000085;
  wdDialogTableSplitCells = $00000089;
  wdDialogTableFormula = $0000015C;
  wdDialogTableAutoFormat = $00000233;
  wdDialogTableFormatCell = $00000264;
  wdDialogViewZoom = $00000241;
  wdDialogNewToolbar = $0000024A;
  wdDialogInsertBreak = $0000009F;
  wdDialogInsertFootnote = $00000172;
  wdDialogInsertSymbol = $000000A2;
  wdDialogInsertPicture = $000000A3;
  wdDialogInsertFile = $000000A4;
  wdDialogInsertDateTime = $000000A5;
  wdDialogInsertNumber = $0000032C;
  wdDialogInsertField = $000000A6;
  wdDialogInsertDatabase = $00000155;
  wdDialogInsertMergeField = $000000A7;
  wdDialogInsertBookmark = $000000A8;
  wdDialogInsertHyperlink = $0000039D;
  wdDialogMarkIndexEntry = $000000A9;
  wdDialogMarkCitation = $000001CF;
  wdDialogEditTOACategory = $00000271;
  wdDialogInsertIndexAndTables = $000001D9;
  wdDialogInsertIndex = $000000AA;
  wdDialogInsertTableOfContents = $000000AB;
  wdDialogMarkTableOfContentsEntry = $000001BA;
  wdDialogInsertTableOfFigures = $000001D8;
  wdDialogInsertTableOfAuthorities = $000001D7;
  wdDialogInsertObject = $000000AC;
  wdDialogFormatCallout = $00000262;
  wdDialogDrawSnapToGrid = $00000279;
  wdDialogDrawAlign = $0000027A;
  wdDialogToolsEnvelopesAndLabels = $0000025F;
  wdDialogToolsCreateEnvelope = $000000AD;
  wdDialogToolsCreateLabels = $000001E9;
  wdDialogToolsProtectDocument = $000001F7;
  wdDialogToolsProtectSection = $00000242;
  wdDialogToolsUnprotectDocument = $00000209;
  wdDialogFormatFont = $000000AE;
  wdDialogFormatParagraph = $000000AF;
  wdDialogFormatSectionLayout = $000000B0;
  wdDialogFormatColumns = $000000B1;
  wdDialogFileDocumentLayout = $000000B2;
  wdDialogFileMacPageSetup = $000002AD;
  wdDialogFilePrintOneCopy = $000001BD;
  wdDialogFileMacPageSetupGX = $000001BC;
  wdDialogFileMacCustomPageSetupGX = $000002E1;
  wdDialogFilePageSetup = $000000B2;
  wdDialogFormatTabs = $000000B3;
  wdDialogFormatStyle = $000000B4;
  wdDialogFormatStyleGallery = $000001F9;
  wdDialogFormatDefineStyleFont = $000000B5;
  wdDialogFormatDefineStylePara = $000000B6;
  wdDialogFormatDefineStyleTabs = $000000B7;
  wdDialogFormatDefineStyleFrame = $000000B8;
  wdDialogFormatDefineStyleBorders = $000000B9;
  wdDialogFormatDefineStyleLang = $000000BA;
  wdDialogFormatPicture = $000000BB;
  wdDialogToolsLanguage = $000000BC;
  wdDialogFormatBordersAndShading = $000000BD;
  wdDialogFormatDrawingObject = $000003C0;
  wdDialogFormatFrame = $000000BE;
  wdDialogFormatDropCap = $000001E8;
  wdDialogFormatBulletsAndNumbering = $00000338;
  wdDialogToolsHyphenation = $000000C3;
  wdDialogToolsBulletsNumbers = $000000C4;
  wdDialogToolsHighlightChanges = $000000C5;
  wdDialogToolsAcceptRejectChanges = $000001FA;
  wdDialogToolsMergeDocuments = $000001B3;
  wdDialogToolsCompareDocuments = $000000C6;
  wdDialogTableSort = $000000C7;
  wdDialogToolsCustomizeMenuBar = $00000267;
  wdDialogToolsCustomize = $00000098;
  wdDialogToolsCustomizeKeyboard = $000001B0;
  wdDialogToolsCustomizeMenus = $000001B1;
  wdDialogListCommands = $000002D3;
  wdDialogToolsOptions = $000003CE;
  wdDialogToolsOptionsGeneral = $000000CB;
  wdDialogToolsAdvancedSettings = $000000CE;
  wdDialogToolsOptionsCompatibility = $0000020D;
  wdDialogToolsOptionsPrint = $000000D0;
  wdDialogToolsOptionsSave = $000000D1;
  wdDialogToolsOptionsSpellingAndGrammar = $000000D3;
  wdDialogToolsSpellingAndGrammar = $0000033C;
  wdDialogToolsThesaurus = $000000C2;
  wdDialogToolsOptionsUserInfo = $000000D5;
  wdDialogToolsOptionsAutoFormat = $000003BF;
  wdDialogToolsOptionsTrackChanges = $00000182;
  wdDialogToolsOptionsEdit = $000000E0;
  wdDialogToolsMacro = $000000D7;
  wdDialogInsertPageNumbers = $00000126;
  wdDialogFormatPageNumber = $0000012A;
  wdDialogNoteOptions = $00000175;
  wdDialogCopyFile = $0000012C;
  wdDialogFormatAddrFonts = $00000067;
  wdDialogFormatRetAddrFonts = $000000DD;
  wdDialogToolsOptionsFileLocations = $000000E1;
  wdDialogToolsCreateDirectory = $00000341;
  wdDialogUpdateTOC = $0000014B;
  wdDialogInsertFormField = $000001E3;
  wdDialogFormFieldOptions = $00000161;
  wdDialogInsertCaption = $00000165;
  wdDialogInsertAutoCaption = $00000167;
  wdDialogInsertAddCaption = $00000192;
  wdDialogInsertCaptionNumbering = $00000166;
  wdDialogInsertCrossReference = $0000016F;
  wdDialogToolsManageFields = $00000277;
  wdDialogToolsAutoManager = $00000393;
  wdDialogToolsAutoCorrect = $0000017A;
  wdDialogToolsAutoCorrectExceptions = $000002FA;
  wdDialogConnect = $000001A4;
  wdDialogToolsOptionsBidi = $00000405;
  wdDialogToolsOptionsView = $000000CC;
  wdDialogInsertSubdocument = $00000247;
  wdDialogFileRoutingSlip = $00000270;
  wdDialogFontSubstitution = $00000245;
  wdDialogEditCreatePublisher = $000002DC;
  wdDialogEditSubscribeTo = $000002DD;
  wdDialogEditPublishOptions = $000002DF;
  wdDialogEditSubscribeOptions = $000002E0;
  wdDialogToolsOptionsTypography = $000002E3;
  wdDialogToolsOptionsAutoFormatAsYouType = $0000030A;
  wdDialogControlRun = $000000EB;
  wdDialogFileVersions = $000003B1;
  wdDialogToolsAutoSummarize = $0000036A;
  wdDialogFileSaveVersion = $000003EF;
  wdDialogWindowActivate = $000000DC;
  wdDialogToolsMacroRecord = $000000D6;
  wdDialogToolsRevisions = $000000C5;
  wdDialogEmailOptions = $0000035F;
  wdDialogWebOptions = $00000382;
  wdDialogFitText = $000003D7;
  wdDialogPhoneticGuide = $000003DA;
  wdDialogHorizontalInVertical = $00000488;
  wdDialogTwoLinesInOne = $00000489;
  wdDialogFormatEncloseCharacters = $0000048A;
  wdDialogFormatTheme = $00000357;
  wdDialogTCSCTranslator = $00000484;

// WdWordDialogHID constants
type
  WdWordDialogHID = TOleEnum;
const
  emptyenum____________ = $00000000;

// WdFieldKind constants
type
  WdFieldKind = TOleEnum;
const
  wdFieldKindNone = $00000000;
  wdFieldKindHot = $00000001;
  wdFieldKindWarm = $00000002;
  wdFieldKindCold = $00000003;

// WdTextFormFieldType constants
type
  WdTextFormFieldType = TOleEnum;
const
  wdRegularText = $00000000;
  wdNumberText = $00000001;
  wdDateText = $00000002;
  wdCurrentDateText = $00000003;
  wdCurrentTimeText = $00000004;
  wdCalculationText = $00000005;

// WdChevronConvertRule constants
type
  WdChevronConvertRule = TOleEnum;
const
  wdNeverConvert = $00000000;
  wdAlwaysConvert = $00000001;
  wdAskToNotConvert = $00000002;
  wdAskToConvert = $00000003;

// WdMailMergeMainDocType constants
type
  WdMailMergeMainDocType = TOleEnum;
const
  wdNotAMergeDocument = $FFFFFFFF;
  wdFormLetters = $00000000;
  wdMailingLabels = $00000001;
  wdEnvelopes = $00000002;
  wdCatalog = $00000003;

// WdMailMergeState constants
type
  WdMailMergeState = TOleEnum;
const
  wdNormalDocument = $00000000;
  wdMainDocumentOnly = $00000001;
  wdMainAndDataSource = $00000002;
  wdMainAndHeader = $00000003;
  wdMainAndSourceAndHeader = $00000004;
  wdDataSource = $00000005;

// WdMailMergeDestination constants
type
  WdMailMergeDestination = TOleEnum;
const
  wdSendToNewDocument = $00000000;
  wdSendToPrinter = $00000001;
  wdSendToEmail = $00000002;
  wdSendToFax = $00000003;

// WdMailMergeActiveRecord constants
type
  WdMailMergeActiveRecord = TOleEnum;
const
  wdNoActiveRecord = $FFFFFFFF;
  wdNextRecord = $FFFFFFFE;
  wdPreviousRecord = $FFFFFFFD;
  wdFirstRecord = $FFFFFFFC;
  wdLastRecord = $FFFFFFFB;

// WdMailMergeDefaultRecord constants
type
  WdMailMergeDefaultRecord = TOleEnum;
const
  wdDefaultFirstRecord = $00000001;
  wdDefaultLastRecord = $FFFFFFF0;

// WdMailMergeDataSource constants
type
  WdMailMergeDataSource = TOleEnum;
const
  wdNoMergeInfo = $FFFFFFFF;
  wdMergeInfoFromWord = $00000000;
  wdMergeInfoFromAccessDDE = $00000001;
  wdMergeInfoFromExcelDDE = $00000002;
  wdMergeInfoFromMSQueryDDE = $00000003;
  wdMergeInfoFromODBC = $00000004;

// WdMailMergeComparison constants
type
  WdMailMergeComparison = TOleEnum;
const
  wdMergeIfEqual = $00000000;
  wdMergeIfNotEqual = $00000001;
  wdMergeIfLessThan = $00000002;
  wdMergeIfGreaterThan = $00000003;
  wdMergeIfLessThanOrEqual = $00000004;
  wdMergeIfGreaterThanOrEqual = $00000005;
  wdMergeIfIsBlank = $00000006;
  wdMergeIfIsNotBlank = $00000007;

// WdBookmarkSortBy constants
type
  WdBookmarkSortBy = TOleEnum;
const
  wdSortByName = $00000000;
  wdSortByLocation = $00000001;

// WdWindowState constants
type
  WdWindowState = TOleEnum;
const
  wdWindowStateNormal = $00000000;
  wdWindowStateMaximize = $00000001;
  wdWindowStateMinimize = $00000002;

// WdPictureLinkType constants
type
  WdPictureLinkType = TOleEnum;
const
  wdLinkNone = $00000000;
  wdLinkDataInDoc = $00000001;
  wdLinkDataOnDisk = $00000002;

// WdLinkType constants
type
  WdLinkType = TOleEnum;
const
  wdLinkTypeOLE = $00000000;
  wdLinkTypePicture = $00000001;
  wdLinkTypeText = $00000002;
  wdLinkTypeReference = $00000003;
  wdLinkTypeInclude = $00000004;
  wdLinkTypeImport = $00000005;
  wdLinkTypeDDE = $00000006;
  wdLinkTypeDDEAuto = $00000007;

// WdWindowType constants
type
  WdWindowType = TOleEnum;
const
  wdWindowDocument = $00000000;
  wdWindowTemplate = $00000001;

// WdViewType constants
type
  WdViewType = TOleEnum;
const
  wdNormalView = $00000001;
  wdOutlineView = $00000002;
  wdPrintView = $00000003;
  wdPrintPreview = $00000004;
  wdMasterView = $00000005;
  wdWebView = $00000006;

// WdSeekView constants
type
  WdSeekView = TOleEnum;
const
  wdSeekMainDocument = $00000000;
  wdSeekPrimaryHeader = $00000001;
  wdSeekFirstPageHeader = $00000002;
  wdSeekEvenPagesHeader = $00000003;
  wdSeekPrimaryFooter = $00000004;
  wdSeekFirstPageFooter = $00000005;
  wdSeekEvenPagesFooter = $00000006;
  wdSeekFootnotes = $00000007;
  wdSeekEndnotes = $00000008;
  wdSeekCurrentPageHeader = $00000009;
  wdSeekCurrentPageFooter = $0000000A;

// WdSpecialPane constants
type
  WdSpecialPane = TOleEnum;
const
  wdPaneNone = $00000000;
  wdPanePrimaryHeader = $00000001;
  wdPaneFirstPageHeader = $00000002;
  wdPaneEvenPagesHeader = $00000003;
  wdPanePrimaryFooter = $00000004;
  wdPaneFirstPageFooter = $00000005;
  wdPaneEvenPagesFooter = $00000006;
  wdPaneFootnotes = $00000007;
  wdPaneEndnotes = $00000008;
  wdPaneFootnoteContinuationNotice = $00000009;
  wdPaneFootnoteContinuationSeparator = $0000000A;
  wdPaneFootnoteSeparator = $0000000B;
  wdPaneEndnoteContinuationNotice = $0000000C;
  wdPaneEndnoteContinuationSeparator = $0000000D;
  wdPaneEndnoteSeparator = $0000000E;
  wdPaneComments = $0000000F;
  wdPaneCurrentPageHeader = $00000010;
  wdPaneCurrentPageFooter = $00000011;

// WdPageFit constants
type
  WdPageFit = TOleEnum;
const
  wdPageFitNone = $00000000;
  wdPageFitFullPage = $00000001;
  wdPageFitBestFit = $00000002;
  wdPageFitTextFit = $00000003;

// WdBrowseTarget constants
type
  WdBrowseTarget = TOleEnum;
const
  wdBrowsePage = $00000001;
  wdBrowseSection = $00000002;
  wdBrowseComment = $00000003;
  wdBrowseFootnote = $00000004;
  wdBrowseEndnote = $00000005;
  wdBrowseField = $00000006;
  wdBrowseTable = $00000007;
  wdBrowseGraphic = $00000008;
  wdBrowseHeading = $00000009;
  wdBrowseEdit = $0000000A;
  wdBrowseFind = $0000000B;
  wdBrowseGoTo = $0000000C;

// WdPaperTray constants
type
  WdPaperTray = TOleEnum;
const
  wdPrinterDefaultBin = $00000000;
  wdPrinterUpperBin = $00000001;
  wdPrinterOnlyBin = $00000001;
  wdPrinterLowerBin = $00000002;
  wdPrinterMiddleBin = $00000003;
  wdPrinterManualFeed = $00000004;
  wdPrinterEnvelopeFeed = $00000005;
  wdPrinterManualEnvelopeFeed = $00000006;
  wdPrinterAutomaticSheetFeed = $00000007;
  wdPrinterTractorFeed = $00000008;
  wdPrinterSmallFormatBin = $00000009;
  wdPrinterLargeFormatBin = $0000000A;
  wdPrinterLargeCapacityBin = $0000000B;
  wdPrinterPaperCassette = $0000000E;
  wdPrinterFormSource = $0000000F;

// WdOrientation constants
type
  WdOrientation = TOleEnum;
const
  wdOrientPortrait = $00000000;
  wdOrientLandscape = $00000001;

// WdSelectionType constants
type
  WdSelectionType = TOleEnum;
const
  wdNoSelection = $00000000;
  wdSelectionIP = $00000001;
  wdSelectionNormal = $00000002;
  wdSelectionFrame = $00000003;
  wdSelectionColumn = $00000004;
  wdSelectionRow = $00000005;
  wdSelectionBlock = $00000006;
  wdSelectionInlineShape = $00000007;
  wdSelectionShape = $00000008;

// WdCaptionLabelID constants
type
  WdCaptionLabelID = TOleEnum;
const
  wdCaptionFigure = $FFFFFFFF;
  wdCaptionTable = $FFFFFFFE;
  wdCaptionEquation = $FFFFFFFD;

// WdReferenceType constants
type
  WdReferenceType = TOleEnum;
const
  wdRefTypeNumberedItem = $00000000;
  wdRefTypeHeading = $00000001;
  wdRefTypeBookmark = $00000002;
  wdRefTypeFootnote = $00000003;
  wdRefTypeEndnote = $00000004;

// WdReferenceKind constants
type
  WdReferenceKind = TOleEnum;
const
  wdContentText = $FFFFFFFF;
  wdNumberRelativeContext = $FFFFFFFE;
  wdNumberNoContext = $FFFFFFFD;
  wdNumberFullContext = $FFFFFFFC;
  wdEntireCaption = $00000002;
  wdOnlyLabelAndNumber = $00000003;
  wdOnlyCaptionText = $00000004;
  wdFootnoteNumber = $00000005;
  wdEndnoteNumber = $00000006;
  wdPageNumber = $00000007;
  wdPosition = $0000000F;
  wdFootnoteNumberFormatted = $00000010;
  wdEndnoteNumberFormatted = $00000011;

// WdIndexFormat constants
type
  WdIndexFormat = TOleEnum;
const
  wdIndexTemplate = $00000000;
  wdIndexClassic = $00000001;
  wdIndexFancy = $00000002;
  wdIndexModern = $00000003;
  wdIndexBulleted = $00000004;
  wdIndexFormal = $00000005;
  wdIndexSimple = $00000006;

// WdIndexType constants
type
  WdIndexType = TOleEnum;
const
  wdIndexIndent = $00000000;
  wdIndexRunin = $00000001;

// WdRevisionsWrap constants
type
  WdRevisionsWrap = TOleEnum;
const
  wdWrapNever = $00000000;
  wdWrapAlways = $00000001;
  wdWrapAsk = $00000002;

// WdRevisionType constants
type
  WdRevisionType = TOleEnum;
const
  wdNoRevision = $00000000;
  wdRevisionInsert = $00000001;
  wdRevisionDelete = $00000002;
  wdRevisionProperty = $00000003;
  wdRevisionParagraphNumber = $00000004;
  wdRevisionDisplayField = $00000005;
  wdRevisionReconcile = $00000006;
  wdRevisionConflict = $00000007;
  wdRevisionStyle = $00000008;
  wdRevisionReplace = $00000009;

// WdRoutingSlipDelivery constants
type
  WdRoutingSlipDelivery = TOleEnum;
const
  wdOneAfterAnother = $00000000;
  wdAllAtOnce = $00000001;

// WdRoutingSlipStatus constants
type
  WdRoutingSlipStatus = TOleEnum;
const
  wdNotYetRouted = $00000000;
  wdRouteInProgress = $00000001;
  wdRouteComplete = $00000002;

// WdSectionStart constants
type
  WdSectionStart = TOleEnum;
const
  wdSectionContinuous = $00000000;
  wdSectionNewColumn = $00000001;
  wdSectionNewPage = $00000002;
  wdSectionEvenPage = $00000003;
  wdSectionOddPage = $00000004;

// WdSaveOptions constants
type
  WdSaveOptions = TOleEnum;
const
  wdDoNotSaveChanges = $00000000;
  wdSaveChanges = $FFFFFFFF;
  wdPromptToSaveChanges = $FFFFFFFE;

// WdDocumentKind constants
type
  WdDocumentKind = TOleEnum;
const
  wdDocumentNotSpecified = $00000000;
  wdDocumentLetter = $00000001;
  wdDocumentEmail = $00000002;

// WdDocumentType constants
type
  WdDocumentType = TOleEnum;
const
  wdTypeDocument = $00000000;
  wdTypeTemplate = $00000001;
  wdTypeFrameset = $00000002;

// WdOriginalFormat constants
type
  WdOriginalFormat = TOleEnum;
const
  wdWordDocument = $00000000;
  wdOriginalDocumentFormat = $00000001;
  wdPromptUser = $00000002;

// WdRelocate constants
type
  WdRelocate = TOleEnum;
const
  wdRelocateUp = $00000000;
  wdRelocateDown = $00000001;

// WdInsertedTextMark constants
type
  WdInsertedTextMark = TOleEnum;
const
  wdInsertedTextMarkNone = $00000000;
  wdInsertedTextMarkBold = $00000001;
  wdInsertedTextMarkItalic = $00000002;
  wdInsertedTextMarkUnderline = $00000003;
  wdInsertedTextMarkDoubleUnderline = $00000004;

// WdRevisedLinesMark constants
type
  WdRevisedLinesMark = TOleEnum;
const
  wdRevisedLinesMarkNone = $00000000;
  wdRevisedLinesMarkLeftBorder = $00000001;
  wdRevisedLinesMarkRightBorder = $00000002;
  wdRevisedLinesMarkOutsideBorder = $00000003;

// WdDeletedTextMark constants
type
  WdDeletedTextMark = TOleEnum;
const
  wdDeletedTextMarkHidden = $00000000;
  wdDeletedTextMarkStrikeThrough = $00000001;
  wdDeletedTextMarkCaret = $00000002;
  wdDeletedTextMarkPound = $00000003;

// WdRevisedPropertiesMark constants
type
  WdRevisedPropertiesMark = TOleEnum;
const
  wdRevisedPropertiesMarkNone = $00000000;
  wdRevisedPropertiesMarkBold = $00000001;
  wdRevisedPropertiesMarkItalic = $00000002;
  wdRevisedPropertiesMarkUnderline = $00000003;
  wdRevisedPropertiesMarkDoubleUnderline = $00000004;

// WdFieldShading constants
type
  WdFieldShading = TOleEnum;
const
  wdFieldShadingNever = $00000000;
  wdFieldShadingAlways = $00000001;
  wdFieldShadingWhenSelected = $00000002;

// WdDefaultFilePath constants
type
  WdDefaultFilePath = TOleEnum;
const
  wdDocumentsPath = $00000000;
  wdPicturesPath = $00000001;
  wdUserTemplatesPath = $00000002;
  wdWorkgroupTemplatesPath = $00000003;
  wdUserOptionsPath = $00000004;
  wdAutoRecoverPath = $00000005;
  wdToolsPath = $00000006;
  wdTutorialPath = $00000007;
  wdStartupPath = $00000008;
  wdProgramPath = $00000009;
  wdGraphicsFiltersPath = $0000000A;
  wdTextConvertersPath = $0000000B;
  wdProofingToolsPath = $0000000C;
  wdTempFilePath = $0000000D;
  wdCurrentFolderPath = $0000000E;
  wdStyleGalleryPath = $0000000F;
  wdBorderArtPath = $00000013;

// WdCompatibility constants
type
  WdCompatibility = TOleEnum;
const
  wdNoTabHangIndent = $00000001;
  wdNoSpaceRaiseLower = $00000002;
  wdPrintColBlack = $00000003;
  wdWrapTrailSpaces = $00000004;
  wdNoColumnBalance = $00000005;
  wdConvMailMergeEsc = $00000006;
  wdSuppressSpBfAfterPgBrk = $00000007;
  wdSuppressTopSpacing = $00000008;
  wdOrigWordTableRules = $00000009;
  wdTransparentMetafiles = $0000000A;
  wdShowBreaksInFrames = $0000000B;
  wdSwapBordersFacingPages = $0000000C;
  wdLeaveBackslashAlone = $0000000D;
  wdExpandShiftReturn = $0000000E;
  wdDontULTrailSpace = $0000000F;
  wdDontBalanceSingleByteDoubleByteWidth = $00000010;
  wdSuppressTopSpacingMac5 = $00000011;
  wdSpacingInWholePoints = $00000012;
  wdPrintBodyTextBeforeHeader = $00000013;
  wdNoLeading = $00000014;
  wdNoSpaceForUL = $00000015;
  wdMWSmallCaps = $00000016;
  wdNoExtraLineSpacing = $00000017;
  wdTruncateFontHeight = $00000018;
  wdSubFontBySize = $00000019;
  wdUsePrinterMetrics = $0000001A;
  wdWW6BorderRules = $0000001B;
  wdExactOnTop = $0000001C;
  wdSuppressBottomSpacing = $0000001D;
  wdWPSpaceWidth = $0000001E;
  wdWPJustification = $0000001F;
  wdLineWrapLikeWord6 = $00000020;
  wdShapeLayoutLikeWW8 = $00000021;
  wdFootnoteLayoutLikeWW8 = $00000022;
  wdDontUseHTMLParagraphAutoSpacing = $00000023;
  wdDontAdjustLineHeightInTable = $00000024;
  wdForgetLastTabAlignment = $00000025;
  wdAutospaceLikeWW7 = $00000026;
  wdAlignTablesRowByRow = $00000027;
  wdLayoutRawTableWidth = $00000028;
  wdLayoutTableRowsApart = $00000029;
  wdUseWord97LineBreakingRules = $0000002A;

// WdPaperSize constants
type
  WdPaperSize = TOleEnum;
const
  wdPaper10x14 = $00000000;
  wdPaper11x17 = $00000001;
  wdPaperLetter = $00000002;
  wdPaperLetterSmall = $00000003;
  wdPaperLegal = $00000004;
  wdPaperExecutive = $00000005;
  wdPaperA3 = $00000006;
  wdPaperA4 = $00000007;
  wdPaperA4Small = $00000008;
  wdPaperA5 = $00000009;
  wdPaperB4 = $0000000A;
  wdPaperB5 = $0000000B;
  wdPaperCSheet = $0000000C;
  wdPaperDSheet = $0000000D;
  wdPaperESheet = $0000000E;
  wdPaperFanfoldLegalGerman = $0000000F;
  wdPaperFanfoldStdGerman = $00000010;
  wdPaperFanfoldUS = $00000011;
  wdPaperFolio = $00000012;
  wdPaperLedger = $00000013;
  wdPaperNote = $00000014;
  wdPaperQuarto = $00000015;
  wdPaperStatement = $00000016;
  wdPaperTabloid = $00000017;
  wdPaperEnvelope9 = $00000018;
  wdPaperEnvelope10 = $00000019;
  wdPaperEnvelope11 = $0000001A;
  wdPaperEnvelope12 = $0000001B;
  wdPaperEnvelope14 = $0000001C;
  wdPaperEnvelopeB4 = $0000001D;
  wdPaperEnvelopeB5 = $0000001E;
  wdPaperEnvelopeB6 = $0000001F;
  wdPaperEnvelopeC3 = $00000020;
  wdPaperEnvelopeC4 = $00000021;
  wdPaperEnvelopeC5 = $00000022;
  wdPaperEnvelopeC6 = $00000023;
  wdPaperEnvelopeC65 = $00000024;
  wdPaperEnvelopeDL = $00000025;
  wdPaperEnvelopeItaly = $00000026;
  wdPaperEnvelopeMonarch = $00000027;
  wdPaperEnvelopePersonal = $00000028;
  wdPaperCustom = $00000029;

// WdCustomLabelPageSize constants
type
  WdCustomLabelPageSize = TOleEnum;
const
  wdCustomLabelLetter = $00000000;
  wdCustomLabelLetterLS = $00000001;
  wdCustomLabelA4 = $00000002;
  wdCustomLabelA4LS = $00000003;
  wdCustomLabelA5 = $00000004;
  wdCustomLabelA5LS = $00000005;
  wdCustomLabelB5 = $00000006;
  wdCustomLabelMini = $00000007;
  wdCustomLabelFanfold = $00000008;
  wdCustomLabelVertHalfSheet = $00000009;
  wdCustomLabelVertHalfSheetLS = $0000000A;
  wdCustomLabelHigaki = $0000000B;
  wdCustomLabelHigakiLS = $0000000C;
  wdCustomLabelB4JIS = $0000000D;

// WdProtectionType constants
type
  WdProtectionType = TOleEnum;
const
  wdNoProtection = $FFFFFFFF;
  wdAllowOnlyRevisions = $00000000;
  wdAllowOnlyComments = $00000001;
  wdAllowOnlyFormFields = $00000002;

// WdPartOfSpeech constants
type
  WdPartOfSpeech = TOleEnum;
const
  wdAdjective = $00000000;
  wdNoun = $00000001;
  wdAdverb = $00000002;
  wdVerb = $00000003;
  wdPronoun = $00000004;
  wdConjunction = $00000005;
  wdPreposition = $00000006;
  wdInterjection = $00000007;
  wdIdiom = $00000008;
  wdOther = $00000009;

// WdSubscriberFormats constants
type
  WdSubscriberFormats = TOleEnum;
const
  wdSubscriberBestFormat = $00000000;
  wdSubscriberRTF = $00000001;
  wdSubscriberText = $00000002;
  wdSubscriberPict = $00000004;

// WdEditionType constants
type
  WdEditionType = TOleEnum;
const
  wdPublisher = $00000000;
  wdSubscriber = $00000001;

// WdEditionOption constants
type
  WdEditionOption = TOleEnum;
const
  wdCancelPublisher = $00000000;
  wdSendPublisher = $00000001;
  wdSelectPublisher = $00000002;
  wdAutomaticUpdate = $00000003;
  wdManualUpdate = $00000004;
  wdChangeAttributes = $00000005;
  wdUpdateSubscriber = $00000006;
  wdOpenSource = $00000007;

// WdRelativeHorizontalPosition constants
type
  WdRelativeHorizontalPosition = TOleEnum;
const
  wdRelativeHorizontalPositionMargin = $00000000;
  wdRelativeHorizontalPositionPage = $00000001;
  wdRelativeHorizontalPositionColumn = $00000002;
  wdRelativeHorizontalPositionCharacter = $00000003;

// WdRelativeVerticalPosition constants
type
  WdRelativeVerticalPosition = TOleEnum;
const
  wdRelativeVerticalPositionMargin = $00000000;
  wdRelativeVerticalPositionPage = $00000001;
  wdRelativeVerticalPositionParagraph = $00000002;
  wdRelativeVerticalPositionLine = $00000003;

// WdHelpType constants
type
  WdHelpType = TOleEnum;
const
  wdHelp = $00000000;
  wdHelpAbout = $00000001;
  wdHelpActiveWindow = $00000002;
  wdHelpContents = $00000003;
  wdHelpExamplesAndDemos = $00000004;
  wdHelpIndex = $00000005;
  wdHelpKeyboard = $00000006;
  wdHelpPSSHelp = $00000007;
  wdHelpQuickPreview = $00000008;
  wdHelpSearch = $00000009;
  wdHelpUsingHelp = $0000000A;
  wdHelpIchitaro = $0000000B;
  wdHelpPE2 = $0000000C;
  wdHelpHWP = $0000000D;

// WdHelpTypeHID constants
type
  WdHelpTypeHID = TOleEnum;
const
  emptyenum_____________ = $00000000;

// WdKeyCategory constants
type
  WdKeyCategory = TOleEnum;
const
  wdKeyCategoryNil = $FFFFFFFF;
  wdKeyCategoryDisable = $00000000;
  wdKeyCategoryCommand = $00000001;
  wdKeyCategoryMacro = $00000002;
  wdKeyCategoryFont = $00000003;
  wdKeyCategoryAutoText = $00000004;
  wdKeyCategoryStyle = $00000005;
  wdKeyCategorySymbol = $00000006;
  wdKeyCategoryPrefix = $00000007;

// WdKey constants
type
  WdKey = TOleEnum;
const
  wdNoKey = $000000FF;
  wdKeyShift = $00000100;
  wdKeyControl = $00000200;
  wdKeyCommand = $00000200;
  wdKeyAlt = $00000400;
  wdKeyOption = $00000400;
  wdKeyA = $00000041;
  wdKeyB = $00000042;
  wdKeyC = $00000043;
  wdKeyD = $00000044;
  wdKeyE = $00000045;
  wdKeyF = $00000046;
  wdKeyG = $00000047;
  wdKeyH = $00000048;
  wdKeyI = $00000049;
  wdKeyJ = $0000004A;
  wdKeyK = $0000004B;
  wdKeyL = $0000004C;
  wdKeyM = $0000004D;
  wdKeyN = $0000004E;
  wdKeyO = $0000004F;
  wdKeyP = $00000050;
  wdKeyQ = $00000051;
  wdKeyR = $00000052;
  wdKeyS = $00000053;
  wdKeyT = $00000054;
  wdKeyU = $00000055;
  wdKeyV = $00000056;
  wdKeyW = $00000057;
  wdKeyX = $00000058;
  wdKeyY = $00000059;
  wdKeyZ = $0000005A;
  wdKey0 = $00000030;
  wdKey1 = $00000031;
  wdKey2 = $00000032;
  wdKey3 = $00000033;
  wdKey4 = $00000034;
  wdKey5 = $00000035;
  wdKey6 = $00000036;
  wdKey7 = $00000037;
  wdKey8 = $00000038;
  wdKey9 = $00000039;
  wdKeyBackspace = $00000008;
  wdKeyTab = $00000009;
  wdKeyNumeric5Special = $0000000C;
  wdKeyReturn = $0000000D;
  wdKeyPause = $00000013;
  wdKeyEsc = $0000001B;
  wdKeySpacebar = $00000020;
  wdKeyPageUp = $00000021;
  wdKeyPageDown = $00000022;
  wdKeyEnd = $00000023;
  wdKeyHome = $00000024;
  wdKeyInsert = $0000002D;
  wdKeyDelete = $0000002E;
  wdKeyNumeric0 = $00000060;
  wdKeyNumeric1 = $00000061;
  wdKeyNumeric2 = $00000062;
  wdKeyNumeric3 = $00000063;
  wdKeyNumeric4 = $00000064;
  wdKeyNumeric5 = $00000065;
  wdKeyNumeric6 = $00000066;
  wdKeyNumeric7 = $00000067;
  wdKeyNumeric8 = $00000068;
  wdKeyNumeric9 = $00000069;
  wdKeyNumericMultiply = $0000006A;
  wdKeyNumericAdd = $0000006B;
  wdKeyNumericSubtract = $0000006D;
  wdKeyNumericDecimal = $0000006E;
  wdKeyNumericDivide = $0000006F;
  wdKeyF1 = $00000070;
  wdKeyF2 = $00000071;
  wdKeyF3 = $00000072;
  wdKeyF4 = $00000073;
  wdKeyF5 = $00000074;
  wdKeyF6 = $00000075;
  wdKeyF7 = $00000076;
  wdKeyF8 = $00000077;
  wdKeyF9 = $00000078;
  wdKeyF10 = $00000079;
  wdKeyF11 = $0000007A;
  wdKeyF12 = $0000007B;
  wdKeyF13 = $0000007C;
  wdKeyF14 = $0000007D;
  wdKeyF15 = $0000007E;
  wdKeyF16 = $0000007F;
  wdKeyScrollLock = $00000091;
  wdKeySemiColon = $000000BA;
  wdKeyEquals = $000000BB;
  wdKeyComma = $000000BC;
  wdKeyHyphen = $000000BD;
  wdKeyPeriod = $000000BE;
  wdKeySlash = $000000BF;
  wdKeyBackSingleQuote = $000000C0;
  wdKeyOpenSquareBrace = $000000DB;
  wdKeyBackSlash = $000000DC;
  wdKeyCloseSquareBrace = $000000DD;
  wdKeySingleQuote = $000000DE;

// WdOLEType constants
type
  WdOLEType = TOleEnum;
const
  wdOLELink = $00000000;
  wdOLEEmbed = $00000001;
  wdOLEControl = $00000002;

// WdOLEVerb constants
type
  WdOLEVerb = TOleEnum;
const
  wdOLEVerbPrimary = $00000000;
  wdOLEVerbShow = $FFFFFFFF;
  wdOLEVerbOpen = $FFFFFFFE;
  wdOLEVerbHide = $FFFFFFFD;
  wdOLEVerbUIActivate = $FFFFFFFC;
  wdOLEVerbInPlaceActivate = $FFFFFFFB;
  wdOLEVerbDiscardUndoState = $FFFFFFFA;

// WdOLEPlacement constants
type
  WdOLEPlacement = TOleEnum;
const
  wdInLine = $00000000;
  wdFloatOverText = $00000001;

// WdEnvelopeOrientation constants
type
  WdEnvelopeOrientation = TOleEnum;
const
  wdLeftPortrait = $00000000;
  wdCenterPortrait = $00000001;
  wdRightPortrait = $00000002;
  wdLeftLandscape = $00000003;
  wdCenterLandscape = $00000004;
  wdRightLandscape = $00000005;
  wdLeftClockwise = $00000006;
  wdCenterClockwise = $00000007;
  wdRightClockwise = $00000008;

// WdLetterStyle constants
type
  WdLetterStyle = TOleEnum;
const
  wdFullBlock = $00000000;
  wdModifiedBlock = $00000001;
  wdSemiBlock = $00000002;

// WdLetterheadLocation constants
type
  WdLetterheadLocation = TOleEnum;
const
  wdLetterTop = $00000000;
  wdLetterBottom = $00000001;
  wdLetterLeft = $00000002;
  wdLetterRight = $00000003;

// WdSalutationType constants
type
  WdSalutationType = TOleEnum;
const
  wdSalutationInformal = $00000000;
  wdSalutationFormal = $00000001;
  wdSalutationBusiness = $00000002;
  wdSalutationOther = $00000003;

// WdSalutationGender constants
type
  WdSalutationGender = TOleEnum;
const
  wdGenderFemale = $00000000;
  wdGenderMale = $00000001;
  wdGenderNeutral = $00000002;
  wdGenderUnknown = $00000003;

// WdMovementType constants
type
  WdMovementType = TOleEnum;
const
  wdMove = $00000000;
  wdExtend = $00000001;

// WdConstants constants
type
  WdConstants = TOleEnum;
const
  wdUndefined = $0098967F;
  wdToggle = $0098967E;
  wdForward = $3FFFFFFF;
  wdBackward = $C0000001;
  wdAutoPosition = $00000000;
  wdFirst = $00000001;
  wdCreatorCode = $4D535744;

// WdPasteDataType constants
type
  WdPasteDataType = TOleEnum;
const
  wdPasteOLEObject = $00000000;
  wdPasteRTF = $00000001;
  wdPasteText = $00000002;
  wdPasteMetafilePicture = $00000003;
  wdPasteBitmap = $00000004;
  wdPasteDeviceIndependentBitmap = $00000005;
  wdPasteHyperlink = $00000007;
  wdPasteShape = $00000008;
  wdPasteEnhancedMetafile = $00000009;
  wdPasteHTML = $0000000A;

// WdPrintOutItem constants
type
  WdPrintOutItem = TOleEnum;
const
  wdPrintDocumentContent = $00000000;
  wdPrintProperties = $00000001;
  wdPrintComments = $00000002;
  wdPrintStyles = $00000003;
  wdPrintAutoTextEntries = $00000004;
  wdPrintKeyAssignments = $00000005;
  wdPrintEnvelope = $00000006;

// WdPrintOutPages constants
type
  WdPrintOutPages = TOleEnum;
const
  wdPrintAllPages = $00000000;
  wdPrintOddPagesOnly = $00000001;
  wdPrintEvenPagesOnly = $00000002;

// WdPrintOutRange constants
type
  WdPrintOutRange = TOleEnum;
const
  wdPrintAllDocument = $00000000;
  wdPrintSelection = $00000001;
  wdPrintCurrentPage = $00000002;
  wdPrintFromTo = $00000003;
  wdPrintRangeOfPages = $00000004;

// WdDictionaryType constants
type
  WdDictionaryType = TOleEnum;
const
  wdSpelling = $00000000;
  wdGrammar = $00000001;
  wdThesaurus = $00000002;
  wdHyphenation = $00000003;
  wdSpellingComplete = $00000004;
  wdSpellingCustom = $00000005;
  wdSpellingLegal = $00000006;
  wdSpellingMedical = $00000007;
  wdHangulHanjaConversion = $00000008;
  wdHangulHanjaConversionCustom = $00000009;

// WdDictionaryTypeHID constants
type
  WdDictionaryTypeHID = TOleEnum;
const
  emptyenum______________ = $00000000;

// WdSpellingWordType constants
type
  WdSpellingWordType = TOleEnum;
const
  wdSpellword = $00000000;
  wdWildcard = $00000001;
  wdAnagram = $00000002;

// WdSpellingErrorType constants
type
  WdSpellingErrorType = TOleEnum;
const
  wdSpellingCorrect = $00000000;
  wdSpellingNotInDictionary = $00000001;
  wdSpellingCapitalization = $00000002;

// WdProofreadingErrorType constants
type
  WdProofreadingErrorType = TOleEnum;
const
  wdSpellingError = $00000000;
  wdGrammaticalError = $00000001;

// WdInlineShapeType constants
type
  WdInlineShapeType = TOleEnum;
const
  wdInlineShapeEmbeddedOLEObject = $00000001;
  wdInlineShapeLinkedOLEObject = $00000002;
  wdInlineShapePicture = $00000003;
  wdInlineShapeLinkedPicture = $00000004;
  wdInlineShapeOLEControlObject = $00000005;
  wdInlineShapeHorizontalLine = $00000006;
  wdInlineShapePictureHorizontalLine = $00000007;
  wdInlineShapeLinkedPictureHorizontalLine = $00000008;
  wdInlineShapePictureBullet = $00000009;
  wdInlineShapeScriptAnchor = $0000000A;
  wdInlineShapeOWSAnchor = $0000000B;

// WdArrangeStyle constants
type
  WdArrangeStyle = TOleEnum;
const
  wdTiled = $00000000;
  wdIcons = $00000001;

// WdSelectionFlags constants
type
  WdSelectionFlags = TOleEnum;
const
  wdSelStartActive = $00000001;
  wdSelAtEOL = $00000002;
  wdSelOvertype = $00000004;
  wdSelActive = $00000008;
  wdSelReplace = $00000010;

// WdAutoVersions constants
type
  WdAutoVersions = TOleEnum;
const
  wdAutoVersionOff = $00000000;
  wdAutoVersionOnClose = $00000001;

// WdOrganizerObject constants
type
  WdOrganizerObject = TOleEnum;
const
  wdOrganizerObjectStyles = $00000000;
  wdOrganizerObjectAutoText = $00000001;
  wdOrganizerObjectCommandBars = $00000002;
  wdOrganizerObjectProjectItems = $00000003;

// WdFindMatch constants
type
  WdFindMatch = TOleEnum;
const
  wdMatchParagraphMark = $0001000F;
  wdMatchTabCharacter = $00000009;
  wdMatchCommentMark = $00000005;
  wdMatchAnyCharacter = $0001003F;
  wdMatchAnyDigit = $0001001F;
  wdMatchAnyLetter = $0001002F;
  wdMatchCaretCharacter = $0000000B;
  wdMatchColumnBreak = $0000000E;
  wdMatchEmDash = $00002014;
  wdMatchEnDash = $00002013;
  wdMatchEndnoteMark = $00010013;
  wdMatchField = $00000013;
  wdMatchFootnoteMark = $00010012;
  wdMatchGraphic = $00000001;
  wdMatchManualLineBreak = $0001000F;
  wdMatchManualPageBreak = $0001001C;
  wdMatchNonbreakingHyphen = $0000001E;
  wdMatchNonbreakingSpace = $000000A0;
  wdMatchOptionalHyphen = $0000001F;
  wdMatchSectionBreak = $0001002C;
  wdMatchWhiteSpace = $00010077;

// WdFindWrap constants
type
  WdFindWrap = TOleEnum;
const
  wdFindStop = $00000000;
  wdFindContinue = $00000001;
  wdFindAsk = $00000002;

// WdInformation constants
type
  WdInformation = TOleEnum;
const
  wdActiveEndAdjustedPageNumber = $00000001;
  wdActiveEndSectionNumber = $00000002;
  wdActiveEndPageNumber = $00000003;
  wdNumberOfPagesInDocument = $00000004;
  wdHorizontalPositionRelativeToPage = $00000005;
  wdVerticalPositionRelativeToPage = $00000006;
  wdHorizontalPositionRelativeToTextBoundary = $00000007;
  wdVerticalPositionRelativeToTextBoundary = $00000008;
  wdFirstCharacterColumnNumber = $00000009;
  wdFirstCharacterLineNumber = $0000000A;
  wdFrameIsSelected = $0000000B;
  wdWithInTable = $0000000C;
  wdStartOfRangeRowNumber = $0000000D;
  wdEndOfRangeRowNumber = $0000000E;
  wdMaximumNumberOfRows = $0000000F;
  wdStartOfRangeColumnNumber = $00000010;
  wdEndOfRangeColumnNumber = $00000011;
  wdMaximumNumberOfColumns = $00000012;
  wdZoomPercentage = $00000013;
  wdSelectionMode = $00000014;
  wdCapsLock = $00000015;
  wdNumLock = $00000016;
  wdOverType = $00000017;
  wdRevisionMarking = $00000018;
  wdInFootnoteEndnotePane = $00000019;
  wdInCommentPane = $0000001A;
  wdInHeaderFooter = $0000001C;
  wdAtEndOfRowMarker = $0000001F;
  wdReferenceOfType = $00000020;
  wdHeaderFooterType = $00000021;
  wdInMasterDocument = $00000022;
  wdInFootnote = $00000023;
  wdInEndnote = $00000024;
  wdInWordMail = $00000025;
  wdInClipboard = $00000026;

// WdWrapType constants
type
  WdWrapType = TOleEnum;
const
  wdWrapSquare = $00000000;
  wdWrapTight = $00000001;
  wdWrapThrough = $00000002;
  wdWrapNone = $00000003;
  wdWrapTopBottom = $00000004;

// WdWrapSideType constants
type
  WdWrapSideType = TOleEnum;
const
  wdWrapBoth = $00000000;
  wdWrapLeft = $00000001;
  wdWrapRight = $00000002;
  wdWrapLargest = $00000003;

// WdOutlineLevel constants
type
  WdOutlineLevel = TOleEnum;
const
  wdOutlineLevel1 = $00000001;
  wdOutlineLevel2 = $00000002;
  wdOutlineLevel3 = $00000003;
  wdOutlineLevel4 = $00000004;
  wdOutlineLevel5 = $00000005;
  wdOutlineLevel6 = $00000006;
  wdOutlineLevel7 = $00000007;
  wdOutlineLevel8 = $00000008;
  wdOutlineLevel9 = $00000009;
  wdOutlineLevelBodyText = $0000000A;

// WdTextOrientation constants
type
  WdTextOrientation = TOleEnum;
const
  wdTextOrientationHorizontal = $00000000;
  wdTextOrientationUpward = $00000002;
  wdTextOrientationDownward = $00000003;
  wdTextOrientationVerticalFarEast = $00000001;
  wdTextOrientationHorizontalRotatedFarEast = $00000004;

// WdTextOrientationHID constants
type
  WdTextOrientationHID = TOleEnum;
const
  emptyenum_______________ = $00000000;

// WdPageBorderArt constants
type
  WdPageBorderArt = TOleEnum;
const
  wdArtApples = $00000001;
  wdArtMapleMuffins = $00000002;
  wdArtCakeSlice = $00000003;
  wdArtCandyCorn = $00000004;
  wdArtIceCreamCones = $00000005;
  wdArtChampagneBottle = $00000006;
  wdArtPartyGlass = $00000007;
  wdArtChristmasTree = $00000008;
  wdArtTrees = $00000009;
  wdArtPalmsColor = $0000000A;
  wdArtBalloons3Colors = $0000000B;
  wdArtBalloonsHotAir = $0000000C;
  wdArtPartyFavor = $0000000D;
  wdArtConfettiStreamers = $0000000E;
  wdArtHearts = $0000000F;
  wdArtHeartBalloon = $00000010;
  wdArtStars3D = $00000011;
  wdArtStarsShadowed = $00000012;
  wdArtStars = $00000013;
  wdArtSun = $00000014;
  wdArtEarth2 = $00000015;
  wdArtEarth1 = $00000016;
  wdArtPeopleHats = $00000017;
  wdArtSombrero = $00000018;
  wdArtPencils = $00000019;
  wdArtPackages = $0000001A;
  wdArtClocks = $0000001B;
  wdArtFirecrackers = $0000001C;
  wdArtRings = $0000001D;
  wdArtMapPins = $0000001E;
  wdArtConfetti = $0000001F;
  wdArtCreaturesButterfly = $00000020;
  wdArtCreaturesLadyBug = $00000021;
  wdArtCreaturesFish = $00000022;
  wdArtBirdsFlight = $00000023;
  wdArtScaredCat = $00000024;
  wdArtBats = $00000025;
  wdArtFlowersRoses = $00000026;
  wdArtFlowersRedRose = $00000027;
  wdArtPoinsettias = $00000028;
  wdArtHolly = $00000029;
  wdArtFlowersTiny = $0000002A;
  wdArtFlowersPansy = $0000002B;
  wdArtFlowersModern2 = $0000002C;
  wdArtFlowersModern1 = $0000002D;
  wdArtWhiteFlowers = $0000002E;
  wdArtVine = $0000002F;
  wdArtFlowersDaisies = $00000030;
  wdArtFlowersBlockPrint = $00000031;
  wdArtDecoArchColor = $00000032;
  wdArtFans = $00000033;
  wdArtFilm = $00000034;
  wdArtLightning1 = $00000035;
  wdArtCompass = $00000036;
  wdArtDoubleD = $00000037;
  wdArtClassicalWave = $00000038;
  wdArtShadowedSquares = $00000039;
  wdArtTwistedLines1 = $0000003A;
  wdArtWaveline = $0000003B;
  wdArtQuadrants = $0000003C;
  wdArtCheckedBarColor = $0000003D;
  wdArtSwirligig = $0000003E;
  wdArtPushPinNote1 = $0000003F;
  wdArtPushPinNote2 = $00000040;
  wdArtPumpkin1 = $00000041;
  wdArtEggsBlack = $00000042;
  wdArtCup = $00000043;
  wdArtHeartGray = $00000044;
  wdArtGingerbreadMan = $00000045;
  wdArtBabyPacifier = $00000046;
  wdArtBabyRattle = $00000047;
  wdArtCabins = $00000048;
  wdArtHouseFunky = $00000049;
  wdArtStarsBlack = $0000004A;
  wdArtSnowflakes = $0000004B;
  wdArtSnowflakeFancy = $0000004C;
  wdArtSkyrocket = $0000004D;
  wdArtSeattle = $0000004E;
  wdArtMusicNotes = $0000004F;
  wdArtPalmsBlack = $00000050;
  wdArtMapleLeaf = $00000051;
  wdArtPaperClips = $00000052;
  wdArtShorebirdTracks = $00000053;
  wdArtPeople = $00000054;
  wdArtPeopleWaving = $00000055;
  wdArtEclipsingSquares2 = $00000056;
  wdArtHypnotic = $00000057;
  wdArtDiamondsGray = $00000058;
  wdArtDecoArch = $00000059;
  wdArtDecoBlocks = $0000005A;
  wdArtCirclesLines = $0000005B;
  wdArtPapyrus = $0000005C;
  wdArtWoodwork = $0000005D;
  wdArtWeavingBraid = $0000005E;
  wdArtWeavingRibbon = $0000005F;
  wdArtWeavingAngles = $00000060;
  wdArtArchedScallops = $00000061;
  wdArtSafari = $00000062;
  wdArtCelticKnotwork = $00000063;
  wdArtCrazyMaze = $00000064;
  wdArtEclipsingSquares1 = $00000065;
  wdArtBirds = $00000066;
  wdArtFlowersTeacup = $00000067;
  wdArtNorthwest = $00000068;
  wdArtSouthwest = $00000069;
  wdArtTribal6 = $0000006A;
  wdArtTribal4 = $0000006B;
  wdArtTribal3 = $0000006C;
  wdArtTribal2 = $0000006D;
  wdArtTribal5 = $0000006E;
  wdArtXIllusions = $0000006F;
  wdArtZanyTriangles = $00000070;
  wdArtPyramids = $00000071;
  wdArtPyramidsAbove = $00000072;
  wdArtConfettiGrays = $00000073;
  wdArtConfettiOutline = $00000074;
  wdArtConfettiWhite = $00000075;
  wdArtMosaic = $00000076;
  wdArtLightning2 = $00000077;
  wdArtHeebieJeebies = $00000078;
  wdArtLightBulb = $00000079;
  wdArtGradient = $0000007A;
  wdArtTriangleParty = $0000007B;
  wdArtTwistedLines2 = $0000007C;
  wdArtMoons = $0000007D;
  wdArtOvals = $0000007E;
  wdArtDoubleDiamonds = $0000007F;
  wdArtChainLink = $00000080;
  wdArtTriangles = $00000081;
  wdArtTribal1 = $00000082;
  wdArtMarqueeToothed = $00000083;
  wdArtSharksTeeth = $00000084;
  wdArtSawtooth = $00000085;
  wdArtSawtoothGray = $00000086;
  wdArtPostageStamp = $00000087;
  wdArtWeavingStrips = $00000088;
  wdArtZigZag = $00000089;
  wdArtCrossStitch = $0000008A;
  wdArtGems = $0000008B;
  wdArtCirclesRectangles = $0000008C;
  wdArtCornerTriangles = $0000008D;
  wdArtCreaturesInsects = $0000008E;
  wdArtZigZagStitch = $0000008F;
  wdArtCheckered = $00000090;
  wdArtCheckedBarBlack = $00000091;
  wdArtMarquee = $00000092;
  wdArtBasicWhiteDots = $00000093;
  wdArtBasicWideMidline = $00000094;
  wdArtBasicWideOutline = $00000095;
  wdArtBasicWideInline = $00000096;
  wdArtBasicThinLines = $00000097;
  wdArtBasicWhiteDashes = $00000098;
  wdArtBasicWhiteSquares = $00000099;
  wdArtBasicBlackSquares = $0000009A;
  wdArtBasicBlackDashes = $0000009B;
  wdArtBasicBlackDots = $0000009C;
  wdArtStarsTop = $0000009D;
  wdArtCertificateBanner = $0000009E;
  wdArtHandmade1 = $0000009F;
  wdArtHandmade2 = $000000A0;
  wdArtTornPaper = $000000A1;
  wdArtTornPaperBlack = $000000A2;
  wdArtCouponCutoutDashes = $000000A3;
  wdArtCouponCutoutDots = $000000A4;

// WdBorderDistanceFrom constants
type
  WdBorderDistanceFrom = TOleEnum;
const
  wdBorderDistanceFromText = $00000000;
  wdBorderDistanceFromPageEdge = $00000001;

// WdReplace constants
type
  WdReplace = TOleEnum;
const
  wdReplaceNone = $00000000;
  wdReplaceOne = $00000001;
  wdReplaceAll = $00000002;

// WdFontBias constants
type
  WdFontBias = TOleEnum;
const
  wdFontBiasDontCare = $000000FF;
  wdFontBiasDefault = $00000000;
  wdFontBiasFareast = $00000001;

// WdBrowserLevel constants
type
  WdBrowserLevel = TOleEnum;
const
  wdBrowserLevelV4 = $00000000;
  wdBrowserLevelMicrosoftInternetExplorer5 = $00000001;

// WdEnclosureType constants
type
  WdEnclosureType = TOleEnum;
const
  wdEnclosureCircle = $00000000;
  wdEnclosureSquare = $00000001;
  wdEnclosureTriangle = $00000002;
  wdEnclosureDiamond = $00000003;

// WdEncloseStyle constants
type
  WdEncloseStyle = TOleEnum;
const
  wdEncloseStyleNone = $00000000;
  wdEncloseStyleSmall = $00000001;
  wdEncloseStyleLarge = $00000002;

// WdHighAnsiText constants
type
  WdHighAnsiText = TOleEnum;
const
  wdHighAnsiIsFarEast = $00000000;
  wdHighAnsiIsHighAnsi = $00000001;
  wdAutoDetectHighAnsiFarEast = $00000002;

// WdLayoutMode constants
type
  WdLayoutMode = TOleEnum;
const
  wdLayoutModeDefault = $00000000;
  wdLayoutModeGrid = $00000001;
  wdLayoutModeLineGrid = $00000002;
  wdLayoutModeGenko = $00000003;

// WdDocumentMedium constants
type
  WdDocumentMedium = TOleEnum;
const
  wdEmailMessage = $00000000;
  wdDocument = $00000001;
  wdWebPage = $00000002;

// WdMailerPriority constants
type
  WdMailerPriority = TOleEnum;
const
  wdPriorityNormal = $00000001;
  wdPriorityLow = $00000002;
  wdPriorityHigh = $00000003;

// WdDocumentViewDirection constants
type
  WdDocumentViewDirection = TOleEnum;
const
  wdDocumentViewRtl = $00000000;
  wdDocumentViewLtr = $00000001;

// WdArabicNumeral constants
type
  WdArabicNumeral = TOleEnum;
const
  wdNumeralArabic = $00000000;
  wdNumeralHindi = $00000001;
  wdNumeralContext = $00000002;
  wdNumeralSystem = $00000003;

// WdMonthNames constants
type
  WdMonthNames = TOleEnum;
const
  wdMonthNamesArabic = $00000000;
  wdMonthNamesEnglish = $00000001;
  wdMonthNamesFrench = $00000002;

// WdCursorMovement constants
type
  WdCursorMovement = TOleEnum;
const
  wdCursorMovementLogical = $00000000;
  wdCursorMovementVisual = $00000001;

// WdVisualSelection constants
type
  WdVisualSelection = TOleEnum;
const
  wdVisualSelectionBlock = $00000000;
  wdVisualSelectionContinuous = $00000001;

// WdTableDirection constants
type
  WdTableDirection = TOleEnum;
const
  wdTableDirectionRtl = $00000000;
  wdTableDirectionLtr = $00000001;

// WdFlowDirection constants
type
  WdFlowDirection = TOleEnum;
const
  wdFlowLtr = $00000000;
  wdFlowRtl = $00000001;

// WdDiacriticColor constants
type
  WdDiacriticColor = TOleEnum;
const
  wdDiacriticColorBidi = $00000000;
  wdDiacriticColorLatin = $00000001;

// WdGutterStyle constants
type
  WdGutterStyle = TOleEnum;
const
  wdGutterPosLeft = $00000000;
  wdGutterPosTop = $00000001;
  wdGutterPosRight = $00000002;

// WdGutterStyleOld constants
type
  WdGutterStyleOld = TOleEnum;
const
  wdGutterStyleLatin = $FFFFFFF6;
  wdGutterStyleBidi = $00000002;

// WdSectionDirection constants
type
  WdSectionDirection = TOleEnum;
const
  wdSectionDirectionRtl = $00000000;
  wdSectionDirectionLtr = $00000001;

// WdDateLanguage constants
type
  WdDateLanguage = TOleEnum;
const
  wdDateLanguageBidi = $0000000A;
  wdDateLanguageLatin = $00000409;

// WdCalendarTypeBi constants
type
  WdCalendarTypeBi = TOleEnum;
const
  wdCalendarTypeBidi = $00000063;
  wdCalendarTypeGregorian = $00000064;

// WdCalendarType constants
type
  WdCalendarType = TOleEnum;
const
  wdCalendarWestern = $00000000;
  wdCalendarArabic = $00000001;
  wdCalendarHebrew = $00000002;
  wdCalendarChina = $00000003;
  wdCalendarJapan = $00000004;
  wdCalendarThai = $00000005;
  wdCalendarKorean = $00000006;

// WdReadingOrder constants
type
  WdReadingOrder = TOleEnum;
const
  wdReadingOrderRtl = $00000000;
  wdReadingOrderLtr = $00000001;

// WdHebSpellStart constants
type
  WdHebSpellStart = TOleEnum;
const
  wdFullScript = $00000000;
  wdPartialScript = $00000001;
  wdMixedScript = $00000002;
  wdMixedAuthorizedScript = $00000003;

// WdAraSpeller constants
type
  WdAraSpeller = TOleEnum;
const
  wdNone = $00000000;
  wdInitalAlef = $00000001;
  wdFinalYaa = $00000002;
  wdBoth = $00000003;

// WdColor constants
type
  WdColor = TOleEnum;
const
  wdColorAutomatic = $FF000000;
  wdColorBlack = $00000000;
  wdColorBlue = $00FF0000;
  wdColorTurquoise = $00FFFF00;
  wdColorBrightGreen = $0000FF00;
  wdColorPink = $00FF00FF;
  wdColorRed = $000000FF;
  wdColorYellow = $0000FFFF;
  wdColorWhite = $00FFFFFF;
  wdColorDarkBlue = $00800000;
  wdColorTeal = $00808000;
  wdColorGreen = $00008000;
  wdColorViolet = $00800080;
  wdColorDarkRed = $00000080;
  wdColorDarkYellow = $00008080;
  wdColorBrown = $00003399;
  wdColorOliveGreen = $00003333;
  wdColorDarkGreen = $00003300;
  wdColorDarkTeal = $00663300;
  wdColorIndigo = $00993333;
  wdColorOrange = $000066FF;
  wdColorBlueGray = $00996666;
  wdColorLightOrange = $000099FF;
  wdColorLime = $0000CC99;
  wdColorSeaGreen = $00669933;
  wdColorAqua = $00CCCC33;
  wdColorLightBlue = $00FF6633;
  wdColorGold = $0000CCFF;
  wdColorSkyBlue = $00FFCC00;
  wdColorPlum = $00663399;
  wdColorRose = $00CC99FF;
  wdColorTan = $0099CCFF;
  wdColorLightYellow = $0099FFFF;
  wdColorLightGreen = $00CCFFCC;
  wdColorLightTurquoise = $00FFFFCC;
  wdColorPaleBlue = $00FFCC99;
  wdColorLavender = $00FF99CC;
  wdColorGray05 = $00F3F3F3;
  wdColorGray10 = $00E6E6E6;
  wdColorGray125 = $00E0E0E0;
  wdColorGray15 = $00D9D9D9;
  wdColorGray20 = $00CCCCCC;
  wdColorGray25 = $00C0C0C0;
  wdColorGray30 = $00B3B3B3;
  wdColorGray35 = $00A6A6A6;
  wdColorGray375 = $00A0A0A0;
  wdColorGray40 = $00999999;
  wdColorGray45 = $008C8C8C;
  wdColorGray50 = $00808080;
  wdColorGray55 = $00737373;
  wdColorGray60 = $00666666;
  wdColorGray625 = $00606060;
  wdColorGray65 = $00595959;
  wdColorGray70 = $004C4C4C;
  wdColorGray75 = $00404040;
  wdColorGray80 = $00333333;
  wdColorGray85 = $00262626;
  wdColorGray875 = $00202020;
  wdColorGray90 = $00191919;
  wdColorGray95 = $000C0C0C;

// WdShapePosition constants
type
  WdShapePosition = TOleEnum;
const
  wdShapeTop = $FFF0BDC1;
  wdShapeLeft = $FFF0BDC2;
  wdShapeBottom = $FFF0BDC3;
  wdShapeRight = $FFF0BDC4;
  wdShapeCenter = $FFF0BDC5;
  wdShapeInside = $FFF0BDC6;
  wdShapeOutside = $FFF0BDC7;

// WdTablePosition constants
type
  WdTablePosition = TOleEnum;
const
  wdTableTop = $FFF0BDC1;
  wdTableLeft = $FFF0BDC2;
  wdTableBottom = $FFF0BDC3;
  wdTableRight = $FFF0BDC4;
  wdTableCenter = $FFF0BDC5;
  wdTableInside = $FFF0BDC6;
  wdTableOutside = $FFF0BDC7;

// WdDefaultListBehavior constants
type
  WdDefaultListBehavior = TOleEnum;
const
  wdWord8ListBehavior = $00000000;
  wdWord9ListBehavior = $00000001;

// WdDefaultTableBehavior constants
type
  WdDefaultTableBehavior = TOleEnum;
const
  wdWord8TableBehavior = $00000000;
  wdWord9TableBehavior = $00000001;

// WdAutoFitBehavior constants
type
  WdAutoFitBehavior = TOleEnum;
const
  wdAutoFitFixed = $00000000;
  wdAutoFitContent = $00000001;
  wdAutoFitWindow = $00000002;

// WdPreferredWidthType constants
type
  WdPreferredWidthType = TOleEnum;
const
  wdPreferredWidthAuto = $00000001;
  wdPreferredWidthPercent = $00000002;
  wdPreferredWidthPoints = $00000003;

// WdFarEastLineBreakLanguageID constants
type
  WdFarEastLineBreakLanguageID = TOleEnum;
const
  wdLineBreakJapanese = $00000411;
  wdLineBreakKorean = $00000412;
  wdLineBreakSimplifiedChinese = $00000804;
  wdLineBreakTraditionalChinese = $00000404;

// WdViewTypeOld constants
type
  WdViewTypeOld = TOleEnum;
const
  wdPageView = $00000003;
  wdOnlineView = $00000006;

// WdFramesetType constants
type
  WdFramesetType = TOleEnum;
const
  wdFramesetTypeFrameset = $00000000;
  wdFramesetTypeFrame = $00000001;

// WdFramesetSizeType constants
type
  WdFramesetSizeType = TOleEnum;
const
  wdFramesetSizeTypePercent = $00000000;
  wdFramesetSizeTypeFixed = $00000001;
  wdFramesetSizeTypeRelative = $00000002;

// WdFramesetNewFrameLocation constants
type
  WdFramesetNewFrameLocation = TOleEnum;
const
  wdFramesetNewFrameAbove = $00000000;
  wdFramesetNewFrameBelow = $00000001;
  wdFramesetNewFrameRight = $00000002;
  wdFramesetNewFrameLeft = $00000003;

// WdScrollbarType constants
type
  WdScrollbarType = TOleEnum;
const
  wdScrollbarTypeAuto = $00000000;
  wdScrollbarTypeYes = $00000001;
  wdScrollbarTypeNo = $00000002;

// WdTwoLinesInOneType constants
type
  WdTwoLinesInOneType = TOleEnum;
const
  wdTwoLinesInOneNone = $00000000;
  wdTwoLinesInOneNoBrackets = $00000001;
  wdTwoLinesInOneParentheses = $00000002;
  wdTwoLinesInOneSquareBrackets = $00000003;
  wdTwoLinesInOneAngleBrackets = $00000004;
  wdTwoLinesInOneCurlyBrackets = $00000005;

// WdHorizontalInVerticalType constants
type
  WdHorizontalInVerticalType = TOleEnum;
const
  wdHorizontalInVerticalNone = $00000000;
  wdHorizontalInVerticalFitInLine = $00000001;
  wdHorizontalInVerticalResizeLine = $00000002;

// WdHorizontalLineAlignment constants
type
  WdHorizontalLineAlignment = TOleEnum;
const
  wdHorizontalLineAlignLeft = $00000000;
  wdHorizontalLineAlignCenter = $00000001;
  wdHorizontalLineAlignRight = $00000002;

// WdHorizontalLineWidthType constants
type
  WdHorizontalLineWidthType = TOleEnum;
const
  wdHorizontalLinePercentWidth = $FFFFFFFF;
  wdHorizontalLineFixedWidth = $FFFFFFFE;

// WdPhoneticGuideAlignmentType constants
type
  WdPhoneticGuideAlignmentType = TOleEnum;
const
  wdPhoneticGuideAlignmentCenter = $00000000;
  wdPhoneticGuideAlignmentZeroOneZero = $00000001;
  wdPhoneticGuideAlignmentOneTwoOne = $00000002;
  wdPhoneticGuideAlignmentLeft = $00000003;
  wdPhoneticGuideAlignmentRight = $00000004;

// WdNewDocumentType constants
type
  WdNewDocumentType = TOleEnum;
const
  wdNewBlankDocument = $00000000;
  wdNewWebPage = $00000001;
  wdNewEmailMessage = $00000002;
  wdNewFrameset = $00000003;

// WdKana constants
type
  WdKana = TOleEnum;
const
  wdKanaKatakana = $00000008;
  wdKanaHiragana = $00000009;

// WdCharacterWidth constants
type
  WdCharacterWidth = TOleEnum;
const
  wdWidthHalfWidth = $00000006;
  wdWidthFullWidth = $00000007;

// WdNumberStyleWordBasicBiDi constants
type
  WdNumberStyleWordBasicBiDi = TOleEnum;
const
  wdListNumberStyleBidi1 = $00000031;
  wdListNumberStyleBidi2 = $00000032;
  wdCaptionNumberStyleBidiLetter1 = $00000031;
  wdCaptionNumberStyleBidiLetter2 = $00000032;
  wdNoteNumberStyleBidiLetter1 = $00000031;
  wdNoteNumberStyleBidiLetter2 = $00000032;
  wdPageNumberStyleBidiLetter1 = $00000031;
  wdPageNumberStyleBidiLetter2 = $00000032;

// WdTCSCConverterDirection constants
type
  WdTCSCConverterDirection = TOleEnum;
const
  wdTCSCConverterDirectionSCTC = $00000000;
  wdTCSCConverterDirectionTCSC = $00000001;
  wdTCSCConverterDirectionAuto = $00000002;

type

// *********************************************************************//
// Forward declaration of interfaces defined in Type Library            //
// *********************************************************************//
  _Application = interface;
  _ApplicationDisp = dispinterface;
//  _Global = interface;
//  _GlobalDisp = dispinterface;
  FontNames = interface;
  FontNamesDisp = dispinterface;
  Languages = interface;
  LanguagesDisp = dispinterface;
  Language = interface;
  LanguageDisp = dispinterface;
  Documents = interface;
  DocumentsDisp = dispinterface;
  _Document = interface;
  _DocumentDisp = dispinterface;
  Template = interface;
  TemplateDisp = dispinterface;
  Templates = interface;
  TemplatesDisp = dispinterface;
  RoutingSlip = interface;
  RoutingSlipDisp = dispinterface;
  Bookmark = interface;
  BookmarkDisp = dispinterface;
  Bookmarks = interface;
  BookmarksDisp = dispinterface;
  Variable = interface;
  VariableDisp = dispinterface;
  Variables = interface;
  VariablesDisp = dispinterface;
  RecentFile = interface;
  RecentFileDisp = dispinterface;
  RecentFiles = interface;
  RecentFilesDisp = dispinterface;
  Window_ = interface;
  Window_Disp = dispinterface;
  Windows = interface;
  WindowsDisp = dispinterface;
  Pane = interface;
  PaneDisp = dispinterface;
  Panes = interface;
  PanesDisp = dispinterface;
  Range = interface;
  RangeDisp = dispinterface;
  ListFormat = interface;
  ListFormatDisp = dispinterface;
  Find = interface;
  FindDisp = dispinterface;
  Replacement = interface;
  ReplacementDisp = dispinterface;
  Characters = interface;
  CharactersDisp = dispinterface;
  Words = interface;
  WordsDisp = dispinterface;
  Sentences = interface;
  SentencesDisp = dispinterface;
  Sections = interface;
  SectionsDisp = dispinterface;
  Section = interface;
  SectionDisp = dispinterface;
  Paragraphs = interface;
  ParagraphsDisp = dispinterface;
  Paragraph = interface;
  ParagraphDisp = dispinterface;
  DropCap = interface;
  DropCapDisp = dispinterface;
  TabStops = interface;
  TabStopsDisp = dispinterface;
  TabStop = interface;
  TabStopDisp = dispinterface;
  _ParagraphFormat = interface;
  _ParagraphFormatDisp = dispinterface;
  _Font = interface;
  _FontDisp = dispinterface;
  Table = interface;
  TableDisp = dispinterface;
  Row = interface;
  RowDisp = dispinterface;
  Column = interface;
  ColumnDisp = dispinterface;
  Cell = interface;
  CellDisp = dispinterface;
  Tables = interface;
  TablesDisp = dispinterface;
  Rows = interface;
  RowsDisp = dispinterface;
  Columns = interface;
  ColumnsDisp = dispinterface;
  Cells = interface;
  CellsDisp = dispinterface;
  AutoCorrect = interface;
  AutoCorrectDisp = dispinterface;
  AutoCorrectEntries = interface;
  AutoCorrectEntriesDisp = dispinterface;
  AutoCorrectEntry = interface;
  AutoCorrectEntryDisp = dispinterface;
  FirstLetterExceptions = interface;
  FirstLetterExceptionsDisp = dispinterface;
  FirstLetterException = interface;
  FirstLetterExceptionDisp = dispinterface;
  TwoInitialCapsExceptions = interface;
  TwoInitialCapsExceptionsDisp = dispinterface;
  TwoInitialCapsException = interface;
  TwoInitialCapsExceptionDisp = dispinterface;
  Footnotes = interface;
  FootnotesDisp = dispinterface;
  Endnotes = interface;
  EndnotesDisp = dispinterface;
  Comments = interface;
  CommentsDisp = dispinterface;
  Footnote = interface;
  FootnoteDisp = dispinterface;
  Endnote = interface;
  EndnoteDisp = dispinterface;
  Comment = interface;
  CommentDisp = dispinterface;
  Borders = interface;
  BordersDisp = dispinterface;
  Border = interface;
  BorderDisp = dispinterface;
  Shading = interface;
  ShadingDisp = dispinterface;
  TextRetrievalMode = interface;
  TextRetrievalModeDisp = dispinterface;
  AutoTextEntries = interface;
  AutoTextEntriesDisp = dispinterface;
  AutoTextEntry = interface;
  AutoTextEntryDisp = dispinterface;
  System_ = interface;
  System_Disp = dispinterface;
  OLEFormat = interface;
  OLEFormatDisp = dispinterface;
  LinkFormat = interface;
  LinkFormatDisp = dispinterface;
  _OLEControl = interface;
  _OLEControlDisp = dispinterface;
  Fields = interface;
  FieldsDisp = dispinterface;
  Field = interface;
  FieldDisp = dispinterface;
  Browser = interface;
  BrowserDisp = dispinterface;
  Styles = interface;
  StylesDisp = dispinterface;
  Style = interface;
  StyleDisp = dispinterface;
  Frames = interface;
  FramesDisp = dispinterface;
  Frame = interface;
  FrameDisp = dispinterface;
  FormFields = interface;
  FormFieldsDisp = dispinterface;
  FormField = interface;
  FormFieldDisp = dispinterface;
  TextInput = interface;
  TextInputDisp = dispinterface;
  CheckBox = interface;
  CheckBoxDisp = dispinterface;
  DropDown = interface;
  DropDownDisp = dispinterface;
  ListEntries = interface;
  ListEntriesDisp = dispinterface;
  ListEntry = interface;
  ListEntryDisp = dispinterface;
  TablesOfFigures = interface;
  TablesOfFiguresDisp = dispinterface;
  TableOfFigures = interface;
  TableOfFiguresDisp = dispinterface;
  MailMerge = interface;
  MailMergeDisp = dispinterface;
  MailMergeFields = interface;
  MailMergeFieldsDisp = dispinterface;
  MailMergeField = interface;
  MailMergeFieldDisp = dispinterface;
  MailMergeDataSource = interface;
  MailMergeDataSourceDisp = dispinterface;
  MailMergeFieldNames = interface;
  MailMergeFieldNamesDisp = dispinterface;
  MailMergeFieldName = interface;
  MailMergeFieldNameDisp = dispinterface;
  MailMergeDataFields = interface;
  MailMergeDataFieldsDisp = dispinterface;
  MailMergeDataField = interface;
  MailMergeDataFieldDisp = dispinterface;
  Envelope = interface;
  EnvelopeDisp = dispinterface;
  MailingLabel = interface;
  MailingLabelDisp = dispinterface;
  CustomLabels = interface;
  CustomLabelsDisp = dispinterface;
  CustomLabel = interface;
  CustomLabelDisp = dispinterface;
  TablesOfContents = interface;
  TablesOfContentsDisp = dispinterface;
  TableOfContents = interface;
  TableOfContentsDisp = dispinterface;
  TablesOfAuthorities = interface;
  TablesOfAuthoritiesDisp = dispinterface;
  TableOfAuthorities = interface;
  TableOfAuthoritiesDisp = dispinterface;
  Dialogs = interface;
  DialogsDisp = dispinterface;
  Dialog = interface;
  DialogDisp = dispinterface;
  PageSetup = interface;
  PageSetupDisp = dispinterface;
  LineNumbering = interface;
  LineNumberingDisp = dispinterface;
  TextColumns = interface;
  TextColumnsDisp = dispinterface;
  TextColumn = interface;
  TextColumnDisp = dispinterface;
  Selection = interface;
  SelectionDisp = dispinterface;
  TablesOfAuthoritiesCategories = interface;
  TablesOfAuthoritiesCategoriesDisp = dispinterface;
  TableOfAuthoritiesCategory = interface;
  TableOfAuthoritiesCategoryDisp = dispinterface;
  CaptionLabels = interface;
  CaptionLabelsDisp = dispinterface;
  CaptionLabel = interface;
  CaptionLabelDisp = dispinterface;
  AutoCaptions = interface;
  AutoCaptionsDisp = dispinterface;
  AutoCaption = interface;
  AutoCaptionDisp = dispinterface;
  Indexes = interface;
  IndexesDisp = dispinterface;
  Index = interface;
  IndexDisp = dispinterface;
  AddIn = interface;
  AddInDisp = dispinterface;
  AddIns = interface;
  AddInsDisp = dispinterface;
  Revisions = interface;
  RevisionsDisp = dispinterface;
  Revision = interface;
  RevisionDisp = dispinterface;
  Task = interface;
  TaskDisp = dispinterface;
  Tasks = interface;
  TasksDisp = dispinterface;
  HeadersFooters = interface;
  HeadersFootersDisp = dispinterface;
  HeaderFooter = interface;
  HeaderFooterDisp = dispinterface;
  PageNumbers = interface;
  PageNumbersDisp = dispinterface;
  PageNumber = interface;
  PageNumberDisp = dispinterface;
  Subdocuments = interface;
  SubdocumentsDisp = dispinterface;
  Subdocument = interface;
  SubdocumentDisp = dispinterface;
  HeadingStyles = interface;
  HeadingStylesDisp = dispinterface;
  HeadingStyle = interface;
  HeadingStyleDisp = dispinterface;
  StoryRanges = interface;
  StoryRangesDisp = dispinterface;
  ListLevel = interface;
  ListLevelDisp = dispinterface;
  ListLevels = interface;
  ListLevelsDisp = dispinterface;
  ListTemplate = interface;
  ListTemplateDisp = dispinterface;
  ListTemplates = interface;
  ListTemplatesDisp = dispinterface;
  ListParagraphs = interface;
  ListParagraphsDisp = dispinterface;
  List = interface;
  ListDisp = dispinterface;
  Lists = interface;
  ListsDisp = dispinterface;
  ListGallery = interface;
  ListGalleryDisp = dispinterface;
  ListGalleries = interface;
  ListGalleriesDisp = dispinterface;
  KeyBindings = interface;
  KeyBindingsDisp = dispinterface;
  KeysBoundTo = interface;
  KeysBoundToDisp = dispinterface;
  KeyBinding = interface;
  KeyBindingDisp = dispinterface;
  FileConverter = interface;
  FileConverterDisp = dispinterface;
  FileConverters = interface;
  FileConvertersDisp = dispinterface;
  SynonymInfo = interface;
  SynonymInfoDisp = dispinterface;
  Hyperlinks = interface;
  HyperlinksDisp = dispinterface;
  Hyperlink = interface;
  HyperlinkDisp = dispinterface;
  Shapes = interface;
  ShapesDisp = dispinterface;
  ShapeRange = interface;
  ShapeRangeDisp = dispinterface;
  GroupShapes = interface;
  GroupShapesDisp = dispinterface;
  Shape = interface;
  ShapeDisp = dispinterface;
  TextFrame = interface;
  TextFrameDisp = dispinterface;
  _LetterContent = interface;
  _LetterContentDisp = dispinterface;
  View = interface;
  ViewDisp = dispinterface;
  Zoom = interface;
  ZoomDisp = dispinterface;
  Zooms = interface;
  ZoomsDisp = dispinterface;
  InlineShape = interface;
  InlineShapeDisp = dispinterface;
  InlineShapes = interface;
  InlineShapesDisp = dispinterface;
  SpellingSuggestions = interface;
  SpellingSuggestionsDisp = dispinterface;
  SpellingSuggestion = interface;
  SpellingSuggestionDisp = dispinterface;
  Dictionaries = interface;
  DictionariesDisp = dispinterface;
  HangulHanjaConversionDictionaries = interface;
  HangulHanjaConversionDictionariesDisp = dispinterface;
  Dictionary = interface;
  DictionaryDisp = dispinterface;
  ReadabilityStatistics = interface;
  ReadabilityStatisticsDisp = dispinterface;
  ReadabilityStatistic = interface;
  ReadabilityStatisticDisp = dispinterface;
  Versions = interface;
  VersionsDisp = dispinterface;
  Version = interface;
  VersionDisp = dispinterface;
  Options = interface;
  OptionsDisp = dispinterface;
  MailMessage = interface;
  MailMessageDisp = dispinterface;
  ProofreadingErrors = interface;
  ProofreadingErrorsDisp = dispinterface;
  Mailer = interface;
  MailerDisp = dispinterface;
  WrapFormat = interface;
  WrapFormatDisp = dispinterface;
  HangulAndAlphabetExceptions = interface;
  HangulAndAlphabetExceptionsDisp = dispinterface;
  HangulAndAlphabetException = interface;
  HangulAndAlphabetExceptionDisp = dispinterface;
  Adjustments = interface;
  AdjustmentsDisp = dispinterface;
  CalloutFormat = interface;
  CalloutFormatDisp = dispinterface;
  ColorFormat = interface;
  ColorFormatDisp = dispinterface;
  ConnectorFormat = interface;
  ConnectorFormatDisp = dispinterface;
  FillFormat = interface;
  FillFormatDisp = dispinterface;
  FreeformBuilder = interface;
  FreeformBuilderDisp = dispinterface;
  LineFormat = interface;
  LineFormatDisp = dispinterface;
  PictureFormat = interface;
  PictureFormatDisp = dispinterface;
  ShadowFormat = interface;
  ShadowFormatDisp = dispinterface;
  ShapeNode = interface;
  ShapeNodeDisp = dispinterface;
  ShapeNodes = interface;
  ShapeNodesDisp = dispinterface;
  TextEffectFormat = interface;
  TextEffectFormatDisp = dispinterface;
  ThreeDFormat = interface;
  ThreeDFormatDisp = dispinterface;
  ApplicationEvents = dispinterface;
  ApplicationEvents2 = dispinterface;
  DocumentEvents = dispinterface;
  OCXEvents = dispinterface;
  IApplicationEvents = interface;
  IApplicationEventsDisp = dispinterface;
  IApplicationEvents2 = interface;
  IApplicationEvents2Disp = dispinterface;
  EmailAuthor = interface;
  EmailAuthorDisp = dispinterface;
  EmailOptions = interface;
  EmailOptionsDisp = dispinterface;
  EmailSignature = interface;
  EmailSignatureDisp = dispinterface;
  Email = interface;
  EmailDisp = dispinterface;
  HorizontalLineFormat = interface;
  HorizontalLineFormatDisp = dispinterface;
  Frameset = interface;
  FramesetDisp = dispinterface;
  DefaultWebOptions = interface;
  DefaultWebOptionsDisp = dispinterface;
  WebOptions = interface;
  WebOptionsDisp = dispinterface;
  OtherCorrectionsExceptions = interface;
  OtherCorrectionsExceptionsDisp = dispinterface;
  OtherCorrectionsException = interface;
  OtherCorrectionsExceptionDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                     //
// (NOTE: Here we map each CoClass to its Default Interface)            //
// *********************************************************************//
//  Global = _Global;
  Document = _Document;
  Font = _Font;
  ParagraphFormat = _ParagraphFormat;
  OLEControl = _OLEControl;
  LetterContent = _LetterContent;
  Application_ = _Application;


// *********************************************************************//
// Declaration of structures, unions and aliases.                       //
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}
  PPSafeArray1 = ^PSafeArray; {*}
  PWordBool1 = ^WordBool; {*}
  PUserType1 = ^TGUID; {*}
  PShortint1 = ^Shortint; {*}
  PPShortint1 = ^PShortint1; {*}
  //PUserType2 = ^DISPPARAMS; {*}


// *********************************************************************//
// Interface: _Application
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020970-0000-0000-C000-000000000046}
// *********************************************************************//
  _Application = interface(IDispatch)
    ['{00020970-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Documents: Documents; safecall;
    function Get_Windows: Windows; safecall;
    function Get_ActiveDocument: Document; safecall;
    function Get_ActiveWindow: Window_; safecall;
    function Get_Selection: Selection; safecall;
    function Get_WordBasic: IDispatch; safecall;
    function Get_RecentFiles: RecentFiles; safecall;
    function Get_NormalTemplate: Template; safecall;
    function Get_System_: System_; safecall;
    function Get_AutoCorrect: AutoCorrect; safecall;
    function Get_FontNames: FontNames; safecall;
    function Get_LandscapeFontNames: FontNames; safecall;
    function Get_PortraitFontNames: FontNames; safecall;
    function Get_Languages: Languages; safecall;
    function Get_Assistant: Assistant; safecall;
    function Get_Browser: Browser; safecall;
    function Get_FileConverters: FileConverters; safecall;
    function Get_MailingLabel: MailingLabel; safecall;
    function Get_Dialogs: Dialogs; safecall;
    function Get_CaptionLabels: CaptionLabels; safecall;
    function Get_AutoCaptions: AutoCaptions; safecall;
    function Get_AddIns: AddIns; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    function Get_Version: WideString; safecall;
    function Get_ScreenUpdating: WordBool; safecall;
    procedure Set_ScreenUpdating(prop: WordBool); safecall;
    function Get_PrintPreview: WordBool; safecall;
    procedure Set_PrintPreview(prop: WordBool); safecall;
    function Get_Tasks: Tasks; safecall;
    function Get_DisplayStatusBar: WordBool; safecall;
    procedure Set_DisplayStatusBar(prop: WordBool); safecall;
    function Get_SpecialMode: WordBool; safecall;
    function Get_UsableWidth: Integer; safecall;
    function Get_UsableHeight: Integer; safecall;
    function Get_MathCoprocessorAvailable: WordBool; safecall;
    function Get_MouseAvailable: WordBool; safecall;
    function Get_International(Index: WdInternationalIndex): OleVariant; safecall;
    function Get_Build: WideString; safecall;
    function Get_CapsLock: WordBool; safecall;
    function Get_NumLock: WordBool; safecall;
    function Get_UserName: WideString; safecall;
    procedure Set_UserName(const prop: WideString); safecall;
    function Get_UserInitials: WideString; safecall;
    procedure Set_UserInitials(const prop: WideString); safecall;
    function Get_UserAddress: WideString; safecall;
    procedure Set_UserAddress(const prop: WideString); safecall;
    function Get_MacroContainer: IDispatch; safecall;
    function Get_DisplayRecentFiles: WordBool; safecall;
    procedure Set_DisplayRecentFiles(prop: WordBool); safecall;
    function Get_CommandBars: CommandBars; safecall;
    function Get_SynonymInfo(const AWord: WideString; var LanguageID: OleVariant): SynonymInfo; safecall;
    function Get_VBE: VBE; safecall;
    function Get_DefaultSaveFormat: WideString; safecall;
    procedure Set_DefaultSaveFormat(const prop: WideString); safecall;
    function Get_ListGalleries: ListGalleries; safecall;
    function Get_ActivePrinter: WideString; safecall;
    procedure Set_ActivePrinter(const prop: WideString); safecall;
    function Get_Templates: Templates; safecall;
    function Get_CustomizationContext: IDispatch; safecall;
    procedure Set_CustomizationContext(const prop: IDispatch); safecall;
    function Get_KeyBindings: KeyBindings; safecall;
    function Get_KeysBoundTo(KeyCategory: WdKeyCategory; const Command: WideString; 
                             var CommandParameter: OleVariant): KeysBoundTo; safecall;
    function Get_FindKey(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const prop: WideString); safecall;
    function Get_Path: WideString; safecall;
    function Get_DisplayScrollBars: WordBool; safecall;
    procedure Set_DisplayScrollBars(prop: WordBool); safecall;
    function Get_StartupPath: WideString; safecall;
    procedure Set_StartupPath(const prop: WideString); safecall;
    function Get_BackgroundSavingStatus: Integer; safecall;
    function Get_BackgroundPrintingStatus: Integer; safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(prop: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(prop: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(prop: Integer); safecall;
    function Get_WindowState: WdWindowState; safecall;
    procedure Set_WindowState(prop: WdWindowState); safecall;
    function Get_DisplayAutoCompleteTips: WordBool; safecall;
    procedure Set_DisplayAutoCompleteTips(prop: WordBool); safecall;
    function Get_Options: Options; safecall;
    function Get_DisplayAlerts: WdAlertLevel; safecall;
    procedure Set_DisplayAlerts(prop: WdAlertLevel); safecall;
    function Get_CustomDictionaries: Dictionaries; safecall;
    function Get_PathSeparator: WideString; safecall;
    procedure Set_StatusBar(const Param1: WideString); safecall;
    function Get_MAPIAvailable: WordBool; safecall;
    function Get_DisplayScreenTips: WordBool; safecall;
    procedure Set_DisplayScreenTips(prop: WordBool); safecall;
    function Get_EnableCancelKey: WdEnableCancelKey; safecall;
    procedure Set_EnableCancelKey(prop: WdEnableCancelKey); safecall;
    function Get_UserControl: WordBool; safecall;
    function Get_FileSearch: FileSearch; safecall;
    function Get_MailSystem: WdMailSystem; safecall;
    function Get_DefaultTableSeparator: WideString; safecall;
    procedure Set_DefaultTableSeparator(const prop: WideString); safecall;
    function Get_ShowVisualBasicEditor: WordBool; safecall;
    procedure Set_ShowVisualBasicEditor(prop: WordBool); safecall;
    function Get_BrowseExtraFileTypes: WideString; safecall;
    procedure Set_BrowseExtraFileTypes(const prop: WideString); safecall;
    function Get_IsObjectValid(const Object_: IDispatch): WordBool; safecall;
    function Get_HangulHanjaDictionaries: HangulHanjaConversionDictionaries; safecall;
    function Get_MailMessage: MailMessage; safecall;
    function Get_FocusInMailHeader: WordBool; safecall;
    procedure Quit(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                   var RouteDocument: OleVariant); safecall;
    procedure ScreenRefresh; safecall;
    procedure PrintOutOld(var Background: OleVariant; var Append: OleVariant; 
                          var Range: OleVariant; var OutputFileName: OleVariant; 
                          var From: OleVariant; var To_: OleVariant; var Item: OleVariant; 
                          var Copies: OleVariant; var Pages: OleVariant; var PageType: OleVariant; 
                          var PrintToFile: OleVariant; var Collate: OleVariant; 
                          var FileName: OleVariant; var ActivePrinterMacGX: OleVariant; 
                          var ManualDuplexPrint: OleVariant); safecall;
    procedure LookupNameProperties(const Name: WideString); safecall;
    procedure SubstituteFont(const UnavailableFont: WideString; const SubstituteFont: WideString); safecall;
    function Repeat_(var Times: OleVariant): WordBool; safecall;
    procedure DDEExecute(Channel: Integer; const Command: WideString); safecall;
    function DDEInitiate(const App: WideString; const Topic: WideString): Integer; safecall;
    procedure DDEPoke(Channel: Integer; const Item: WideString; const Data: WideString); safecall;
    function DDERequest(Channel: Integer; const Item: WideString): WideString; safecall;
    procedure DDETerminate(Channel: Integer); safecall;
    procedure DDETerminateAll; safecall;
    function BuildKeyCode(Arg1: WdKey; var Arg2: OleVariant; var Arg3: OleVariant; 
                          var Arg4: OleVariant): Integer; safecall;
    function KeyString(KeyCode: Integer; var KeyCode2: OleVariant): WideString; safecall;
    procedure OrganizerCopy(const Source: WideString; const Destination: WideString; 
                            const Name: WideString; Object_: WdOrganizerObject); safecall;
    procedure OrganizerDelete(const Source: WideString; const Name: WideString; 
                              Object_: WdOrganizerObject); safecall;
    procedure OrganizerRename(const Source: WideString; const Name: WideString; 
                              const NewName: WideString; Object_: WdOrganizerObject); safecall;
    procedure AddAddress(var TagID: PSafeArray; var Value: PSafeArray); safecall;
    function GetAddress(var Name: OleVariant; var AddressProperties: OleVariant; 
                        var UseAutoText: OleVariant; var DisplaySelectDialog: OleVariant; 
                        var SelectDialog: OleVariant; var CheckNamesDialog: OleVariant; 
                        var RecentAddressesChoice: OleVariant; var UpdateRecentAddresses: OleVariant): WideString; safecall;
    function CheckGrammar(const String_: WideString): WordBool; safecall;
    function CheckSpelling(const Word: WideString; var CustomDictionary: OleVariant; 
                           var IgnoreUppercase: OleVariant; var MainDictionary: OleVariant; 
                           var CustomDictionary2: OleVariant; var CustomDictionary3: OleVariant; 
                           var CustomDictionary4: OleVariant; var CustomDictionary5: OleVariant; 
                           var CustomDictionary6: OleVariant; var CustomDictionary7: OleVariant; 
                           var CustomDictionary8: OleVariant; var CustomDictionary9: OleVariant; 
                           var CustomDictionary10: OleVariant): WordBool; safecall;
    procedure ResetIgnoreAll; safecall;
    function GetSpellingSuggestions(const Word: WideString; var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; safecall;
    procedure GoBack; safecall;
    procedure Help(var HelpType: OleVariant); safecall;
    procedure AutomaticChange; safecall;
    procedure ShowMe; safecall;
    procedure HelpTool; safecall;
    function NewWindow: Window_; safecall;
    procedure ListCommands(ListAllCommands: WordBool); safecall;
    procedure ShowClipboard; safecall;
    procedure OnTime(var When: OleVariant; const Name: WideString; var Tolerance: OleVariant); safecall;
    procedure NextLetter; safecall;
    function MountVolume(const Zone: WideString; const Server: WideString; 
                         const Volume: WideString; var User: OleVariant; 
                         var UserPassword: OleVariant; var VolumePassword: OleVariant): Smallint; safecall;
    function CleanString(const String_: WideString): WideString; safecall;
    procedure SendFax; safecall;
    procedure ChangeFileOpenDirectory(const Path: WideString); safecall;
    procedure RunOld(const MacroName: WideString); safecall;
    procedure GoForward; safecall;
    procedure Move(Left: Integer; Top: Integer); safecall;
    procedure Resize(Width: Integer; Height: Integer); safecall;
    function InchesToPoints(Inches: Single): Single; safecall;
    function CentimetersToPoints(Centimeters: Single): Single; safecall;
    function MillimetersToPoints(Millimeters: Single): Single; safecall;
    function PicasToPoints(Picas: Single): Single; safecall;
    function LinesToPoints(Lines: Single): Single; safecall;
    function PointsToInches(Points: Single): Single; safecall;
    function PointsToCentimeters(Points: Single): Single; safecall;
    function PointsToMillimeters(Points: Single): Single; safecall;
    function PointsToPicas(Points: Single): Single; safecall;
    function PointsToLines(Points: Single): Single; safecall;
    procedure Activate; safecall;
    function PointsToPixels(Points: Single; var fVertical: OleVariant): Single; safecall;
    function PixelsToPoints(Pixels: Single; var fVertical: OleVariant): Single; safecall;
    procedure KeyboardLatin; safecall;
    procedure KeyboardBidi; safecall;
    procedure ToggleKeyboard; safecall;
    function Keyboard(LangId: Integer): Integer; safecall;
    function ProductCode: WideString; safecall;
    function DefaultWebOptions: DefaultWebOptions; safecall;
    procedure DiscussionSupport(var Range: OleVariant; var cid: OleVariant; var piCSE: OleVariant); safecall;
    procedure SetDefaultTheme(const Name: WideString; DocumentType: WdDocumentMedium); safecall;
    function GetDefaultTheme(DocumentType: WdDocumentMedium): WideString; safecall;
    function Get_EmailOptions: EmailOptions; safecall;
    function Get_Language: MsoLanguageID; safecall;
    function Get_COMAddIns: COMAddIns; safecall;
    function Get_CheckLanguage: WordBool; safecall;
    procedure Set_CheckLanguage(prop: WordBool); safecall;
    function Get_LanguageSettings: LanguageSettings; safecall;
    function Get_Dummy1: WordBool; safecall;
    function Get_AnswerWizard: AnswerWizard; safecall;
    function Get_FeatureInstall: MsoFeatureInstall; safecall;
    procedure Set_FeatureInstall(prop: MsoFeatureInstall); safecall;
    procedure PrintOut(var Background: OleVariant; var Append: OleVariant; var Range: OleVariant; 
                       var OutputFileName: OleVariant; var From: OleVariant; var To_: OleVariant; 
                       var Item: OleVariant; var Copies: OleVariant; var Pages: OleVariant; 
                       var PageType: OleVariant; var PrintToFile: OleVariant; 
                       var Collate: OleVariant; var FileName: OleVariant; 
                       var ActivePrinterMacGX: OleVariant; var ManualDuplexPrint: OleVariant; 
                       var PrintZoomColumn: OleVariant; var PrintZoomRow: OleVariant; 
                       var PrintZoomPaperWidth: OleVariant; var PrintZoomPaperHeight: OleVariant); safecall;
    function Run(const MacroName: WideString; var varg1: OleVariant; var varg2: OleVariant; 
                 var varg3: OleVariant; var varg4: OleVariant; var varg5: OleVariant; 
                 var varg6: OleVariant; var varg7: OleVariant; var varg8: OleVariant; 
                 var varg9: OleVariant; var varg10: OleVariant; var varg11: OleVariant; 
                 var varg12: OleVariant; var varg13: OleVariant; var varg14: OleVariant; 
                 var varg15: OleVariant; var varg16: OleVariant; var varg17: OleVariant; 
                 var varg18: OleVariant; var varg19: OleVariant; var varg20: OleVariant; 
                 var varg21: OleVariant; var varg22: OleVariant; var varg23: OleVariant; 
                 var varg24: OleVariant; var varg25: OleVariant; var varg26: OleVariant; 
                 var varg27: OleVariant; var varg28: OleVariant; var varg29: OleVariant; 
                 var varg30: OleVariant): OleVariant; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Documents: Documents read Get_Documents;
    property Windows: Windows read Get_Windows;
    property ActiveDocument: Document read Get_ActiveDocument;
    property ActiveWindow: Window_ read Get_ActiveWindow;
    property Selection: Selection read Get_Selection;
    property WordBasic: IDispatch read Get_WordBasic;
    property RecentFiles: RecentFiles read Get_RecentFiles;
    property NormalTemplate: Template read Get_NormalTemplate;
    property System_: System_ read Get_System_;
    property AutoCorrect: AutoCorrect read Get_AutoCorrect;
    property FontNames: FontNames read Get_FontNames;
    property LandscapeFontNames: FontNames read Get_LandscapeFontNames;
    property PortraitFontNames: FontNames read Get_PortraitFontNames;
    property Languages: Languages read Get_Languages;
    property Assistant: Assistant read Get_Assistant;
    property Browser: Browser read Get_Browser;
    property FileConverters: FileConverters read Get_FileConverters;
    property MailingLabel: MailingLabel read Get_MailingLabel;
    property Dialogs: Dialogs read Get_Dialogs;
    property CaptionLabels: CaptionLabels read Get_CaptionLabels;
    property AutoCaptions: AutoCaptions read Get_AutoCaptions;
    property AddIns: AddIns read Get_AddIns;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Version: WideString read Get_Version;
    property ScreenUpdating: WordBool read Get_ScreenUpdating write Set_ScreenUpdating;
    property PrintPreview: WordBool read Get_PrintPreview write Set_PrintPreview;
    property Tasks: Tasks read Get_Tasks;
    property DisplayStatusBar: WordBool read Get_DisplayStatusBar write Set_DisplayStatusBar;
    property SpecialMode: WordBool read Get_SpecialMode;
    property UsableWidth: Integer read Get_UsableWidth;
    property UsableHeight: Integer read Get_UsableHeight;
    property MathCoprocessorAvailable: WordBool read Get_MathCoprocessorAvailable;
    property MouseAvailable: WordBool read Get_MouseAvailable;
    property International[Index: WdInternationalIndex]: OleVariant read Get_International;
    property Build: WideString read Get_Build;
    property CapsLock: WordBool read Get_CapsLock;
    property NumLock: WordBool read Get_NumLock;
    property UserName: WideString read Get_UserName write Set_UserName;
    property UserInitials: WideString read Get_UserInitials write Set_UserInitials;
    property UserAddress: WideString read Get_UserAddress write Set_UserAddress;
    property MacroContainer: IDispatch read Get_MacroContainer;
    property DisplayRecentFiles: WordBool read Get_DisplayRecentFiles write Set_DisplayRecentFiles;
    property CommandBars: CommandBars read Get_CommandBars;
    property SynonymInfo[const AWord: WideString; var LanguageID: OleVariant]: SynonymInfo read Get_SynonymInfo;
    property VBE: VBE read Get_VBE;
    property DefaultSaveFormat: WideString read Get_DefaultSaveFormat write Set_DefaultSaveFormat;
    property ListGalleries: ListGalleries read Get_ListGalleries;
    property ActivePrinter: WideString read Get_ActivePrinter write Set_ActivePrinter;
    property Templates: Templates read Get_Templates;
    property CustomizationContext: IDispatch read Get_CustomizationContext write Set_CustomizationContext;
    property KeyBindings: KeyBindings read Get_KeyBindings;
    property KeysBoundTo[KeyCategory: WdKeyCategory; const Command: WideString; 
                         var CommandParameter: OleVariant]: KeysBoundTo read Get_KeysBoundTo;
    property FindKey[KeyCode: Integer; var KeyCode2: OleVariant]: KeyBinding read Get_FindKey;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Path: WideString read Get_Path;
    property DisplayScrollBars: WordBool read Get_DisplayScrollBars write Set_DisplayScrollBars;
    property StartupPath: WideString read Get_StartupPath write Set_StartupPath;
    property BackgroundSavingStatus: Integer read Get_BackgroundSavingStatus;
    property BackgroundPrintingStatus: Integer read Get_BackgroundPrintingStatus;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property WindowState: WdWindowState read Get_WindowState write Set_WindowState;
    property DisplayAutoCompleteTips: WordBool read Get_DisplayAutoCompleteTips write Set_DisplayAutoCompleteTips;
    property Options: Options read Get_Options;
    property DisplayAlerts: WdAlertLevel read Get_DisplayAlerts write Set_DisplayAlerts;
    property CustomDictionaries: Dictionaries read Get_CustomDictionaries;
    property PathSeparator: WideString read Get_PathSeparator;
    property StatusBar: WideString write Set_StatusBar;
    property MAPIAvailable: WordBool read Get_MAPIAvailable;
    property DisplayScreenTips: WordBool read Get_DisplayScreenTips write Set_DisplayScreenTips;
    property EnableCancelKey: WdEnableCancelKey read Get_EnableCancelKey write Set_EnableCancelKey;
    property UserControl: WordBool read Get_UserControl;
    property FileSearch: FileSearch read Get_FileSearch;
    property MailSystem: WdMailSystem read Get_MailSystem;
    property DefaultTableSeparator: WideString read Get_DefaultTableSeparator write Set_DefaultTableSeparator;
    property ShowVisualBasicEditor: WordBool read Get_ShowVisualBasicEditor write Set_ShowVisualBasicEditor;
    property BrowseExtraFileTypes: WideString read Get_BrowseExtraFileTypes write Set_BrowseExtraFileTypes;
    property IsObjectValid[const Object_: IDispatch]: WordBool read Get_IsObjectValid;
    property HangulHanjaDictionaries: HangulHanjaConversionDictionaries read Get_HangulHanjaDictionaries;
    property MailMessage: MailMessage read Get_MailMessage;
    property FocusInMailHeader: WordBool read Get_FocusInMailHeader;
    property EmailOptions: EmailOptions read Get_EmailOptions;
    property Language: MsoLanguageID read Get_Language;
    property COMAddIns: COMAddIns read Get_COMAddIns;
    property CheckLanguage: WordBool read Get_CheckLanguage write Set_CheckLanguage;
    property LanguageSettings: LanguageSettings read Get_LanguageSettings;
    property Dummy1: WordBool read Get_Dummy1;
    property AnswerWizard: AnswerWizard read Get_AnswerWizard;
    property FeatureInstall: MsoFeatureInstall read Get_FeatureInstall write Set_FeatureInstall;
  end;

// *********************************************************************//
// DispIntf:  _ApplicationDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020970-0000-0000-C000-000000000046}
// *********************************************************************//
  _ApplicationDisp = dispinterface
    ['{00020970-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Documents: Documents readonly dispid 6;
    property Windows: Windows readonly dispid 2;
    property ActiveDocument: Document readonly dispid 3;
    property ActiveWindow: Window_ readonly dispid 4;
    property Selection: Selection readonly dispid 5;
    property WordBasic: IDispatch readonly dispid 1;
    property RecentFiles: RecentFiles readonly dispid 7;
    property NormalTemplate: Template readonly dispid 8;
    property System_: System_ readonly dispid 9;
    property AutoCorrect: AutoCorrect readonly dispid 10;
    property FontNames: FontNames readonly dispid 11;
    property LandscapeFontNames: FontNames readonly dispid 12;
    property PortraitFontNames: FontNames readonly dispid 13;
    property Languages: Languages readonly dispid 14;
    property Assistant: Assistant readonly dispid 15;
    property Browser: Browser readonly dispid 16;
    property FileConverters: FileConverters readonly dispid 17;
    property MailingLabel: MailingLabel readonly dispid 18;
    property Dialogs: Dialogs readonly dispid 19;
    property CaptionLabels: CaptionLabels readonly dispid 20;
    property AutoCaptions: AutoCaptions readonly dispid 21;
    property AddIns: AddIns readonly dispid 22;
    property Visible: WordBool dispid 23;
    property Version: WideString readonly dispid 24;
    property ScreenUpdating: WordBool dispid 26;
    property PrintPreview: WordBool dispid 27;
    property Tasks: Tasks readonly dispid 28;
    property DisplayStatusBar: WordBool dispid 29;
    property SpecialMode: WordBool readonly dispid 30;
    property UsableWidth: Integer readonly dispid 33;
    property UsableHeight: Integer readonly dispid 34;
    property MathCoprocessorAvailable: WordBool readonly dispid 36;
    property MouseAvailable: WordBool readonly dispid 37;
    property International[Index: WdInternationalIndex]: OleVariant readonly dispid 46;
    property Build: WideString readonly dispid 47;
    property CapsLock: WordBool readonly dispid 48;
    property NumLock: WordBool readonly dispid 49;
    property UserName: WideString dispid 52;
    property UserInitials: WideString dispid 53;
    property UserAddress: WideString dispid 54;
    property MacroContainer: IDispatch readonly dispid 55;
    property DisplayRecentFiles: WordBool dispid 56;
    property CommandBars: CommandBars readonly dispid 57;
    property SynonymInfo[const AWord: WideString; var LanguageID: OleVariant]: SynonymInfo readonly dispid 59;
    property VBE: VBE readonly dispid 61;
    property DefaultSaveFormat: WideString dispid 64;
    property ListGalleries: ListGalleries readonly dispid 65;
    property ActivePrinter: WideString dispid 66;
    property Templates: Templates readonly dispid 67;
    property CustomizationContext: IDispatch dispid 68;
    property KeyBindings: KeyBindings readonly dispid 69;
    property KeysBoundTo[KeyCategory: WdKeyCategory; const Command: WideString; 
                         var CommandParameter: OleVariant]: KeysBoundTo readonly dispid 70;
    property FindKey[KeyCode: Integer; var KeyCode2: OleVariant]: KeyBinding readonly dispid 71;
    property Caption: WideString dispid 80;
    property Path: WideString readonly dispid 81;
    property DisplayScrollBars: WordBool dispid 82;
    property StartupPath: WideString dispid 83;
    property BackgroundSavingStatus: Integer readonly dispid 85;
    property BackgroundPrintingStatus: Integer readonly dispid 86;
    property Left: Integer dispid 87;
    property Top: Integer dispid 88;
    property Width: Integer dispid 89;
    property Height: Integer dispid 90;
    property WindowState: WdWindowState dispid 91;
    property DisplayAutoCompleteTips: WordBool dispid 92;
    property Options: Options readonly dispid 93;
    property DisplayAlerts: WdAlertLevel dispid 94;
    property CustomDictionaries: Dictionaries readonly dispid 95;
    property PathSeparator: WideString readonly dispid 96;
    property StatusBar: WideString writeonly dispid 97;
    property MAPIAvailable: WordBool readonly dispid 98;
    property DisplayScreenTips: WordBool dispid 99;
    property EnableCancelKey: WdEnableCancelKey dispid 100;
    property UserControl: WordBool readonly dispid 101;
    property FileSearch: FileSearch readonly dispid 103;
    property MailSystem: WdMailSystem readonly dispid 104;
    property DefaultTableSeparator: WideString dispid 105;
    property ShowVisualBasicEditor: WordBool dispid 106;
    property BrowseExtraFileTypes: WideString dispid 108;
    property IsObjectValid[const Object_: IDispatch]: WordBool readonly dispid 109;
    property HangulHanjaDictionaries: HangulHanjaConversionDictionaries readonly dispid 110;
    property MailMessage: MailMessage readonly dispid 348;
    property FocusInMailHeader: WordBool readonly dispid 386;
    procedure Quit(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                   var RouteDocument: OleVariant); dispid 1105;
    procedure ScreenRefresh; dispid 301;
    procedure PrintOutOld(var Background: OleVariant; var Append: OleVariant; 
                          var Range: OleVariant; var OutputFileName: OleVariant; 
                          var From: OleVariant; var To_: OleVariant; var Item: OleVariant; 
                          var Copies: OleVariant; var Pages: OleVariant; var PageType: OleVariant; 
                          var PrintToFile: OleVariant; var Collate: OleVariant; 
                          var FileName: OleVariant; var ActivePrinterMacGX: OleVariant; 
                          var ManualDuplexPrint: OleVariant); dispid 302;
    procedure LookupNameProperties(const Name: WideString); dispid 303;
    procedure SubstituteFont(const UnavailableFont: WideString; const SubstituteFont: WideString); dispid 304;
    function Repeat_(var Times: OleVariant): WordBool; dispid 305;
    procedure DDEExecute(Channel: Integer; const Command: WideString); dispid 310;
    function DDEInitiate(const App: WideString; const Topic: WideString): Integer; dispid 311;
    procedure DDEPoke(Channel: Integer; const Item: WideString; const Data: WideString); dispid 312;
    function DDERequest(Channel: Integer; const Item: WideString): WideString; dispid 313;
    procedure DDETerminate(Channel: Integer); dispid 314;
    procedure DDETerminateAll; dispid 315;
    function BuildKeyCode(Arg1: WdKey; var Arg2: OleVariant; var Arg3: OleVariant; 
                          var Arg4: OleVariant): Integer; dispid 316;
    function KeyString(KeyCode: Integer; var KeyCode2: OleVariant): WideString; dispid 317;
    procedure OrganizerCopy(const Source: WideString; const Destination: WideString; 
                            const Name: WideString; Object_: WdOrganizerObject); dispid 318;
    procedure OrganizerDelete(const Source: WideString; const Name: WideString; 
                              Object_: WdOrganizerObject); dispid 319;
    procedure OrganizerRename(const Source: WideString; const Name: WideString; 
                              const NewName: WideString; Object_: WdOrganizerObject); dispid 320;
    procedure AddAddress(var TagID: {??PSafeArray} OleVariant; var Value: {??PSafeArray} OleVariant); dispid 321;
    function GetAddress(var Name: OleVariant; var AddressProperties: OleVariant; 
                        var UseAutoText: OleVariant; var DisplaySelectDialog: OleVariant; 
                        var SelectDialog: OleVariant; var CheckNamesDialog: OleVariant; 
                        var RecentAddressesChoice: OleVariant; var UpdateRecentAddresses: OleVariant): WideString; dispid 322;
    function CheckGrammar(const String_: WideString): WordBool; dispid 323;
    function CheckSpelling(const Word: WideString; var CustomDictionary: OleVariant; 
                           var IgnoreUppercase: OleVariant; var MainDictionary: OleVariant; 
                           var CustomDictionary2: OleVariant; var CustomDictionary3: OleVariant; 
                           var CustomDictionary4: OleVariant; var CustomDictionary5: OleVariant; 
                           var CustomDictionary6: OleVariant; var CustomDictionary7: OleVariant; 
                           var CustomDictionary8: OleVariant; var CustomDictionary9: OleVariant; 
                           var CustomDictionary10: OleVariant): WordBool; dispid 324;
    procedure ResetIgnoreAll; dispid 326;
    function GetSpellingSuggestions(const Word: WideString; var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; dispid 327;
    procedure GoBack; dispid 328;
    procedure Help(var HelpType: OleVariant); dispid 329;
    procedure AutomaticChange; dispid 330;
    procedure ShowMe; dispid 331;
    procedure HelpTool; dispid 332;
    function NewWindow: Window_; dispid 345;
    procedure ListCommands(ListAllCommands: WordBool); dispid 346;
    procedure ShowClipboard; dispid 349;
    procedure OnTime(var When: OleVariant; const Name: WideString; var Tolerance: OleVariant); dispid 350;
    procedure NextLetter; dispid 351;
    function MountVolume(const Zone: WideString; const Server: WideString; 
                         const Volume: WideString; var User: OleVariant; 
                         var UserPassword: OleVariant; var VolumePassword: OleVariant): Smallint; dispid 353;
    function CleanString(const String_: WideString): WideString; dispid 354;
    procedure SendFax; dispid 356;
    procedure ChangeFileOpenDirectory(const Path: WideString); dispid 357;
    procedure RunOld(const MacroName: WideString); dispid 358;
    procedure GoForward; dispid 359;
    procedure Move(Left: Integer; Top: Integer); dispid 360;
    procedure Resize(Width: Integer; Height: Integer); dispid 361;
    function InchesToPoints(Inches: Single): Single; dispid 370;
    function CentimetersToPoints(Centimeters: Single): Single; dispid 371;
    function MillimetersToPoints(Millimeters: Single): Single; dispid 372;
    function PicasToPoints(Picas: Single): Single; dispid 373;
    function LinesToPoints(Lines: Single): Single; dispid 374;
    function PointsToInches(Points: Single): Single; dispid 380;
    function PointsToCentimeters(Points: Single): Single; dispid 381;
    function PointsToMillimeters(Points: Single): Single; dispid 382;
    function PointsToPicas(Points: Single): Single; dispid 383;
    function PointsToLines(Points: Single): Single; dispid 384;
    procedure Activate; dispid 385;
    function PointsToPixels(Points: Single; var fVertical: OleVariant): Single; dispid 387;
    function PixelsToPoints(Pixels: Single; var fVertical: OleVariant): Single; dispid 388;
    procedure KeyboardLatin; dispid 400;
    procedure KeyboardBidi; dispid 401;
    procedure ToggleKeyboard; dispid 402;
    function Keyboard(LangId: Integer): Integer; dispid 446;
    function ProductCode: WideString; dispid 404;
    function DefaultWebOptions: DefaultWebOptions; dispid 405;
    procedure DiscussionSupport(var Range: OleVariant; var cid: OleVariant; var piCSE: OleVariant); dispid 407;
    procedure SetDefaultTheme(const Name: WideString; DocumentType: WdDocumentMedium); dispid 414;
    function GetDefaultTheme(DocumentType: WdDocumentMedium): WideString; dispid 416;
    property EmailOptions: EmailOptions readonly dispid 389;
    property Language: MsoLanguageID readonly dispid 391;
    property COMAddIns: COMAddIns readonly dispid 111;
    property CheckLanguage: WordBool dispid 112;
    property LanguageSettings: LanguageSettings readonly dispid 403;
    property Dummy1: WordBool readonly dispid 406;
    property AnswerWizard: AnswerWizard readonly dispid 409;
    property FeatureInstall: MsoFeatureInstall dispid 447;
    procedure PrintOut(var Background: OleVariant; var Append: OleVariant; var Range: OleVariant; 
                       var OutputFileName: OleVariant; var From: OleVariant; var To_: OleVariant; 
                       var Item: OleVariant; var Copies: OleVariant; var Pages: OleVariant; 
                       var PageType: OleVariant; var PrintToFile: OleVariant; 
                       var Collate: OleVariant; var FileName: OleVariant; 
                       var ActivePrinterMacGX: OleVariant; var ManualDuplexPrint: OleVariant; 
                       var PrintZoomColumn: OleVariant; var PrintZoomRow: OleVariant; 
                       var PrintZoomPaperWidth: OleVariant; var PrintZoomPaperHeight: OleVariant); dispid 444;
    function Run(const MacroName: WideString; var varg1: OleVariant; var varg2: OleVariant; 
                 var varg3: OleVariant; var varg4: OleVariant; var varg5: OleVariant; 
                 var varg6: OleVariant; var varg7: OleVariant; var varg8: OleVariant; 
                 var varg9: OleVariant; var varg10: OleVariant; var varg11: OleVariant; 
                 var varg12: OleVariant; var varg13: OleVariant; var varg14: OleVariant; 
                 var varg15: OleVariant; var varg16: OleVariant; var varg17: OleVariant; 
                 var varg18: OleVariant; var varg19: OleVariant; var varg20: OleVariant; 
                 var varg21: OleVariant; var varg22: OleVariant; var varg23: OleVariant; 
                 var varg24: OleVariant; var varg25: OleVariant; var varg26: OleVariant; 
                 var varg27: OleVariant; var varg28: OleVariant; var varg29: OleVariant; 
                 var varg30: OleVariant): OleVariant; dispid 445;
  end;

// *********************************************************************//
// Interface: _Global
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B9-0000-0000-C000-000000000046}
// *********************************************************************//
(*  _Global = interface(IDispatch)
    ['{000209B9-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Documents: Documents; safecall;
    function Get_Windows: Windows; safecall;
    function Get_ActiveDocument: Document; safecall;
    function Get_ActiveWindow: Window_; safecall;
    function Get_Selection: Selection; safecall;
    function Get_WordBasic: IDispatch; safecall;
    function Get_PrintPreview: WordBool; safecall;
    procedure Set_PrintPreview(prop: WordBool); safecall;
    function Get_RecentFiles: RecentFiles; safecall;
    function Get_NormalTemplate: Template; safecall;
    function Get_System_: System_; safecall;
    function Get_AutoCorrect: AutoCorrect; safecall;
    function Get_FontNames: FontNames; safecall;
    function Get_LandscapeFontNames: FontNames; safecall;
    function Get_PortraitFontNames: FontNames; safecall;
    function Get_Languages: Languages; safecall;
    function Get_Assistant: Assistant; safecall;
    function Get_FileConverters: FileConverters; safecall;
    function Get_Dialogs: Dialogs; safecall;
    function Get_CaptionLabels: CaptionLabels; safecall;
    function Get_AutoCaptions: AutoCaptions; safecall;
    function Get_AddIns: AddIns; safecall;
    function Get_Tasks: Tasks; safecall;
    function Get_MacroContainer: IDispatch; safecall;
    function Get_CommandBars: CommandBars; safecall;
    function Get_SynonymInfo(const AWord: WideString; var LanguageID: OleVariant): SynonymInfo; safecall;
    function Get_VBE: VBE; safecall;
    function Get_ListGalleries: ListGalleries; safecall;
    function Get_ActivePrinter: WideString; safecall;
    procedure Set_ActivePrinter(const prop: WideString); safecall;
    function Get_Templates: Templates; safecall;
    function Get_CustomizationContext: IDispatch; safecall;
    procedure Set_CustomizationContext(const prop: IDispatch); safecall;
    function Get_KeyBindings: KeyBindings; safecall;
    function Get_KeysBoundTo(KeyCategory: WdKeyCategory; const Command: WideString; 
                             var CommandParameter: OleVariant): KeysBoundTo; safecall;
    function Get_FindKey(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    function Get_Options: Options; safecall;
    function Get_CustomDictionaries: Dictionaries; safecall;
    procedure Set_StatusBar(const Param1: WideString); safecall;
    function Get_ShowVisualBasicEditor: WordBool; safecall;
    procedure Set_ShowVisualBasicEditor(prop: WordBool); safecall;
    function Get_IsObjectValid(const Object_: IDispatch): WordBool; safecall;
    function Get_HangulHanjaDictionaries: HangulHanjaConversionDictionaries; safecall;
    function Repeat_(var Times: OleVariant): WordBool; safecall;
    procedure DDEExecute(Channel: Integer; const Command: WideString); safecall;
    function DDEInitiate(const App: WideString; const Topic: WideString): Integer; safecall;
    procedure DDEPoke(Channel: Integer; const Item: WideString; const Data: WideString); safecall;
    function DDERequest(Channel: Integer; const Item: WideString): WideString; safecall;
    procedure DDETerminate(Channel: Integer); safecall;
    procedure DDETerminateAll; safecall;
    function BuildKeyCode(Arg1: WdKey; var Arg2: OleVariant; var Arg3: OleVariant; 
                          var Arg4: OleVariant): Integer; safecall;
    function KeyString(KeyCode: Integer; var KeyCode2: OleVariant): WideString; safecall;
    function CheckSpelling(const Word: WideString; var CustomDictionary: OleVariant; 
                           var IgnoreUppercase: OleVariant; var MainDictionary: OleVariant; 
                           var CustomDictionary2: OleVariant; var CustomDictionary3: OleVariant; 
                           var CustomDictionary4: OleVariant; var CustomDictionary5: OleVariant; 
                           var CustomDictionary6: OleVariant; var CustomDictionary7: OleVariant; 
                           var CustomDictionary8: OleVariant; var CustomDictionary9: OleVariant; 
                           var CustomDictionary10: OleVariant): WordBool; safecall;
    function GetSpellingSuggestions(const Word: WideString; var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; safecall;
    procedure Help(var HelpType: OleVariant); safecall;
    function NewWindow: Window_; safecall;
    function CleanString(const String_: WideString): WideString; safecall;
    procedure ChangeFileOpenDirectory(const Path: WideString); safecall;
    function InchesToPoints(Inches: Single): Single; safecall;
    function CentimetersToPoints(Centimeters: Single): Single; safecall;
    function MillimetersToPoints(Millimeters: Single): Single; safecall;
    function PicasToPoints(Picas: Single): Single; safecall;
    function LinesToPoints(Lines: Single): Single; safecall;
    function PointsToInches(Points: Single): Single; safecall;
    function PointsToCentimeters(Points: Single): Single; safecall;
    function PointsToMillimeters(Points: Single): Single; safecall;
    function PointsToPicas(Points: Single): Single; safecall;
    function PointsToLines(Points: Single): Single; safecall;
    function PointsToPixels(Points: Single; var fVertical: OleVariant): Single; safecall;
    function PixelsToPoints(Pixels: Single; var fVertical: OleVariant): Single; safecall;
    function Get_LanguageSettings: LanguageSettings; safecall;
    function Get_AnswerWizard: AnswerWizard; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Documents: Documents read Get_Documents;
    property Windows: Windows read Get_Windows;
    property ActiveDocument: Document read Get_ActiveDocument;
    property ActiveWindow: Window_ read Get_ActiveWindow;
    property Selection: Selection read Get_Selection;
    property WordBasic: IDispatch read Get_WordBasic;
    property PrintPreview: WordBool read Get_PrintPreview write Set_PrintPreview;
    property RecentFiles: RecentFiles read Get_RecentFiles;
    property NormalTemplate: Template read Get_NormalTemplate;
    property System_: System_ read Get_System_;
    property AutoCorrect: AutoCorrect read Get_AutoCorrect;
    property FontNames: FontNames read Get_FontNames;
    property LandscapeFontNames: FontNames read Get_LandscapeFontNames;
    property PortraitFontNames: FontNames read Get_PortraitFontNames;
    property Languages: Languages read Get_Languages;
    property Assistant: Assistant read Get_Assistant;
    property FileConverters: FileConverters read Get_FileConverters;
    property Dialogs: Dialogs read Get_Dialogs;
    property CaptionLabels: CaptionLabels read Get_CaptionLabels;
    property AutoCaptions: AutoCaptions read Get_AutoCaptions;
    property AddIns: AddIns read Get_AddIns;
    property Tasks: Tasks read Get_Tasks;
    property MacroContainer: IDispatch read Get_MacroContainer;
    property CommandBars: CommandBars read Get_CommandBars;
    property SynonymInfo[const AWord: WideString; var LanguageID: OleVariant]: SynonymInfo read Get_SynonymInfo;
    property VBE: VBE read Get_VBE;
    property ListGalleries: ListGalleries read Get_ListGalleries;
    property ActivePrinter: WideString read Get_ActivePrinter write Set_ActivePrinter;
    property Templates: Templates read Get_Templates;
    property CustomizationContext: IDispatch read Get_CustomizationContext write Set_CustomizationContext;
    property KeyBindings: KeyBindings read Get_KeyBindings;
    property KeysBoundTo[KeyCategory: WdKeyCategory; const Command: WideString;
                         var CommandParameter: OleVariant]: KeysBoundTo read Get_KeysBoundTo;
    property FindKey[KeyCode: Integer; var KeyCode2: OleVariant]: KeyBinding read Get_FindKey;
    property Options: Options read Get_Options;
    property CustomDictionaries: Dictionaries read Get_CustomDictionaries;
    property StatusBar: WideString write Set_StatusBar;
    property ShowVisualBasicEditor: WordBool read Get_ShowVisualBasicEditor write Set_ShowVisualBasicEditor;
    property IsObjectValid[const Object_: IDispatch]: WordBool read Get_IsObjectValid;
    property HangulHanjaDictionaries: HangulHanjaConversionDictionaries read Get_HangulHanjaDictionaries;
    property LanguageSettings: LanguageSettings read Get_LanguageSettings;
    property AnswerWizard: AnswerWizard read Get_AnswerWizard;
  end;
*)
// *********************************************************************//
// DispIntf:  _GlobalDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B9-0000-0000-C000-000000000046}
// *********************************************************************//
(*  _GlobalDisp = dispinterface
    ['{000209B9-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Documents: Documents readonly dispid 1;
    property Windows: Windows readonly dispid 2;
    property ActiveDocument: Document readonly dispid 3;
    property ActiveWindow: Window_ readonly dispid 4;
    property Selection: Selection readonly dispid 5;
    property WordBasic: IDispatch readonly dispid 6;
    property PrintPreview: WordBool dispid 27;
    property RecentFiles: RecentFiles readonly dispid 7;
    property NormalTemplate: Template readonly dispid 8;
    property System_: System_ readonly dispid 9;
    property AutoCorrect: AutoCorrect readonly dispid 10;
    property FontNames: FontNames readonly dispid 11;
    property LandscapeFontNames: FontNames readonly dispid 12;
    property PortraitFontNames: FontNames readonly dispid 13;
    property Languages: Languages readonly dispid 14;
    property Assistant: Assistant readonly dispid 15;
    property FileConverters: FileConverters readonly dispid 17;
    property Dialogs: Dialogs readonly dispid 19;
    property CaptionLabels: CaptionLabels readonly dispid 20;
    property AutoCaptions: AutoCaptions readonly dispid 21;
    property AddIns: AddIns readonly dispid 22;
    property Tasks: Tasks readonly dispid 28;
    property MacroContainer: IDispatch readonly dispid 55;
    property CommandBars: CommandBars readonly dispid 57;
    property SynonymInfo[const AWord: WideString; var LanguageID: OleVariant]: SynonymInfo readonly dispid 59;
    property VBE: VBE readonly dispid 61;
    property ListGalleries: ListGalleries readonly dispid 65;
    property ActivePrinter: WideString dispid 66;
    property Templates: Templates readonly dispid 67;
    property CustomizationContext: IDispatch dispid 68;
    property KeyBindings: KeyBindings readonly dispid 69;
    property KeysBoundTo[KeyCategory: WdKeyCategory; const Command: WideString; 
                         var CommandParameter: OleVariant]: KeysBoundTo readonly dispid 70;
    property FindKey[KeyCode: Integer; var KeyCode2: OleVariant]: KeyBinding readonly dispid 71;
    property Options: Options readonly dispid 93;
    property CustomDictionaries: Dictionaries readonly dispid 95;
    property StatusBar: WideString writeonly dispid 97;
    property ShowVisualBasicEditor: WordBool dispid 104;
    property IsObjectValid[const Object_: IDispatch]: WordBool readonly dispid 109;
    property HangulHanjaDictionaries: HangulHanjaConversionDictionaries readonly dispid 110;
    function Repeat_(var Times: OleVariant): WordBool; dispid 305;
    procedure DDEExecute(Channel: Integer; const Command: WideString); dispid 310;
    function DDEInitiate(const App: WideString; const Topic: WideString): Integer; dispid 311;
    procedure DDEPoke(Channel: Integer; const Item: WideString; const Data: WideString); dispid 312;
    function DDERequest(Channel: Integer; const Item: WideString): WideString; dispid 313;
    procedure DDETerminate(Channel: Integer); dispid 314;
    procedure DDETerminateAll; dispid 315;
    function BuildKeyCode(Arg1: WdKey; var Arg2: OleVariant; var Arg3: OleVariant; 
                          var Arg4: OleVariant): Integer; dispid 316;
    function KeyString(KeyCode: Integer; var KeyCode2: OleVariant): WideString; dispid 317;
    function CheckSpelling(const Word: WideString; var CustomDictionary: OleVariant; 
                           var IgnoreUppercase: OleVariant; var MainDictionary: OleVariant; 
                           var CustomDictionary2: OleVariant; var CustomDictionary3: OleVariant; 
                           var CustomDictionary4: OleVariant; var CustomDictionary5: OleVariant; 
                           var CustomDictionary6: OleVariant; var CustomDictionary7: OleVariant; 
                           var CustomDictionary8: OleVariant; var CustomDictionary9: OleVariant; 
                           var CustomDictionary10: OleVariant): WordBool; dispid 324;
    function GetSpellingSuggestions(const Word: WideString; var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant;
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; dispid 327;
    procedure Help(var HelpType: OleVariant); dispid 329;
    function NewWindow: Window_; dispid 345;
    function CleanString(const String_: WideString): WideString; dispid 354;
    procedure ChangeFileOpenDirectory(const Path: WideString); dispid 355;
    function InchesToPoints(Inches: Single): Single; dispid 370;
    function CentimetersToPoints(Centimeters: Single): Single; dispid 371;
    function MillimetersToPoints(Millimeters: Single): Single; dispid 372;
    function PicasToPoints(Picas: Single): Single; dispid 373;
    function LinesToPoints(Lines: Single): Single; dispid 374;
    function PointsToInches(Points: Single): Single; dispid 380;
    function PointsToCentimeters(Points: Single): Single; dispid 381;
    function PointsToMillimeters(Points: Single): Single; dispid 382;
    function PointsToPicas(Points: Single): Single; dispid 383;
    function PointsToLines(Points: Single): Single; dispid 384;
    function PointsToPixels(Points: Single; var fVertical: OleVariant): Single; dispid 385;
    function PixelsToPoints(Pixels: Single; var fVertical: OleVariant): Single; dispid 386;
    property LanguageSettings: LanguageSettings readonly dispid 111;
    property AnswerWizard: AnswerWizard readonly dispid 112;
  end;
*)
// *********************************************************************//
// Interface: FontNames
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096F-0000-0000-C000-000000000046}
// *********************************************************************//
  FontNames = interface(IDispatch)
    ['{0002096F-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): WideString; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  FontNamesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096F-0000-0000-C000-000000000046}
// *********************************************************************//
  FontNamesDisp = dispinterface
    ['{0002096F-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): WideString; dispid 0;
  end;

// *********************************************************************//
// Interface: Languages
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096E-0000-0000-C000-000000000046}
// *********************************************************************//
  Languages = interface(IDispatch)
    ['{0002096E-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): Language; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  LanguagesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096E-0000-0000-C000-000000000046}
// *********************************************************************//
  LanguagesDisp = dispinterface
    ['{0002096E-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): Language; dispid 0;
  end;

// *********************************************************************//
// Interface: Language
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096D-0000-0000-C000-000000000046}
// *********************************************************************//
  Language = interface(IDispatch)
    ['{0002096D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ID: WdLanguageID; safecall;
    function Get_NameLocal: WideString; safecall;
    function Get_Name: WideString; safecall;
    function Get_ActiveGrammarDictionary: Dictionary; safecall;
    function Get_ActiveHyphenationDictionary: Dictionary; safecall;
    function Get_ActiveSpellingDictionary: Dictionary; safecall;
    function Get_ActiveThesaurusDictionary: Dictionary; safecall;
    function Get_DefaultWritingStyle: WideString; safecall;
    procedure Set_DefaultWritingStyle(const prop: WideString); safecall;
    function Get_WritingStyleList: OleVariant; safecall;
    function Get_SpellingDictionaryType: WdDictionaryType; safecall;
    procedure Set_SpellingDictionaryType(prop: WdDictionaryType); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ID: WdLanguageID read Get_ID;
    property NameLocal: WideString read Get_NameLocal;
    property Name: WideString read Get_Name;
    property ActiveGrammarDictionary: Dictionary read Get_ActiveGrammarDictionary;
    property ActiveHyphenationDictionary: Dictionary read Get_ActiveHyphenationDictionary;
    property ActiveSpellingDictionary: Dictionary read Get_ActiveSpellingDictionary;
    property ActiveThesaurusDictionary: Dictionary read Get_ActiveThesaurusDictionary;
    property DefaultWritingStyle: WideString read Get_DefaultWritingStyle write Set_DefaultWritingStyle;
    property WritingStyleList: OleVariant read Get_WritingStyleList;
    property SpellingDictionaryType: WdDictionaryType read Get_SpellingDictionaryType write Set_SpellingDictionaryType;
  end;

// *********************************************************************//
// DispIntf:  LanguageDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096D-0000-0000-C000-000000000046}
// *********************************************************************//
  LanguageDisp = dispinterface
    ['{0002096D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ID: WdLanguageID readonly dispid 10;
    property NameLocal: WideString readonly dispid 0;
    property Name: WideString readonly dispid 12;
    property ActiveGrammarDictionary: Dictionary readonly dispid 13;
    property ActiveHyphenationDictionary: Dictionary readonly dispid 14;
    property ActiveSpellingDictionary: Dictionary readonly dispid 15;
    property ActiveThesaurusDictionary: Dictionary readonly dispid 16;
    property DefaultWritingStyle: WideString dispid 17;
    property WritingStyleList: OleVariant readonly dispid 18;
    property SpellingDictionaryType: WdDictionaryType dispid 19;
  end;

// *********************************************************************//
// Interface: Documents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096C-0000-0000-C000-000000000046}
// *********************************************************************//
  Documents = interface(IDispatch)
    ['{0002096C-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): Document; safecall;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); safecall;
    function AddOld(var Template: OleVariant; var NewTemplate: OleVariant): Document; safecall;
    function OpenOld(var FileName: OleVariant; var ConfirmConversions: OleVariant; 
                     var ReadOnly: OleVariant; var AddToRecentFiles: OleVariant; 
                     var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                     var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                     var WritePasswordTemplate: OleVariant; var Format: OleVariant): Document; safecall;
    procedure Save(var NoPrompt: OleVariant; var OriginalFormat: OleVariant); safecall;
    function Add(var Template: OleVariant; var NewTemplate: OleVariant; 
                 var DocumentType: OleVariant; var Visible: OleVariant): Document; safecall;
    function Open(var FileName: OleVariant; var ConfirmConversions: OleVariant; 
                  var ReadOnly: OleVariant; var AddToRecentFiles: OleVariant; 
                  var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                  var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                  var WritePasswordTemplate: OleVariant; var Format: OleVariant; 
                  var Encoding: OleVariant; var Visible: OleVariant): Document; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  DocumentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096C-0000-0000-C000-000000000046}
// *********************************************************************//
  DocumentsDisp = dispinterface
    ['{0002096C-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): Document; dispid 0;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); dispid 1105;
    function AddOld(var Template: OleVariant; var NewTemplate: OleVariant): Document; dispid 11;
    function OpenOld(var FileName: OleVariant; var ConfirmConversions: OleVariant; 
                     var ReadOnly: OleVariant; var AddToRecentFiles: OleVariant; 
                     var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                     var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                     var WritePasswordTemplate: OleVariant; var Format: OleVariant): Document; dispid 12;
    procedure Save(var NoPrompt: OleVariant; var OriginalFormat: OleVariant); dispid 13;
    function Add(var Template: OleVariant; var NewTemplate: OleVariant; 
                 var DocumentType: OleVariant; var Visible: OleVariant): Document; dispid 14;
    function Open(var FileName: OleVariant; var ConfirmConversions: OleVariant; 
                  var ReadOnly: OleVariant; var AddToRecentFiles: OleVariant; 
                  var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                  var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                  var WritePasswordTemplate: OleVariant; var Format: OleVariant; 
                  var Encoding: OleVariant; var Visible: OleVariant): Document; dispid 15;
  end;

// *********************************************************************//
// Interface: _Document
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {0002096B-0000-0000-C000-000000000046}
// *********************************************************************//
  _Document = interface(IDispatch)
    ['{0002096B-0000-0000-C000-000000000046}']
    function Get_Name: WideString; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_BuiltInDocumentProperties: IDispatch; safecall;
    function Get_CustomDocumentProperties: IDispatch; safecall;
    function Get_Path: WideString; safecall;
    function Get_Bookmarks: Bookmarks; safecall;
    function Get_Tables: Tables; safecall;
    function Get_Footnotes: Footnotes; safecall;
    function Get_Endnotes: Endnotes; safecall;
    function Get_Comments: Comments; safecall;
    function Get_Type_: WdDocumentType; safecall;
    function Get_AutoHyphenation: WordBool; safecall;
    procedure Set_AutoHyphenation(prop: WordBool); safecall;
    function Get_HyphenateCaps: WordBool; safecall;
    procedure Set_HyphenateCaps(prop: WordBool); safecall;
    function Get_HyphenationZone: Integer; safecall;
    procedure Set_HyphenationZone(prop: Integer); safecall;
    function Get_ConsecutiveHyphensLimit: Integer; safecall;
    procedure Set_ConsecutiveHyphensLimit(prop: Integer); safecall;
    function Get_Sections: Sections; safecall;
    function Get_Paragraphs: Paragraphs; safecall;
    function Get_Words: Words; safecall;
    function Get_Sentences: Sentences; safecall;
    function Get_Characters: Characters; safecall;
    function Get_Fields: Fields; safecall;
    function Get_FormFields: FormFields; safecall;
    function Get_Styles: Styles; safecall;
    function Get_Frames: Frames; safecall;
    function Get_TablesOfFigures: TablesOfFigures; safecall;
    function Get_Variables: Variables; safecall;
    function Get_MailMerge: MailMerge; safecall;
    function Get_Envelope: Envelope; safecall;
    function Get_FullName: WideString; safecall;
    function Get_Revisions: Revisions; safecall;
    function Get_TablesOfContents: TablesOfContents; safecall;
    function Get_TablesOfAuthorities: TablesOfAuthorities; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_Windows: Windows; safecall;
    function Get_HasRoutingSlip: WordBool; safecall;
    procedure Set_HasRoutingSlip(prop: WordBool); safecall;
    function Get_RoutingSlip: RoutingSlip; safecall;
    function Get_Routed: WordBool; safecall;
    function Get_TablesOfAuthoritiesCategories: TablesOfAuthoritiesCategories; safecall;
    function Get_Indexes: Indexes; safecall;
    function Get_Saved: WordBool; safecall;
    procedure Set_Saved(prop: WordBool); safecall;
    function Get_Content: Range; safecall;
    function Get_ActiveWindow: Window_; safecall;
    function Get_Kind: WdDocumentKind; safecall;
    procedure Set_Kind(prop: WdDocumentKind); safecall;
    function Get_ReadOnly: WordBool; safecall;
    function Get_Subdocuments: Subdocuments; safecall;
    function Get_IsMasterDocument: WordBool; safecall;
    function Get_DefaultTabStop: Single; safecall;
    procedure Set_DefaultTabStop(prop: Single); safecall;
    function Get_EmbedTrueTypeFonts: WordBool; safecall;
    procedure Set_EmbedTrueTypeFonts(prop: WordBool); safecall;
    function Get_SaveFormsData: WordBool; safecall;
    procedure Set_SaveFormsData(prop: WordBool); safecall;
    function Get_ReadOnlyRecommended: WordBool; safecall;
    procedure Set_ReadOnlyRecommended(prop: WordBool); safecall;
    function Get_SaveSubsetFonts: WordBool; safecall;
    procedure Set_SaveSubsetFonts(prop: WordBool); safecall;
    function Get_Compatibility(Type_: WdCompatibility): WordBool; safecall;
    procedure Set_Compatibility(Type_: WdCompatibility; prop: WordBool); safecall;
    function Get_StoryRanges: StoryRanges; safecall;
    function Get_CommandBars: CommandBars; safecall;
    function Get_IsSubdocument: WordBool; safecall;
    function Get_SaveFormat: Integer; safecall;
    function Get_ProtectionType: WdProtectionType; safecall;
    function Get_Hyperlinks: Hyperlinks; safecall;
    function Get_Shapes: Shapes; safecall;
    function Get_ListTemplates: ListTemplates; safecall;
    function Get_Lists: Lists; safecall;
    function Get_UpdateStylesOnOpen: WordBool; safecall;
    procedure Set_UpdateStylesOnOpen(prop: WordBool); safecall;
    function Get_AttachedTemplate: OleVariant; safecall;
    procedure Set_AttachedTemplate(var prop: OleVariant); safecall;
    function Get_InlineShapes: InlineShapes; safecall;
    function Get_Background: Shape; safecall;
    procedure Set_Background(const prop: Shape); safecall;
    function Get_GrammarChecked: WordBool; safecall;
    procedure Set_GrammarChecked(prop: WordBool); safecall;
    function Get_SpellingChecked: WordBool; safecall;
    procedure Set_SpellingChecked(prop: WordBool); safecall;
    function Get_ShowGrammaticalErrors: WordBool; safecall;
    procedure Set_ShowGrammaticalErrors(prop: WordBool); safecall;
    function Get_ShowSpellingErrors: WordBool; safecall;
    procedure Set_ShowSpellingErrors(prop: WordBool); safecall;
    function Get_Versions: Versions; safecall;
    function Get_ShowSummary: WordBool; safecall;
    procedure Set_ShowSummary(prop: WordBool); safecall;
    function Get_SummaryViewMode: WdSummaryMode; safecall;
    procedure Set_SummaryViewMode(prop: WdSummaryMode); safecall;
    function Get_SummaryLength: Integer; safecall;
    procedure Set_SummaryLength(prop: Integer); safecall;
    function Get_PrintFractionalWidths: WordBool; safecall;
    procedure Set_PrintFractionalWidths(prop: WordBool); safecall;
    function Get_PrintPostScriptOverText: WordBool; safecall;
    procedure Set_PrintPostScriptOverText(prop: WordBool); safecall;
    function Get_Container: IDispatch; safecall;
    function Get_PrintFormsData: WordBool; safecall;
    procedure Set_PrintFormsData(prop: WordBool); safecall;
    function Get_ListParagraphs: ListParagraphs; safecall;
    procedure Set_Password(const Param1: WideString); safecall;
    procedure Set_WritePassword(const Param1: WideString); safecall;
    function Get_HasPassword: WordBool; safecall;
    function Get_WriteReserved: WordBool; safecall;
    function Get_ActiveWritingStyle(var LanguageID: OleVariant): WideString; safecall;
    procedure Set_ActiveWritingStyle(var LanguageID: OleVariant; const prop: WideString); safecall;
    function Get_UserControl: WordBool; safecall;
    procedure Set_UserControl(prop: WordBool); safecall;
    function Get_HasMailer: WordBool; safecall;
    procedure Set_HasMailer(prop: WordBool); safecall;
    function Get_Mailer: Mailer; safecall;
    function Get_ReadabilityStatistics: ReadabilityStatistics; safecall;
    function Get_GrammaticalErrors: ProofreadingErrors; safecall;
    function Get_SpellingErrors: ProofreadingErrors; safecall;
    function Get_VBProject: VBProject; safecall;
    function Get_FormsDesign: WordBool; safecall;
    function Get__CodeName: WideString; safecall;
    procedure Set__CodeName(const prop: WideString); safecall;
    function Get_CodeName: WideString; safecall;
    function Get_SnapToGrid: WordBool; safecall;
    procedure Set_SnapToGrid(prop: WordBool); safecall;
    function Get_SnapToShapes: WordBool; safecall;
    procedure Set_SnapToShapes(prop: WordBool); safecall;
    function Get_GridDistanceHorizontal: Single; safecall;
    procedure Set_GridDistanceHorizontal(prop: Single); safecall;
    function Get_GridDistanceVertical: Single; safecall;
    procedure Set_GridDistanceVertical(prop: Single); safecall;
    function Get_GridOriginHorizontal: Single; safecall;
    procedure Set_GridOriginHorizontal(prop: Single); safecall;
    function Get_GridOriginVertical: Single; safecall;
    procedure Set_GridOriginVertical(prop: Single); safecall;
    function Get_GridSpaceBetweenHorizontalLines: Integer; safecall;
    procedure Set_GridSpaceBetweenHorizontalLines(prop: Integer); safecall;
    function Get_GridSpaceBetweenVerticalLines: Integer; safecall;
    procedure Set_GridSpaceBetweenVerticalLines(prop: Integer); safecall;
    function Get_GridOriginFromMargin: WordBool; safecall;
    procedure Set_GridOriginFromMargin(prop: WordBool); safecall;
    function Get_KerningByAlgorithm: WordBool; safecall;
    procedure Set_KerningByAlgorithm(prop: WordBool); safecall;
    function Get_JustificationMode: WdJustificationMode; safecall;
    procedure Set_JustificationMode(prop: WdJustificationMode); safecall;
    function Get_FarEastLineBreakLevel: WdFarEastLineBreakLevel; safecall;
    procedure Set_FarEastLineBreakLevel(prop: WdFarEastLineBreakLevel); safecall;
    function Get_NoLineBreakBefore: WideString; safecall;
    procedure Set_NoLineBreakBefore(const prop: WideString); safecall;
    function Get_NoLineBreakAfter: WideString; safecall;
    procedure Set_NoLineBreakAfter(const prop: WideString); safecall;
    function Get_TrackRevisions: WordBool; safecall;
    procedure Set_TrackRevisions(prop: WordBool); safecall;
    function Get_PrintRevisions: WordBool; safecall;
    procedure Set_PrintRevisions(prop: WordBool); safecall;
    function Get_ShowRevisions: WordBool; safecall;
    procedure Set_ShowRevisions(prop: WordBool); safecall;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); safecall;
    procedure SaveAs(var FileName: OleVariant; var FileFormat: OleVariant; 
                     var LockComments: OleVariant; var Password: OleVariant; 
                     var AddToRecentFiles: OleVariant; var WritePassword: OleVariant; 
                     var ReadOnlyRecommended: OleVariant; var EmbedTrueTypeFonts: OleVariant; 
                     var SaveNativePictureFormat: OleVariant; var SaveFormsData: OleVariant; 
                     var SaveAsAOCELetter: OleVariant); safecall;
    procedure Repaginate; safecall;
    procedure FitToPages; safecall;
    procedure ManualHyphenation; safecall;
    procedure Select; safecall;
    procedure DataForm; safecall;
    procedure Route; safecall;
    procedure Save; safecall;
    procedure PrintOutOld(var Background: OleVariant; var Append: OleVariant; 
                          var Range: OleVariant; var OutputFileName: OleVariant; 
                          var From: OleVariant; var To_: OleVariant; var Item: OleVariant; 
                          var Copies: OleVariant; var Pages: OleVariant; var PageType: OleVariant; 
                          var PrintToFile: OleVariant; var Collate: OleVariant; 
                          var ActivePrinterMacGX: OleVariant; var ManualDuplexPrint: OleVariant); safecall;
    procedure SendMail; safecall;
    function Range(var Start: OleVariant; var End_: OleVariant): Range; safecall;
    procedure RunAutoMacro(Which: WdAutoMacros); safecall;
    procedure Activate; safecall;
    procedure PrintPreview; safecall;
    function GoTo_(var What: OleVariant; var Which: OleVariant; var Count: OleVariant; 
                   var Name: OleVariant): Range; safecall;
    function Undo(var Times: OleVariant): WordBool; safecall;
    function Redo(var Times: OleVariant): WordBool; safecall;
    function ComputeStatistics(Statistic: WdStatistic; var IncludeFootnotesAndEndnotes: OleVariant): Integer; safecall;
    procedure MakeCompatibilityDefault; safecall;
    procedure Protect(Type_: WdProtectionType; var NoReset: OleVariant; var Password: OleVariant); safecall;
    procedure Unprotect(var Password: OleVariant); safecall;
    procedure EditionOptions(Type_: WdEditionType; Option: WdEditionOption; const Name: WideString; 
                             var Format: OleVariant); safecall;
    procedure RunLetterWizard(var LetterContent: OleVariant; var WizardMode: OleVariant); safecall;
    function GetLetterContent: LetterContent; safecall;
    procedure SetLetterContent(var LetterContent: OleVariant); safecall;
    procedure CopyStylesFromTemplate(const Template: WideString); safecall;
    procedure UpdateStyles; safecall;
    procedure CheckGrammar; safecall;
    procedure CheckSpelling(var CustomDictionary: OleVariant; var IgnoreUppercase: OleVariant; 
                            var AlwaysSuggest: OleVariant; var CustomDictionary2: OleVariant; 
                            var CustomDictionary3: OleVariant; var CustomDictionary4: OleVariant; 
                            var CustomDictionary5: OleVariant; var CustomDictionary6: OleVariant; 
                            var CustomDictionary7: OleVariant; var CustomDictionary8: OleVariant; 
                            var CustomDictionary9: OleVariant; var CustomDictionary10: OleVariant); safecall;
    procedure FollowHyperlink(var Address: OleVariant; var SubAddress: OleVariant; 
                              var NewWindow: OleVariant; var AddHistory: OleVariant; 
                              var ExtraInfo: OleVariant; var Method: OleVariant; 
                              var HeaderInfo: OleVariant); safecall;
    procedure AddToFavorites; safecall;
    procedure Reload; safecall;
    function AutoSummarize(var Length: OleVariant; var Mode: OleVariant; 
                           var UpdateProperties: OleVariant): Range; safecall;
    procedure RemoveNumbers(var NumberType: OleVariant); safecall;
    procedure ConvertNumbersToText(var NumberType: OleVariant); safecall;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; safecall;
    procedure Post; safecall;
    procedure ToggleFormsDesign; safecall;
    procedure Compare(const Name: WideString); safecall;
    procedure UpdateSummaryProperties; safecall;
    function GetCrossReferenceItems(var ReferenceType: OleVariant): OleVariant; safecall;
    procedure AutoFormat; safecall;
    procedure ViewCode; safecall;
    procedure ViewPropertyBrowser; safecall;
    procedure ForwardMailer; safecall;
    procedure Reply; safecall;
    procedure ReplyAll; safecall;
    procedure SendMailer(var FileFormat: OleVariant; var Priority: OleVariant); safecall;
    procedure UndoClear; safecall;
    procedure PresentIt; safecall;
    procedure SendFax(const Address: WideString; var Subject: OleVariant); safecall;
    procedure Merge(const FileName: WideString); safecall;
    procedure ClosePrintPreview; safecall;
    procedure CheckConsistency; safecall;
    function CreateLetterContent(const DateFormat: WideString; IncludeHeaderFooter: WordBool; 
                                 const PageDesign: WideString; LetterStyle: WdLetterStyle; 
                                 Letterhead: WordBool; LetterheadLocation: WdLetterheadLocation; 
                                 LetterheadSize: Single; const RecipientName: WideString; 
                                 const RecipientAddress: WideString; const Salutation: WideString; 
                                 SalutationType: WdSalutationType; 
                                 const RecipientReference: WideString; 
                                 const MailingInstructions: WideString; 
                                 const AttentionLine: WideString; const Subject: WideString; 
                                 const CCList: WideString; const ReturnAddress: WideString; 
                                 const SenderName: WideString; const Closing: WideString; 
                                 const SenderCompany: WideString; const SenderJobTitle: WideString; 
                                 const SenderInitials: WideString; EnclosureNumber: Integer; 
                                 var InfoBlock: OleVariant; var RecipientCode: OleVariant; 
                                 var RecipientGender: OleVariant; 
                                 var ReturnAddressShortForm: OleVariant; 
                                 var SenderCity: OleVariant; var SenderCode: OleVariant; 
                                 var SenderGender: OleVariant; var SenderReference: OleVariant): LetterContent; safecall;
    procedure AcceptAllRevisions; safecall;
    procedure RejectAllRevisions; safecall;
    procedure DetectLanguage; safecall;
    procedure ApplyTheme(const Name: WideString); safecall;
    procedure RemoveTheme; safecall;
    procedure WebPagePreview; safecall;
    procedure ReloadAs(Encoding: MsoEncoding); safecall;
    function Get_ActiveTheme: WideString; safecall;
    function Get_ActiveThemeDisplayName: WideString; safecall;
    function Get_Email: Email; safecall;
    function Get_Scripts: Scripts; safecall;
    function Get_LanguageDetected: WordBool; safecall;
    procedure Set_LanguageDetected(prop: WordBool); safecall;
    function Get_FarEastLineBreakLanguage: WdFarEastLineBreakLanguageID; safecall;
    procedure Set_FarEastLineBreakLanguage(prop: WdFarEastLineBreakLanguageID); safecall;
    function Get_Frameset: Frameset; safecall;
    function Get_ClickAndTypeParagraphStyle: OleVariant; safecall;
    procedure Set_ClickAndTypeParagraphStyle(var prop: OleVariant); safecall;
    function Get_HTMLProject: HTMLProject; safecall;
    function Get_WebOptions: WebOptions; safecall;
    function Get_OpenEncoding: MsoEncoding; safecall;
    function Get_SaveEncoding: MsoEncoding; safecall;
    procedure Set_SaveEncoding(prop: MsoEncoding); safecall;
    function Get_OptimizeForWord97: WordBool; safecall;
    procedure Set_OptimizeForWord97(prop: WordBool); safecall;
    function Get_VBASigned: WordBool; safecall;
    procedure PrintOut(var Background: OleVariant; var Append: OleVariant; var Range: OleVariant; 
                       var OutputFileName: OleVariant; var From: OleVariant; var To_: OleVariant; 
                       var Item: OleVariant; var Copies: OleVariant; var Pages: OleVariant; 
                       var PageType: OleVariant; var PrintToFile: OleVariant; 
                       var Collate: OleVariant; var ActivePrinterMacGX: OleVariant; 
                       var ManualDuplexPrint: OleVariant; var PrintZoomColumn: OleVariant; 
                       var PrintZoomRow: OleVariant; var PrintZoomPaperWidth: OleVariant; 
                       var PrintZoomPaperHeight: OleVariant); safecall;
    procedure sblt(const s: WideString); safecall;
    property Name: WideString read Get_Name;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property BuiltInDocumentProperties: IDispatch read Get_BuiltInDocumentProperties;
    property CustomDocumentProperties: IDispatch read Get_CustomDocumentProperties;
    property Path: WideString read Get_Path;
    property Bookmarks: Bookmarks read Get_Bookmarks;
    property Tables: Tables read Get_Tables;
    property Footnotes: Footnotes read Get_Footnotes;
    property Endnotes: Endnotes read Get_Endnotes;
    property Comments: Comments read Get_Comments;
    property Type_: WdDocumentType read Get_Type_;
    property AutoHyphenation: WordBool read Get_AutoHyphenation write Set_AutoHyphenation;
    property HyphenateCaps: WordBool read Get_HyphenateCaps write Set_HyphenateCaps;
    property HyphenationZone: Integer read Get_HyphenationZone write Set_HyphenationZone;
    property ConsecutiveHyphensLimit: Integer read Get_ConsecutiveHyphensLimit write Set_ConsecutiveHyphensLimit;
    property Sections: Sections read Get_Sections;
    property Paragraphs: Paragraphs read Get_Paragraphs;
    property Words: Words read Get_Words;
    property Sentences: Sentences read Get_Sentences;
    property Characters: Characters read Get_Characters;
    property Fields: Fields read Get_Fields;
    property FormFields: FormFields read Get_FormFields;
    property Styles: Styles read Get_Styles;
    property Frames: Frames read Get_Frames;
    property TablesOfFigures: TablesOfFigures read Get_TablesOfFigures;
    property Variables: Variables read Get_Variables;
    property MailMerge: MailMerge read Get_MailMerge;
    property Envelope: Envelope read Get_Envelope;
    property FullName: WideString read Get_FullName;
    property Revisions: Revisions read Get_Revisions;
    property TablesOfContents: TablesOfContents read Get_TablesOfContents;
    property TablesOfAuthorities: TablesOfAuthorities read Get_TablesOfAuthorities;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property Windows: Windows read Get_Windows;
    property HasRoutingSlip: WordBool read Get_HasRoutingSlip write Set_HasRoutingSlip;
    property RoutingSlip: RoutingSlip read Get_RoutingSlip;
    property Routed: WordBool read Get_Routed;
    property TablesOfAuthoritiesCategories: TablesOfAuthoritiesCategories read Get_TablesOfAuthoritiesCategories;
    property Indexes: Indexes read Get_Indexes;
    property Saved: WordBool read Get_Saved write Set_Saved;
    property Content: Range read Get_Content;
    property ActiveWindow: Window_ read Get_ActiveWindow;
    property Kind: WdDocumentKind read Get_Kind write Set_Kind;
    property ReadOnly: WordBool read Get_ReadOnly;
    property Subdocuments: Subdocuments read Get_Subdocuments;
    property IsMasterDocument: WordBool read Get_IsMasterDocument;
    property DefaultTabStop: Single read Get_DefaultTabStop write Set_DefaultTabStop;
    property EmbedTrueTypeFonts: WordBool read Get_EmbedTrueTypeFonts write Set_EmbedTrueTypeFonts;
    property SaveFormsData: WordBool read Get_SaveFormsData write Set_SaveFormsData;
    property ReadOnlyRecommended: WordBool read Get_ReadOnlyRecommended write Set_ReadOnlyRecommended;
    property SaveSubsetFonts: WordBool read Get_SaveSubsetFonts write Set_SaveSubsetFonts;
    property Compatibility[Type_: WdCompatibility]: WordBool read Get_Compatibility write Set_Compatibility;
    property StoryRanges: StoryRanges read Get_StoryRanges;
    property CommandBars: CommandBars read Get_CommandBars;
    property IsSubdocument: WordBool read Get_IsSubdocument;
    property SaveFormat: Integer read Get_SaveFormat;
    property ProtectionType: WdProtectionType read Get_ProtectionType;
    property Hyperlinks: Hyperlinks read Get_Hyperlinks;
    property Shapes: Shapes read Get_Shapes;
    property ListTemplates: ListTemplates read Get_ListTemplates;
    property Lists: Lists read Get_Lists;
    property UpdateStylesOnOpen: WordBool read Get_UpdateStylesOnOpen write Set_UpdateStylesOnOpen;
    property InlineShapes: InlineShapes read Get_InlineShapes;
    property Background: Shape read Get_Background write Set_Background;
    property GrammarChecked: WordBool read Get_GrammarChecked write Set_GrammarChecked;
    property SpellingChecked: WordBool read Get_SpellingChecked write Set_SpellingChecked;
    property ShowGrammaticalErrors: WordBool read Get_ShowGrammaticalErrors write Set_ShowGrammaticalErrors;
    property ShowSpellingErrors: WordBool read Get_ShowSpellingErrors write Set_ShowSpellingErrors;
    property Versions: Versions read Get_Versions;
    property ShowSummary: WordBool read Get_ShowSummary write Set_ShowSummary;
    property SummaryViewMode: WdSummaryMode read Get_SummaryViewMode write Set_SummaryViewMode;
    property SummaryLength: Integer read Get_SummaryLength write Set_SummaryLength;
    property PrintFractionalWidths: WordBool read Get_PrintFractionalWidths write Set_PrintFractionalWidths;
    property PrintPostScriptOverText: WordBool read Get_PrintPostScriptOverText write Set_PrintPostScriptOverText;
    property Container: IDispatch read Get_Container;
    property PrintFormsData: WordBool read Get_PrintFormsData write Set_PrintFormsData;
    property ListParagraphs: ListParagraphs read Get_ListParagraphs;
    property Password: WideString write Set_Password;
    property WritePassword: WideString write Set_WritePassword;
    property HasPassword: WordBool read Get_HasPassword;
    property WriteReserved: WordBool read Get_WriteReserved;
    property ActiveWritingStyle[var LanguageID: OleVariant]: WideString read Get_ActiveWritingStyle write Set_ActiveWritingStyle;
    property UserControl: WordBool read Get_UserControl write Set_UserControl;
    property HasMailer: WordBool read Get_HasMailer write Set_HasMailer;
    property Mailer: Mailer read Get_Mailer;
    property ReadabilityStatistics: ReadabilityStatistics read Get_ReadabilityStatistics;
    property GrammaticalErrors: ProofreadingErrors read Get_GrammaticalErrors;
    property SpellingErrors: ProofreadingErrors read Get_SpellingErrors;
    property VBProject: VBProject read Get_VBProject;
    property FormsDesign: WordBool read Get_FormsDesign;
    property _CodeName: WideString read Get__CodeName write Set__CodeName;
    property CodeName: WideString read Get_CodeName;
    property SnapToGrid: WordBool read Get_SnapToGrid write Set_SnapToGrid;
    property SnapToShapes: WordBool read Get_SnapToShapes write Set_SnapToShapes;
    property GridDistanceHorizontal: Single read Get_GridDistanceHorizontal write Set_GridDistanceHorizontal;
    property GridDistanceVertical: Single read Get_GridDistanceVertical write Set_GridDistanceVertical;
    property GridOriginHorizontal: Single read Get_GridOriginHorizontal write Set_GridOriginHorizontal;
    property GridOriginVertical: Single read Get_GridOriginVertical write Set_GridOriginVertical;
    property GridSpaceBetweenHorizontalLines: Integer read Get_GridSpaceBetweenHorizontalLines write Set_GridSpaceBetweenHorizontalLines;
    property GridSpaceBetweenVerticalLines: Integer read Get_GridSpaceBetweenVerticalLines write Set_GridSpaceBetweenVerticalLines;
    property GridOriginFromMargin: WordBool read Get_GridOriginFromMargin write Set_GridOriginFromMargin;
    property KerningByAlgorithm: WordBool read Get_KerningByAlgorithm write Set_KerningByAlgorithm;
    property JustificationMode: WdJustificationMode read Get_JustificationMode write Set_JustificationMode;
    property FarEastLineBreakLevel: WdFarEastLineBreakLevel read Get_FarEastLineBreakLevel write Set_FarEastLineBreakLevel;
    property NoLineBreakBefore: WideString read Get_NoLineBreakBefore write Set_NoLineBreakBefore;
    property NoLineBreakAfter: WideString read Get_NoLineBreakAfter write Set_NoLineBreakAfter;
    property TrackRevisions: WordBool read Get_TrackRevisions write Set_TrackRevisions;
    property PrintRevisions: WordBool read Get_PrintRevisions write Set_PrintRevisions;
    property ShowRevisions: WordBool read Get_ShowRevisions write Set_ShowRevisions;
    property ActiveTheme: WideString read Get_ActiveTheme;
    property ActiveThemeDisplayName: WideString read Get_ActiveThemeDisplayName;
    property Email: Email read Get_Email;
    property Scripts: Scripts read Get_Scripts;
    property LanguageDetected: WordBool read Get_LanguageDetected write Set_LanguageDetected;
    property FarEastLineBreakLanguage: WdFarEastLineBreakLanguageID read Get_FarEastLineBreakLanguage write Set_FarEastLineBreakLanguage;
    property Frameset: Frameset read Get_Frameset;
    property HTMLProject: HTMLProject read Get_HTMLProject;
    property WebOptions: WebOptions read Get_WebOptions;
    property OpenEncoding: MsoEncoding read Get_OpenEncoding;
    property SaveEncoding: MsoEncoding read Get_SaveEncoding write Set_SaveEncoding;
    property OptimizeForWord97: WordBool read Get_OptimizeForWord97 write Set_OptimizeForWord97;
    property VBASigned: WordBool read Get_VBASigned;
  end;

// *********************************************************************//
// DispIntf:  _DocumentDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {0002096B-0000-0000-C000-000000000046}
// *********************************************************************//
  _DocumentDisp = dispinterface
    ['{0002096B-0000-0000-C000-000000000046}']
    property Name: WideString readonly dispid 0;
    property Application_: Application_ readonly dispid 1;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property BuiltInDocumentProperties: IDispatch readonly dispid 1000;
    property CustomDocumentProperties: IDispatch readonly dispid 2;
    property Path: WideString readonly dispid 3;
    property Bookmarks: Bookmarks readonly dispid 4;
    property Tables: Tables readonly dispid 6;
    property Footnotes: Footnotes readonly dispid 7;
    property Endnotes: Endnotes readonly dispid 8;
    property Comments: Comments readonly dispid 9;
    property Type_: WdDocumentType readonly dispid 10;
    property AutoHyphenation: WordBool dispid 11;
    property HyphenateCaps: WordBool dispid 12;
    property HyphenationZone: Integer dispid 13;
    property ConsecutiveHyphensLimit: Integer dispid 14;
    property Sections: Sections readonly dispid 15;
    property Paragraphs: Paragraphs readonly dispid 16;
    property Words: Words readonly dispid 17;
    property Sentences: Sentences readonly dispid 18;
    property Characters: Characters readonly dispid 19;
    property Fields: Fields readonly dispid 20;
    property FormFields: FormFields readonly dispid 21;
    property Styles: Styles readonly dispid 22;
    property Frames: Frames readonly dispid 23;
    property TablesOfFigures: TablesOfFigures readonly dispid 25;
    property Variables: Variables readonly dispid 26;
    property MailMerge: MailMerge readonly dispid 27;
    property Envelope: Envelope readonly dispid 28;
    property FullName: WideString readonly dispid 29;
    property Revisions: Revisions readonly dispid 30;
    property TablesOfContents: TablesOfContents readonly dispid 31;
    property TablesOfAuthorities: TablesOfAuthorities readonly dispid 32;
    property PageSetup: PageSetup dispid 1101;
    property Windows: Windows readonly dispid 34;
    property HasRoutingSlip: WordBool dispid 35;
    property RoutingSlip: RoutingSlip readonly dispid 36;
    property Routed: WordBool readonly dispid 37;
    property TablesOfAuthoritiesCategories: TablesOfAuthoritiesCategories readonly dispid 38;
    property Indexes: Indexes readonly dispid 39;
    property Saved: WordBool dispid 40;
    property Content: Range readonly dispid 41;
    property ActiveWindow: Window_ readonly dispid 42;
    property Kind: WdDocumentKind dispid 43;
    property ReadOnly: WordBool readonly dispid 44;
    property Subdocuments: Subdocuments readonly dispid 45;
    property IsMasterDocument: WordBool readonly dispid 46;
    property DefaultTabStop: Single dispid 48;
    property EmbedTrueTypeFonts: WordBool dispid 50;
    property SaveFormsData: WordBool dispid 51;
    property ReadOnlyRecommended: WordBool dispid 52;
    property SaveSubsetFonts: WordBool dispid 53;
    property Compatibility[Type_: WdCompatibility]: WordBool dispid 55;
    property StoryRanges: StoryRanges readonly dispid 56;
    property CommandBars: CommandBars readonly dispid 57;
    property IsSubdocument: WordBool readonly dispid 58;
    property SaveFormat: Integer readonly dispid 59;
    property ProtectionType: WdProtectionType readonly dispid 60;
    property Hyperlinks: Hyperlinks readonly dispid 61;
    property Shapes: Shapes readonly dispid 62;
    property ListTemplates: ListTemplates readonly dispid 63;
    property Lists: Lists readonly dispid 64;
    property UpdateStylesOnOpen: WordBool dispid 66;
    function AttachedTemplate: OleVariant; dispid 67;
    property InlineShapes: InlineShapes readonly dispid 68;
    property Background: Shape dispid 69;
    property GrammarChecked: WordBool dispid 70;
    property SpellingChecked: WordBool dispid 71;
    property ShowGrammaticalErrors: WordBool dispid 72;
    property ShowSpellingErrors: WordBool dispid 73;
    property Versions: Versions readonly dispid 75;
    property ShowSummary: WordBool dispid 76;
    property SummaryViewMode: WdSummaryMode dispid 77;
    property SummaryLength: Integer dispid 78;
    property PrintFractionalWidths: WordBool dispid 79;
    property PrintPostScriptOverText: WordBool dispid 80;
    property Container: IDispatch readonly dispid 82;
    property PrintFormsData: WordBool dispid 83;
    property ListParagraphs: ListParagraphs readonly dispid 84;
    property Password: WideString writeonly dispid 85;
    property WritePassword: WideString writeonly dispid 86;
    property HasPassword: WordBool readonly dispid 87;
    property WriteReserved: WordBool readonly dispid 88;
    property ActiveWritingStyle[var LanguageID: OleVariant]: WideString dispid 90;
    property UserControl: WordBool dispid 92;
    property HasMailer: WordBool dispid 93;
    property Mailer: Mailer readonly dispid 94;
    property ReadabilityStatistics: ReadabilityStatistics readonly dispid 96;
    property GrammaticalErrors: ProofreadingErrors readonly dispid 97;
    property SpellingErrors: ProofreadingErrors readonly dispid 98;
    property VBProject: VBProject readonly dispid 99;
    property FormsDesign: WordBool readonly dispid 100;
    property _CodeName: WideString dispid -2147418112;
    property CodeName: WideString readonly dispid 262;
    property SnapToGrid: WordBool dispid 300;
    property SnapToShapes: WordBool dispid 301;
    property GridDistanceHorizontal: Single dispid 302;
    property GridDistanceVertical: Single dispid 303;
    property GridOriginHorizontal: Single dispid 304;
    property GridOriginVertical: Single dispid 305;
    property GridSpaceBetweenHorizontalLines: Integer dispid 306;
    property GridSpaceBetweenVerticalLines: Integer dispid 307;
    property GridOriginFromMargin: WordBool dispid 308;
    property KerningByAlgorithm: WordBool dispid 309;
    property JustificationMode: WdJustificationMode dispid 310;
    property FarEastLineBreakLevel: WdFarEastLineBreakLevel dispid 311;
    property NoLineBreakBefore: WideString dispid 312;
    property NoLineBreakAfter: WideString dispid 313;
    property TrackRevisions: WordBool dispid 314;
    property PrintRevisions: WordBool dispid 315;
    property ShowRevisions: WordBool dispid 316;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); dispid 1105;
    procedure SaveAs(var FileName: OleVariant; var FileFormat: OleVariant; 
                     var LockComments: OleVariant; var Password: OleVariant; 
                     var AddToRecentFiles: OleVariant; var WritePassword: OleVariant; 
                     var ReadOnlyRecommended: OleVariant; var EmbedTrueTypeFonts: OleVariant; 
                     var SaveNativePictureFormat: OleVariant; var SaveFormsData: OleVariant; 
                     var SaveAsAOCELetter: OleVariant); dispid 102;
    procedure Repaginate; dispid 103;
    procedure FitToPages; dispid 104;
    procedure ManualHyphenation; dispid 105;
    procedure Select; dispid 65535;
    procedure DataForm; dispid 106;
    procedure Route; dispid 107;
    procedure Save; dispid 108;
    procedure PrintOutOld(var Background: OleVariant; var Append: OleVariant; 
                          var Range: OleVariant; var OutputFileName: OleVariant; 
                          var From: OleVariant; var To_: OleVariant; var Item: OleVariant; 
                          var Copies: OleVariant; var Pages: OleVariant; var PageType: OleVariant; 
                          var PrintToFile: OleVariant; var Collate: OleVariant; 
                          var ActivePrinterMacGX: OleVariant; var ManualDuplexPrint: OleVariant); dispid 109;
    procedure SendMail; dispid 110;
    function Range(var Start: OleVariant; var End_: OleVariant): Range; dispid 2000;
    procedure RunAutoMacro(Which: WdAutoMacros); dispid 112;
    procedure Activate; dispid 113;
    procedure PrintPreview; dispid 114;
    function GoTo_(var What: OleVariant; var Which: OleVariant; var Count: OleVariant; 
                   var Name: OleVariant): Range; dispid 115;
    function Undo(var Times: OleVariant): WordBool; dispid 116;
    function Redo(var Times: OleVariant): WordBool; dispid 117;
    function ComputeStatistics(Statistic: WdStatistic; var IncludeFootnotesAndEndnotes: OleVariant): Integer; dispid 118;
    procedure MakeCompatibilityDefault; dispid 119;
    procedure Protect(Type_: WdProtectionType; var NoReset: OleVariant; var Password: OleVariant); dispid 120;
    procedure Unprotect(var Password: OleVariant); dispid 121;
    procedure EditionOptions(Type_: WdEditionType; Option: WdEditionOption; const Name: WideString; 
                             var Format: OleVariant); dispid 122;
    procedure RunLetterWizard(var LetterContent: OleVariant; var WizardMode: OleVariant); dispid 123;
    function GetLetterContent: LetterContent; dispid 124;
    procedure SetLetterContent(var LetterContent: OleVariant); dispid 125;
    procedure CopyStylesFromTemplate(const Template: WideString); dispid 126;
    procedure UpdateStyles; dispid 127;
    procedure CheckGrammar; dispid 131;
    procedure CheckSpelling(var CustomDictionary: OleVariant; var IgnoreUppercase: OleVariant; 
                            var AlwaysSuggest: OleVariant; var CustomDictionary2: OleVariant; 
                            var CustomDictionary3: OleVariant; var CustomDictionary4: OleVariant; 
                            var CustomDictionary5: OleVariant; var CustomDictionary6: OleVariant; 
                            var CustomDictionary7: OleVariant; var CustomDictionary8: OleVariant; 
                            var CustomDictionary9: OleVariant; var CustomDictionary10: OleVariant); dispid 132;
    procedure FollowHyperlink(var Address: OleVariant; var SubAddress: OleVariant; 
                              var NewWindow: OleVariant; var AddHistory: OleVariant; 
                              var ExtraInfo: OleVariant; var Method: OleVariant; 
                              var HeaderInfo: OleVariant); dispid 135;
    procedure AddToFavorites; dispid 136;
    procedure Reload; dispid 137;
    function AutoSummarize(var Length: OleVariant; var Mode: OleVariant; 
                           var UpdateProperties: OleVariant): Range; dispid 138;
    procedure RemoveNumbers(var NumberType: OleVariant); dispid 140;
    procedure ConvertNumbersToText(var NumberType: OleVariant); dispid 141;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; dispid 142;
    procedure Post; dispid 143;
    procedure ToggleFormsDesign; dispid 144;
    procedure Compare(const Name: WideString); dispid 145;
    procedure UpdateSummaryProperties; dispid 146;
    function GetCrossReferenceItems(var ReferenceType: OleVariant): OleVariant; dispid 147;
    procedure AutoFormat; dispid 148;
    procedure ViewCode; dispid 149;
    procedure ViewPropertyBrowser; dispid 150;
    procedure ForwardMailer; dispid 250;
    procedure Reply; dispid 251;
    procedure ReplyAll; dispid 252;
    procedure SendMailer(var FileFormat: OleVariant; var Priority: OleVariant); dispid 253;
    procedure UndoClear; dispid 254;
    procedure PresentIt; dispid 255;
    procedure SendFax(const Address: WideString; var Subject: OleVariant); dispid 256;
    procedure Merge(const FileName: WideString); dispid 257;
    procedure ClosePrintPreview; dispid 258;
    procedure CheckConsistency; dispid 259;
    function CreateLetterContent(const DateFormat: WideString; IncludeHeaderFooter: WordBool; 
                                 const PageDesign: WideString; LetterStyle: WdLetterStyle; 
                                 Letterhead: WordBool; LetterheadLocation: WdLetterheadLocation; 
                                 LetterheadSize: Single; const RecipientName: WideString; 
                                 const RecipientAddress: WideString; const Salutation: WideString; 
                                 SalutationType: WdSalutationType; 
                                 const RecipientReference: WideString; 
                                 const MailingInstructions: WideString; 
                                 const AttentionLine: WideString; const Subject: WideString; 
                                 const CCList: WideString; const ReturnAddress: WideString; 
                                 const SenderName: WideString; const Closing: WideString; 
                                 const SenderCompany: WideString; const SenderJobTitle: WideString; 
                                 const SenderInitials: WideString; EnclosureNumber: Integer; 
                                 var InfoBlock: OleVariant; var RecipientCode: OleVariant; 
                                 var RecipientGender: OleVariant; 
                                 var ReturnAddressShortForm: OleVariant; 
                                 var SenderCity: OleVariant; var SenderCode: OleVariant; 
                                 var SenderGender: OleVariant; var SenderReference: OleVariant): LetterContent; dispid 260;
    procedure AcceptAllRevisions; dispid 317;
    procedure RejectAllRevisions; dispid 318;
    procedure DetectLanguage; dispid 151;
    procedure ApplyTheme(const Name: WideString); dispid 322;
    procedure RemoveTheme; dispid 323;
    procedure WebPagePreview; dispid 325;
    procedure ReloadAs(Encoding: MsoEncoding); dispid 331;
    property ActiveTheme: WideString readonly dispid 540;
    property ActiveThemeDisplayName: WideString readonly dispid 541;
    property Email: Email readonly dispid 319;
    property Scripts: Scripts readonly dispid 320;
    property LanguageDetected: WordBool dispid 321;
    property FarEastLineBreakLanguage: WdFarEastLineBreakLanguageID dispid 326;
    property Frameset: Frameset readonly dispid 327;
    function ClickAndTypeParagraphStyle: OleVariant; dispid 328;
    property HTMLProject: HTMLProject readonly dispid 329;
    property WebOptions: WebOptions readonly dispid 330;
    property OpenEncoding: MsoEncoding readonly dispid 332;
    property SaveEncoding: MsoEncoding dispid 333;
    property OptimizeForWord97: WordBool dispid 334;
    property VBASigned: WordBool readonly dispid 335;
    procedure PrintOut(var Background: OleVariant; var Append: OleVariant; var Range: OleVariant; 
                       var OutputFileName: OleVariant; var From: OleVariant; var To_: OleVariant; 
                       var Item: OleVariant; var Copies: OleVariant; var Pages: OleVariant; 
                       var PageType: OleVariant; var PrintToFile: OleVariant; 
                       var Collate: OleVariant; var ActivePrinterMacGX: OleVariant; 
                       var ManualDuplexPrint: OleVariant; var PrintZoomColumn: OleVariant; 
                       var PrintZoomRow: OleVariant; var PrintZoomPaperWidth: OleVariant; 
                       var PrintZoomPaperHeight: OleVariant); dispid 444;
    procedure sblt(const s: WideString); dispid 445;
  end;

// *********************************************************************//
// Interface: Template
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096A-0000-0000-C000-000000000046}
// *********************************************************************//
  Template = interface(IDispatch)
    ['{0002096A-0000-0000-C000-000000000046}']
    function Get_Name: WideString; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Path: WideString; safecall;
    function Get_AutoTextEntries: AutoTextEntries; safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_Saved: WordBool; safecall;
    procedure Set_Saved(prop: WordBool); safecall;
    function Get_Type_: WdTemplateType; safecall;
    function Get_FullName: WideString; safecall;
    function Get_BuiltInDocumentProperties: IDispatch; safecall;
    function Get_CustomDocumentProperties: IDispatch; safecall;
    function Get_ListTemplates: ListTemplates; safecall;
    function Get_LanguageIDFarEast: WdLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WdLanguageID); safecall;
    function Get_VBProject: VBProject; safecall;
    function Get_KerningByAlgorithm: WordBool; safecall;
    procedure Set_KerningByAlgorithm(prop: WordBool); safecall;
    function Get_JustificationMode: WdJustificationMode; safecall;
    procedure Set_JustificationMode(prop: WdJustificationMode); safecall;
    function Get_FarEastLineBreakLevel: WdFarEastLineBreakLevel; safecall;
    procedure Set_FarEastLineBreakLevel(prop: WdFarEastLineBreakLevel); safecall;
    function Get_NoLineBreakBefore: WideString; safecall;
    procedure Set_NoLineBreakBefore(const prop: WideString); safecall;
    function Get_NoLineBreakAfter: WideString; safecall;
    procedure Set_NoLineBreakAfter(const prop: WideString); safecall;
    function OpenAsDocument: Document; safecall;
    procedure Save; safecall;
    function Get_NoProofing: Integer; safecall;
    procedure Set_NoProofing(prop: Integer); safecall;
    function Get_FarEastLineBreakLanguage: WdFarEastLineBreakLanguageID; safecall;
    procedure Set_FarEastLineBreakLanguage(prop: WdFarEastLineBreakLanguageID); safecall;
    property Name: WideString read Get_Name;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Path: WideString read Get_Path;
    property AutoTextEntries: AutoTextEntries read Get_AutoTextEntries;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property Saved: WordBool read Get_Saved write Set_Saved;
    property Type_: WdTemplateType read Get_Type_;
    property FullName: WideString read Get_FullName;
    property BuiltInDocumentProperties: IDispatch read Get_BuiltInDocumentProperties;
    property CustomDocumentProperties: IDispatch read Get_CustomDocumentProperties;
    property ListTemplates: ListTemplates read Get_ListTemplates;
    property LanguageIDFarEast: WdLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property VBProject: VBProject read Get_VBProject;
    property KerningByAlgorithm: WordBool read Get_KerningByAlgorithm write Set_KerningByAlgorithm;
    property JustificationMode: WdJustificationMode read Get_JustificationMode write Set_JustificationMode;
    property FarEastLineBreakLevel: WdFarEastLineBreakLevel read Get_FarEastLineBreakLevel write Set_FarEastLineBreakLevel;
    property NoLineBreakBefore: WideString read Get_NoLineBreakBefore write Set_NoLineBreakBefore;
    property NoLineBreakAfter: WideString read Get_NoLineBreakAfter write Set_NoLineBreakAfter;
    property NoProofing: Integer read Get_NoProofing write Set_NoProofing;
    property FarEastLineBreakLanguage: WdFarEastLineBreakLanguageID read Get_FarEastLineBreakLanguage write Set_FarEastLineBreakLanguage;
  end;

// *********************************************************************//
// DispIntf:  TemplateDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096A-0000-0000-C000-000000000046}
// *********************************************************************//
  TemplateDisp = dispinterface
    ['{0002096A-0000-0000-C000-000000000046}']
    property Name: WideString readonly dispid 0;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Path: WideString readonly dispid 1;
    property AutoTextEntries: AutoTextEntries readonly dispid 2;
    property LanguageID: WdLanguageID dispid 4;
    property Saved: WordBool dispid 5;
    property Type_: WdTemplateType readonly dispid 6;
    property FullName: WideString readonly dispid 7;
    property BuiltInDocumentProperties: IDispatch readonly dispid 8;
    property CustomDocumentProperties: IDispatch readonly dispid 9;
    property ListTemplates: ListTemplates readonly dispid 10;
    property LanguageIDFarEast: WdLanguageID dispid 11;
    property VBProject: VBProject readonly dispid 99;
    property KerningByAlgorithm: WordBool dispid 12;
    property JustificationMode: WdJustificationMode dispid 13;
    property FarEastLineBreakLevel: WdFarEastLineBreakLevel dispid 14;
    property NoLineBreakBefore: WideString dispid 15;
    property NoLineBreakAfter: WideString dispid 16;
    function OpenAsDocument: Document; dispid 100;
    procedure Save; dispid 101;
    property NoProofing: Integer dispid 17;
    property FarEastLineBreakLanguage: WdFarEastLineBreakLanguageID dispid 18;
  end;

// *********************************************************************//
// Interface: Templates
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A2-0000-0000-C000-000000000046}
// *********************************************************************//
  Templates = interface(IDispatch)
    ['{000209A2-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): Template; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  TemplatesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A2-0000-0000-C000-000000000046}
// *********************************************************************//
  TemplatesDisp = dispinterface
    ['{000209A2-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): Template; dispid 0;
  end;

// *********************************************************************//
// Interface: RoutingSlip
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020969-0000-0000-C000-000000000046}
// *********************************************************************//
  RoutingSlip = interface(IDispatch)
    ['{00020969-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const prop: WideString); safecall;
    function Get_Message: WideString; safecall;
    procedure Set_Message(const prop: WideString); safecall;
    function Get_Delivery: WdRoutingSlipDelivery; safecall;
    procedure Set_Delivery(prop: WdRoutingSlipDelivery); safecall;
    function Get_TrackStatus: WordBool; safecall;
    procedure Set_TrackStatus(prop: WordBool); safecall;
    function Get_Protect: WdProtectionType; safecall;
    procedure Set_Protect(prop: WdProtectionType); safecall;
    function Get_ReturnWhenDone: WordBool; safecall;
    procedure Set_ReturnWhenDone(prop: WordBool); safecall;
    function Get_Status: WdRoutingSlipStatus; safecall;
    function Get_Recipients(var Index: OleVariant): OleVariant; safecall;
    procedure Reset; safecall;
    procedure AddRecipient(const Recipient: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Subject: WideString read Get_Subject write Set_Subject;
    property Message: WideString read Get_Message write Set_Message;
    property Delivery: WdRoutingSlipDelivery read Get_Delivery write Set_Delivery;
    property TrackStatus: WordBool read Get_TrackStatus write Set_TrackStatus;
    property Protect: WdProtectionType read Get_Protect write Set_Protect;
    property ReturnWhenDone: WordBool read Get_ReturnWhenDone write Set_ReturnWhenDone;
    property Status: WdRoutingSlipStatus read Get_Status;
    property Recipients[var Index: OleVariant]: OleVariant read Get_Recipients;
  end;

// *********************************************************************//
// DispIntf:  RoutingSlipDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020969-0000-0000-C000-000000000046}
// *********************************************************************//
  RoutingSlipDisp = dispinterface
    ['{00020969-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Subject: WideString dispid 1;
    property Message: WideString dispid 2;
    property Delivery: WdRoutingSlipDelivery dispid 3;
    property TrackStatus: WordBool dispid 4;
    property Protect: WdProtectionType dispid 5;
    property ReturnWhenDone: WordBool dispid 6;
    property Status: WdRoutingSlipStatus readonly dispid 7;
    property Recipients[var Index: OleVariant]: OleVariant readonly dispid 9;
    procedure Reset; dispid 101;
    procedure AddRecipient(const Recipient: WideString); dispid 102;
  end;

// *********************************************************************//
// Interface: Bookmark
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020968-0000-0000-C000-000000000046}
// *********************************************************************//
  Bookmark = interface(IDispatch)
    ['{00020968-0000-0000-C000-000000000046}']
    function Get_Name: WideString; safecall;
    function Get_Range: Range; safecall;
    function Get_Empty: WordBool; safecall;
    function Get_Start: Integer; safecall;
    procedure Set_Start(prop: Integer); safecall;
    function Get_End_: Integer; safecall;
    procedure Set_End_(prop: Integer); safecall;
    function Get_Column: WordBool; safecall;
    function Get_StoryType: WdStoryType; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    function Copy(const Name: WideString): Bookmark; safecall;
    property Name: WideString read Get_Name;
    property Range: Range read Get_Range;
    property Empty: WordBool read Get_Empty;
    property Start: Integer read Get_Start write Set_Start;
    property End_: Integer read Get_End_ write Set_End_;
    property Column: WordBool read Get_Column;
    property StoryType: WdStoryType read Get_StoryType;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  BookmarkDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020968-0000-0000-C000-000000000046}
// *********************************************************************//
  BookmarkDisp = dispinterface
    ['{00020968-0000-0000-C000-000000000046}']
    property Name: WideString readonly dispid 0;
    property Range: Range readonly dispid 1;
    property Empty: WordBool readonly dispid 2;
    property Start: Integer dispid 3;
    property End_: Integer dispid 4;
    property Column: WordBool readonly dispid 5;
    property StoryType: WdStoryType readonly dispid 6;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure Select; dispid 65535;
    procedure Delete; dispid 11;
    function Copy(const Name: WideString): Bookmark; dispid 12;
  end;

// *********************************************************************//
// Interface: Bookmarks
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020967-0000-0000-C000-000000000046}
// *********************************************************************//
  Bookmarks = interface(IDispatch)
    ['{00020967-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_DefaultSorting: WdBookmarkSortBy; safecall;
    procedure Set_DefaultSorting(prop: WdBookmarkSortBy); safecall;
    function Get_ShowHidden: WordBool; safecall;
    procedure Set_ShowHidden(prop: WordBool); safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): Bookmark; safecall;
    function Add(const Name: WideString; var Range: OleVariant): Bookmark; safecall;
    function Exists(const Name: WideString): WordBool; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property DefaultSorting: WdBookmarkSortBy read Get_DefaultSorting write Set_DefaultSorting;
    property ShowHidden: WordBool read Get_ShowHidden write Set_ShowHidden;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  BookmarksDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020967-0000-0000-C000-000000000046}
// *********************************************************************//
  BookmarksDisp = dispinterface
    ['{00020967-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property DefaultSorting: WdBookmarkSortBy dispid 3;
    property ShowHidden: WordBool dispid 4;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): Bookmark; dispid 0;
    function Add(const Name: WideString; var Range: OleVariant): Bookmark; dispid 5;
    function Exists(const Name: WideString): WordBool; dispid 6;
  end;

// *********************************************************************//
// Interface: Variable
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020966-0000-0000-C000-000000000046}
// *********************************************************************//
  Variable = interface(IDispatch)
    ['{00020966-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Value: WideString; safecall;
    procedure Set_Value(const prop: WideString); safecall;
    function Get_Index: Integer; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Value: WideString read Get_Value write Set_Value;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  VariableDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020966-0000-0000-C000-000000000046}
// *********************************************************************//
  VariableDisp = dispinterface
    ['{00020966-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 1;
    property Value: WideString dispid 0;
    property Index: Integer readonly dispid 2;
    procedure Delete; dispid 11;
  end;

// *********************************************************************//
// Interface: Variables
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020965-0000-0000-C000-000000000046}
// *********************************************************************//
  Variables = interface(IDispatch)
    ['{00020965-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): Variable; safecall;
    function Add(const Name: WideString; var Value: OleVariant): Variable; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  VariablesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020965-0000-0000-C000-000000000046}
// *********************************************************************//
  VariablesDisp = dispinterface
    ['{00020965-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): Variable; dispid 0;
    function Add(const Name: WideString; var Value: OleVariant): Variable; dispid 7;
  end;

// *********************************************************************//
// Interface: RecentFile
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020964-0000-0000-C000-000000000046}
// *********************************************************************//
  RecentFile = interface(IDispatch)
    ['{00020964-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    function Get_ReadOnly: WordBool; safecall;
    procedure Set_ReadOnly(prop: WordBool); safecall;
    function Get_Path: WideString; safecall;
    function Open: Document; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
    property ReadOnly: WordBool read Get_ReadOnly write Set_ReadOnly;
    property Path: WideString read Get_Path;
  end;

// *********************************************************************//
// DispIntf:  RecentFileDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020964-0000-0000-C000-000000000046}
// *********************************************************************//
  RecentFileDisp = dispinterface
    ['{00020964-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Index: Integer readonly dispid 1;
    property ReadOnly: WordBool dispid 2;
    property Path: WideString readonly dispid 3;
    function Open: Document; dispid 4;
    procedure Delete; dispid 5;
  end;

// *********************************************************************//
// Interface: RecentFiles
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020963-0000-0000-C000-000000000046}
// *********************************************************************//
  RecentFiles = interface(IDispatch)
    ['{00020963-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Maximum: Integer; safecall;
    procedure Set_Maximum(prop: Integer); safecall;
    function Item(Index: Integer): RecentFile; safecall;
    function Add(var Document: OleVariant; var ReadOnly: OleVariant): RecentFile; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Maximum: Integer read Get_Maximum write Set_Maximum;
  end;

// *********************************************************************//
// DispIntf:  RecentFilesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020963-0000-0000-C000-000000000046}
// *********************************************************************//
  RecentFilesDisp = dispinterface
    ['{00020963-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Maximum: Integer dispid 2;
    function Item(Index: Integer): RecentFile; dispid 0;
    function Add(var Document: OleVariant; var ReadOnly: OleVariant): RecentFile; dispid 3;
  end;

// *********************************************************************//
// Interface: Window_
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020962-0000-0000-C000-000000000046}
// *********************************************************************//
  Window_ = interface(IDispatch)
    ['{00020962-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ActivePane: Pane; safecall;
    function Get_Document: Document; safecall;
    function Get_Panes: Panes; safecall;
    function Get_Selection: Selection; safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(prop: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(prop: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(prop: Integer); safecall;
    function Get_Split: WordBool; safecall;
    procedure Set_Split(prop: WordBool); safecall;
    function Get_SplitVertical: Integer; safecall;
    procedure Set_SplitVertical(prop: Integer); safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const prop: WideString); safecall;
    function Get_WindowState: WdWindowState; safecall;
    procedure Set_WindowState(prop: WdWindowState); safecall;
    function Get_DisplayRulers: WordBool; safecall;
    procedure Set_DisplayRulers(prop: WordBool); safecall;
    function Get_DisplayVerticalRuler: WordBool; safecall;
    procedure Set_DisplayVerticalRuler(prop: WordBool); safecall;
    function Get_View: View; safecall;
    function Get_Type_: WdWindowType; safecall;
    function Get_Next: Window_; safecall;
    function Get_Previous: Window_; safecall;
    function Get_WindowNumber: Integer; safecall;
    function Get_DisplayVerticalScrollBar: WordBool; safecall;
    procedure Set_DisplayVerticalScrollBar(prop: WordBool); safecall;
    function Get_DisplayHorizontalScrollBar: WordBool; safecall;
    procedure Set_DisplayHorizontalScrollBar(prop: WordBool); safecall;
    function Get_StyleAreaWidth: Single; safecall;
    procedure Set_StyleAreaWidth(prop: Single); safecall;
    function Get_DisplayScreenTips: WordBool; safecall;
    procedure Set_DisplayScreenTips(prop: WordBool); safecall;
    function Get_HorizontalPercentScrolled: Integer; safecall;
    procedure Set_HorizontalPercentScrolled(prop: Integer); safecall;
    function Get_VerticalPercentScrolled: Integer; safecall;
    procedure Set_VerticalPercentScrolled(prop: Integer); safecall;
    function Get_DocumentMap: WordBool; safecall;
    procedure Set_DocumentMap(prop: WordBool); safecall;
    function Get_Active: WordBool; safecall;
    function Get_DocumentMapPercentWidth: Integer; safecall;
    procedure Set_DocumentMapPercentWidth(prop: Integer); safecall;
    function Get_Index: Integer; safecall;
    function Get_IMEMode: WdIMEMode; safecall;
    procedure Set_IMEMode(prop: WdIMEMode); safecall;
    procedure Activate; safecall;
    procedure Close(var SaveChanges: OleVariant; var RouteDocument: OleVariant); safecall;
    procedure LargeScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); safecall;
    procedure SmallScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); safecall;
    function NewWindow: Window_; safecall;
    procedure PrintOutOld(var Background: OleVariant; var Append: OleVariant; 
                          var Range: OleVariant; var OutputFileName: OleVariant; 
                          var From: OleVariant; var To_: OleVariant; var Item: OleVariant; 
                          var Copies: OleVariant; var Pages: OleVariant; var PageType: OleVariant; 
                          var PrintToFile: OleVariant; var Collate: OleVariant; 
                          var ActivePrinterMacGX: OleVariant; var ManualDuplexPrint: OleVariant); safecall;
    procedure PageScroll(var Down: OleVariant; var Up: OleVariant); safecall;
    procedure SetFocus; safecall;
    function RangeFromPoint(x: Integer; y: Integer): IDispatch; safecall;
    procedure ScrollIntoView(const obj: IDispatch; var Start: OleVariant); safecall;
    procedure GetPoint(out ScreenPixelsLeft: Integer; out ScreenPixelsTop: Integer; 
                       out ScreenPixelsWidth: Integer; out ScreenPixelsHeight: Integer; 
                       const obj: IDispatch); safecall;
    procedure PrintOut(var Background: OleVariant; var Append: OleVariant; var Range: OleVariant; 
                       var OutputFileName: OleVariant; var From: OleVariant; var To_: OleVariant; 
                       var Item: OleVariant; var Copies: OleVariant; var Pages: OleVariant; 
                       var PageType: OleVariant; var PrintToFile: OleVariant; 
                       var Collate: OleVariant; var ActivePrinterMacGX: OleVariant; 
                       var ManualDuplexPrint: OleVariant; var PrintZoomColumn: OleVariant; 
                       var PrintZoomRow: OleVariant; var PrintZoomPaperWidth: OleVariant; 
                       var PrintZoomPaperHeight: OleVariant); safecall;
    function Get_UsableWidth: Integer; safecall;
    function Get_UsableHeight: Integer; safecall;
    function Get_EnvelopeVisible: WordBool; safecall;
    procedure Set_EnvelopeVisible(prop: WordBool); safecall;
    function Get_DisplayRightRuler: WordBool; safecall;
    procedure Set_DisplayRightRuler(prop: WordBool); safecall;
    function Get_DisplayLeftScrollBar: WordBool; safecall;
    procedure Set_DisplayLeftScrollBar(prop: WordBool); safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ActivePane: Pane read Get_ActivePane;
    property Document: Document read Get_Document;
    property Panes: Panes read Get_Panes;
    property Selection: Selection read Get_Selection;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property Split: WordBool read Get_Split write Set_Split;
    property SplitVertical: Integer read Get_SplitVertical write Set_SplitVertical;
    property Caption: WideString read Get_Caption write Set_Caption;
    property WindowState: WdWindowState read Get_WindowState write Set_WindowState;
    property DisplayRulers: WordBool read Get_DisplayRulers write Set_DisplayRulers;
    property DisplayVerticalRuler: WordBool read Get_DisplayVerticalRuler write Set_DisplayVerticalRuler;
    property View: View read Get_View;
    property Type_: WdWindowType read Get_Type_;
    property Next: Window_ read Get_Next;
    property Previous: Window_ read Get_Previous;
    property WindowNumber: Integer read Get_WindowNumber;
    property DisplayVerticalScrollBar: WordBool read Get_DisplayVerticalScrollBar write Set_DisplayVerticalScrollBar;
    property DisplayHorizontalScrollBar: WordBool read Get_DisplayHorizontalScrollBar write Set_DisplayHorizontalScrollBar;
    property StyleAreaWidth: Single read Get_StyleAreaWidth write Set_StyleAreaWidth;
    property DisplayScreenTips: WordBool read Get_DisplayScreenTips write Set_DisplayScreenTips;
    property HorizontalPercentScrolled: Integer read Get_HorizontalPercentScrolled write Set_HorizontalPercentScrolled;
    property VerticalPercentScrolled: Integer read Get_VerticalPercentScrolled write Set_VerticalPercentScrolled;
    property DocumentMap: WordBool read Get_DocumentMap write Set_DocumentMap;
    property Active: WordBool read Get_Active;
    property DocumentMapPercentWidth: Integer read Get_DocumentMapPercentWidth write Set_DocumentMapPercentWidth;
    property Index: Integer read Get_Index;
    property IMEMode: WdIMEMode read Get_IMEMode write Set_IMEMode;
    property UsableWidth: Integer read Get_UsableWidth;
    property UsableHeight: Integer read Get_UsableHeight;
    property EnvelopeVisible: WordBool read Get_EnvelopeVisible write Set_EnvelopeVisible;
    property DisplayRightRuler: WordBool read Get_DisplayRightRuler write Set_DisplayRightRuler;
    property DisplayLeftScrollBar: WordBool read Get_DisplayLeftScrollBar write Set_DisplayLeftScrollBar;
    property Visible: WordBool read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  Window_Disp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020962-0000-0000-C000-000000000046}
// *********************************************************************//
  Window_Disp = dispinterface
    ['{00020962-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ActivePane: Pane readonly dispid 1;
    property Document: Document readonly dispid 2;
    property Panes: Panes readonly dispid 3;
    property Selection: Selection readonly dispid 4;
    property Left: Integer dispid 5;
    property Top: Integer dispid 6;
    property Width: Integer dispid 7;
    property Height: Integer dispid 8;
    property Split: WordBool dispid 9;
    property SplitVertical: Integer dispid 10;
    property Caption: WideString dispid 0;
    property WindowState: WdWindowState dispid 11;
    property DisplayRulers: WordBool dispid 12;
    property DisplayVerticalRuler: WordBool dispid 13;
    property View: View readonly dispid 14;
    property Type_: WdWindowType readonly dispid 15;
    property Next: Window_ readonly dispid 16;
    property Previous: Window_ readonly dispid 17;
    property WindowNumber: Integer readonly dispid 18;
    property DisplayVerticalScrollBar: WordBool dispid 19;
    property DisplayHorizontalScrollBar: WordBool dispid 20;
    property StyleAreaWidth: Single dispid 21;
    property DisplayScreenTips: WordBool dispid 22;
    property HorizontalPercentScrolled: Integer dispid 23;
    property VerticalPercentScrolled: Integer dispid 24;
    property DocumentMap: WordBool dispid 25;
    property Active: WordBool readonly dispid 26;
    property DocumentMapPercentWidth: Integer dispid 27;
    property Index: Integer readonly dispid 28;
    property IMEMode: WdIMEMode dispid 30;
    procedure Activate; dispid 100;
    procedure Close(var SaveChanges: OleVariant; var RouteDocument: OleVariant); dispid 102;
    procedure LargeScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); dispid 103;
    procedure SmallScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); dispid 104;
    function NewWindow: Window_; dispid 105;
    procedure PrintOutOld(var Background: OleVariant; var Append: OleVariant; 
                          var Range: OleVariant; var OutputFileName: OleVariant; 
                          var From: OleVariant; var To_: OleVariant; var Item: OleVariant; 
                          var Copies: OleVariant; var Pages: OleVariant; var PageType: OleVariant; 
                          var PrintToFile: OleVariant; var Collate: OleVariant; 
                          var ActivePrinterMacGX: OleVariant; var ManualDuplexPrint: OleVariant); dispid 107;
    procedure PageScroll(var Down: OleVariant; var Up: OleVariant); dispid 108;
    procedure SetFocus; dispid 109;
    function RangeFromPoint(x: Integer; y: Integer): IDispatch; dispid 110;
    procedure ScrollIntoView(const obj: IDispatch; var Start: OleVariant); dispid 111;
    procedure GetPoint(out ScreenPixelsLeft: Integer; out ScreenPixelsTop: Integer; 
                       out ScreenPixelsWidth: Integer; out ScreenPixelsHeight: Integer; 
                       const obj: IDispatch); dispid 112;
    procedure PrintOut(var Background: OleVariant; var Append: OleVariant; var Range: OleVariant; 
                       var OutputFileName: OleVariant; var From: OleVariant; var To_: OleVariant; 
                       var Item: OleVariant; var Copies: OleVariant; var Pages: OleVariant; 
                       var PageType: OleVariant; var PrintToFile: OleVariant; 
                       var Collate: OleVariant; var ActivePrinterMacGX: OleVariant; 
                       var ManualDuplexPrint: OleVariant; var PrintZoomColumn: OleVariant; 
                       var PrintZoomRow: OleVariant; var PrintZoomPaperWidth: OleVariant; 
                       var PrintZoomPaperHeight: OleVariant); dispid 444;
    property UsableWidth: Integer readonly dispid 31;
    property UsableHeight: Integer readonly dispid 32;
    property EnvelopeVisible: WordBool dispid 33;
    property DisplayRightRuler: WordBool dispid 35;
    property DisplayLeftScrollBar: WordBool dispid 34;
    property Visible: WordBool dispid 36;
  end;

// *********************************************************************//
// Interface: Windows
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020961-0000-0000-C000-000000000046}
// *********************************************************************//
  Windows = interface(IDispatch)
    ['{00020961-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): Window_; safecall;
    function Add(var Window_: OleVariant): Window_; safecall;
    procedure Arrange(var ArrangeStyle: OleVariant); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  WindowsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020961-0000-0000-C000-000000000046}
// *********************************************************************//
  WindowsDisp = dispinterface
    ['{00020961-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): Window_; dispid 0;
    function Add(var Window_: OleVariant): Window_; dispid 10;
    procedure Arrange(var ArrangeStyle: OleVariant); dispid 11;
  end;

// *********************************************************************//
// Interface: Pane
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020960-0000-0000-C000-000000000046}
// *********************************************************************//
  Pane = interface(IDispatch)
    ['{00020960-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Document: Document; safecall;
    function Get_Selection: Selection; safecall;
    function Get_DisplayRulers: WordBool; safecall;
    procedure Set_DisplayRulers(prop: WordBool); safecall;
    function Get_DisplayVerticalRuler: WordBool; safecall;
    procedure Set_DisplayVerticalRuler(prop: WordBool); safecall;
    function Get_Zooms: Zooms; safecall;
    function Get_Index: Integer; safecall;
    function Get_View: View; safecall;
    function Get_Next: Pane; safecall;
    function Get_Previous: Pane; safecall;
    function Get_HorizontalPercentScrolled: Integer; safecall;
    procedure Set_HorizontalPercentScrolled(prop: Integer); safecall;
    function Get_VerticalPercentScrolled: Integer; safecall;
    procedure Set_VerticalPercentScrolled(prop: Integer); safecall;
    function Get_MinimumFontSize: Integer; safecall;
    procedure Set_MinimumFontSize(prop: Integer); safecall;
    function Get_BrowseToWindow: WordBool; safecall;
    procedure Set_BrowseToWindow(prop: WordBool); safecall;
    function Get_BrowseWidth: Integer; safecall;
    procedure Activate; safecall;
    procedure Close; safecall;
    procedure LargeScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); safecall;
    procedure SmallScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); safecall;
    procedure AutoScroll(Velocity: Integer); safecall;
    procedure PageScroll(var Down: OleVariant; var Up: OleVariant); safecall;
    procedure NewFrameset; safecall;
    procedure TOCInFrameset; safecall;
    function Get_Frameset: Frameset; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Document: Document read Get_Document;
    property Selection: Selection read Get_Selection;
    property DisplayRulers: WordBool read Get_DisplayRulers write Set_DisplayRulers;
    property DisplayVerticalRuler: WordBool read Get_DisplayVerticalRuler write Set_DisplayVerticalRuler;
    property Zooms: Zooms read Get_Zooms;
    property Index: Integer read Get_Index;
    property View: View read Get_View;
    property Next: Pane read Get_Next;
    property Previous: Pane read Get_Previous;
    property HorizontalPercentScrolled: Integer read Get_HorizontalPercentScrolled write Set_HorizontalPercentScrolled;
    property VerticalPercentScrolled: Integer read Get_VerticalPercentScrolled write Set_VerticalPercentScrolled;
    property MinimumFontSize: Integer read Get_MinimumFontSize write Set_MinimumFontSize;
    property BrowseToWindow: WordBool read Get_BrowseToWindow write Set_BrowseToWindow;
    property BrowseWidth: Integer read Get_BrowseWidth;
    property Frameset: Frameset read Get_Frameset;
  end;

// *********************************************************************//
// DispIntf:  PaneDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020960-0000-0000-C000-000000000046}
// *********************************************************************//
  PaneDisp = dispinterface
    ['{00020960-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Document: Document readonly dispid 1;
    property Selection: Selection readonly dispid 3;
    property DisplayRulers: WordBool dispid 4;
    property DisplayVerticalRuler: WordBool dispid 5;
    property Zooms: Zooms readonly dispid 7;
    property Index: Integer readonly dispid 9;
    property View: View readonly dispid 10;
    property Next: Pane readonly dispid 11;
    property Previous: Pane readonly dispid 12;
    property HorizontalPercentScrolled: Integer dispid 13;
    property VerticalPercentScrolled: Integer dispid 14;
    property MinimumFontSize: Integer dispid 15;
    property BrowseToWindow: WordBool dispid 16;
    property BrowseWidth: Integer readonly dispid 17;
    procedure Activate; dispid 100;
    procedure Close; dispid 101;
    procedure LargeScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); dispid 102;
    procedure SmallScroll(var Down: OleVariant; var Up: OleVariant; var ToRight: OleVariant; 
                          var ToLeft: OleVariant); dispid 103;
    procedure AutoScroll(Velocity: Integer); dispid 104;
    procedure PageScroll(var Down: OleVariant; var Up: OleVariant); dispid 105;
    procedure NewFrameset; dispid 106;
    procedure TOCInFrameset; dispid 107;
    property Frameset: Frameset readonly dispid 18;
  end;

// *********************************************************************//
// Interface: Panes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095F-0000-0000-C000-000000000046}
// *********************************************************************//
  Panes = interface(IDispatch)
    ['{0002095F-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): Pane; safecall;
    function Add(var SplitVertical: OleVariant): Pane; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  PanesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095F-0000-0000-C000-000000000046}
// *********************************************************************//
  PanesDisp = dispinterface
    ['{0002095F-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): Pane; dispid 0;
    function Add(var SplitVertical: OleVariant): Pane; dispid 3;
  end;

// *********************************************************************//
// Interface: Range
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095E-0000-0000-C000-000000000046}
// *********************************************************************//
  Range = interface(IDispatch)
    ['{0002095E-0000-0000-C000-000000000046}']
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_FormattedText: Range; safecall;
    procedure Set_FormattedText(const prop: Range); safecall;
    function Get_Start: Integer; safecall;
    procedure Set_Start(prop: Integer); safecall;
    function Get_End_: Integer; safecall;
    procedure Set_End_(prop: Integer); safecall;
    function Get_Font: Font; safecall;
    procedure Set_Font(const prop: Font); safecall;
    function Get_Duplicate: Range; safecall;
    function Get_StoryType: WdStoryType; safecall;
    function Get_Tables: Tables; safecall;
    function Get_Words: Words; safecall;
    function Get_Sentences: Sentences; safecall;
    function Get_Characters: Characters; safecall;
    function Get_Footnotes: Footnotes; safecall;
    function Get_Endnotes: Endnotes; safecall;
    function Get_Comments: Comments; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Sections: Sections; safecall;
    function Get_Paragraphs: Paragraphs; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_TextRetrievalMode: TextRetrievalMode; safecall;
    procedure Set_TextRetrievalMode(const prop: TextRetrievalMode); safecall;
    function Get_Fields: Fields; safecall;
    function Get_FormFields: FormFields; safecall;
    function Get_Frames: Frames; safecall;
    function Get_ParagraphFormat: ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: ParagraphFormat); safecall;
    function Get_ListFormat: ListFormat; safecall;
    function Get_Bookmarks: Bookmarks; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Bold: Integer; safecall;
    procedure Set_Bold(prop: Integer); safecall;
    function Get_Italic: Integer; safecall;
    procedure Set_Italic(prop: Integer); safecall;
    function Get_Underline: WdUnderline; safecall;
    procedure Set_Underline(prop: WdUnderline); safecall;
    function Get_EmphasisMark: WdEmphasisMark; safecall;
    procedure Set_EmphasisMark(prop: WdEmphasisMark); safecall;
    function Get_DisableCharacterSpaceGrid: WordBool; safecall;
    procedure Set_DisableCharacterSpaceGrid(prop: WordBool); safecall;
    function Get_Revisions: Revisions; safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_StoryLength: Integer; safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_SynonymInfo: SynonymInfo; safecall;
    function Get_Hyperlinks: Hyperlinks; safecall;
    function Get_ListParagraphs: ListParagraphs; safecall;
    function Get_Subdocuments: Subdocuments; safecall;
    function Get_GrammarChecked: WordBool; safecall;
    procedure Set_GrammarChecked(prop: WordBool); safecall;
    function Get_SpellingChecked: WordBool; safecall;
    procedure Set_SpellingChecked(prop: WordBool); safecall;
    function Get_HighlightColorIndex: WdColorIndex; safecall;
    procedure Set_HighlightColorIndex(prop: WdColorIndex); safecall;
    function Get_Columns: Columns; safecall;
    function Get_Rows: Rows; safecall;
    function Get_CanEdit: Integer; safecall;
    function Get_CanPaste: Integer; safecall;
    function Get_IsEndOfRowMark: WordBool; safecall;
    function Get_BookmarkID: Integer; safecall;
    function Get_PreviousBookmarkID: Integer; safecall;
    function Get_Find: Find; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_ShapeRange: ShapeRange; safecall;
    function Get_Case_: WdCharacterCase; safecall;
    procedure Set_Case_(prop: WdCharacterCase); safecall;
    function Get_Information(Type_: WdInformation): OleVariant; safecall;
    function Get_ReadabilityStatistics: ReadabilityStatistics; safecall;
    function Get_GrammaticalErrors: ProofreadingErrors; safecall;
    function Get_SpellingErrors: ProofreadingErrors; safecall;
    function Get_Orientation: WdTextOrientation; safecall;
    procedure Set_Orientation(prop: WdTextOrientation); safecall;
    function Get_InlineShapes: InlineShapes; safecall;
    function Get_NextStoryRange: Range; safecall;
    function Get_LanguageIDFarEast: WdLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WdLanguageID); safecall;
    function Get_LanguageIDOther: WdLanguageID; safecall;
    procedure Set_LanguageIDOther(prop: WdLanguageID); safecall;
    procedure Select; safecall;
    procedure SetRange(Start: Integer; End_: Integer); safecall;
    procedure Collapse(var Direction: OleVariant); safecall;
    procedure InsertBefore(const Text: WideString); safecall;
    procedure InsertAfter(const Text: WideString); safecall;
    function Next(var Unit_: OleVariant; var Count: OleVariant): Range; safecall;
    function Previous(var Unit_: OleVariant; var Count: OleVariant): Range; safecall;
    function StartOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; safecall;
    function EndOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; safecall;
    function Move(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveStart(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveEnd(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveWhile(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveStartWhile(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveEndWhile(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveUntil(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveStartUntil(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveEndUntil(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    procedure Cut; safecall;
    procedure Copy; safecall;
    procedure Paste; safecall;
    procedure InsertBreak(var Type_: OleVariant); safecall;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); safecall;
    function InStory(const Range: Range): WordBool; safecall;
    function InRange(const Range: Range): WordBool; safecall;
    function Delete(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    procedure WholeStory; safecall;
    function Expand(var Unit_: OleVariant): Integer; safecall;
    procedure InsertParagraph; safecall;
    procedure InsertParagraphAfter; safecall;
    function ConvertToTableOld(var Separator: OleVariant; var NumRows: OleVariant; 
                               var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                               var Format: OleVariant; var ApplyBorders: OleVariant; 
                               var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                               var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                               var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                               var ApplyLastColumn: OleVariant; var AutoFit: OleVariant): Table; safecall;
    procedure InsertDateTimeOld(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                                var InsertAsFullWidth: OleVariant); safecall;
    procedure InsertSymbol(CharacterNumber: Integer; var Font: OleVariant; var Unicode: OleVariant; 
                           var Bias: OleVariant); safecall;
    procedure InsertCrossReference(var ReferenceType: OleVariant; ReferenceKind: WdReferenceKind; 
                                   var ReferenceItem: OleVariant; 
                                   var InsertAsHyperlink: OleVariant; 
                                   var IncludePosition: OleVariant); safecall;
    procedure InsertCaption(var Label_: OleVariant; var Title: OleVariant; 
                            var TitleAutoText: OleVariant; var Position: OleVariant); safecall;
    procedure CopyAsPicture; safecall;
    procedure SortOld(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                      var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                      var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                      var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                      var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                      var SortColumn: OleVariant; var Separator: OleVariant; 
                      var CaseSensitive: OleVariant; var LanguageID: OleVariant); safecall;
    procedure SortAscending; safecall;
    procedure SortDescending; safecall;
    function IsEqual(const Range: Range): WordBool; safecall;
    function Calculate: Single; safecall;
    function GoTo_(var What: OleVariant; var Which: OleVariant; var Count: OleVariant; 
                   var Name: OleVariant): Range; safecall;
    function GoToNext(What: WdGoToItem): Range; safecall;
    function GoToPrevious(What: WdGoToItem): Range; safecall;
    procedure PasteSpecial(var IconIndex: OleVariant; var Link: OleVariant; 
                           var Placement: OleVariant; var DisplayAsIcon: OleVariant; 
                           var DataType: OleVariant; var IconFileName: OleVariant; 
                           var IconLabel: OleVariant); safecall;
    procedure LookupNameProperties; safecall;
    function ComputeStatistics(Statistic: WdStatistic): Integer; safecall;
    procedure Relocate(Direction: Integer); safecall;
    procedure CheckSynonyms; safecall;
    procedure SubscribeTo(const Edition: WideString; var Format: OleVariant); safecall;
    procedure CreatePublisher(var Edition: OleVariant; var ContainsPICT: OleVariant; 
                              var ContainsRTF: OleVariant; var ContainsText: OleVariant); safecall;
    procedure InsertAutoText; safecall;
    procedure InsertDatabase(var Format: OleVariant; var Style: OleVariant; 
                             var LinkToSource: OleVariant; var Connection: OleVariant; 
                             var SQLStatement: OleVariant; var SQLStatement1: OleVariant; 
                             var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                             var WritePasswordDocument: OleVariant; 
                             var WritePasswordTemplate: OleVariant; var DataSource: OleVariant; 
                             var From: OleVariant; var To_: OleVariant; 
                             var IncludeFields: OleVariant); safecall;
    procedure AutoFormat; safecall;
    procedure CheckGrammar; safecall;
    procedure CheckSpelling(var CustomDictionary: OleVariant; var IgnoreUppercase: OleVariant; 
                            var AlwaysSuggest: OleVariant; var CustomDictionary2: OleVariant; 
                            var CustomDictionary3: OleVariant; var CustomDictionary4: OleVariant; 
                            var CustomDictionary5: OleVariant; var CustomDictionary6: OleVariant; 
                            var CustomDictionary7: OleVariant; var CustomDictionary8: OleVariant; 
                            var CustomDictionary9: OleVariant; var CustomDictionary10: OleVariant); safecall;
    function GetSpellingSuggestions(var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; safecall;
    procedure InsertParagraphBefore; safecall;
    procedure NextSubdocument; safecall;
    procedure PreviousSubdocument; safecall;
    procedure ConvertHangulAndHanja(var ConversionsMode: OleVariant; 
                                    var FastConversion: OleVariant; 
                                    var CheckHangulEnding: OleVariant; 
                                    var EnableRecentOrdering: OleVariant; 
                                    var CustomDictionary: OleVariant); safecall;
    procedure PasteAsNestedTable; safecall;
    procedure ModifyEnclosure(var Style: OleVariant; var Symbol: OleVariant; 
                              var EnclosedText: OleVariant); safecall;
    procedure PhoneticGuide(const Text: WideString; Alignment: WdPhoneticGuideAlignmentType; 
                            Raise_: Integer; FontSize: Integer; const FontName: WideString); safecall;
    procedure InsertDateTime(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                             var InsertAsFullWidth: OleVariant; var DateLanguage: OleVariant; 
                             var CalendarType: OleVariant); safecall;
    procedure Sort(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                   var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                   var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                   var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                   var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                   var SortColumn: OleVariant; var Separator: OleVariant; 
                   var CaseSensitive: OleVariant; var BidiSort: OleVariant; 
                   var IgnoreThe: OleVariant; var IgnoreKashida: OleVariant; 
                   var IgnoreDiacritics: OleVariant; var IgnoreHe: OleVariant; 
                   var LanguageID: OleVariant); safecall;
    procedure DetectLanguage; safecall;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; safecall;
    procedure TCSCConverter(WdTCSCConverterDirection: WdTCSCConverterDirection; 
                            CommonTerms: WordBool; UseVariants: WordBool); safecall;
    function Get_LanguageDetected: WordBool; safecall;
    procedure Set_LanguageDetected(prop: WordBool); safecall;
    function Get_FitTextWidth: Single; safecall;
    procedure Set_FitTextWidth(prop: Single); safecall;
    function Get_HorizontalInVertical: WdHorizontalInVerticalType; safecall;
    procedure Set_HorizontalInVertical(prop: WdHorizontalInVerticalType); safecall;
    function Get_TwoLinesInOne: WdTwoLinesInOneType; safecall;
    procedure Set_TwoLinesInOne(prop: WdTwoLinesInOneType); safecall;
    function Get_CombineCharacters: WordBool; safecall;
    procedure Set_CombineCharacters(prop: WordBool); safecall;
    function Get_NoProofing: Integer; safecall;
    procedure Set_NoProofing(prop: Integer); safecall;
    function Get_TopLevelTables: Tables; safecall;
    function Get_Scripts: Scripts; safecall;
    function Get_CharacterWidth: WdCharacterWidth; safecall;
    procedure Set_CharacterWidth(prop: WdCharacterWidth); safecall;
    function Get_Kana: WdKana; safecall;
    procedure Set_Kana(prop: WdKana); safecall;
    function Get_BoldBi: Integer; safecall;
    procedure Set_BoldBi(prop: Integer); safecall;
    function Get_ItalicBi: Integer; safecall;
    procedure Set_ItalicBi(prop: Integer); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    property Text: WideString read Get_Text write Set_Text;
    property FormattedText: Range read Get_FormattedText write Set_FormattedText;
    property Start: Integer read Get_Start write Set_Start;
    property End_: Integer read Get_End_ write Set_End_;
    property Font: Font read Get_Font write Set_Font;
    property Duplicate: Range read Get_Duplicate;
    property StoryType: WdStoryType read Get_StoryType;
    property Tables: Tables read Get_Tables;
    property Words: Words read Get_Words;
    property Sentences: Sentences read Get_Sentences;
    property Characters: Characters read Get_Characters;
    property Footnotes: Footnotes read Get_Footnotes;
    property Endnotes: Endnotes read Get_Endnotes;
    property Comments: Comments read Get_Comments;
    property Cells: Cells read Get_Cells;
    property Sections: Sections read Get_Sections;
    property Paragraphs: Paragraphs read Get_Paragraphs;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property TextRetrievalMode: TextRetrievalMode read Get_TextRetrievalMode write Set_TextRetrievalMode;
    property Fields: Fields read Get_Fields;
    property FormFields: FormFields read Get_FormFields;
    property Frames: Frames read Get_Frames;
    property ParagraphFormat: ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property ListFormat: ListFormat read Get_ListFormat;
    property Bookmarks: Bookmarks read Get_Bookmarks;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Bold: Integer read Get_Bold write Set_Bold;
    property Italic: Integer read Get_Italic write Set_Italic;
    property Underline: WdUnderline read Get_Underline write Set_Underline;
    property EmphasisMark: WdEmphasisMark read Get_EmphasisMark write Set_EmphasisMark;
    property DisableCharacterSpaceGrid: WordBool read Get_DisableCharacterSpaceGrid write Set_DisableCharacterSpaceGrid;
    property Revisions: Revisions read Get_Revisions;
    property StoryLength: Integer read Get_StoryLength;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property SynonymInfo: SynonymInfo read Get_SynonymInfo;
    property Hyperlinks: Hyperlinks read Get_Hyperlinks;
    property ListParagraphs: ListParagraphs read Get_ListParagraphs;
    property Subdocuments: Subdocuments read Get_Subdocuments;
    property GrammarChecked: WordBool read Get_GrammarChecked write Set_GrammarChecked;
    property SpellingChecked: WordBool read Get_SpellingChecked write Set_SpellingChecked;
    property HighlightColorIndex: WdColorIndex read Get_HighlightColorIndex write Set_HighlightColorIndex;
    property Columns: Columns read Get_Columns;
    property Rows: Rows read Get_Rows;
    property CanEdit: Integer read Get_CanEdit;
    property CanPaste: Integer read Get_CanPaste;
    property IsEndOfRowMark: WordBool read Get_IsEndOfRowMark;
    property BookmarkID: Integer read Get_BookmarkID;
    property PreviousBookmarkID: Integer read Get_PreviousBookmarkID;
    property Find: Find read Get_Find;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property ShapeRange: ShapeRange read Get_ShapeRange;
    property Case_: WdCharacterCase read Get_Case_ write Set_Case_;
    property Information[Type_: WdInformation]: OleVariant read Get_Information;
    property ReadabilityStatistics: ReadabilityStatistics read Get_ReadabilityStatistics;
    property GrammaticalErrors: ProofreadingErrors read Get_GrammaticalErrors;
    property SpellingErrors: ProofreadingErrors read Get_SpellingErrors;
    property Orientation: WdTextOrientation read Get_Orientation write Set_Orientation;
    property InlineShapes: InlineShapes read Get_InlineShapes;
    property NextStoryRange: Range read Get_NextStoryRange;
    property LanguageIDFarEast: WdLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property LanguageIDOther: WdLanguageID read Get_LanguageIDOther write Set_LanguageIDOther;
    property LanguageDetected: WordBool read Get_LanguageDetected write Set_LanguageDetected;
    property FitTextWidth: Single read Get_FitTextWidth write Set_FitTextWidth;
    property HorizontalInVertical: WdHorizontalInVerticalType read Get_HorizontalInVertical write Set_HorizontalInVertical;
    property TwoLinesInOne: WdTwoLinesInOneType read Get_TwoLinesInOne write Set_TwoLinesInOne;
    property CombineCharacters: WordBool read Get_CombineCharacters write Set_CombineCharacters;
    property NoProofing: Integer read Get_NoProofing write Set_NoProofing;
    property TopLevelTables: Tables read Get_TopLevelTables;
    property Scripts: Scripts read Get_Scripts;
    property CharacterWidth: WdCharacterWidth read Get_CharacterWidth write Set_CharacterWidth;
    property Kana: WdKana read Get_Kana write Set_Kana;
    property BoldBi: Integer read Get_BoldBi write Set_BoldBi;
    property ItalicBi: Integer read Get_ItalicBi write Set_ItalicBi;
    property ID: WideString read Get_ID write Set_ID;
  end;

// *********************************************************************//
// DispIntf:  RangeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095E-0000-0000-C000-000000000046}
// *********************************************************************//
  RangeDisp = dispinterface
    ['{0002095E-0000-0000-C000-000000000046}']
    property Text: WideString dispid 0;
    property FormattedText: Range dispid 2;
    property Start: Integer dispid 3;
    property End_: Integer dispid 4;
    property Font: Font dispid 5;
    property Duplicate: Range readonly dispid 6;
    property StoryType: WdStoryType readonly dispid 7;
    property Tables: Tables readonly dispid 50;
    property Words: Words readonly dispid 51;
    property Sentences: Sentences readonly dispid 52;
    property Characters: Characters readonly dispid 53;
    property Footnotes: Footnotes readonly dispid 54;
    property Endnotes: Endnotes readonly dispid 55;
    property Comments: Comments readonly dispid 56;
    property Cells: Cells readonly dispid 57;
    property Sections: Sections readonly dispid 58;
    property Paragraphs: Paragraphs readonly dispid 59;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 61;
    property TextRetrievalMode: TextRetrievalMode dispid 62;
    property Fields: Fields readonly dispid 64;
    property FormFields: FormFields readonly dispid 65;
    property Frames: Frames readonly dispid 66;
    property ParagraphFormat: ParagraphFormat dispid 1102;
    property ListFormat: ListFormat readonly dispid 68;
    property Bookmarks: Bookmarks readonly dispid 75;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Bold: Integer dispid 130;
    property Italic: Integer dispid 131;
    property Underline: WdUnderline dispid 139;
    property EmphasisMark: WdEmphasisMark dispid 140;
    property DisableCharacterSpaceGrid: WordBool dispid 141;
    property Revisions: Revisions readonly dispid 150;
    function Style: OleVariant; dispid 151;
    property StoryLength: Integer readonly dispid 152;
    property LanguageID: WdLanguageID dispid 153;
    property SynonymInfo: SynonymInfo readonly dispid 155;
    property Hyperlinks: Hyperlinks readonly dispid 156;
    property ListParagraphs: ListParagraphs readonly dispid 157;
    property Subdocuments: Subdocuments readonly dispid 159;
    property GrammarChecked: WordBool dispid 260;
    property SpellingChecked: WordBool dispid 261;
    property HighlightColorIndex: WdColorIndex dispid 301;
    property Columns: Columns readonly dispid 302;
    property Rows: Rows readonly dispid 303;
    property CanEdit: Integer readonly dispid 304;
    property CanPaste: Integer readonly dispid 305;
    property IsEndOfRowMark: WordBool readonly dispid 307;
    property BookmarkID: Integer readonly dispid 308;
    property PreviousBookmarkID: Integer readonly dispid 309;
    property Find: Find readonly dispid 262;
    property PageSetup: PageSetup dispid 1101;
    property ShapeRange: ShapeRange readonly dispid 311;
    property Case_: WdCharacterCase dispid 312;
    property Information[Type_: WdInformation]: OleVariant readonly dispid 313;
    property ReadabilityStatistics: ReadabilityStatistics readonly dispid 314;
    property GrammaticalErrors: ProofreadingErrors readonly dispid 315;
    property SpellingErrors: ProofreadingErrors readonly dispid 316;
    property Orientation: WdTextOrientation dispid 317;
    property InlineShapes: InlineShapes readonly dispid 319;
    property NextStoryRange: Range readonly dispid 320;
    property LanguageIDFarEast: WdLanguageID dispid 321;
    property LanguageIDOther: WdLanguageID dispid 322;
    procedure Select; dispid 65535;
    procedure SetRange(Start: Integer; End_: Integer); dispid 100;
    procedure Collapse(var Direction: OleVariant); dispid 101;
    procedure InsertBefore(const Text: WideString); dispid 102;
    procedure InsertAfter(const Text: WideString); dispid 104;
    function Next(var Unit_: OleVariant; var Count: OleVariant): Range; dispid 105;
    function Previous(var Unit_: OleVariant; var Count: OleVariant): Range; dispid 106;
    function StartOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; dispid 107;
    function EndOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; dispid 108;
    function Move(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 109;
    function MoveStart(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 110;
    function MoveEnd(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 111;
    function MoveWhile(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 112;
    function MoveStartWhile(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 113;
    function MoveEndWhile(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 114;
    function MoveUntil(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 115;
    function MoveStartUntil(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 116;
    function MoveEndUntil(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 117;
    procedure Cut; dispid 119;
    procedure Copy; dispid 120;
    procedure Paste; dispid 121;
    procedure InsertBreak(var Type_: OleVariant); dispid 122;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); dispid 123;
    function InStory(const Range: Range): WordBool; dispid 125;
    function InRange(const Range: Range): WordBool; dispid 126;
    function Delete(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 127;
    procedure WholeStory; dispid 128;
    function Expand(var Unit_: OleVariant): Integer; dispid 129;
    procedure InsertParagraph; dispid 160;
    procedure InsertParagraphAfter; dispid 161;
    function ConvertToTableOld(var Separator: OleVariant; var NumRows: OleVariant; 
                               var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                               var Format: OleVariant; var ApplyBorders: OleVariant; 
                               var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                               var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                               var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                               var ApplyLastColumn: OleVariant; var AutoFit: OleVariant): Table; dispid 162;
    procedure InsertDateTimeOld(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                                var InsertAsFullWidth: OleVariant); dispid 163;
    procedure InsertSymbol(CharacterNumber: Integer; var Font: OleVariant; var Unicode: OleVariant; 
                           var Bias: OleVariant); dispid 164;
    procedure InsertCrossReference(var ReferenceType: OleVariant; ReferenceKind: WdReferenceKind; 
                                   var ReferenceItem: OleVariant; 
                                   var InsertAsHyperlink: OleVariant; 
                                   var IncludePosition: OleVariant); dispid 165;
    procedure InsertCaption(var Label_: OleVariant; var Title: OleVariant; 
                            var TitleAutoText: OleVariant; var Position: OleVariant); dispid 166;
    procedure CopyAsPicture; dispid 167;
    procedure SortOld(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                      var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                      var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                      var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                      var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                      var SortColumn: OleVariant; var Separator: OleVariant; 
                      var CaseSensitive: OleVariant; var LanguageID: OleVariant); dispid 168;
    procedure SortAscending; dispid 169;
    procedure SortDescending; dispid 170;
    function IsEqual(const Range: Range): WordBool; dispid 171;
    function Calculate: Single; dispid 172;
    function GoTo_(var What: OleVariant; var Which: OleVariant; var Count: OleVariant; 
                   var Name: OleVariant): Range; dispid 173;
    function GoToNext(What: WdGoToItem): Range; dispid 174;
    function GoToPrevious(What: WdGoToItem): Range; dispid 175;
    procedure PasteSpecial(var IconIndex: OleVariant; var Link: OleVariant; 
                           var Placement: OleVariant; var DisplayAsIcon: OleVariant; 
                           var DataType: OleVariant; var IconFileName: OleVariant; 
                           var IconLabel: OleVariant); dispid 176;
    procedure LookupNameProperties; dispid 177;
    function ComputeStatistics(Statistic: WdStatistic): Integer; dispid 178;
    procedure Relocate(Direction: Integer); dispid 179;
    procedure CheckSynonyms; dispid 180;
    procedure SubscribeTo(const Edition: WideString; var Format: OleVariant); dispid 181;
    procedure CreatePublisher(var Edition: OleVariant; var ContainsPICT: OleVariant; 
                              var ContainsRTF: OleVariant; var ContainsText: OleVariant); dispid 182;
    procedure InsertAutoText; dispid 183;
    procedure InsertDatabase(var Format: OleVariant; var Style: OleVariant; 
                             var LinkToSource: OleVariant; var Connection: OleVariant; 
                             var SQLStatement: OleVariant; var SQLStatement1: OleVariant; 
                             var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                             var WritePasswordDocument: OleVariant; 
                             var WritePasswordTemplate: OleVariant; var DataSource: OleVariant; 
                             var From: OleVariant; var To_: OleVariant; 
                             var IncludeFields: OleVariant); dispid 194;
    procedure AutoFormat; dispid 195;
    procedure CheckGrammar; dispid 204;
    procedure CheckSpelling(var CustomDictionary: OleVariant; var IgnoreUppercase: OleVariant; 
                            var AlwaysSuggest: OleVariant; var CustomDictionary2: OleVariant; 
                            var CustomDictionary3: OleVariant; var CustomDictionary4: OleVariant; 
                            var CustomDictionary5: OleVariant; var CustomDictionary6: OleVariant; 
                            var CustomDictionary7: OleVariant; var CustomDictionary8: OleVariant; 
                            var CustomDictionary9: OleVariant; var CustomDictionary10: OleVariant); dispid 205;
    function GetSpellingSuggestions(var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; dispid 209;
    procedure InsertParagraphBefore; dispid 212;
    procedure NextSubdocument; dispid 219;
    procedure PreviousSubdocument; dispid 220;
    procedure ConvertHangulAndHanja(var ConversionsMode: OleVariant; 
                                    var FastConversion: OleVariant; 
                                    var CheckHangulEnding: OleVariant; 
                                    var EnableRecentOrdering: OleVariant; 
                                    var CustomDictionary: OleVariant); dispid 221;
    procedure PasteAsNestedTable; dispid 222;
    procedure ModifyEnclosure(var Style: OleVariant; var Symbol: OleVariant; 
                              var EnclosedText: OleVariant); dispid 223;
    procedure PhoneticGuide(const Text: WideString; Alignment: WdPhoneticGuideAlignmentType; 
                            Raise_: Integer; FontSize: Integer; const FontName: WideString); dispid 224;
    procedure InsertDateTime(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                             var InsertAsFullWidth: OleVariant; var DateLanguage: OleVariant; 
                             var CalendarType: OleVariant); dispid 444;
    procedure Sort(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                   var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                   var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                   var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                   var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                   var SortColumn: OleVariant; var Separator: OleVariant; 
                   var CaseSensitive: OleVariant; var BidiSort: OleVariant; 
                   var IgnoreThe: OleVariant; var IgnoreKashida: OleVariant; 
                   var IgnoreDiacritics: OleVariant; var IgnoreHe: OleVariant; 
                   var LanguageID: OleVariant); dispid 484;
    procedure DetectLanguage; dispid 203;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; dispid 498;
    procedure TCSCConverter(WdTCSCConverterDirection: WdTCSCConverterDirection; 
                            CommonTerms: WordBool; UseVariants: WordBool); dispid 499;
    property LanguageDetected: WordBool dispid 263;
    property FitTextWidth: Single dispid 264;
    property HorizontalInVertical: WdHorizontalInVerticalType dispid 265;
    property TwoLinesInOne: WdTwoLinesInOneType dispid 266;
    property CombineCharacters: WordBool dispid 267;
    property NoProofing: Integer dispid 323;
    property TopLevelTables: Tables readonly dispid 324;
    property Scripts: Scripts readonly dispid 325;
    property CharacterWidth: WdCharacterWidth dispid 326;
    property Kana: WdKana dispid 327;
    property BoldBi: Integer dispid 400;
    property ItalicBi: Integer dispid 401;
    property ID: WideString dispid 405;
  end;

// *********************************************************************//
// Interface: ListFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C0-0000-0000-C000-000000000046}
// *********************************************************************//
  ListFormat = interface(IDispatch)
    ['{000209C0-0000-0000-C000-000000000046}']
    function Get_ListLevelNumber: Integer; safecall;
    procedure Set_ListLevelNumber(prop: Integer); safecall;
    function Get_List: List; safecall;
    function Get_ListTemplate: ListTemplate; safecall;
    function Get_ListValue: Integer; safecall;
    function Get_SingleList: WordBool; safecall;
    function Get_SingleListTemplate: WordBool; safecall;
    function Get_ListType: WdListType; safecall;
    function Get_ListString: WideString; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WdContinue; safecall;
    procedure RemoveNumbers(var NumberType: OleVariant); safecall;
    procedure ConvertNumbersToText(var NumberType: OleVariant); safecall;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; safecall;
    procedure ApplyBulletDefaultOld; safecall;
    procedure ApplyNumberDefaultOld; safecall;
    procedure ApplyOutlineNumberDefaultOld; safecall;
    procedure ApplyListTemplateOld(const ListTemplate: ListTemplate; 
                                   var ContinuePreviousList: OleVariant; var ApplyTo: OleVariant); safecall;
    procedure ListOutdent; safecall;
    procedure ListIndent; safecall;
    procedure ApplyBulletDefault(var DefaultListBehavior: OleVariant); safecall;
    procedure ApplyNumberDefault(var DefaultListBehavior: OleVariant); safecall;
    procedure ApplyOutlineNumberDefault(var DefaultListBehavior: OleVariant); safecall;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                var ContinuePreviousList: OleVariant; var ApplyTo: OleVariant; 
                                var DefaultListBehavior: OleVariant); safecall;
    property ListLevelNumber: Integer read Get_ListLevelNumber write Set_ListLevelNumber;
    property List: List read Get_List;
    property ListTemplate: ListTemplate read Get_ListTemplate;
    property ListValue: Integer read Get_ListValue;
    property SingleList: WordBool read Get_SingleList;
    property SingleListTemplate: WordBool read Get_SingleListTemplate;
    property ListType: WdListType read Get_ListType;
    property ListString: WideString read Get_ListString;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C0-0000-0000-C000-000000000046}
// *********************************************************************//
  ListFormatDisp = dispinterface
    ['{000209C0-0000-0000-C000-000000000046}']
    property ListLevelNumber: Integer dispid 68;
    property List: List readonly dispid 69;
    property ListTemplate: ListTemplate readonly dispid 70;
    property ListValue: Integer readonly dispid 71;
    property SingleList: WordBool readonly dispid 72;
    property SingleListTemplate: WordBool readonly dispid 73;
    property ListType: WdListType readonly dispid 74;
    property ListString: WideString readonly dispid 75;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WdContinue; dispid 184;
    procedure RemoveNumbers(var NumberType: OleVariant); dispid 185;
    procedure ConvertNumbersToText(var NumberType: OleVariant); dispid 186;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; dispid 187;
    procedure ApplyBulletDefaultOld; dispid 188;
    procedure ApplyNumberDefaultOld; dispid 189;
    procedure ApplyOutlineNumberDefaultOld; dispid 190;
    procedure ApplyListTemplateOld(const ListTemplate: ListTemplate; 
                                   var ContinuePreviousList: OleVariant; var ApplyTo: OleVariant); dispid 191;
    procedure ListOutdent; dispid 210;
    procedure ListIndent; dispid 211;
    procedure ApplyBulletDefault(var DefaultListBehavior: OleVariant); dispid 212;
    procedure ApplyNumberDefault(var DefaultListBehavior: OleVariant); dispid 213;
    procedure ApplyOutlineNumberDefault(var DefaultListBehavior: OleVariant); dispid 214;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                var ContinuePreviousList: OleVariant; var ApplyTo: OleVariant; 
                                var DefaultListBehavior: OleVariant); dispid 215;
  end;

// *********************************************************************//
// Interface: Find
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B0-0000-0000-C000-000000000046}
// *********************************************************************//
  Find = interface(IDispatch)
    ['{000209B0-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Forward: WordBool; safecall;
    procedure Set_Forward(prop: WordBool); safecall;
    function Get_Font: Font; safecall;
    procedure Set_Font(const prop: Font); safecall;
    function Get_Found: WordBool; safecall;
    function Get_MatchAllWordForms: WordBool; safecall;
    procedure Set_MatchAllWordForms(prop: WordBool); safecall;
    function Get_MatchCase: WordBool; safecall;
    procedure Set_MatchCase(prop: WordBool); safecall;
    function Get_MatchWildcards: WordBool; safecall;
    procedure Set_MatchWildcards(prop: WordBool); safecall;
    function Get_MatchSoundsLike: WordBool; safecall;
    procedure Set_MatchSoundsLike(prop: WordBool); safecall;
    function Get_MatchWholeWord: WordBool; safecall;
    procedure Set_MatchWholeWord(prop: WordBool); safecall;
    function Get_MatchFuzzy: WordBool; safecall;
    procedure Set_MatchFuzzy(prop: WordBool); safecall;
    function Get_MatchByte: WordBool; safecall;
    procedure Set_MatchByte(prop: WordBool); safecall;
    function Get_ParagraphFormat: ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: ParagraphFormat); safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_Highlight: Integer; safecall;
    procedure Set_Highlight(prop: Integer); safecall;
    function Get_Replacement: Replacement; safecall;
    function Get_Frame: Frame; safecall;
    function Get_Wrap: WdFindWrap; safecall;
    procedure Set_Wrap(prop: WdFindWrap); safecall;
    function Get_Format: WordBool; safecall;
    procedure Set_Format(prop: WordBool); safecall;
    function Get_LanguageIDFarEast: WdLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WdLanguageID); safecall;
    function Get_LanguageIDOther: WdLanguageID; safecall;
    procedure Set_LanguageIDOther(prop: WdLanguageID); safecall;
    function Get_CorrectHangulEndings: WordBool; safecall;
    procedure Set_CorrectHangulEndings(prop: WordBool); safecall;
    function ExecuteOld(var FindText: OleVariant; var MatchCase: OleVariant; 
                        var MatchWholeWord: OleVariant; var MatchWildcards: OleVariant; 
                        var MatchSoundsLike: OleVariant; var MatchAllWordForms: OleVariant; 
                        var Forward: OleVariant; var Wrap: OleVariant; var Format: OleVariant; 
                        var ReplaceWith: OleVariant; var Replace: OleVariant): WordBool; safecall;
    procedure ClearFormatting; safecall;
    procedure SetAllFuzzyOptions; safecall;
    procedure ClearAllFuzzyOptions; safecall;
    function Execute(var FindText: OleVariant; var MatchCase: OleVariant; 
                     var MatchWholeWord: OleVariant; var MatchWildcards: OleVariant; 
                     var MatchSoundsLike: OleVariant; var MatchAllWordForms: OleVariant; 
                     var Forward: OleVariant; var Wrap: OleVariant; var Format: OleVariant; 
                     var ReplaceWith: OleVariant; var Replace: OleVariant; 
                     var MatchKashida: OleVariant; var MatchDiacritics: OleVariant; 
                     var MatchAlefHamza: OleVariant; var MatchControl: OleVariant): WordBool; safecall;
    function Get_NoProofing: Integer; safecall;
    procedure Set_NoProofing(prop: Integer); safecall;
    function Get_MatchKashida: WordBool; safecall;
    procedure Set_MatchKashida(prop: WordBool); safecall;
    function Get_MatchDiacritics: WordBool; safecall;
    procedure Set_MatchDiacritics(prop: WordBool); safecall;
    function Get_MatchAlefHamza: WordBool; safecall;
    procedure Set_MatchAlefHamza(prop: WordBool); safecall;
    function Get_MatchControl: WordBool; safecall;
    procedure Set_MatchControl(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Forward: WordBool read Get_Forward write Set_Forward;
    property Font: Font read Get_Font write Set_Font;
    property Found: WordBool read Get_Found;
    property MatchAllWordForms: WordBool read Get_MatchAllWordForms write Set_MatchAllWordForms;
    property MatchCase: WordBool read Get_MatchCase write Set_MatchCase;
    property MatchWildcards: WordBool read Get_MatchWildcards write Set_MatchWildcards;
    property MatchSoundsLike: WordBool read Get_MatchSoundsLike write Set_MatchSoundsLike;
    property MatchWholeWord: WordBool read Get_MatchWholeWord write Set_MatchWholeWord;
    property MatchFuzzy: WordBool read Get_MatchFuzzy write Set_MatchFuzzy;
    property MatchByte: WordBool read Get_MatchByte write Set_MatchByte;
    property ParagraphFormat: ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property Text: WideString read Get_Text write Set_Text;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property Highlight: Integer read Get_Highlight write Set_Highlight;
    property Replacement: Replacement read Get_Replacement;
    property Frame: Frame read Get_Frame;
    property Wrap: WdFindWrap read Get_Wrap write Set_Wrap;
    property Format: WordBool read Get_Format write Set_Format;
    property LanguageIDFarEast: WdLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property LanguageIDOther: WdLanguageID read Get_LanguageIDOther write Set_LanguageIDOther;
    property CorrectHangulEndings: WordBool read Get_CorrectHangulEndings write Set_CorrectHangulEndings;
    property NoProofing: Integer read Get_NoProofing write Set_NoProofing;
    property MatchKashida: WordBool read Get_MatchKashida write Set_MatchKashida;
    property MatchDiacritics: WordBool read Get_MatchDiacritics write Set_MatchDiacritics;
    property MatchAlefHamza: WordBool read Get_MatchAlefHamza write Set_MatchAlefHamza;
    property MatchControl: WordBool read Get_MatchControl write Set_MatchControl;
  end;

// *********************************************************************//
// DispIntf:  FindDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B0-0000-0000-C000-000000000046}
// *********************************************************************//
  FindDisp = dispinterface
    ['{000209B0-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Forward: WordBool dispid 10;
    property Font: Font dispid 11;
    property Found: WordBool readonly dispid 12;
    property MatchAllWordForms: WordBool dispid 13;
    property MatchCase: WordBool dispid 14;
    property MatchWildcards: WordBool dispid 15;
    property MatchSoundsLike: WordBool dispid 16;
    property MatchWholeWord: WordBool dispid 17;
    property MatchFuzzy: WordBool dispid 40;
    property MatchByte: WordBool dispid 41;
    property ParagraphFormat: ParagraphFormat dispid 18;
    function Style: OleVariant; dispid 19;
    property Text: WideString dispid 22;
    property LanguageID: WdLanguageID dispid 23;
    property Highlight: Integer dispid 24;
    property Replacement: Replacement readonly dispid 25;
    property Frame: Frame readonly dispid 26;
    property Wrap: WdFindWrap dispid 27;
    property Format: WordBool dispid 28;
    property LanguageIDFarEast: WdLanguageID dispid 29;
    property LanguageIDOther: WdLanguageID dispid 60;
    property CorrectHangulEndings: WordBool dispid 61;
    function ExecuteOld(var FindText: OleVariant; var MatchCase: OleVariant; 
                        var MatchWholeWord: OleVariant; var MatchWildcards: OleVariant; 
                        var MatchSoundsLike: OleVariant; var MatchAllWordForms: OleVariant; 
                        var Forward: OleVariant; var Wrap: OleVariant; var Format: OleVariant; 
                        var ReplaceWith: OleVariant; var Replace: OleVariant): WordBool; dispid 30;
    procedure ClearFormatting; dispid 31;
    procedure SetAllFuzzyOptions; dispid 32;
    procedure ClearAllFuzzyOptions; dispid 33;
    function Execute(var FindText: OleVariant; var MatchCase: OleVariant; 
                     var MatchWholeWord: OleVariant; var MatchWildcards: OleVariant; 
                     var MatchSoundsLike: OleVariant; var MatchAllWordForms: OleVariant; 
                     var Forward: OleVariant; var Wrap: OleVariant; var Format: OleVariant; 
                     var ReplaceWith: OleVariant; var Replace: OleVariant; 
                     var MatchKashida: OleVariant; var MatchDiacritics: OleVariant; 
                     var MatchAlefHamza: OleVariant; var MatchControl: OleVariant): WordBool; dispid 444;
    property NoProofing: Integer dispid 34;
    property MatchKashida: WordBool dispid 100;
    property MatchDiacritics: WordBool dispid 101;
    property MatchAlefHamza: WordBool dispid 102;
    property MatchControl: WordBool dispid 103;
  end;

// *********************************************************************//
// Interface: Replacement
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B1-0000-0000-C000-000000000046}
// *********************************************************************//
  Replacement = interface(IDispatch)
    ['{000209B1-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Font: Font; safecall;
    procedure Set_Font(const prop: Font); safecall;
    function Get_ParagraphFormat: ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: ParagraphFormat); safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_Highlight: Integer; safecall;
    procedure Set_Highlight(prop: Integer); safecall;
    function Get_Frame: Frame; safecall;
    function Get_LanguageIDFarEast: WdLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WdLanguageID); safecall;
    procedure ClearFormatting; safecall;
    function Get_NoProofing: Integer; safecall;
    procedure Set_NoProofing(prop: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Font: Font read Get_Font write Set_Font;
    property ParagraphFormat: ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property Text: WideString read Get_Text write Set_Text;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property Highlight: Integer read Get_Highlight write Set_Highlight;
    property Frame: Frame read Get_Frame;
    property LanguageIDFarEast: WdLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property NoProofing: Integer read Get_NoProofing write Set_NoProofing;
  end;

// *********************************************************************//
// DispIntf:  ReplacementDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B1-0000-0000-C000-000000000046}
// *********************************************************************//
  ReplacementDisp = dispinterface
    ['{000209B1-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Font: Font dispid 10;
    property ParagraphFormat: ParagraphFormat dispid 11;
    function Style: OleVariant; dispid 12;
    property Text: WideString dispid 15;
    property LanguageID: WdLanguageID dispid 16;
    property Highlight: Integer dispid 17;
    property Frame: Frame readonly dispid 18;
    property LanguageIDFarEast: WdLanguageID dispid 19;
    procedure ClearFormatting; dispid 20;
    property NoProofing: Integer dispid 21;
  end;

// *********************************************************************//
// Interface: Characters
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095D-0000-0000-C000-000000000046}
// *********************************************************************//
  Characters = interface(IDispatch)
    ['{0002095D-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Range; safecall;
    function Get_Last: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): Range; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Range read Get_First;
    property Last: Range read Get_Last;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  CharactersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095D-0000-0000-C000-000000000046}
// *********************************************************************//
  CharactersDisp = dispinterface
    ['{0002095D-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Range readonly dispid 3;
    property Last: Range readonly dispid 4;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): Range; dispid 0;
  end;

// *********************************************************************//
// Interface: Words
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095C-0000-0000-C000-000000000046}
// *********************************************************************//
  Words = interface(IDispatch)
    ['{0002095C-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Range; safecall;
    function Get_Last: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): Range; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Range read Get_First;
    property Last: Range read Get_Last;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  WordsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095C-0000-0000-C000-000000000046}
// *********************************************************************//
  WordsDisp = dispinterface
    ['{0002095C-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Range readonly dispid 3;
    property Last: Range readonly dispid 4;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): Range; dispid 0;
  end;

// *********************************************************************//
// Interface: Sentences
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095B-0000-0000-C000-000000000046}
// *********************************************************************//
  Sentences = interface(IDispatch)
    ['{0002095B-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Range; safecall;
    function Get_Last: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): Range; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Range read Get_First;
    property Last: Range read Get_Last;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  SentencesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095B-0000-0000-C000-000000000046}
// *********************************************************************//
  SentencesDisp = dispinterface
    ['{0002095B-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Range readonly dispid 3;
    property Last: Range readonly dispid 4;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): Range; dispid 0;
  end;

// *********************************************************************//
// Interface: Sections
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095A-0000-0000-C000-000000000046}
// *********************************************************************//
  Sections = interface(IDispatch)
    ['{0002095A-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Section; safecall;
    function Get_Last: Section; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Item(Index: Integer): Section; safecall;
    function Add(var Range: OleVariant; var Start: OleVariant): Section; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Section read Get_First;
    property Last: Section read Get_Last;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
  end;

// *********************************************************************//
// DispIntf:  SectionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095A-0000-0000-C000-000000000046}
// *********************************************************************//
  SectionsDisp = dispinterface
    ['{0002095A-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Section readonly dispid 3;
    property Last: Section readonly dispid 4;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property PageSetup: PageSetup dispid 1101;
    function Item(Index: Integer): Section; dispid 0;
    function Add(var Range: OleVariant; var Start: OleVariant): Section; dispid 5;
  end;

// *********************************************************************//
// Interface: Section
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020959-0000-0000-C000-000000000046}
// *********************************************************************//
  Section = interface(IDispatch)
    ['{00020959-0000-0000-C000-000000000046}']
    function Get_Range: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_Headers: HeadersFooters; safecall;
    function Get_Footers: HeadersFooters; safecall;
    function Get_ProtectedForForms: WordBool; safecall;
    procedure Set_ProtectedForForms(prop: WordBool); safecall;
    function Get_Index: Integer; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    property Range: Range read Get_Range;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property Headers: HeadersFooters read Get_Headers;
    property Footers: HeadersFooters read Get_Footers;
    property ProtectedForForms: WordBool read Get_ProtectedForForms write Set_ProtectedForForms;
    property Index: Integer read Get_Index;
    property Borders: Borders read Get_Borders write Set_Borders;
  end;

// *********************************************************************//
// DispIntf:  SectionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020959-0000-0000-C000-000000000046}
// *********************************************************************//
  SectionDisp = dispinterface
    ['{00020959-0000-0000-C000-000000000046}']
    property Range: Range readonly dispid 0;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property PageSetup: PageSetup dispid 1101;
    property Headers: HeadersFooters readonly dispid 121;
    property Footers: HeadersFooters readonly dispid 122;
    property ProtectedForForms: WordBool dispid 123;
    property Index: Integer readonly dispid 124;
    property Borders: Borders dispid 1100;
  end;

// *********************************************************************//
// Interface: Paragraphs
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020958-0000-0000-C000-000000000046}
// *********************************************************************//
  Paragraphs = interface(IDispatch)
    ['{00020958-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Paragraph; safecall;
    function Get_Last: Paragraph; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Format: ParagraphFormat; safecall;
    procedure Set_Format(const prop: ParagraphFormat); safecall;
    function Get_TabStops: TabStops; safecall;
    procedure Set_TabStops(const prop: TabStops); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Alignment: WdParagraphAlignment; safecall;
    procedure Set_Alignment(prop: WdParagraphAlignment); safecall;
    function Get_KeepTogether: Integer; safecall;
    procedure Set_KeepTogether(prop: Integer); safecall;
    function Get_KeepWithNext: Integer; safecall;
    procedure Set_KeepWithNext(prop: Integer); safecall;
    function Get_PageBreakBefore: Integer; safecall;
    procedure Set_PageBreakBefore(prop: Integer); safecall;
    function Get_NoLineNumber: Integer; safecall;
    procedure Set_NoLineNumber(prop: Integer); safecall;
    function Get_RightIndent: Single; safecall;
    procedure Set_RightIndent(prop: Single); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_FirstLineIndent: Single; safecall;
    procedure Set_FirstLineIndent(prop: Single); safecall;
    function Get_LineSpacing: Single; safecall;
    procedure Set_LineSpacing(prop: Single); safecall;
    function Get_LineSpacingRule: WdLineSpacing; safecall;
    procedure Set_LineSpacingRule(prop: WdLineSpacing); safecall;
    function Get_SpaceBefore: Single; safecall;
    procedure Set_SpaceBefore(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    function Get_Hyphenation: Integer; safecall;
    procedure Set_Hyphenation(prop: Integer); safecall;
    function Get_WidowControl: Integer; safecall;
    procedure Set_WidowControl(prop: Integer); safecall;
    function Get_Shading: Shading; safecall;
    function Get_FarEastLineBreakControl: Integer; safecall;
    procedure Set_FarEastLineBreakControl(prop: Integer); safecall;
    function Get_WordWrap: Integer; safecall;
    procedure Set_WordWrap(prop: Integer); safecall;
    function Get_HangingPunctuation: Integer; safecall;
    procedure Set_HangingPunctuation(prop: Integer); safecall;
    function Get_HalfWidthPunctuationOnTopOfLine: Integer; safecall;
    procedure Set_HalfWidthPunctuationOnTopOfLine(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndAlpha: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndAlpha(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndDigit: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndDigit(prop: Integer); safecall;
    function Get_BaseLineAlignment: WdBaselineAlignment; safecall;
    procedure Set_BaseLineAlignment(prop: WdBaselineAlignment); safecall;
    function Get_AutoAdjustRightIndent: Integer; safecall;
    procedure Set_AutoAdjustRightIndent(prop: Integer); safecall;
    function Get_DisableLineHeightGrid: Integer; safecall;
    procedure Set_DisableLineHeightGrid(prop: Integer); safecall;
    function Get_OutlineLevel: WdOutlineLevel; safecall;
    procedure Set_OutlineLevel(prop: WdOutlineLevel); safecall;
    function Item(Index: Integer): Paragraph; safecall;
    function Add(var Range: OleVariant): Paragraph; safecall;
    procedure CloseUp; safecall;
    procedure OpenUp; safecall;
    procedure OpenOrCloseUp; safecall;
    procedure TabHangingIndent(Count: Smallint); safecall;
    procedure TabIndent(Count: Smallint); safecall;
    procedure Reset; safecall;
    procedure Space1; safecall;
    procedure Space15; safecall;
    procedure Space2; safecall;
    procedure IndentCharWidth(Count: Smallint); safecall;
    procedure IndentFirstLineCharWidth(Count: Smallint); safecall;
    procedure OutlinePromote; safecall;
    procedure OutlineDemote; safecall;
    procedure OutlineDemoteToBody; safecall;
    procedure Indent; safecall;
    procedure Outdent; safecall;
    function Get_CharacterUnitRightIndent: Single; safecall;
    procedure Set_CharacterUnitRightIndent(prop: Single); safecall;
    function Get_CharacterUnitLeftIndent: Single; safecall;
    procedure Set_CharacterUnitLeftIndent(prop: Single); safecall;
    function Get_CharacterUnitFirstLineIndent: Single; safecall;
    procedure Set_CharacterUnitFirstLineIndent(prop: Single); safecall;
    function Get_LineUnitBefore: Single; safecall;
    procedure Set_LineUnitBefore(prop: Single); safecall;
    function Get_LineUnitAfter: Single; safecall;
    procedure Set_LineUnitAfter(prop: Single); safecall;
    function Get_ReadingOrder: WdReadingOrder; safecall;
    procedure Set_ReadingOrder(prop: WdReadingOrder); safecall;
    function Get_SpaceBeforeAuto: Integer; safecall;
    procedure Set_SpaceBeforeAuto(prop: Integer); safecall;
    function Get_SpaceAfterAuto: Integer; safecall;
    procedure Set_SpaceAfterAuto(prop: Integer); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Paragraph read Get_First;
    property Last: Paragraph read Get_Last;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Format: ParagraphFormat read Get_Format write Set_Format;
    property TabStops: TabStops read Get_TabStops write Set_TabStops;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Alignment: WdParagraphAlignment read Get_Alignment write Set_Alignment;
    property KeepTogether: Integer read Get_KeepTogether write Set_KeepTogether;
    property KeepWithNext: Integer read Get_KeepWithNext write Set_KeepWithNext;
    property PageBreakBefore: Integer read Get_PageBreakBefore write Set_PageBreakBefore;
    property NoLineNumber: Integer read Get_NoLineNumber write Set_NoLineNumber;
    property RightIndent: Single read Get_RightIndent write Set_RightIndent;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property FirstLineIndent: Single read Get_FirstLineIndent write Set_FirstLineIndent;
    property LineSpacing: Single read Get_LineSpacing write Set_LineSpacing;
    property LineSpacingRule: WdLineSpacing read Get_LineSpacingRule write Set_LineSpacingRule;
    property SpaceBefore: Single read Get_SpaceBefore write Set_SpaceBefore;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
    property Hyphenation: Integer read Get_Hyphenation write Set_Hyphenation;
    property WidowControl: Integer read Get_WidowControl write Set_WidowControl;
    property Shading: Shading read Get_Shading;
    property FarEastLineBreakControl: Integer read Get_FarEastLineBreakControl write Set_FarEastLineBreakControl;
    property WordWrap: Integer read Get_WordWrap write Set_WordWrap;
    property HangingPunctuation: Integer read Get_HangingPunctuation write Set_HangingPunctuation;
    property HalfWidthPunctuationOnTopOfLine: Integer read Get_HalfWidthPunctuationOnTopOfLine write Set_HalfWidthPunctuationOnTopOfLine;
    property AddSpaceBetweenFarEastAndAlpha: Integer read Get_AddSpaceBetweenFarEastAndAlpha write Set_AddSpaceBetweenFarEastAndAlpha;
    property AddSpaceBetweenFarEastAndDigit: Integer read Get_AddSpaceBetweenFarEastAndDigit write Set_AddSpaceBetweenFarEastAndDigit;
    property BaseLineAlignment: WdBaselineAlignment read Get_BaseLineAlignment write Set_BaseLineAlignment;
    property AutoAdjustRightIndent: Integer read Get_AutoAdjustRightIndent write Set_AutoAdjustRightIndent;
    property DisableLineHeightGrid: Integer read Get_DisableLineHeightGrid write Set_DisableLineHeightGrid;
    property OutlineLevel: WdOutlineLevel read Get_OutlineLevel write Set_OutlineLevel;
    property CharacterUnitRightIndent: Single read Get_CharacterUnitRightIndent write Set_CharacterUnitRightIndent;
    property CharacterUnitLeftIndent: Single read Get_CharacterUnitLeftIndent write Set_CharacterUnitLeftIndent;
    property CharacterUnitFirstLineIndent: Single read Get_CharacterUnitFirstLineIndent write Set_CharacterUnitFirstLineIndent;
    property LineUnitBefore: Single read Get_LineUnitBefore write Set_LineUnitBefore;
    property LineUnitAfter: Single read Get_LineUnitAfter write Set_LineUnitAfter;
    property ReadingOrder: WdReadingOrder read Get_ReadingOrder write Set_ReadingOrder;
    property SpaceBeforeAuto: Integer read Get_SpaceBeforeAuto write Set_SpaceBeforeAuto;
    property SpaceAfterAuto: Integer read Get_SpaceAfterAuto write Set_SpaceAfterAuto;
  end;

// *********************************************************************//
// DispIntf:  ParagraphsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020958-0000-0000-C000-000000000046}
// *********************************************************************//
  ParagraphsDisp = dispinterface
    ['{00020958-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Paragraph readonly dispid 3;
    property Last: Paragraph readonly dispid 4;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Format: ParagraphFormat dispid 1102;
    property TabStops: TabStops dispid 1103;
    property Borders: Borders dispid 1100;
    function Style: OleVariant; dispid 100;
    property Alignment: WdParagraphAlignment dispid 101;
    property KeepTogether: Integer dispid 102;
    property KeepWithNext: Integer dispid 103;
    property PageBreakBefore: Integer dispid 104;
    property NoLineNumber: Integer dispid 105;
    property RightIndent: Single dispid 106;
    property LeftIndent: Single dispid 107;
    property FirstLineIndent: Single dispid 108;
    property LineSpacing: Single dispid 109;
    property LineSpacingRule: WdLineSpacing dispid 110;
    property SpaceBefore: Single dispid 111;
    property SpaceAfter: Single dispid 112;
    property Hyphenation: Integer dispid 113;
    property WidowControl: Integer dispid 114;
    property Shading: Shading readonly dispid 116;
    property FarEastLineBreakControl: Integer dispid 117;
    property WordWrap: Integer dispid 118;
    property HangingPunctuation: Integer dispid 119;
    property HalfWidthPunctuationOnTopOfLine: Integer dispid 120;
    property AddSpaceBetweenFarEastAndAlpha: Integer dispid 121;
    property AddSpaceBetweenFarEastAndDigit: Integer dispid 122;
    property BaseLineAlignment: WdBaselineAlignment dispid 123;
    property AutoAdjustRightIndent: Integer dispid 124;
    property DisableLineHeightGrid: Integer dispid 125;
    property OutlineLevel: WdOutlineLevel dispid 202;
    function Item(Index: Integer): Paragraph; dispid 0;
    function Add(var Range: OleVariant): Paragraph; dispid 5;
    procedure CloseUp; dispid 301;
    procedure OpenUp; dispid 302;
    procedure OpenOrCloseUp; dispid 303;
    procedure TabHangingIndent(Count: Smallint); dispid 304;
    procedure TabIndent(Count: Smallint); dispid 306;
    procedure Reset; dispid 312;
    procedure Space1; dispid 313;
    procedure Space15; dispid 314;
    procedure Space2; dispid 315;
    procedure IndentCharWidth(Count: Smallint); dispid 320;
    procedure IndentFirstLineCharWidth(Count: Smallint); dispid 322;
    procedure OutlinePromote; dispid 324;
    procedure OutlineDemote; dispid 325;
    procedure OutlineDemoteToBody; dispid 326;
    procedure Indent; dispid 333;
    procedure Outdent; dispid 334;
    property CharacterUnitRightIndent: Single dispid 126;
    property CharacterUnitLeftIndent: Single dispid 127;
    property CharacterUnitFirstLineIndent: Single dispid 128;
    property LineUnitBefore: Single dispid 129;
    property LineUnitAfter: Single dispid 130;
    property ReadingOrder: WdReadingOrder dispid 131;
    property SpaceBeforeAuto: Integer dispid 132;
    property SpaceAfterAuto: Integer dispid 133;
  end;

// *********************************************************************//
// Interface: Paragraph
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020957-0000-0000-C000-000000000046}
// *********************************************************************//
  Paragraph = interface(IDispatch)
    ['{00020957-0000-0000-C000-000000000046}']
    function Get_Range: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Format: ParagraphFormat; safecall;
    procedure Set_Format(const prop: ParagraphFormat); safecall;
    function Get_TabStops: TabStops; safecall;
    procedure Set_TabStops(const prop: TabStops); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_DropCap: DropCap; safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Alignment: WdParagraphAlignment; safecall;
    procedure Set_Alignment(prop: WdParagraphAlignment); safecall;
    function Get_KeepTogether: Integer; safecall;
    procedure Set_KeepTogether(prop: Integer); safecall;
    function Get_KeepWithNext: Integer; safecall;
    procedure Set_KeepWithNext(prop: Integer); safecall;
    function Get_PageBreakBefore: Integer; safecall;
    procedure Set_PageBreakBefore(prop: Integer); safecall;
    function Get_NoLineNumber: Integer; safecall;
    procedure Set_NoLineNumber(prop: Integer); safecall;
    function Get_RightIndent: Single; safecall;
    procedure Set_RightIndent(prop: Single); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_FirstLineIndent: Single; safecall;
    procedure Set_FirstLineIndent(prop: Single); safecall;
    function Get_LineSpacing: Single; safecall;
    procedure Set_LineSpacing(prop: Single); safecall;
    function Get_LineSpacingRule: WdLineSpacing; safecall;
    procedure Set_LineSpacingRule(prop: WdLineSpacing); safecall;
    function Get_SpaceBefore: Single; safecall;
    procedure Set_SpaceBefore(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    function Get_Hyphenation: Integer; safecall;
    procedure Set_Hyphenation(prop: Integer); safecall;
    function Get_WidowControl: Integer; safecall;
    procedure Set_WidowControl(prop: Integer); safecall;
    function Get_Shading: Shading; safecall;
    function Get_FarEastLineBreakControl: Integer; safecall;
    procedure Set_FarEastLineBreakControl(prop: Integer); safecall;
    function Get_WordWrap: Integer; safecall;
    procedure Set_WordWrap(prop: Integer); safecall;
    function Get_HangingPunctuation: Integer; safecall;
    procedure Set_HangingPunctuation(prop: Integer); safecall;
    function Get_HalfWidthPunctuationOnTopOfLine: Integer; safecall;
    procedure Set_HalfWidthPunctuationOnTopOfLine(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndAlpha: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndAlpha(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndDigit: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndDigit(prop: Integer); safecall;
    function Get_BaseLineAlignment: WdBaselineAlignment; safecall;
    procedure Set_BaseLineAlignment(prop: WdBaselineAlignment); safecall;
    function Get_AutoAdjustRightIndent: Integer; safecall;
    procedure Set_AutoAdjustRightIndent(prop: Integer); safecall;
    function Get_DisableLineHeightGrid: Integer; safecall;
    procedure Set_DisableLineHeightGrid(prop: Integer); safecall;
    function Get_OutlineLevel: WdOutlineLevel; safecall;
    procedure Set_OutlineLevel(prop: WdOutlineLevel); safecall;
    procedure CloseUp; safecall;
    procedure OpenUp; safecall;
    procedure OpenOrCloseUp; safecall;
    procedure TabHangingIndent(Count: Smallint); safecall;
    procedure TabIndent(Count: Smallint); safecall;
    procedure Reset; safecall;
    procedure Space1; safecall;
    procedure Space15; safecall;
    procedure Space2; safecall;
    procedure IndentCharWidth(Count: Smallint); safecall;
    procedure IndentFirstLineCharWidth(Count: Smallint); safecall;
    function Next(var Count: OleVariant): Paragraph; safecall;
    function Previous(var Count: OleVariant): Paragraph; safecall;
    procedure OutlinePromote; safecall;
    procedure OutlineDemote; safecall;
    procedure OutlineDemoteToBody; safecall;
    procedure Indent; safecall;
    procedure Outdent; safecall;
    function Get_CharacterUnitRightIndent: Single; safecall;
    procedure Set_CharacterUnitRightIndent(prop: Single); safecall;
    function Get_CharacterUnitLeftIndent: Single; safecall;
    procedure Set_CharacterUnitLeftIndent(prop: Single); safecall;
    function Get_CharacterUnitFirstLineIndent: Single; safecall;
    procedure Set_CharacterUnitFirstLineIndent(prop: Single); safecall;
    function Get_LineUnitBefore: Single; safecall;
    procedure Set_LineUnitBefore(prop: Single); safecall;
    function Get_LineUnitAfter: Single; safecall;
    procedure Set_LineUnitAfter(prop: Single); safecall;
    function Get_ReadingOrder: WdReadingOrder; safecall;
    procedure Set_ReadingOrder(prop: WdReadingOrder); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    function Get_SpaceBeforeAuto: Integer; safecall;
    procedure Set_SpaceBeforeAuto(prop: Integer); safecall;
    function Get_SpaceAfterAuto: Integer; safecall;
    procedure Set_SpaceAfterAuto(prop: Integer); safecall;
    property Range: Range read Get_Range;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Format: ParagraphFormat read Get_Format write Set_Format;
    property TabStops: TabStops read Get_TabStops write Set_TabStops;
    property Borders: Borders read Get_Borders write Set_Borders;
    property DropCap: DropCap read Get_DropCap;
    property Alignment: WdParagraphAlignment read Get_Alignment write Set_Alignment;
    property KeepTogether: Integer read Get_KeepTogether write Set_KeepTogether;
    property KeepWithNext: Integer read Get_KeepWithNext write Set_KeepWithNext;
    property PageBreakBefore: Integer read Get_PageBreakBefore write Set_PageBreakBefore;
    property NoLineNumber: Integer read Get_NoLineNumber write Set_NoLineNumber;
    property RightIndent: Single read Get_RightIndent write Set_RightIndent;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property FirstLineIndent: Single read Get_FirstLineIndent write Set_FirstLineIndent;
    property LineSpacing: Single read Get_LineSpacing write Set_LineSpacing;
    property LineSpacingRule: WdLineSpacing read Get_LineSpacingRule write Set_LineSpacingRule;
    property SpaceBefore: Single read Get_SpaceBefore write Set_SpaceBefore;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
    property Hyphenation: Integer read Get_Hyphenation write Set_Hyphenation;
    property WidowControl: Integer read Get_WidowControl write Set_WidowControl;
    property Shading: Shading read Get_Shading;
    property FarEastLineBreakControl: Integer read Get_FarEastLineBreakControl write Set_FarEastLineBreakControl;
    property WordWrap: Integer read Get_WordWrap write Set_WordWrap;
    property HangingPunctuation: Integer read Get_HangingPunctuation write Set_HangingPunctuation;
    property HalfWidthPunctuationOnTopOfLine: Integer read Get_HalfWidthPunctuationOnTopOfLine write Set_HalfWidthPunctuationOnTopOfLine;
    property AddSpaceBetweenFarEastAndAlpha: Integer read Get_AddSpaceBetweenFarEastAndAlpha write Set_AddSpaceBetweenFarEastAndAlpha;
    property AddSpaceBetweenFarEastAndDigit: Integer read Get_AddSpaceBetweenFarEastAndDigit write Set_AddSpaceBetweenFarEastAndDigit;
    property BaseLineAlignment: WdBaselineAlignment read Get_BaseLineAlignment write Set_BaseLineAlignment;
    property AutoAdjustRightIndent: Integer read Get_AutoAdjustRightIndent write Set_AutoAdjustRightIndent;
    property DisableLineHeightGrid: Integer read Get_DisableLineHeightGrid write Set_DisableLineHeightGrid;
    property OutlineLevel: WdOutlineLevel read Get_OutlineLevel write Set_OutlineLevel;
    property CharacterUnitRightIndent: Single read Get_CharacterUnitRightIndent write Set_CharacterUnitRightIndent;
    property CharacterUnitLeftIndent: Single read Get_CharacterUnitLeftIndent write Set_CharacterUnitLeftIndent;
    property CharacterUnitFirstLineIndent: Single read Get_CharacterUnitFirstLineIndent write Set_CharacterUnitFirstLineIndent;
    property LineUnitBefore: Single read Get_LineUnitBefore write Set_LineUnitBefore;
    property LineUnitAfter: Single read Get_LineUnitAfter write Set_LineUnitAfter;
    property ReadingOrder: WdReadingOrder read Get_ReadingOrder write Set_ReadingOrder;
    property ID: WideString read Get_ID write Set_ID;
    property SpaceBeforeAuto: Integer read Get_SpaceBeforeAuto write Set_SpaceBeforeAuto;
    property SpaceAfterAuto: Integer read Get_SpaceAfterAuto write Set_SpaceAfterAuto;
  end;

// *********************************************************************//
// DispIntf:  ParagraphDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020957-0000-0000-C000-000000000046}
// *********************************************************************//
  ParagraphDisp = dispinterface
    ['{00020957-0000-0000-C000-000000000046}']
    property Range: Range readonly dispid 0;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Format: ParagraphFormat dispid 1102;
    property TabStops: TabStops dispid 1103;
    property Borders: Borders dispid 1100;
    property DropCap: DropCap readonly dispid 13;
    function Style: OleVariant; dispid 100;
    property Alignment: WdParagraphAlignment dispid 101;
    property KeepTogether: Integer dispid 102;
    property KeepWithNext: Integer dispid 103;
    property PageBreakBefore: Integer dispid 104;
    property NoLineNumber: Integer dispid 105;
    property RightIndent: Single dispid 106;
    property LeftIndent: Single dispid 107;
    property FirstLineIndent: Single dispid 108;
    property LineSpacing: Single dispid 109;
    property LineSpacingRule: WdLineSpacing dispid 110;
    property SpaceBefore: Single dispid 111;
    property SpaceAfter: Single dispid 112;
    property Hyphenation: Integer dispid 113;
    property WidowControl: Integer dispid 114;
    property Shading: Shading readonly dispid 116;
    property FarEastLineBreakControl: Integer dispid 117;
    property WordWrap: Integer dispid 118;
    property HangingPunctuation: Integer dispid 119;
    property HalfWidthPunctuationOnTopOfLine: Integer dispid 120;
    property AddSpaceBetweenFarEastAndAlpha: Integer dispid 121;
    property AddSpaceBetweenFarEastAndDigit: Integer dispid 122;
    property BaseLineAlignment: WdBaselineAlignment dispid 123;
    property AutoAdjustRightIndent: Integer dispid 124;
    property DisableLineHeightGrid: Integer dispid 125;
    property OutlineLevel: WdOutlineLevel dispid 202;
    procedure CloseUp; dispid 301;
    procedure OpenUp; dispid 302;
    procedure OpenOrCloseUp; dispid 303;
    procedure TabHangingIndent(Count: Smallint); dispid 304;
    procedure TabIndent(Count: Smallint); dispid 306;
    procedure Reset; dispid 312;
    procedure Space1; dispid 313;
    procedure Space15; dispid 314;
    procedure Space2; dispid 315;
    procedure IndentCharWidth(Count: Smallint); dispid 320;
    procedure IndentFirstLineCharWidth(Count: Smallint); dispid 322;
    function Next(var Count: OleVariant): Paragraph; dispid 324;
    function Previous(var Count: OleVariant): Paragraph; dispid 325;
    procedure OutlinePromote; dispid 326;
    procedure OutlineDemote; dispid 327;
    procedure OutlineDemoteToBody; dispid 328;
    procedure Indent; dispid 333;
    procedure Outdent; dispid 334;
    property CharacterUnitRightIndent: Single dispid 126;
    property CharacterUnitLeftIndent: Single dispid 127;
    property CharacterUnitFirstLineIndent: Single dispid 128;
    property LineUnitBefore: Single dispid 129;
    property LineUnitAfter: Single dispid 130;
    property ReadingOrder: WdReadingOrder dispid 203;
    property ID: WideString dispid 204;
    property SpaceBeforeAuto: Integer dispid 132;
    property SpaceAfterAuto: Integer dispid 133;
  end;

// *********************************************************************//
// Interface: DropCap
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020956-0000-0000-C000-000000000046}
// *********************************************************************//
  DropCap = interface(IDispatch)
    ['{00020956-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Position: WdDropPosition; safecall;
    procedure Set_Position(prop: WdDropPosition); safecall;
    function Get_FontName: WideString; safecall;
    procedure Set_FontName(const prop: WideString); safecall;
    function Get_LinesToDrop: Integer; safecall;
    procedure Set_LinesToDrop(prop: Integer); safecall;
    function Get_DistanceFromText: Single; safecall;
    procedure Set_DistanceFromText(prop: Single); safecall;
    procedure Clear; safecall;
    procedure Enable; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Position: WdDropPosition read Get_Position write Set_Position;
    property FontName: WideString read Get_FontName write Set_FontName;
    property LinesToDrop: Integer read Get_LinesToDrop write Set_LinesToDrop;
    property DistanceFromText: Single read Get_DistanceFromText write Set_DistanceFromText;
  end;

// *********************************************************************//
// DispIntf:  DropCapDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020956-0000-0000-C000-000000000046}
// *********************************************************************//
  DropCapDisp = dispinterface
    ['{00020956-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Position: WdDropPosition dispid 10;
    property FontName: WideString dispid 11;
    property LinesToDrop: Integer dispid 12;
    property DistanceFromText: Single dispid 13;
    procedure Clear; dispid 100;
    procedure Enable; dispid 101;
  end;

// *********************************************************************//
// Interface: TabStops
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020955-0000-0000-C000-000000000046}
// *********************************************************************//
  TabStops = interface(IDispatch)
    ['{00020955-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): TabStop; safecall;
    function Add(Position: Single; var Alignment: OleVariant; var Leader: OleVariant): TabStop; safecall;
    procedure ClearAll; safecall;
    function Before(Position: Single): TabStop; safecall;
    function After(Position: Single): TabStop; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  TabStopsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020955-0000-0000-C000-000000000046}
// *********************************************************************//
  TabStopsDisp = dispinterface
    ['{00020955-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): TabStop; dispid 0;
    function Add(Position: Single; var Alignment: OleVariant; var Leader: OleVariant): TabStop; dispid 100;
    procedure ClearAll; dispid 101;
    function Before(Position: Single): TabStop; dispid 102;
    function After(Position: Single): TabStop; dispid 103;
  end;

// *********************************************************************//
// Interface: TabStop
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020954-0000-0000-C000-000000000046}
// *********************************************************************//
  TabStop = interface(IDispatch)
    ['{00020954-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Alignment: WdTabAlignment; safecall;
    procedure Set_Alignment(prop: WdTabAlignment); safecall;
    function Get_Leader: WdTabLeader; safecall;
    procedure Set_Leader(prop: WdTabLeader); safecall;
    function Get_Position: Single; safecall;
    procedure Set_Position(prop: Single); safecall;
    function Get_CustomTab: WordBool; safecall;
    function Get_Next: TabStop; safecall;
    function Get_Previous: TabStop; safecall;
    procedure Clear; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Alignment: WdTabAlignment read Get_Alignment write Set_Alignment;
    property Leader: WdTabLeader read Get_Leader write Set_Leader;
    property Position: Single read Get_Position write Set_Position;
    property CustomTab: WordBool read Get_CustomTab;
    property Next: TabStop read Get_Next;
    property Previous: TabStop read Get_Previous;
  end;

// *********************************************************************//
// DispIntf:  TabStopDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020954-0000-0000-C000-000000000046}
// *********************************************************************//
  TabStopDisp = dispinterface
    ['{00020954-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Alignment: WdTabAlignment dispid 100;
    property Leader: WdTabLeader dispid 101;
    property Position: Single dispid 102;
    property CustomTab: WordBool readonly dispid 103;
    property Next: TabStop readonly dispid 104;
    property Previous: TabStop readonly dispid 105;
    procedure Clear; dispid 200;
  end;

// *********************************************************************//
// Interface: _ParagraphFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020953-0000-0000-C000-000000000046}
// *********************************************************************//
  _ParagraphFormat = interface(IDispatch)
    ['{00020953-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Duplicate: ParagraphFormat; safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Alignment: WdParagraphAlignment; safecall;
    procedure Set_Alignment(prop: WdParagraphAlignment); safecall;
    function Get_KeepTogether: Integer; safecall;
    procedure Set_KeepTogether(prop: Integer); safecall;
    function Get_KeepWithNext: Integer; safecall;
    procedure Set_KeepWithNext(prop: Integer); safecall;
    function Get_PageBreakBefore: Integer; safecall;
    procedure Set_PageBreakBefore(prop: Integer); safecall;
    function Get_NoLineNumber: Integer; safecall;
    procedure Set_NoLineNumber(prop: Integer); safecall;
    function Get_RightIndent: Single; safecall;
    procedure Set_RightIndent(prop: Single); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_FirstLineIndent: Single; safecall;
    procedure Set_FirstLineIndent(prop: Single); safecall;
    function Get_LineSpacing: Single; safecall;
    procedure Set_LineSpacing(prop: Single); safecall;
    function Get_LineSpacingRule: WdLineSpacing; safecall;
    procedure Set_LineSpacingRule(prop: WdLineSpacing); safecall;
    function Get_SpaceBefore: Single; safecall;
    procedure Set_SpaceBefore(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    function Get_Hyphenation: Integer; safecall;
    procedure Set_Hyphenation(prop: Integer); safecall;
    function Get_WidowControl: Integer; safecall;
    procedure Set_WidowControl(prop: Integer); safecall;
    function Get_FarEastLineBreakControl: Integer; safecall;
    procedure Set_FarEastLineBreakControl(prop: Integer); safecall;
    function Get_WordWrap: Integer; safecall;
    procedure Set_WordWrap(prop: Integer); safecall;
    function Get_HangingPunctuation: Integer; safecall;
    procedure Set_HangingPunctuation(prop: Integer); safecall;
    function Get_HalfWidthPunctuationOnTopOfLine: Integer; safecall;
    procedure Set_HalfWidthPunctuationOnTopOfLine(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndAlpha: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndAlpha(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndDigit: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndDigit(prop: Integer); safecall;
    function Get_BaseLineAlignment: WdBaselineAlignment; safecall;
    procedure Set_BaseLineAlignment(prop: WdBaselineAlignment); safecall;
    function Get_AutoAdjustRightIndent: Integer; safecall;
    procedure Set_AutoAdjustRightIndent(prop: Integer); safecall;
    function Get_DisableLineHeightGrid: Integer; safecall;
    procedure Set_DisableLineHeightGrid(prop: Integer); safecall;
    function Get_TabStops: TabStops; safecall;
    procedure Set_TabStops(const prop: TabStops); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_OutlineLevel: WdOutlineLevel; safecall;
    procedure Set_OutlineLevel(prop: WdOutlineLevel); safecall;
    procedure CloseUp; safecall;
    procedure OpenUp; safecall;
    procedure OpenOrCloseUp; safecall;
    procedure TabHangingIndent(Count: Smallint); safecall;
    procedure TabIndent(Count: Smallint); safecall;
    procedure Reset; safecall;
    procedure Space1; safecall;
    procedure Space15; safecall;
    procedure Space2; safecall;
    procedure IndentCharWidth(Count: Smallint); safecall;
    procedure IndentFirstLineCharWidth(Count: Smallint); safecall;
    function Get_CharacterUnitRightIndent: Single; safecall;
    procedure Set_CharacterUnitRightIndent(prop: Single); safecall;
    function Get_CharacterUnitLeftIndent: Single; safecall;
    procedure Set_CharacterUnitLeftIndent(prop: Single); safecall;
    function Get_CharacterUnitFirstLineIndent: Single; safecall;
    procedure Set_CharacterUnitFirstLineIndent(prop: Single); safecall;
    function Get_LineUnitBefore: Single; safecall;
    procedure Set_LineUnitBefore(prop: Single); safecall;
    function Get_LineUnitAfter: Single; safecall;
    procedure Set_LineUnitAfter(prop: Single); safecall;
    function Get_ReadingOrder: WdReadingOrder; safecall;
    procedure Set_ReadingOrder(prop: WdReadingOrder); safecall;
    function Get_SpaceBeforeAuto: Integer; safecall;
    procedure Set_SpaceBeforeAuto(prop: Integer); safecall;
    function Get_SpaceAfterAuto: Integer; safecall;
    procedure Set_SpaceAfterAuto(prop: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Duplicate: ParagraphFormat read Get_Duplicate;
    property Alignment: WdParagraphAlignment read Get_Alignment write Set_Alignment;
    property KeepTogether: Integer read Get_KeepTogether write Set_KeepTogether;
    property KeepWithNext: Integer read Get_KeepWithNext write Set_KeepWithNext;
    property PageBreakBefore: Integer read Get_PageBreakBefore write Set_PageBreakBefore;
    property NoLineNumber: Integer read Get_NoLineNumber write Set_NoLineNumber;
    property RightIndent: Single read Get_RightIndent write Set_RightIndent;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property FirstLineIndent: Single read Get_FirstLineIndent write Set_FirstLineIndent;
    property LineSpacing: Single read Get_LineSpacing write Set_LineSpacing;
    property LineSpacingRule: WdLineSpacing read Get_LineSpacingRule write Set_LineSpacingRule;
    property SpaceBefore: Single read Get_SpaceBefore write Set_SpaceBefore;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
    property Hyphenation: Integer read Get_Hyphenation write Set_Hyphenation;
    property WidowControl: Integer read Get_WidowControl write Set_WidowControl;
    property FarEastLineBreakControl: Integer read Get_FarEastLineBreakControl write Set_FarEastLineBreakControl;
    property WordWrap: Integer read Get_WordWrap write Set_WordWrap;
    property HangingPunctuation: Integer read Get_HangingPunctuation write Set_HangingPunctuation;
    property HalfWidthPunctuationOnTopOfLine: Integer read Get_HalfWidthPunctuationOnTopOfLine write Set_HalfWidthPunctuationOnTopOfLine;
    property AddSpaceBetweenFarEastAndAlpha: Integer read Get_AddSpaceBetweenFarEastAndAlpha write Set_AddSpaceBetweenFarEastAndAlpha;
    property AddSpaceBetweenFarEastAndDigit: Integer read Get_AddSpaceBetweenFarEastAndDigit write Set_AddSpaceBetweenFarEastAndDigit;
    property BaseLineAlignment: WdBaselineAlignment read Get_BaseLineAlignment write Set_BaseLineAlignment;
    property AutoAdjustRightIndent: Integer read Get_AutoAdjustRightIndent write Set_AutoAdjustRightIndent;
    property DisableLineHeightGrid: Integer read Get_DisableLineHeightGrid write Set_DisableLineHeightGrid;
    property TabStops: TabStops read Get_TabStops write Set_TabStops;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property OutlineLevel: WdOutlineLevel read Get_OutlineLevel write Set_OutlineLevel;
    property CharacterUnitRightIndent: Single read Get_CharacterUnitRightIndent write Set_CharacterUnitRightIndent;
    property CharacterUnitLeftIndent: Single read Get_CharacterUnitLeftIndent write Set_CharacterUnitLeftIndent;
    property CharacterUnitFirstLineIndent: Single read Get_CharacterUnitFirstLineIndent write Set_CharacterUnitFirstLineIndent;
    property LineUnitBefore: Single read Get_LineUnitBefore write Set_LineUnitBefore;
    property LineUnitAfter: Single read Get_LineUnitAfter write Set_LineUnitAfter;
    property ReadingOrder: WdReadingOrder read Get_ReadingOrder write Set_ReadingOrder;
    property SpaceBeforeAuto: Integer read Get_SpaceBeforeAuto write Set_SpaceBeforeAuto;
    property SpaceAfterAuto: Integer read Get_SpaceAfterAuto write Set_SpaceAfterAuto;
  end;

// *********************************************************************//
// DispIntf:  _ParagraphFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020953-0000-0000-C000-000000000046}
// *********************************************************************//
  _ParagraphFormatDisp = dispinterface
    ['{00020953-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Duplicate: ParagraphFormat readonly dispid 10;
    function Style: OleVariant; dispid 100;
    property Alignment: WdParagraphAlignment dispid 101;
    property KeepTogether: Integer dispid 102;
    property KeepWithNext: Integer dispid 103;
    property PageBreakBefore: Integer dispid 104;
    property NoLineNumber: Integer dispid 105;
    property RightIndent: Single dispid 106;
    property LeftIndent: Single dispid 107;
    property FirstLineIndent: Single dispid 108;
    property LineSpacing: Single dispid 109;
    property LineSpacingRule: WdLineSpacing dispid 110;
    property SpaceBefore: Single dispid 111;
    property SpaceAfter: Single dispid 112;
    property Hyphenation: Integer dispid 113;
    property WidowControl: Integer dispid 114;
    property FarEastLineBreakControl: Integer dispid 117;
    property WordWrap: Integer dispid 118;
    property HangingPunctuation: Integer dispid 119;
    property HalfWidthPunctuationOnTopOfLine: Integer dispid 120;
    property AddSpaceBetweenFarEastAndAlpha: Integer dispid 121;
    property AddSpaceBetweenFarEastAndDigit: Integer dispid 122;
    property BaseLineAlignment: WdBaselineAlignment dispid 123;
    property AutoAdjustRightIndent: Integer dispid 124;
    property DisableLineHeightGrid: Integer dispid 125;
    property TabStops: TabStops dispid 1103;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 1101;
    property OutlineLevel: WdOutlineLevel dispid 202;
    procedure CloseUp; dispid 301;
    procedure OpenUp; dispid 302;
    procedure OpenOrCloseUp; dispid 303;
    procedure TabHangingIndent(Count: Smallint); dispid 304;
    procedure TabIndent(Count: Smallint); dispid 306;
    procedure Reset; dispid 312;
    procedure Space1; dispid 313;
    procedure Space15; dispid 314;
    procedure Space2; dispid 315;
    procedure IndentCharWidth(Count: Smallint); dispid 320;
    procedure IndentFirstLineCharWidth(Count: Smallint); dispid 322;
    property CharacterUnitRightIndent: Single dispid 126;
    property CharacterUnitLeftIndent: Single dispid 127;
    property CharacterUnitFirstLineIndent: Single dispid 128;
    property LineUnitBefore: Single dispid 129;
    property LineUnitAfter: Single dispid 130;
    property ReadingOrder: WdReadingOrder dispid 131;
    property SpaceBeforeAuto: Integer dispid 132;
    property SpaceAfterAuto: Integer dispid 133;
  end;

// *********************************************************************//
// Interface: _Font
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020952-0000-0000-C000-000000000046}
// *********************************************************************//
  _Font = interface(IDispatch)
    ['{00020952-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Duplicate: Font; safecall;
    function Get_Bold: Integer; safecall;
    procedure Set_Bold(prop: Integer); safecall;
    function Get_Italic: Integer; safecall;
    procedure Set_Italic(prop: Integer); safecall;
    function Get_Hidden: Integer; safecall;
    procedure Set_Hidden(prop: Integer); safecall;
    function Get_SmallCaps: Integer; safecall;
    procedure Set_SmallCaps(prop: Integer); safecall;
    function Get_AllCaps: Integer; safecall;
    procedure Set_AllCaps(prop: Integer); safecall;
    function Get_StrikeThrough: Integer; safecall;
    procedure Set_StrikeThrough(prop: Integer); safecall;
    function Get_DoubleStrikeThrough: Integer; safecall;
    procedure Set_DoubleStrikeThrough(prop: Integer); safecall;
    function Get_ColorIndex: WdColorIndex; safecall;
    procedure Set_ColorIndex(prop: WdColorIndex); safecall;
    function Get_Subscript: Integer; safecall;
    procedure Set_Subscript(prop: Integer); safecall;
    function Get_Superscript: Integer; safecall;
    procedure Set_Superscript(prop: Integer); safecall;
    function Get_Underline: WdUnderline; safecall;
    procedure Set_Underline(prop: WdUnderline); safecall;
    function Get_Size: Single; safecall;
    procedure Set_Size(prop: Single); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Position: Integer; safecall;
    procedure Set_Position(prop: Integer); safecall;
    function Get_Spacing: Single; safecall;
    procedure Set_Spacing(prop: Single); safecall;
    function Get_Scaling: Integer; safecall;
    procedure Set_Scaling(prop: Integer); safecall;
    function Get_Shadow: Integer; safecall;
    procedure Set_Shadow(prop: Integer); safecall;
    function Get_Outline: Integer; safecall;
    procedure Set_Outline(prop: Integer); safecall;
    function Get_Emboss: Integer; safecall;
    procedure Set_Emboss(prop: Integer); safecall;
    function Get_Kerning: Single; safecall;
    procedure Set_Kerning(prop: Single); safecall;
    function Get_Engrave: Integer; safecall;
    procedure Set_Engrave(prop: Integer); safecall;
    function Get_Animation: WdAnimation; safecall;
    procedure Set_Animation(prop: WdAnimation); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_EmphasisMark: WdEmphasisMark; safecall;
    procedure Set_EmphasisMark(prop: WdEmphasisMark); safecall;
    function Get_DisableCharacterSpaceGrid: WordBool; safecall;
    procedure Set_DisableCharacterSpaceGrid(prop: WordBool); safecall;
    function Get_NameFarEast: WideString; safecall;
    procedure Set_NameFarEast(const prop: WideString); safecall;
    function Get_NameAscii: WideString; safecall;
    procedure Set_NameAscii(const prop: WideString); safecall;
    function Get_NameOther: WideString; safecall;
    procedure Set_NameOther(const prop: WideString); safecall;
    procedure Grow; safecall;
    procedure Shrink; safecall;
    procedure Reset; safecall;
    procedure SetAsTemplateDefault; safecall;
    function Get_Color: WdColor; safecall;
    procedure Set_Color(prop: WdColor); safecall;
    function Get_BoldBi: Integer; safecall;
    procedure Set_BoldBi(prop: Integer); safecall;
    function Get_ItalicBi: Integer; safecall;
    procedure Set_ItalicBi(prop: Integer); safecall;
    function Get_SizeBi: Single; safecall;
    procedure Set_SizeBi(prop: Single); safecall;
    function Get_NameBi: WideString; safecall;
    procedure Set_NameBi(const prop: WideString); safecall;
    function Get_ColorIndexBi: WdColorIndex; safecall;
    procedure Set_ColorIndexBi(prop: WdColorIndex); safecall;
    function Get_DiacriticColor: WdColor; safecall;
    procedure Set_DiacriticColor(prop: WdColor); safecall;
    function Get_UnderlineColor: WdColor; safecall;
    procedure Set_UnderlineColor(prop: WdColor); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Duplicate: Font read Get_Duplicate;
    property Bold: Integer read Get_Bold write Set_Bold;
    property Italic: Integer read Get_Italic write Set_Italic;
    property Hidden: Integer read Get_Hidden write Set_Hidden;
    property SmallCaps: Integer read Get_SmallCaps write Set_SmallCaps;
    property AllCaps: Integer read Get_AllCaps write Set_AllCaps;
    property StrikeThrough: Integer read Get_StrikeThrough write Set_StrikeThrough;
    property DoubleStrikeThrough: Integer read Get_DoubleStrikeThrough write Set_DoubleStrikeThrough;
    property ColorIndex: WdColorIndex read Get_ColorIndex write Set_ColorIndex;
    property Subscript: Integer read Get_Subscript write Set_Subscript;
    property Superscript: Integer read Get_Superscript write Set_Superscript;
    property Underline: WdUnderline read Get_Underline write Set_Underline;
    property Size: Single read Get_Size write Set_Size;
    property Name: WideString read Get_Name write Set_Name;
    property Position: Integer read Get_Position write Set_Position;
    property Spacing: Single read Get_Spacing write Set_Spacing;
    property Scaling: Integer read Get_Scaling write Set_Scaling;
    property Shadow: Integer read Get_Shadow write Set_Shadow;
    property Outline: Integer read Get_Outline write Set_Outline;
    property Emboss: Integer read Get_Emboss write Set_Emboss;
    property Kerning: Single read Get_Kerning write Set_Kerning;
    property Engrave: Integer read Get_Engrave write Set_Engrave;
    property Animation: WdAnimation read Get_Animation write Set_Animation;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property EmphasisMark: WdEmphasisMark read Get_EmphasisMark write Set_EmphasisMark;
    property DisableCharacterSpaceGrid: WordBool read Get_DisableCharacterSpaceGrid write Set_DisableCharacterSpaceGrid;
    property NameFarEast: WideString read Get_NameFarEast write Set_NameFarEast;
    property NameAscii: WideString read Get_NameAscii write Set_NameAscii;
    property NameOther: WideString read Get_NameOther write Set_NameOther;
    property Color: WdColor read Get_Color write Set_Color;
    property BoldBi: Integer read Get_BoldBi write Set_BoldBi;
    property ItalicBi: Integer read Get_ItalicBi write Set_ItalicBi;
    property SizeBi: Single read Get_SizeBi write Set_SizeBi;
    property NameBi: WideString read Get_NameBi write Set_NameBi;
    property ColorIndexBi: WdColorIndex read Get_ColorIndexBi write Set_ColorIndexBi;
    property DiacriticColor: WdColor read Get_DiacriticColor write Set_DiacriticColor;
    property UnderlineColor: WdColor read Get_UnderlineColor write Set_UnderlineColor;
  end;

// *********************************************************************//
// DispIntf:  _FontDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020952-0000-0000-C000-000000000046}
// *********************************************************************//
  _FontDisp = dispinterface
    ['{00020952-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Duplicate: Font readonly dispid 10;
    property Bold: Integer dispid 130;
    property Italic: Integer dispid 131;
    property Hidden: Integer dispid 132;
    property SmallCaps: Integer dispid 133;
    property AllCaps: Integer dispid 134;
    property StrikeThrough: Integer dispid 135;
    property DoubleStrikeThrough: Integer dispid 136;
    property ColorIndex: WdColorIndex dispid 137;
    property Subscript: Integer dispid 138;
    property Superscript: Integer dispid 139;
    property Underline: WdUnderline dispid 140;
    property Size: Single dispid 141;
    property Name: WideString dispid 142;
    property Position: Integer dispid 143;
    property Spacing: Single dispid 144;
    property Scaling: Integer dispid 145;
    property Shadow: Integer dispid 146;
    property Outline: Integer dispid 147;
    property Emboss: Integer dispid 148;
    property Kerning: Single dispid 149;
    property Engrave: Integer dispid 150;
    property Animation: WdAnimation dispid 151;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 153;
    property EmphasisMark: WdEmphasisMark dispid 154;
    property DisableCharacterSpaceGrid: WordBool dispid 155;
    property NameFarEast: WideString dispid 156;
    property NameAscii: WideString dispid 157;
    property NameOther: WideString dispid 158;
    procedure Grow; dispid 100;
    procedure Shrink; dispid 101;
    procedure Reset; dispid 102;
    procedure SetAsTemplateDefault; dispid 103;
    property Color: WdColor dispid 159;
    property BoldBi: Integer dispid 160;
    property ItalicBi: Integer dispid 161;
    property SizeBi: Single dispid 162;
    property NameBi: WideString dispid 163;
    property ColorIndexBi: WdColorIndex dispid 164;
    property DiacriticColor: WdColor dispid 165;
    property UnderlineColor: WdColor dispid 166;
  end;

// *********************************************************************//
// Interface: Table
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020951-0000-0000-C000-000000000046}
// *********************************************************************//
  Table = interface(IDispatch)
    ['{00020951-0000-0000-C000-000000000046}']
    function Get_Range: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Columns: Columns; safecall;
    function Get_Rows: Rows; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_Uniform: WordBool; safecall;
    function Get_AutoFormatType: Integer; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SortOld(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                      var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                      var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                      var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                      var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                      var CaseSensitive: OleVariant; var LanguageID: OleVariant); safecall;
    procedure SortAscending; safecall;
    procedure SortDescending; safecall;
    procedure AutoFormat(var Format: OleVariant; var ApplyBorders: OleVariant; 
                         var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                         var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                         var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                         var ApplyLastColumn: OleVariant; var AutoFit: OleVariant); safecall;
    procedure UpdateAutoFormat; safecall;
    function ConvertToTextOld(var Separator: OleVariant): Range; safecall;
    function Cell(Row: Integer; Column: Integer): Cell; safecall;
    function Split(var BeforeRow: OleVariant): Table; safecall;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; safecall;
    procedure AutoFitBehavior(Behavior: WdAutoFitBehavior); safecall;
    procedure Sort(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                   var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                   var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                   var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                   var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                   var CaseSensitive: OleVariant; var BidiSort: OleVariant; 
                   var IgnoreThe: OleVariant; var IgnoreKashida: OleVariant; 
                   var IgnoreDiacritics: OleVariant; var IgnoreHe: OleVariant; 
                   var LanguageID: OleVariant); safecall;
    function Get_Tables: Tables; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_AllowPageBreaks: WordBool; safecall;
    procedure Set_AllowPageBreaks(prop: WordBool); safecall;
    function Get_AllowAutoFit: WordBool; safecall;
    procedure Set_AllowAutoFit(prop: WordBool); safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WdPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WdPreferredWidthType); safecall;
    function Get_TopPadding: Single; safecall;
    procedure Set_TopPadding(prop: Single); safecall;
    function Get_BottomPadding: Single; safecall;
    procedure Set_BottomPadding(prop: Single); safecall;
    function Get_LeftPadding: Single; safecall;
    procedure Set_LeftPadding(prop: Single); safecall;
    function Get_RightPadding: Single; safecall;
    procedure Set_RightPadding(prop: Single); safecall;
    function Get_Spacing: Single; safecall;
    procedure Set_Spacing(prop: Single); safecall;
    function Get_TableDirection: WdTableDirection; safecall;
    procedure Set_TableDirection(prop: WdTableDirection); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    property Range: Range read Get_Range;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Columns: Columns read Get_Columns;
    property Rows: Rows read Get_Rows;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property Uniform: WordBool read Get_Uniform;
    property AutoFormatType: Integer read Get_AutoFormatType;
    property Tables: Tables read Get_Tables;
    property NestingLevel: Integer read Get_NestingLevel;
    property AllowPageBreaks: WordBool read Get_AllowPageBreaks write Set_AllowPageBreaks;
    property AllowAutoFit: WordBool read Get_AllowAutoFit write Set_AllowAutoFit;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WdPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
    property TopPadding: Single read Get_TopPadding write Set_TopPadding;
    property BottomPadding: Single read Get_BottomPadding write Set_BottomPadding;
    property LeftPadding: Single read Get_LeftPadding write Set_LeftPadding;
    property RightPadding: Single read Get_RightPadding write Set_RightPadding;
    property Spacing: Single read Get_Spacing write Set_Spacing;
    property TableDirection: WdTableDirection read Get_TableDirection write Set_TableDirection;
    property ID: WideString read Get_ID write Set_ID;
  end;

// *********************************************************************//
// DispIntf:  TableDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020951-0000-0000-C000-000000000046}
// *********************************************************************//
  TableDisp = dispinterface
    ['{00020951-0000-0000-C000-000000000046}']
    property Range: Range readonly dispid 0;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Columns: Columns readonly dispid 100;
    property Rows: Rows readonly dispid 101;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 104;
    property Uniform: WordBool readonly dispid 105;
    property AutoFormatType: Integer readonly dispid 106;
    procedure Select; dispid 200;
    procedure Delete; dispid 9;
    procedure SortOld(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                      var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                      var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                      var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                      var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                      var CaseSensitive: OleVariant; var LanguageID: OleVariant); dispid 10;
    procedure SortAscending; dispid 12;
    procedure SortDescending; dispid 13;
    procedure AutoFormat(var Format: OleVariant; var ApplyBorders: OleVariant; 
                         var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                         var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                         var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                         var ApplyLastColumn: OleVariant; var AutoFit: OleVariant); dispid 14;
    procedure UpdateAutoFormat; dispid 15;
    function ConvertToTextOld(var Separator: OleVariant): Range; dispid 16;
    function Cell(Row: Integer; Column: Integer): Cell; dispid 17;
    function Split(var BeforeRow: OleVariant): Table; dispid 18;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; dispid 19;
    procedure AutoFitBehavior(Behavior: WdAutoFitBehavior); dispid 20;
    procedure Sort(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                   var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                   var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                   var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                   var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                   var CaseSensitive: OleVariant; var BidiSort: OleVariant; 
                   var IgnoreThe: OleVariant; var IgnoreKashida: OleVariant; 
                   var IgnoreDiacritics: OleVariant; var IgnoreHe: OleVariant; 
                   var LanguageID: OleVariant); dispid 23;
    property Tables: Tables readonly dispid 107;
    property NestingLevel: Integer readonly dispid 108;
    property AllowPageBreaks: WordBool dispid 109;
    property AllowAutoFit: WordBool dispid 110;
    property PreferredWidth: Single dispid 111;
    property PreferredWidthType: WdPreferredWidthType dispid 112;
    property TopPadding: Single dispid 113;
    property BottomPadding: Single dispid 114;
    property LeftPadding: Single dispid 115;
    property RightPadding: Single dispid 116;
    property Spacing: Single dispid 117;
    property TableDirection: WdTableDirection dispid 118;
    property ID: WideString dispid 119;
  end;

// *********************************************************************//
// Interface: Row
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020950-0000-0000-C000-000000000046}
// *********************************************************************//
  Row = interface(IDispatch)
    ['{00020950-0000-0000-C000-000000000046}']
    function Get_Range: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_AllowBreakAcrossPages: Integer; safecall;
    procedure Set_AllowBreakAcrossPages(prop: Integer); safecall;
    function Get_Alignment: WdRowAlignment; safecall;
    procedure Set_Alignment(prop: WdRowAlignment); safecall;
    function Get_HeadingFormat: Integer; safecall;
    procedure Set_HeadingFormat(prop: Integer); safecall;
    function Get_SpaceBetweenColumns: Single; safecall;
    procedure Set_SpaceBetweenColumns(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WdRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WdRowHeightRule); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_IsLast: WordBool; safecall;
    function Get_IsFirst: WordBool; safecall;
    function Get_Index: Integer; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_Next: Row; safecall;
    function Get_Previous: Row; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WdRulerStyle); safecall;
    procedure SetHeight(RowHeight: Single; HeightRule: WdRowHeightRule); safecall;
    function ConvertToTextOld(var Separator: OleVariant): Range; safecall;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    property Range: Range read Get_Range;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property AllowBreakAcrossPages: Integer read Get_AllowBreakAcrossPages write Set_AllowBreakAcrossPages;
    property Alignment: WdRowAlignment read Get_Alignment write Set_Alignment;
    property HeadingFormat: Integer read Get_HeadingFormat write Set_HeadingFormat;
    property SpaceBetweenColumns: Single read Get_SpaceBetweenColumns write Set_SpaceBetweenColumns;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WdRowHeightRule read Get_HeightRule write Set_HeightRule;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property IsLast: WordBool read Get_IsLast;
    property IsFirst: WordBool read Get_IsFirst;
    property Index: Integer read Get_Index;
    property Cells: Cells read Get_Cells;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property Next: Row read Get_Next;
    property Previous: Row read Get_Previous;
    property NestingLevel: Integer read Get_NestingLevel;
    property ID: WideString read Get_ID write Set_ID;
  end;

// *********************************************************************//
// DispIntf:  RowDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020950-0000-0000-C000-000000000046}
// *********************************************************************//
  RowDisp = dispinterface
    ['{00020950-0000-0000-C000-000000000046}']
    property Range: Range readonly dispid 0;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property AllowBreakAcrossPages: Integer dispid 3;
    property Alignment: WdRowAlignment dispid 4;
    property HeadingFormat: Integer dispid 5;
    property SpaceBetweenColumns: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WdRowHeightRule dispid 8;
    property LeftIndent: Single dispid 9;
    property IsLast: WordBool readonly dispid 10;
    property IsFirst: WordBool readonly dispid 11;
    property Index: Integer readonly dispid 12;
    property Cells: Cells readonly dispid 100;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 103;
    property Next: Row readonly dispid 104;
    property Previous: Row readonly dispid 105;
    procedure Select; dispid 65535;
    procedure Delete; dispid 200;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WdRulerStyle); dispid 202;
    procedure SetHeight(RowHeight: Single; HeightRule: WdRowHeightRule); dispid 203;
    function ConvertToTextOld(var Separator: OleVariant): Range; dispid 16;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; dispid 18;
    property NestingLevel: Integer readonly dispid 106;
    property ID: WideString dispid 107;
  end;

// *********************************************************************//
// Interface: Column
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094F-0000-0000-C000-000000000046}
// *********************************************************************//
  Column = interface(IDispatch)
    ['{0002094F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_IsFirst: WordBool; safecall;
    function Get_IsLast: WordBool; safecall;
    function Get_Index: Integer; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_Next: Column; safecall;
    function Get_Previous: Column; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); safecall;
    procedure AutoFit; safecall;
    procedure SortOld(var ExcludeHeader: OleVariant; var SortFieldType: OleVariant; 
                      var SortOrder: OleVariant; var CaseSensitive: OleVariant; 
                      var LanguageID: OleVariant); safecall;
    procedure Sort(var ExcludeHeader: OleVariant; var SortFieldType: OleVariant; 
                   var SortOrder: OleVariant; var CaseSensitive: OleVariant; 
                   var BidiSort: OleVariant; var IgnoreThe: OleVariant; 
                   var IgnoreKashida: OleVariant; var IgnoreDiacritics: OleVariant; 
                   var IgnoreHe: OleVariant; var LanguageID: OleVariant); safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WdPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WdPreferredWidthType); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Width: Single read Get_Width write Set_Width;
    property IsFirst: WordBool read Get_IsFirst;
    property IsLast: WordBool read Get_IsLast;
    property Index: Integer read Get_Index;
    property Cells: Cells read Get_Cells;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property Next: Column read Get_Next;
    property Previous: Column read Get_Previous;
    property NestingLevel: Integer read Get_NestingLevel;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WdPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
  end;

// *********************************************************************//
// DispIntf:  ColumnDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094F-0000-0000-C000-000000000046}
// *********************************************************************//
  ColumnDisp = dispinterface
    ['{0002094F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Width: Single dispid 3;
    property IsFirst: WordBool readonly dispid 4;
    property IsLast: WordBool readonly dispid 5;
    property Index: Integer readonly dispid 6;
    property Cells: Cells readonly dispid 100;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 102;
    property Next: Column readonly dispid 103;
    property Previous: Column readonly dispid 104;
    procedure Select; dispid 65535;
    procedure Delete; dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); dispid 201;
    procedure AutoFit; dispid 202;
    procedure SortOld(var ExcludeHeader: OleVariant; var SortFieldType: OleVariant; 
                      var SortOrder: OleVariant; var CaseSensitive: OleVariant; 
                      var LanguageID: OleVariant); dispid 203;
    procedure Sort(var ExcludeHeader: OleVariant; var SortFieldType: OleVariant; 
                   var SortOrder: OleVariant; var CaseSensitive: OleVariant; 
                   var BidiSort: OleVariant; var IgnoreThe: OleVariant; 
                   var IgnoreKashida: OleVariant; var IgnoreDiacritics: OleVariant; 
                   var IgnoreHe: OleVariant; var LanguageID: OleVariant); dispid 204;
    property NestingLevel: Integer readonly dispid 105;
    property PreferredWidth: Single dispid 106;
    property PreferredWidthType: WdPreferredWidthType dispid 107;
  end;

// *********************************************************************//
// Interface: Cell
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094E-0000-0000-C000-000000000046}
// *********************************************************************//
  Cell = interface(IDispatch)
    ['{0002094E-0000-0000-C000-000000000046}']
    function Get_Range: Range; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_RowIndex: Integer; safecall;
    function Get_ColumnIndex: Integer; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WdRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WdRowHeightRule); safecall;
    function Get_VerticalAlignment: WdCellVerticalAlignment; safecall;
    procedure Set_VerticalAlignment(prop: WdCellVerticalAlignment); safecall;
    function Get_Column: Column; safecall;
    function Get_Row: Row; safecall;
    function Get_Next: Cell; safecall;
    function Get_Previous: Cell; safecall;
    function Get_Shading: Shading; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    procedure Select; safecall;
    procedure Delete(var ShiftCells: OleVariant); safecall;
    procedure Formula(var Formula: OleVariant; var NumFormat: OleVariant); safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); safecall;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WdRowHeightRule); safecall;
    procedure Merge(const MergeTo: Cell); safecall;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant); safecall;
    procedure AutoSum; safecall;
    function Get_Tables: Tables; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_WordWrap: WordBool; safecall;
    procedure Set_WordWrap(prop: WordBool); safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_FitText: WordBool; safecall;
    procedure Set_FitText(prop: WordBool); safecall;
    function Get_TopPadding: Single; safecall;
    procedure Set_TopPadding(prop: Single); safecall;
    function Get_BottomPadding: Single; safecall;
    procedure Set_BottomPadding(prop: Single); safecall;
    function Get_LeftPadding: Single; safecall;
    procedure Set_LeftPadding(prop: Single); safecall;
    function Get_RightPadding: Single; safecall;
    procedure Set_RightPadding(prop: Single); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    function Get_PreferredWidthType: WdPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WdPreferredWidthType); safecall;
    property Range: Range read Get_Range;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property RowIndex: Integer read Get_RowIndex;
    property ColumnIndex: Integer read Get_ColumnIndex;
    property Width: Single read Get_Width write Set_Width;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WdRowHeightRule read Get_HeightRule write Set_HeightRule;
    property VerticalAlignment: WdCellVerticalAlignment read Get_VerticalAlignment write Set_VerticalAlignment;
    property Column: Column read Get_Column;
    property Row: Row read Get_Row;
    property Next: Cell read Get_Next;
    property Previous: Cell read Get_Previous;
    property Shading: Shading read Get_Shading;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Tables: Tables read Get_Tables;
    property NestingLevel: Integer read Get_NestingLevel;
    property WordWrap: WordBool read Get_WordWrap write Set_WordWrap;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property FitText: WordBool read Get_FitText write Set_FitText;
    property TopPadding: Single read Get_TopPadding write Set_TopPadding;
    property BottomPadding: Single read Get_BottomPadding write Set_BottomPadding;
    property LeftPadding: Single read Get_LeftPadding write Set_LeftPadding;
    property RightPadding: Single read Get_RightPadding write Set_RightPadding;
    property ID: WideString read Get_ID write Set_ID;
    property PreferredWidthType: WdPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
  end;

// *********************************************************************//
// DispIntf:  CellDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094E-0000-0000-C000-000000000046}
// *********************************************************************//
  CellDisp = dispinterface
    ['{0002094E-0000-0000-C000-000000000046}']
    property Range: Range readonly dispid 0;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property RowIndex: Integer readonly dispid 4;
    property ColumnIndex: Integer readonly dispid 5;
    property Width: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WdRowHeightRule dispid 8;
    property VerticalAlignment: WdCellVerticalAlignment dispid 1104;
    property Column: Column readonly dispid 101;
    property Row: Row readonly dispid 102;
    property Next: Cell readonly dispid 103;
    property Previous: Cell readonly dispid 104;
    property Shading: Shading readonly dispid 105;
    property Borders: Borders dispid 1100;
    procedure Select; dispid 65535;
    procedure Delete(var ShiftCells: OleVariant); dispid 200;
    procedure Formula(var Formula: OleVariant; var NumFormat: OleVariant); dispid 201;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); dispid 202;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WdRowHeightRule); dispid 203;
    procedure Merge(const MergeTo: Cell); dispid 204;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant); dispid 205;
    procedure AutoSum; dispid 206;
    property Tables: Tables readonly dispid 106;
    property NestingLevel: Integer readonly dispid 107;
    property WordWrap: WordBool dispid 108;
    property PreferredWidth: Single dispid 109;
    property FitText: WordBool dispid 110;
    property TopPadding: Single dispid 111;
    property BottomPadding: Single dispid 112;
    property LeftPadding: Single dispid 113;
    property RightPadding: Single dispid 114;
    property ID: WideString dispid 115;
    property PreferredWidthType: WdPreferredWidthType dispid 116;
  end;

// *********************************************************************//
// Interface: Tables
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094D-0000-0000-C000-000000000046}
// *********************************************************************//
  Tables = interface(IDispatch)
    ['{0002094D-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): Table; safecall;
    function AddOld(const Range: Range; NumRows: Integer; NumColumns: Integer): Table; safecall;
    function Add(const Range: Range; NumRows: Integer; NumColumns: Integer; 
                 var DefaultTableBehavior: OleVariant; var AutoFitBehavior: OleVariant): Table; safecall;
    function Get_NestingLevel: Integer; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property NestingLevel: Integer read Get_NestingLevel;
  end;

// *********************************************************************//
// DispIntf:  TablesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094D-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesDisp = dispinterface
    ['{0002094D-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): Table; dispid 0;
    function AddOld(const Range: Range; NumRows: Integer; NumColumns: Integer): Table; dispid 4;
    function Add(const Range: Range; NumRows: Integer; NumColumns: Integer; 
                 var DefaultTableBehavior: OleVariant; var AutoFitBehavior: OleVariant): Table; dispid 200;
    property NestingLevel: Integer readonly dispid 100;
  end;

// *********************************************************************//
// Interface: Rows
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094C-0000-0000-C000-000000000046}
// *********************************************************************//
  Rows = interface(IDispatch)
    ['{0002094C-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_AllowBreakAcrossPages: Integer; safecall;
    procedure Set_AllowBreakAcrossPages(prop: Integer); safecall;
    function Get_Alignment: WdRowAlignment; safecall;
    procedure Set_Alignment(prop: WdRowAlignment); safecall;
    function Get_HeadingFormat: Integer; safecall;
    procedure Set_HeadingFormat(prop: Integer); safecall;
    function Get_SpaceBetweenColumns: Single; safecall;
    procedure Set_SpaceBetweenColumns(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WdRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WdRowHeightRule); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_First: Row; safecall;
    function Get_Last: Row; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Item(Index: Integer): Row; safecall;
    function Add(var BeforeRow: OleVariant): Row; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WdRulerStyle); safecall;
    procedure SetHeight(RowHeight: Single; HeightRule: WdRowHeightRule); safecall;
    function ConvertToTextOld(var Separator: OleVariant): Range; safecall;
    procedure DistributeHeight; safecall;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; safecall;
    function Get_WrapAroundText: Integer; safecall;
    procedure Set_WrapAroundText(prop: Integer); safecall;
    function Get_DistanceTop: Single; safecall;
    procedure Set_DistanceTop(prop: Single); safecall;
    function Get_DistanceBottom: Single; safecall;
    procedure Set_DistanceBottom(prop: Single); safecall;
    function Get_DistanceLeft: Single; safecall;
    procedure Set_DistanceLeft(prop: Single); safecall;
    function Get_DistanceRight: Single; safecall;
    procedure Set_DistanceRight(prop: Single); safecall;
    function Get_HorizontalPosition: Single; safecall;
    procedure Set_HorizontalPosition(prop: Single); safecall;
    function Get_VerticalPosition: Single; safecall;
    procedure Set_VerticalPosition(prop: Single); safecall;
    function Get_RelativeHorizontalPosition: WdRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WdRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WdRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WdRelativeVerticalPosition); safecall;
    function Get_AllowOverlap: Integer; safecall;
    procedure Set_AllowOverlap(prop: Integer); safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_TableDirection: WdTableDirection; safecall;
    procedure Set_TableDirection(prop: WdTableDirection); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property AllowBreakAcrossPages: Integer read Get_AllowBreakAcrossPages write Set_AllowBreakAcrossPages;
    property Alignment: WdRowAlignment read Get_Alignment write Set_Alignment;
    property HeadingFormat: Integer read Get_HeadingFormat write Set_HeadingFormat;
    property SpaceBetweenColumns: Single read Get_SpaceBetweenColumns write Set_SpaceBetweenColumns;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WdRowHeightRule read Get_HeightRule write Set_HeightRule;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property First: Row read Get_First;
    property Last: Row read Get_Last;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property WrapAroundText: Integer read Get_WrapAroundText write Set_WrapAroundText;
    property DistanceTop: Single read Get_DistanceTop write Set_DistanceTop;
    property DistanceBottom: Single read Get_DistanceBottom write Set_DistanceBottom;
    property DistanceLeft: Single read Get_DistanceLeft write Set_DistanceLeft;
    property DistanceRight: Single read Get_DistanceRight write Set_DistanceRight;
    property HorizontalPosition: Single read Get_HorizontalPosition write Set_HorizontalPosition;
    property VerticalPosition: Single read Get_VerticalPosition write Set_VerticalPosition;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WdRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property AllowOverlap: Integer read Get_AllowOverlap write Set_AllowOverlap;
    property NestingLevel: Integer read Get_NestingLevel;
    property TableDirection: WdTableDirection read Get_TableDirection write Set_TableDirection;
  end;

// *********************************************************************//
// DispIntf:  RowsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094C-0000-0000-C000-000000000046}
// *********************************************************************//
  RowsDisp = dispinterface
    ['{0002094C-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property AllowBreakAcrossPages: Integer dispid 3;
    property Alignment: WdRowAlignment dispid 4;
    property HeadingFormat: Integer dispid 5;
    property SpaceBetweenColumns: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WdRowHeightRule dispid 8;
    property LeftIndent: Single dispid 9;
    property First: Row readonly dispid 10;
    property Last: Row readonly dispid 11;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 102;
    function Item(Index: Integer): Row; dispid 0;
    function Add(var BeforeRow: OleVariant): Row; dispid 100;
    procedure Select; dispid 199;
    procedure Delete; dispid 200;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WdRulerStyle); dispid 202;
    procedure SetHeight(RowHeight: Single; HeightRule: WdRowHeightRule); dispid 203;
    function ConvertToTextOld(var Separator: OleVariant): Range; dispid 16;
    procedure DistributeHeight; dispid 206;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; dispid 210;
    property WrapAroundText: Integer dispid 12;
    property DistanceTop: Single dispid 13;
    property DistanceBottom: Single dispid 14;
    property DistanceLeft: Single dispid 20;
    property DistanceRight: Single dispid 21;
    property HorizontalPosition: Single dispid 15;
    property VerticalPosition: Single dispid 17;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition dispid 18;
    property RelativeVerticalPosition: WdRelativeVerticalPosition dispid 19;
    property AllowOverlap: Integer dispid 22;
    property NestingLevel: Integer readonly dispid 103;
    property TableDirection: WdTableDirection dispid 104;
  end;

// *********************************************************************//
// Interface: Columns
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094B-0000-0000-C000-000000000046}
// *********************************************************************//
  Columns = interface(IDispatch)
    ['{0002094B-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_First: Column; safecall;
    function Get_Last: Column; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Item(Index: Integer): Column; safecall;
    function Add(var BeforeColumn: OleVariant): Column; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); safecall;
    procedure AutoFit; safecall;
    procedure DistributeWidth; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WdPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WdPreferredWidthType); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property First: Column read Get_First;
    property Last: Column read Get_Last;
    property Width: Single read Get_Width write Set_Width;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property NestingLevel: Integer read Get_NestingLevel;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WdPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
  end;

// *********************************************************************//
// DispIntf:  ColumnsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094B-0000-0000-C000-000000000046}
// *********************************************************************//
  ColumnsDisp = dispinterface
    ['{0002094B-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property First: Column readonly dispid 100;
    property Last: Column readonly dispid 101;
    property Width: Single dispid 3;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 103;
    function Item(Index: Integer): Column; dispid 0;
    function Add(var BeforeColumn: OleVariant): Column; dispid 5;
    procedure Select; dispid 199;
    procedure Delete; dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); dispid 201;
    procedure AutoFit; dispid 202;
    procedure DistributeWidth; dispid 203;
    property NestingLevel: Integer readonly dispid 104;
    property PreferredWidth: Single dispid 105;
    property PreferredWidthType: WdPreferredWidthType dispid 106;
  end;

// *********************************************************************//
// Interface: Cells
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094A-0000-0000-C000-000000000046}
// *********************************************************************//
  Cells = interface(IDispatch)
    ['{0002094A-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WdRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WdRowHeightRule); safecall;
    function Get_VerticalAlignment: WdCellVerticalAlignment; safecall;
    procedure Set_VerticalAlignment(prop: WdCellVerticalAlignment); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Item(Index: Integer): Cell; safecall;
    function Add(var BeforeCell: OleVariant): Cell; safecall;
    procedure Delete(var ShiftCells: OleVariant); safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); safecall;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WdRowHeightRule); safecall;
    procedure Merge; safecall;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant; 
                    var MergeBeforeSplit: OleVariant); safecall;
    procedure DistributeHeight; safecall;
    procedure DistributeWidth; safecall;
    procedure AutoFit; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WdPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WdPreferredWidthType); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Width: Single read Get_Width write Set_Width;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WdRowHeightRule read Get_HeightRule write Set_HeightRule;
    property VerticalAlignment: WdCellVerticalAlignment read Get_VerticalAlignment write Set_VerticalAlignment;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property NestingLevel: Integer read Get_NestingLevel;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WdPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
  end;

// *********************************************************************//
// DispIntf:  CellsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094A-0000-0000-C000-000000000046}
// *********************************************************************//
  CellsDisp = dispinterface
    ['{0002094A-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Width: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WdRowHeightRule dispid 8;
    property VerticalAlignment: WdCellVerticalAlignment dispid 1104;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 101;
    function Item(Index: Integer): Cell; dispid 0;
    function Add(var BeforeCell: OleVariant): Cell; dispid 4;
    procedure Delete(var ShiftCells: OleVariant); dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WdRulerStyle); dispid 202;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WdRowHeightRule); dispid 203;
    procedure Merge; dispid 204;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant; 
                    var MergeBeforeSplit: OleVariant); dispid 205;
    procedure DistributeHeight; dispid 206;
    procedure DistributeWidth; dispid 207;
    procedure AutoFit; dispid 208;
    property NestingLevel: Integer readonly dispid 102;
    property PreferredWidth: Single dispid 103;
    property PreferredWidthType: WdPreferredWidthType dispid 104;
  end;

// *********************************************************************//
// Interface: AutoCorrect
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020949-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCorrect = interface(IDispatch)
    ['{00020949-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_CorrectDays: WordBool; safecall;
    procedure Set_CorrectDays(prop: WordBool); safecall;
    function Get_CorrectInitialCaps: WordBool; safecall;
    procedure Set_CorrectInitialCaps(prop: WordBool); safecall;
    function Get_CorrectSentenceCaps: WordBool; safecall;
    procedure Set_CorrectSentenceCaps(prop: WordBool); safecall;
    function Get_ReplaceText: WordBool; safecall;
    procedure Set_ReplaceText(prop: WordBool); safecall;
    function Get_Entries: AutoCorrectEntries; safecall;
    function Get_FirstLetterExceptions: FirstLetterExceptions; safecall;
    function Get_FirstLetterAutoAdd: WordBool; safecall;
    procedure Set_FirstLetterAutoAdd(prop: WordBool); safecall;
    function Get_TwoInitialCapsExceptions: TwoInitialCapsExceptions; safecall;
    function Get_TwoInitialCapsAutoAdd: WordBool; safecall;
    procedure Set_TwoInitialCapsAutoAdd(prop: WordBool); safecall;
    function Get_CorrectCapsLock: WordBool; safecall;
    procedure Set_CorrectCapsLock(prop: WordBool); safecall;
    function Get_CorrectHangulAndAlphabet: WordBool; safecall;
    procedure Set_CorrectHangulAndAlphabet(prop: WordBool); safecall;
    function Get_HangulAndAlphabetExceptions: HangulAndAlphabetExceptions; safecall;
    function Get_HangulAndAlphabetAutoAdd: WordBool; safecall;
    procedure Set_HangulAndAlphabetAutoAdd(prop: WordBool); safecall;
    function Get_ReplaceTextFromSpellingChecker: WordBool; safecall;
    procedure Set_ReplaceTextFromSpellingChecker(prop: WordBool); safecall;
    function Get_OtherCorrectionsAutoAdd: WordBool; safecall;
    procedure Set_OtherCorrectionsAutoAdd(prop: WordBool); safecall;
    function Get_OtherCorrectionsExceptions: OtherCorrectionsExceptions; safecall;
    function Get_CorrectKeyboardSetting: WordBool; safecall;
    procedure Set_CorrectKeyboardSetting(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property CorrectDays: WordBool read Get_CorrectDays write Set_CorrectDays;
    property CorrectInitialCaps: WordBool read Get_CorrectInitialCaps write Set_CorrectInitialCaps;
    property CorrectSentenceCaps: WordBool read Get_CorrectSentenceCaps write Set_CorrectSentenceCaps;
    property ReplaceText: WordBool read Get_ReplaceText write Set_ReplaceText;
    property Entries: AutoCorrectEntries read Get_Entries;
    property FirstLetterExceptions: FirstLetterExceptions read Get_FirstLetterExceptions;
    property FirstLetterAutoAdd: WordBool read Get_FirstLetterAutoAdd write Set_FirstLetterAutoAdd;
    property TwoInitialCapsExceptions: TwoInitialCapsExceptions read Get_TwoInitialCapsExceptions;
    property TwoInitialCapsAutoAdd: WordBool read Get_TwoInitialCapsAutoAdd write Set_TwoInitialCapsAutoAdd;
    property CorrectCapsLock: WordBool read Get_CorrectCapsLock write Set_CorrectCapsLock;
    property CorrectHangulAndAlphabet: WordBool read Get_CorrectHangulAndAlphabet write Set_CorrectHangulAndAlphabet;
    property HangulAndAlphabetExceptions: HangulAndAlphabetExceptions read Get_HangulAndAlphabetExceptions;
    property HangulAndAlphabetAutoAdd: WordBool read Get_HangulAndAlphabetAutoAdd write Set_HangulAndAlphabetAutoAdd;
    property ReplaceTextFromSpellingChecker: WordBool read Get_ReplaceTextFromSpellingChecker write Set_ReplaceTextFromSpellingChecker;
    property OtherCorrectionsAutoAdd: WordBool read Get_OtherCorrectionsAutoAdd write Set_OtherCorrectionsAutoAdd;
    property OtherCorrectionsExceptions: OtherCorrectionsExceptions read Get_OtherCorrectionsExceptions;
    property CorrectKeyboardSetting: WordBool read Get_CorrectKeyboardSetting write Set_CorrectKeyboardSetting;
  end;

// *********************************************************************//
// DispIntf:  AutoCorrectDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020949-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCorrectDisp = dispinterface
    ['{00020949-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property CorrectDays: WordBool dispid 1;
    property CorrectInitialCaps: WordBool dispid 2;
    property CorrectSentenceCaps: WordBool dispid 3;
    property ReplaceText: WordBool dispid 4;
    property Entries: AutoCorrectEntries readonly dispid 6;
    property FirstLetterExceptions: FirstLetterExceptions readonly dispid 7;
    property FirstLetterAutoAdd: WordBool dispid 8;
    property TwoInitialCapsExceptions: TwoInitialCapsExceptions readonly dispid 9;
    property TwoInitialCapsAutoAdd: WordBool dispid 10;
    property CorrectCapsLock: WordBool dispid 11;
    property CorrectHangulAndAlphabet: WordBool dispid 12;
    property HangulAndAlphabetExceptions: HangulAndAlphabetExceptions readonly dispid 13;
    property HangulAndAlphabetAutoAdd: WordBool dispid 14;
    property ReplaceTextFromSpellingChecker: WordBool dispid 15;
    property OtherCorrectionsAutoAdd: WordBool dispid 16;
    property OtherCorrectionsExceptions: OtherCorrectionsExceptions readonly dispid 17;
    property CorrectKeyboardSetting: WordBool dispid 18;
  end;

// *********************************************************************//
// Interface: AutoCorrectEntries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020948-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCorrectEntries = interface(IDispatch)
    ['{00020948-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): AutoCorrectEntry; safecall;
    function Add(const Name: WideString; const Value: WideString): AutoCorrectEntry; safecall;
    function AddRichText(const Name: WideString; const Range: Range): AutoCorrectEntry; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AutoCorrectEntriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020948-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCorrectEntriesDisp = dispinterface
    ['{00020948-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): AutoCorrectEntry; dispid 0;
    function Add(const Name: WideString; const Value: WideString): AutoCorrectEntry; dispid 101;
    function AddRichText(const Name: WideString; const Range: Range): AutoCorrectEntry; dispid 102;
  end;

// *********************************************************************//
// Interface: AutoCorrectEntry
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020947-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCorrectEntry = interface(IDispatch)
    ['{00020947-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Value: WideString; safecall;
    procedure Set_Value(const prop: WideString); safecall;
    function Get_RichText: WordBool; safecall;
    procedure Delete; safecall;
    procedure Apply(const Range: Range); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name write Set_Name;
    property Value: WideString read Get_Value write Set_Value;
    property RichText: WordBool read Get_RichText;
  end;

// *********************************************************************//
// DispIntf:  AutoCorrectEntryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020947-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCorrectEntryDisp = dispinterface
    ['{00020947-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString dispid 2;
    property Value: WideString dispid 3;
    property RichText: WordBool readonly dispid 4;
    procedure Delete; dispid 101;
    procedure Apply(const Range: Range); dispid 102;
  end;

// *********************************************************************//
// Interface: FirstLetterExceptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020946-0000-0000-C000-000000000046}
// *********************************************************************//
  FirstLetterExceptions = interface(IDispatch)
    ['{00020946-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): FirstLetterException; safecall;
    function Add(const Name: WideString): FirstLetterException; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  FirstLetterExceptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020946-0000-0000-C000-000000000046}
// *********************************************************************//
  FirstLetterExceptionsDisp = dispinterface
    ['{00020946-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): FirstLetterException; dispid 0;
    function Add(const Name: WideString): FirstLetterException; dispid 101;
  end;

// *********************************************************************//
// Interface: FirstLetterException
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020945-0000-0000-C000-000000000046}
// *********************************************************************//
  FirstLetterException = interface(IDispatch)
    ['{00020945-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  FirstLetterExceptionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020945-0000-0000-C000-000000000046}
// *********************************************************************//
  FirstLetterExceptionDisp = dispinterface
    ['{00020945-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString readonly dispid 2;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: TwoInitialCapsExceptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020944-0000-0000-C000-000000000046}
// *********************************************************************//
  TwoInitialCapsExceptions = interface(IDispatch)
    ['{00020944-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): TwoInitialCapsException; safecall;
    function Add(const Name: WideString): TwoInitialCapsException; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  TwoInitialCapsExceptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020944-0000-0000-C000-000000000046}
// *********************************************************************//
  TwoInitialCapsExceptionsDisp = dispinterface
    ['{00020944-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): TwoInitialCapsException; dispid 0;
    function Add(const Name: WideString): TwoInitialCapsException; dispid 101;
  end;

// *********************************************************************//
// Interface: TwoInitialCapsException
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020943-0000-0000-C000-000000000046}
// *********************************************************************//
  TwoInitialCapsException = interface(IDispatch)
    ['{00020943-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  TwoInitialCapsExceptionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020943-0000-0000-C000-000000000046}
// *********************************************************************//
  TwoInitialCapsExceptionDisp = dispinterface
    ['{00020943-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString readonly dispid 2;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: Footnotes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020942-0000-0000-C000-000000000046}
// *********************************************************************//
  Footnotes = interface(IDispatch)
    ['{00020942-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Location: WdFootnoteLocation; safecall;
    procedure Set_Location(prop: WdFootnoteLocation); safecall;
    function Get_NumberStyle: WdNoteNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WdNoteNumberStyle); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_NumberingRule: WdNumberingRule; safecall;
    procedure Set_NumberingRule(prop: WdNumberingRule); safecall;
    function Get_Separator: Range; safecall;
    function Get_ContinuationSeparator: Range; safecall;
    function Get_ContinuationNotice: Range; safecall;
    function Item(Index: Integer): Footnote; safecall;
    function Add(const Range: Range; var Reference: OleVariant; var Text: OleVariant): Footnote; safecall;
    procedure Convert; safecall;
    procedure SwapWithEndnotes; safecall;
    procedure ResetSeparator; safecall;
    procedure ResetContinuationSeparator; safecall;
    procedure ResetContinuationNotice; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Location: WdFootnoteLocation read Get_Location write Set_Location;
    property NumberStyle: WdNoteNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property NumberingRule: WdNumberingRule read Get_NumberingRule write Set_NumberingRule;
    property Separator: Range read Get_Separator;
    property ContinuationSeparator: Range read Get_ContinuationSeparator;
    property ContinuationNotice: Range read Get_ContinuationNotice;
  end;

// *********************************************************************//
// DispIntf:  FootnotesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020942-0000-0000-C000-000000000046}
// *********************************************************************//
  FootnotesDisp = dispinterface
    ['{00020942-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Location: WdFootnoteLocation dispid 100;
    property NumberStyle: WdNoteNumberStyle dispid 101;
    property StartingNumber: Integer dispid 102;
    property NumberingRule: WdNumberingRule dispid 103;
    property Separator: Range readonly dispid 104;
    property ContinuationSeparator: Range readonly dispid 105;
    property ContinuationNotice: Range readonly dispid 106;
    function Item(Index: Integer): Footnote; dispid 0;
    function Add(const Range: Range; var Reference: OleVariant; var Text: OleVariant): Footnote; dispid 4;
    procedure Convert; dispid 5;
    procedure SwapWithEndnotes; dispid 6;
    procedure ResetSeparator; dispid 7;
    procedure ResetContinuationSeparator; dispid 8;
    procedure ResetContinuationNotice; dispid 9;
  end;

// *********************************************************************//
// Interface: Endnotes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020941-0000-0000-C000-000000000046}
// *********************************************************************//
  Endnotes = interface(IDispatch)
    ['{00020941-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Location: WdEndnoteLocation; safecall;
    procedure Set_Location(prop: WdEndnoteLocation); safecall;
    function Get_NumberStyle: WdNoteNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WdNoteNumberStyle); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_NumberingRule: WdNumberingRule; safecall;
    procedure Set_NumberingRule(prop: WdNumberingRule); safecall;
    function Get_Separator: Range; safecall;
    function Get_ContinuationSeparator: Range; safecall;
    function Get_ContinuationNotice: Range; safecall;
    function Item(Index: Integer): Endnote; safecall;
    function Add(const Range: Range; var Reference: OleVariant; var Text: OleVariant): Endnote; safecall;
    procedure Convert; safecall;
    procedure SwapWithFootnotes; safecall;
    procedure ResetSeparator; safecall;
    procedure ResetContinuationSeparator; safecall;
    procedure ResetContinuationNotice; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Location: WdEndnoteLocation read Get_Location write Set_Location;
    property NumberStyle: WdNoteNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property NumberingRule: WdNumberingRule read Get_NumberingRule write Set_NumberingRule;
    property Separator: Range read Get_Separator;
    property ContinuationSeparator: Range read Get_ContinuationSeparator;
    property ContinuationNotice: Range read Get_ContinuationNotice;
  end;

// *********************************************************************//
// DispIntf:  EndnotesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020941-0000-0000-C000-000000000046}
// *********************************************************************//
  EndnotesDisp = dispinterface
    ['{00020941-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Location: WdEndnoteLocation dispid 100;
    property NumberStyle: WdNoteNumberStyle dispid 101;
    property StartingNumber: Integer dispid 102;
    property NumberingRule: WdNumberingRule dispid 103;
    property Separator: Range readonly dispid 104;
    property ContinuationSeparator: Range readonly dispid 105;
    property ContinuationNotice: Range readonly dispid 106;
    function Item(Index: Integer): Endnote; dispid 0;
    function Add(const Range: Range; var Reference: OleVariant; var Text: OleVariant): Endnote; dispid 4;
    procedure Convert; dispid 5;
    procedure SwapWithFootnotes; dispid 6;
    procedure ResetSeparator; dispid 7;
    procedure ResetContinuationSeparator; dispid 8;
    procedure ResetContinuationNotice; dispid 9;
  end;

// *********************************************************************//
// Interface: Comments
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020940-0000-0000-C000-000000000046}
// *********************************************************************//
  Comments = interface(IDispatch)
    ['{00020940-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ShowBy: WideString; safecall;
    procedure Set_ShowBy(const prop: WideString); safecall;
    function Item(Index: Integer): Comment; safecall;
    function Add(const Range: Range; var Text: OleVariant): Comment; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ShowBy: WideString read Get_ShowBy write Set_ShowBy;
  end;

// *********************************************************************//
// DispIntf:  CommentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020940-0000-0000-C000-000000000046}
// *********************************************************************//
  CommentsDisp = dispinterface
    ['{00020940-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ShowBy: WideString dispid 1003;
    function Item(Index: Integer): Comment; dispid 0;
    function Add(const Range: Range; var Text: OleVariant): Comment; dispid 4;
  end;

// *********************************************************************//
// Interface: Footnote
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093F-0000-0000-C000-000000000046}
// *********************************************************************//
  Footnote = interface(IDispatch)
    ['{0002093F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Range: Range; safecall;
    function Get_Reference: Range; safecall;
    function Get_Index: Integer; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Range: Range read Get_Range;
    property Reference: Range read Get_Reference;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  FootnoteDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093F-0000-0000-C000-000000000046}
// *********************************************************************//
  FootnoteDisp = dispinterface
    ['{0002093F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Range: Range readonly dispid 4;
    property Reference: Range readonly dispid 5;
    property Index: Integer readonly dispid 6;
    procedure Delete; dispid 10;
  end;

// *********************************************************************//
// Interface: Endnote
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093E-0000-0000-C000-000000000046}
// *********************************************************************//
  Endnote = interface(IDispatch)
    ['{0002093E-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Range: Range; safecall;
    function Get_Reference: Range; safecall;
    function Get_Index: Integer; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Range: Range read Get_Range;
    property Reference: Range read Get_Reference;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  EndnoteDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093E-0000-0000-C000-000000000046}
// *********************************************************************//
  EndnoteDisp = dispinterface
    ['{0002093E-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Range: Range readonly dispid 4;
    property Reference: Range readonly dispid 5;
    property Index: Integer readonly dispid 6;
    procedure Delete; dispid 10;
  end;

// *********************************************************************//
// Interface: Comment
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093D-0000-0000-C000-000000000046}
// *********************************************************************//
  Comment = interface(IDispatch)
    ['{0002093D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Range: Range; safecall;
    function Get_Reference: Range; safecall;
    function Get_Scope: Range; safecall;
    function Get_Index: Integer; safecall;
    function Get_Author: WideString; safecall;
    procedure Set_Author(const prop: WideString); safecall;
    function Get_Initial: WideString; safecall;
    procedure Set_Initial(const prop: WideString); safecall;
    function Get_ShowTip: WordBool; safecall;
    procedure Set_ShowTip(prop: WordBool); safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Range: Range read Get_Range;
    property Reference: Range read Get_Reference;
    property Scope: Range read Get_Scope;
    property Index: Integer read Get_Index;
    property Author: WideString read Get_Author write Set_Author;
    property Initial: WideString read Get_Initial write Set_Initial;
    property ShowTip: WordBool read Get_ShowTip write Set_ShowTip;
  end;

// *********************************************************************//
// DispIntf:  CommentDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093D-0000-0000-C000-000000000046}
// *********************************************************************//
  CommentDisp = dispinterface
    ['{0002093D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Range: Range readonly dispid 1003;
    property Reference: Range readonly dispid 1004;
    property Scope: Range readonly dispid 1005;
    property Index: Integer readonly dispid 1006;
    property Author: WideString dispid 1007;
    property Initial: WideString dispid 1008;
    property ShowTip: WordBool dispid 1009;
    procedure Delete; dispid 10;
  end;

// *********************************************************************//
// Interface: Borders
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093C-0000-0000-C000-000000000046}
// *********************************************************************//
  Borders = interface(IDispatch)
    ['{0002093C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Enable: Integer; safecall;
    procedure Set_Enable(prop: Integer); safecall;
    function Get_DistanceFromTop: Integer; safecall;
    procedure Set_DistanceFromTop(prop: Integer); safecall;
    function Get_Shadow: WordBool; safecall;
    procedure Set_Shadow(prop: WordBool); safecall;
    function Get_InsideLineStyle: WdLineStyle; safecall;
    procedure Set_InsideLineStyle(prop: WdLineStyle); safecall;
    function Get_OutsideLineStyle: WdLineStyle; safecall;
    procedure Set_OutsideLineStyle(prop: WdLineStyle); safecall;
    function Get_InsideLineWidth: WdLineWidth; safecall;
    procedure Set_InsideLineWidth(prop: WdLineWidth); safecall;
    function Get_OutsideLineWidth: WdLineWidth; safecall;
    procedure Set_OutsideLineWidth(prop: WdLineWidth); safecall;
    function Get_InsideColorIndex: WdColorIndex; safecall;
    procedure Set_InsideColorIndex(prop: WdColorIndex); safecall;
    function Get_OutsideColorIndex: WdColorIndex; safecall;
    procedure Set_OutsideColorIndex(prop: WdColorIndex); safecall;
    function Get_DistanceFromLeft: Integer; safecall;
    procedure Set_DistanceFromLeft(prop: Integer); safecall;
    function Get_DistanceFromBottom: Integer; safecall;
    procedure Set_DistanceFromBottom(prop: Integer); safecall;
    function Get_DistanceFromRight: Integer; safecall;
    procedure Set_DistanceFromRight(prop: Integer); safecall;
    function Get_AlwaysInFront: WordBool; safecall;
    procedure Set_AlwaysInFront(prop: WordBool); safecall;
    function Get_SurroundHeader: WordBool; safecall;
    procedure Set_SurroundHeader(prop: WordBool); safecall;
    function Get_SurroundFooter: WordBool; safecall;
    procedure Set_SurroundFooter(prop: WordBool); safecall;
    function Get_JoinBorders: WordBool; safecall;
    procedure Set_JoinBorders(prop: WordBool); safecall;
    function Get_HasHorizontal: WordBool; safecall;
    function Get_HasVertical: WordBool; safecall;
    function Get_DistanceFrom: WdBorderDistanceFrom; safecall;
    procedure Set_DistanceFrom(prop: WdBorderDistanceFrom); safecall;
    function Get_EnableFirstPageInSection: WordBool; safecall;
    procedure Set_EnableFirstPageInSection(prop: WordBool); safecall;
    function Get_EnableOtherPagesInSection: WordBool; safecall;
    procedure Set_EnableOtherPagesInSection(prop: WordBool); safecall;
    function Item(Index: WdBorderType): Border; safecall;
    procedure ApplyPageBordersToAllSections; safecall;
    function Get_InsideColor: WdColor; safecall;
    procedure Set_InsideColor(prop: WdColor); safecall;
    function Get_OutsideColor: WdColor; safecall;
    procedure Set_OutsideColor(prop: WdColor); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Enable: Integer read Get_Enable write Set_Enable;
    property DistanceFromTop: Integer read Get_DistanceFromTop write Set_DistanceFromTop;
    property Shadow: WordBool read Get_Shadow write Set_Shadow;
    property InsideLineStyle: WdLineStyle read Get_InsideLineStyle write Set_InsideLineStyle;
    property OutsideLineStyle: WdLineStyle read Get_OutsideLineStyle write Set_OutsideLineStyle;
    property InsideLineWidth: WdLineWidth read Get_InsideLineWidth write Set_InsideLineWidth;
    property OutsideLineWidth: WdLineWidth read Get_OutsideLineWidth write Set_OutsideLineWidth;
    property InsideColorIndex: WdColorIndex read Get_InsideColorIndex write Set_InsideColorIndex;
    property OutsideColorIndex: WdColorIndex read Get_OutsideColorIndex write Set_OutsideColorIndex;
    property DistanceFromLeft: Integer read Get_DistanceFromLeft write Set_DistanceFromLeft;
    property DistanceFromBottom: Integer read Get_DistanceFromBottom write Set_DistanceFromBottom;
    property DistanceFromRight: Integer read Get_DistanceFromRight write Set_DistanceFromRight;
    property AlwaysInFront: WordBool read Get_AlwaysInFront write Set_AlwaysInFront;
    property SurroundHeader: WordBool read Get_SurroundHeader write Set_SurroundHeader;
    property SurroundFooter: WordBool read Get_SurroundFooter write Set_SurroundFooter;
    property JoinBorders: WordBool read Get_JoinBorders write Set_JoinBorders;
    property HasHorizontal: WordBool read Get_HasHorizontal;
    property HasVertical: WordBool read Get_HasVertical;
    property DistanceFrom: WdBorderDistanceFrom read Get_DistanceFrom write Set_DistanceFrom;
    property EnableFirstPageInSection: WordBool read Get_EnableFirstPageInSection write Set_EnableFirstPageInSection;
    property EnableOtherPagesInSection: WordBool read Get_EnableOtherPagesInSection write Set_EnableOtherPagesInSection;
    property InsideColor: WdColor read Get_InsideColor write Set_InsideColor;
    property OutsideColor: WdColor read Get_OutsideColor write Set_OutsideColor;
  end;

// *********************************************************************//
// DispIntf:  BordersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093C-0000-0000-C000-000000000046}
// *********************************************************************//
  BordersDisp = dispinterface
    ['{0002093C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Enable: Integer dispid 2;
    property DistanceFromTop: Integer dispid 4;
    property Shadow: WordBool dispid 5;
    property InsideLineStyle: WdLineStyle dispid 6;
    property OutsideLineStyle: WdLineStyle dispid 7;
    property InsideLineWidth: WdLineWidth dispid 8;
    property OutsideLineWidth: WdLineWidth dispid 9;
    property InsideColorIndex: WdColorIndex dispid 10;
    property OutsideColorIndex: WdColorIndex dispid 11;
    property DistanceFromLeft: Integer dispid 20;
    property DistanceFromBottom: Integer dispid 21;
    property DistanceFromRight: Integer dispid 22;
    property AlwaysInFront: WordBool dispid 23;
    property SurroundHeader: WordBool dispid 24;
    property SurroundFooter: WordBool dispid 25;
    property JoinBorders: WordBool dispid 26;
    property HasHorizontal: WordBool readonly dispid 27;
    property HasVertical: WordBool readonly dispid 28;
    property DistanceFrom: WdBorderDistanceFrom dispid 29;
    property EnableFirstPageInSection: WordBool dispid 30;
    property EnableOtherPagesInSection: WordBool dispid 31;
    function Item(Index: WdBorderType): Border; dispid 0;
    procedure ApplyPageBordersToAllSections; dispid 2000;
    property InsideColor: WdColor dispid 32;
    property OutsideColor: WdColor dispid 33;
  end;

// *********************************************************************//
// Interface: Border
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093B-0000-0000-C000-000000000046}
// *********************************************************************//
  Border = interface(IDispatch)
    ['{0002093B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    function Get_ColorIndex: WdColorIndex; safecall;
    procedure Set_ColorIndex(prop: WdColorIndex); safecall;
    function Get_Inside: WordBool; safecall;
    function Get_LineStyle: WdLineStyle; safecall;
    procedure Set_LineStyle(prop: WdLineStyle); safecall;
    function Get_LineWidth: WdLineWidth; safecall;
    procedure Set_LineWidth(prop: WdLineWidth); safecall;
    function Get_ArtStyle: WdPageBorderArt; safecall;
    procedure Set_ArtStyle(prop: WdPageBorderArt); safecall;
    function Get_ArtWidth: Integer; safecall;
    procedure Set_ArtWidth(prop: Integer); safecall;
    function Get_Color: WdColor; safecall;
    procedure Set_Color(prop: WdColor); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property ColorIndex: WdColorIndex read Get_ColorIndex write Set_ColorIndex;
    property Inside: WordBool read Get_Inside;
    property LineStyle: WdLineStyle read Get_LineStyle write Set_LineStyle;
    property LineWidth: WdLineWidth read Get_LineWidth write Set_LineWidth;
    property ArtStyle: WdPageBorderArt read Get_ArtStyle write Set_ArtStyle;
    property ArtWidth: Integer read Get_ArtWidth write Set_ArtWidth;
    property Color: WdColor read Get_Color write Set_Color;
  end;

// *********************************************************************//
// DispIntf:  BorderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093B-0000-0000-C000-000000000046}
// *********************************************************************//
  BorderDisp = dispinterface
    ['{0002093B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Visible: WordBool dispid 0;
    property ColorIndex: WdColorIndex dispid 1;
    property Inside: WordBool readonly dispid 2;
    property LineStyle: WdLineStyle dispid 3;
    property LineWidth: WdLineWidth dispid 4;
    property ArtStyle: WdPageBorderArt dispid 5;
    property ArtWidth: Integer dispid 6;
    property Color: WdColor dispid 7;
  end;

// *********************************************************************//
// Interface: Shading
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093A-0000-0000-C000-000000000046}
// *********************************************************************//
  Shading = interface(IDispatch)
    ['{0002093A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ForegroundPatternColorIndex: WdColorIndex; safecall;
    procedure Set_ForegroundPatternColorIndex(prop: WdColorIndex); safecall;
    function Get_BackgroundPatternColorIndex: WdColorIndex; safecall;
    procedure Set_BackgroundPatternColorIndex(prop: WdColorIndex); safecall;
    function Get_Texture: WdTextureIndex; safecall;
    procedure Set_Texture(prop: WdTextureIndex); safecall;
    function Get_ForegroundPatternColor: WdColor; safecall;
    procedure Set_ForegroundPatternColor(prop: WdColor); safecall;
    function Get_BackgroundPatternColor: WdColor; safecall;
    procedure Set_BackgroundPatternColor(prop: WdColor); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ForegroundPatternColorIndex: WdColorIndex read Get_ForegroundPatternColorIndex write Set_ForegroundPatternColorIndex;
    property BackgroundPatternColorIndex: WdColorIndex read Get_BackgroundPatternColorIndex write Set_BackgroundPatternColorIndex;
    property Texture: WdTextureIndex read Get_Texture write Set_Texture;
    property ForegroundPatternColor: WdColor read Get_ForegroundPatternColor write Set_ForegroundPatternColor;
    property BackgroundPatternColor: WdColor read Get_BackgroundPatternColor write Set_BackgroundPatternColor;
  end;

// *********************************************************************//
// DispIntf:  ShadingDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093A-0000-0000-C000-000000000046}
// *********************************************************************//
  ShadingDisp = dispinterface
    ['{0002093A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ForegroundPatternColorIndex: WdColorIndex dispid 1;
    property BackgroundPatternColorIndex: WdColorIndex dispid 2;
    property Texture: WdTextureIndex dispid 3;
    property ForegroundPatternColor: WdColor dispid 4;
    property BackgroundPatternColor: WdColor dispid 5;
  end;

// *********************************************************************//
// Interface: TextRetrievalMode
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020939-0000-0000-C000-000000000046}
// *********************************************************************//
  TextRetrievalMode = interface(IDispatch)
    ['{00020939-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ViewType: WdViewType; safecall;
    procedure Set_ViewType(prop: WdViewType); safecall;
    function Get_Duplicate: TextRetrievalMode; safecall;
    function Get_IncludeHiddenText: WordBool; safecall;
    procedure Set_IncludeHiddenText(prop: WordBool); safecall;
    function Get_IncludeFieldCodes: WordBool; safecall;
    procedure Set_IncludeFieldCodes(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ViewType: WdViewType read Get_ViewType write Set_ViewType;
    property Duplicate: TextRetrievalMode read Get_Duplicate;
    property IncludeHiddenText: WordBool read Get_IncludeHiddenText write Set_IncludeHiddenText;
    property IncludeFieldCodes: WordBool read Get_IncludeFieldCodes write Set_IncludeFieldCodes;
  end;

// *********************************************************************//
// DispIntf:  TextRetrievalModeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020939-0000-0000-C000-000000000046}
// *********************************************************************//
  TextRetrievalModeDisp = dispinterface
    ['{00020939-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ViewType: WdViewType dispid 0;
    property Duplicate: TextRetrievalMode readonly dispid 1;
    property IncludeHiddenText: WordBool dispid 2;
    property IncludeFieldCodes: WordBool dispid 3;
  end;

// *********************************************************************//
// Interface: AutoTextEntries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020937-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoTextEntries = interface(IDispatch)
    ['{00020937-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): AutoTextEntry; safecall;
    function Add(const Name: WideString; const Range: Range): AutoTextEntry; safecall;
    function AppendToSpike(const Range: Range): AutoTextEntry; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AutoTextEntriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020937-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoTextEntriesDisp = dispinterface
    ['{00020937-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): AutoTextEntry; dispid 0;
    function Add(const Name: WideString; const Range: Range): AutoTextEntry; dispid 101;
    function AppendToSpike(const Range: Range): AutoTextEntry; dispid 102;
  end;

// *********************************************************************//
// Interface: AutoTextEntry
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020936-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoTextEntry = interface(IDispatch)
    ['{00020936-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_StyleName: WideString; safecall;
    function Get_Value: WideString; safecall;
    procedure Set_Value(const prop: WideString); safecall;
    procedure Delete; safecall;
    function Insert(const Where: Range; var RichText: OleVariant): Range; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name write Set_Name;
    property StyleName: WideString read Get_StyleName;
    property Value: WideString read Get_Value write Set_Value;
  end;

// *********************************************************************//
// DispIntf:  AutoTextEntryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020936-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoTextEntryDisp = dispinterface
    ['{00020936-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString dispid 2;
    property StyleName: WideString readonly dispid 3;
    property Value: WideString dispid 0;
    procedure Delete; dispid 101;
    function Insert(const Where: Range; var RichText: OleVariant): Range; dispid 102;
  end;

// *********************************************************************//
// Interface: System_
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020935-0000-0000-C000-000000000046}
// *********************************************************************//
  System_ = interface(IDispatch)
    ['{00020935-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_OperatingSystem: WideString; safecall;
    function Get_ProcessorType: WideString; safecall;
    function Get_Version: WideString; safecall;
    function Get_FreeDiskSpace: Integer; safecall;
    function Get_Country: WdCountry; safecall;
    function Get_LanguageDesignation: WideString; safecall;
    function Get_HorizontalResolution: Integer; safecall;
    function Get_VerticalResolution: Integer; safecall;
    function Get_ProfileString(const Section: WideString; const Key: WideString): WideString; safecall;
    procedure Set_ProfileString(const Section: WideString; const Key: WideString; 
                                const prop: WideString); safecall;
    function Get_PrivateProfileString(const FileName: WideString; const Section: WideString; 
                                      const Key: WideString): WideString; safecall;
    procedure Set_PrivateProfileString(const FileName: WideString; const Section: WideString; 
                                       const Key: WideString; const prop: WideString); safecall;
    function Get_MathCoprocessorInstalled: WordBool; safecall;
    function Get_ComputerType: WideString; safecall;
    function Get_MacintoshName: WideString; safecall;
    function Get_QuickDrawInstalled: WordBool; safecall;
    function Get_Cursor: WdCursorType; safecall;
    procedure Set_Cursor(prop: WdCursorType); safecall;
    procedure MSInfo; safecall;
    procedure Connect(const Path: WideString; var Drive: OleVariant; var Password: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property OperatingSystem: WideString read Get_OperatingSystem;
    property ProcessorType: WideString read Get_ProcessorType;
    property Version: WideString read Get_Version;
    property FreeDiskSpace: Integer read Get_FreeDiskSpace;
    property Country: WdCountry read Get_Country;
    property LanguageDesignation: WideString read Get_LanguageDesignation;
    property HorizontalResolution: Integer read Get_HorizontalResolution;
    property VerticalResolution: Integer read Get_VerticalResolution;
    property ProfileString[const Section: WideString; const Key: WideString]: WideString read Get_ProfileString write Set_ProfileString;
    property PrivateProfileString[const FileName: WideString; const Section: WideString; 
                                  const Key: WideString]: WideString read Get_PrivateProfileString write Set_PrivateProfileString;
    property MathCoprocessorInstalled: WordBool read Get_MathCoprocessorInstalled;
    property ComputerType: WideString read Get_ComputerType;
    property MacintoshName: WideString read Get_MacintoshName;
    property QuickDrawInstalled: WordBool read Get_QuickDrawInstalled;
    property Cursor: WdCursorType read Get_Cursor write Set_Cursor;
  end;

// *********************************************************************//
// DispIntf:  System_Disp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020935-0000-0000-C000-000000000046}
// *********************************************************************//
  System_Disp = dispinterface
    ['{00020935-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property OperatingSystem: WideString readonly dispid 1;
    property ProcessorType: WideString readonly dispid 2;
    property Version: WideString readonly dispid 3;
    property FreeDiskSpace: Integer readonly dispid 4;
    property Country: WdCountry readonly dispid 5;
    property LanguageDesignation: WideString readonly dispid 6;
    property HorizontalResolution: Integer readonly dispid 7;
    property VerticalResolution: Integer readonly dispid 8;
    property ProfileString[const Section: WideString; const Key: WideString]: WideString dispid 9;
    property PrivateProfileString[const FileName: WideString; const Section: WideString; 
                                  const Key: WideString]: WideString dispid 10;
    property MathCoprocessorInstalled: WordBool readonly dispid 11;
    property ComputerType: WideString readonly dispid 12;
    property MacintoshName: WideString readonly dispid 14;
    property QuickDrawInstalled: WordBool readonly dispid 15;
    property Cursor: WdCursorType dispid 16;
    procedure MSInfo; dispid 101;
    procedure Connect(const Path: WideString; var Drive: OleVariant; var Password: OleVariant); dispid 102;
  end;

// *********************************************************************//
// Interface: OLEFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020933-0000-0000-C000-000000000046}
// *********************************************************************//
  OLEFormat = interface(IDispatch)
    ['{00020933-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ClassType: WideString; safecall;
    procedure Set_ClassType(const prop: WideString); safecall;
    function Get_DisplayAsIcon: WordBool; safecall;
    procedure Set_DisplayAsIcon(prop: WordBool); safecall;
    function Get_IconName: WideString; safecall;
    procedure Set_IconName(const prop: WideString); safecall;
    function Get_IconPath: WideString; safecall;
    function Get_IconIndex: Integer; safecall;
    procedure Set_IconIndex(prop: Integer); safecall;
    function Get_IconLabel: WideString; safecall;
    procedure Set_IconLabel(const prop: WideString); safecall;
    function Get_Label_: WideString; safecall;
    function Get_Object_: IDispatch; safecall;
    function Get_ProgID: WideString; safecall;
    procedure Activate; safecall;
    procedure Edit; safecall;
    procedure Open; safecall;
    procedure DoVerb(var VerbIndex: OleVariant); safecall;
    procedure ConvertTo(var ClassType: OleVariant; var DisplayAsIcon: OleVariant; 
                        var IconFileName: OleVariant; var IconIndex: OleVariant; 
                        var IconLabel: OleVariant); safecall;
    procedure ActivateAs(const ClassType: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ClassType: WideString read Get_ClassType write Set_ClassType;
    property DisplayAsIcon: WordBool read Get_DisplayAsIcon write Set_DisplayAsIcon;
    property IconName: WideString read Get_IconName write Set_IconName;
    property IconPath: WideString read Get_IconPath;
    property IconIndex: Integer read Get_IconIndex write Set_IconIndex;
    property IconLabel: WideString read Get_IconLabel write Set_IconLabel;
    property Label_: WideString read Get_Label_;
    property Object_: IDispatch read Get_Object_;
    property ProgID: WideString read Get_ProgID;
  end;

// *********************************************************************//
// DispIntf:  OLEFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020933-0000-0000-C000-000000000046}
// *********************************************************************//
  OLEFormatDisp = dispinterface
    ['{00020933-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ClassType: WideString dispid 2;
    property DisplayAsIcon: WordBool dispid 3;
    property IconName: WideString dispid 7;
    property IconPath: WideString readonly dispid 8;
    property IconIndex: Integer dispid 9;
    property IconLabel: WideString dispid 10;
    property Label_: WideString readonly dispid 12;
    property Object_: IDispatch readonly dispid 14;
    property ProgID: WideString readonly dispid 22;
    procedure Activate; dispid 104;
    procedure Edit; dispid 106;
    procedure Open; dispid 107;
    procedure DoVerb(var VerbIndex: OleVariant); dispid 109;
    procedure ConvertTo(var ClassType: OleVariant; var DisplayAsIcon: OleVariant; 
                        var IconFileName: OleVariant; var IconIndex: OleVariant; 
                        var IconLabel: OleVariant); dispid 110;
    procedure ActivateAs(const ClassType: WideString); dispid 111;
  end;

// *********************************************************************//
// Interface: LinkFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020931-0000-0000-C000-000000000046}
// *********************************************************************//
  LinkFormat = interface(IDispatch)
    ['{00020931-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_AutoUpdate: WordBool; safecall;
    procedure Set_AutoUpdate(prop: WordBool); safecall;
    function Get_SourceName: WideString; safecall;
    function Get_SourcePath: WideString; safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(prop: WordBool); safecall;
    function Get_Type_: WdLinkType; safecall;
    function Get_SourceFullName: WideString; safecall;
    procedure Set_SourceFullName(const prop: WideString); safecall;
    function Get_SavePictureWithDocument: WordBool; safecall;
    procedure Set_SavePictureWithDocument(prop: WordBool); safecall;
    procedure BreakLink; safecall;
    procedure Update; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property AutoUpdate: WordBool read Get_AutoUpdate write Set_AutoUpdate;
    property SourceName: WideString read Get_SourceName;
    property SourcePath: WideString read Get_SourcePath;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property Type_: WdLinkType read Get_Type_;
    property SourceFullName: WideString read Get_SourceFullName write Set_SourceFullName;
    property SavePictureWithDocument: WordBool read Get_SavePictureWithDocument write Set_SavePictureWithDocument;
  end;

// *********************************************************************//
// DispIntf:  LinkFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020931-0000-0000-C000-000000000046}
// *********************************************************************//
  LinkFormatDisp = dispinterface
    ['{00020931-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property AutoUpdate: WordBool dispid 1;
    property SourceName: WideString readonly dispid 4;
    property SourcePath: WideString readonly dispid 5;
    property Locked: WordBool dispid 13;
    property Type_: WdLinkType readonly dispid 16;
    property SourceFullName: WideString dispid 21;
    property SavePictureWithDocument: WordBool dispid 22;
    procedure BreakLink; dispid 104;
    procedure Update; dispid 105;
  end;

// *********************************************************************//
// Interface: _OLEControl
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {000209A4-0000-0000-C000-000000000046}
// *********************************************************************//
  _OLEControl = interface(IDispatch)
    ['{000209A4-0000-0000-C000-000000000046}']
    function Get_Left: Single; safecall;
    procedure Set_Left(prop: Single); safecall;
    function Get_Top: Single; safecall;
    procedure Set_Top(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Automation: IDispatch; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    procedure Activate; safecall;
    function Get_AltHTML: WideString; safecall;
    procedure Set_AltHTML(const prop: WideString); safecall;
    property Left: Single read Get_Left write Set_Left;
    property Top: Single read Get_Top write Set_Top;
    property Height: Single read Get_Height write Set_Height;
    property Width: Single read Get_Width write Set_Width;
    property Name: WideString read Get_Name write Set_Name;
    property Automation: IDispatch read Get_Automation;
    property AltHTML: WideString read Get_AltHTML write Set_AltHTML;
  end;

// *********************************************************************//
// DispIntf:  _OLEControlDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {000209A4-0000-0000-C000-000000000046}
// *********************************************************************//
  _OLEControlDisp = dispinterface
    ['{000209A4-0000-0000-C000-000000000046}']
    property Left: Single dispid -2147417853;
    property Top: Single dispid -2147417852;
    property Height: Single dispid -2147417851;
    property Width: Single dispid -2147417850;
    property Name: WideString dispid -2147418112;
    property Automation: IDispatch readonly dispid -2147417849;
    procedure Select; dispid -2147417568;
    procedure Copy; dispid -2147417560;
    procedure Cut; dispid -2147417559;
    procedure Delete; dispid -2147417520;
    procedure Activate; dispid -2147417519;
    property AltHTML: WideString dispid -2147415101;
  end;

// *********************************************************************//
// Interface: Fields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020930-0000-0000-C000-000000000046}
// *********************************************************************//
  Fields = interface(IDispatch)
    ['{00020930-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get_Locked: Integer; safecall;
    procedure Set_Locked(prop: Integer); safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(Index: Integer): Field; safecall;
    procedure ToggleShowCodes; safecall;
    function Update: Integer; safecall;
    procedure Unlink; safecall;
    procedure UpdateSource; safecall;
    function Add(const Range: Range; var Type_: OleVariant; var Text: OleVariant; 
                 var PreserveFormatting: OleVariant): Field; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property Locked: Integer read Get_Locked write Set_Locked;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  FieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020930-0000-0000-C000-000000000046}
// *********************************************************************//
  FieldsDisp = dispinterface
    ['{00020930-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property Locked: Integer dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(Index: Integer): Field; dispid 0;
    procedure ToggleShowCodes; dispid 100;
    function Update: Integer; dispid 101;
    procedure Unlink; dispid 102;
    procedure UpdateSource; dispid 104;
    function Add(const Range: Range; var Type_: OleVariant; var Text: OleVariant; 
                 var PreserveFormatting: OleVariant): Field; dispid 105;
  end;

// *********************************************************************//
// Interface: Field
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092F-0000-0000-C000-000000000046}
// *********************************************************************//
  Field = interface(IDispatch)
    ['{0002092F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Code: Range; safecall;
    procedure Set_Code(const prop: Range); safecall;
    function Get_Type_: WdFieldType; safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(prop: WordBool); safecall;
    function Get_Kind: WdFieldKind; safecall;
    function Get_Result_: Range; safecall;
    procedure Set_Result_(const prop: Range); safecall;
    function Get_Data: WideString; safecall;
    procedure Set_Data(const prop: WideString); safecall;
    function Get_Next: Field; safecall;
    function Get_Previous: Field; safecall;
    function Get_Index: Integer; safecall;
    function Get_ShowCodes: WordBool; safecall;
    procedure Set_ShowCodes(prop: WordBool); safecall;
    function Get_LinkFormat: LinkFormat; safecall;
    function Get_OLEFormat: OLEFormat; safecall;
    function Get_InlineShape: InlineShape; safecall;
    procedure Select; safecall;
    function Update: WordBool; safecall;
    procedure Unlink; safecall;
    procedure UpdateSource; safecall;
    procedure DoClick; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Code: Range read Get_Code write Set_Code;
    property Type_: WdFieldType read Get_Type_;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property Kind: WdFieldKind read Get_Kind;
    property Result_: Range read Get_Result_ write Set_Result_;
    property Data: WideString read Get_Data write Set_Data;
    property Next: Field read Get_Next;
    property Previous: Field read Get_Previous;
    property Index: Integer read Get_Index;
    property ShowCodes: WordBool read Get_ShowCodes write Set_ShowCodes;
    property LinkFormat: LinkFormat read Get_LinkFormat;
    property OLEFormat: OLEFormat read Get_OLEFormat;
    property InlineShape: InlineShape read Get_InlineShape;
  end;

// *********************************************************************//
// DispIntf:  FieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092F-0000-0000-C000-000000000046}
// *********************************************************************//
  FieldDisp = dispinterface
    ['{0002092F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Code: Range dispid 0;
    property Type_: WdFieldType readonly dispid 1;
    property Locked: WordBool dispid 2;
    property Kind: WdFieldKind readonly dispid 3;
    property Result_: Range dispid 4;
    property Data: WideString dispid 5;
    property Next: Field readonly dispid 6;
    property Previous: Field readonly dispid 7;
    property Index: Integer readonly dispid 8;
    property ShowCodes: WordBool dispid 9;
    property LinkFormat: LinkFormat readonly dispid 10;
    property OLEFormat: OLEFormat readonly dispid 11;
    property InlineShape: InlineShape readonly dispid 12;
    procedure Select; dispid 65535;
    function Update: WordBool; dispid 101;
    procedure Unlink; dispid 102;
    procedure UpdateSource; dispid 103;
    procedure DoClick; dispid 104;
    procedure Copy; dispid 105;
    procedure Cut; dispid 106;
    procedure Delete; dispid 107;
  end;

// *********************************************************************//
// Interface: Browser
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092E-0000-0000-C000-000000000046}
// *********************************************************************//
  Browser = interface(IDispatch)
    ['{0002092E-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Target: WdBrowseTarget; safecall;
    procedure Set_Target(prop: WdBrowseTarget); safecall;
    procedure Next; safecall;
    procedure Previous; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Target: WdBrowseTarget read Get_Target write Set_Target;
  end;

// *********************************************************************//
// DispIntf:  BrowserDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092E-0000-0000-C000-000000000046}
// *********************************************************************//
  BrowserDisp = dispinterface
    ['{0002092E-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Target: WdBrowseTarget dispid 1;
    procedure Next; dispid 101;
    procedure Previous; dispid 102;
  end;

// *********************************************************************//
// Interface: Styles
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092D-0000-0000-C000-000000000046}
// *********************************************************************//
  Styles = interface(IDispatch)
    ['{0002092D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): Style; safecall;
    function Add(const Name: WideString; var Type_: OleVariant): Style; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  StylesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092D-0000-0000-C000-000000000046}
// *********************************************************************//
  StylesDisp = dispinterface
    ['{0002092D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): Style; dispid 0;
    function Add(const Name: WideString; var Type_: OleVariant): Style; dispid 100;
  end;

// *********************************************************************//
// Interface: Style
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092C-0000-0000-C000-000000000046}
// *********************************************************************//
  Style = interface(IDispatch)
    ['{0002092C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_NameLocal: WideString; safecall;
    procedure Set_NameLocal(const prop: WideString); safecall;
    function Get_BaseStyle: OleVariant; safecall;
    procedure Set_BaseStyle(var prop: OleVariant); safecall;
    function Get_Description: WideString; safecall;
    function Get_Type_: WdStyleType; safecall;
    function Get_BuiltIn: WordBool; safecall;
    function Get_NextParagraphStyle: OleVariant; safecall;
    procedure Set_NextParagraphStyle(var prop: OleVariant); safecall;
    function Get_InUse: WordBool; safecall;
    function Get_Shading: Shading; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_ParagraphFormat: ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: ParagraphFormat); safecall;
    function Get_Font: Font; safecall;
    procedure Set_Font(const prop: Font); safecall;
    function Get_Frame: Frame; safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_AutomaticallyUpdate: WordBool; safecall;
    procedure Set_AutomaticallyUpdate(prop: WordBool); safecall;
    function Get_ListTemplate: ListTemplate; safecall;
    function Get_ListLevelNumber: Integer; safecall;
    function Get_LanguageIDFarEast: WdLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WdLanguageID); safecall;
    function Get_Hidden: WordBool; safecall;
    procedure Set_Hidden(prop: WordBool); safecall;
    procedure Delete; safecall;
    procedure LinkToListTemplate(const ListTemplate: ListTemplate; var ListLevelNumber: OleVariant); safecall;
    function Get_NoProofing: Integer; safecall;
    procedure Set_NoProofing(prop: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property NameLocal: WideString read Get_NameLocal write Set_NameLocal;
    property Description: WideString read Get_Description;
    property Type_: WdStyleType read Get_Type_;
    property BuiltIn: WordBool read Get_BuiltIn;
    property InUse: WordBool read Get_InUse;
    property Shading: Shading read Get_Shading;
    property Borders: Borders read Get_Borders write Set_Borders;
    property ParagraphFormat: ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property Font: Font read Get_Font write Set_Font;
    property Frame: Frame read Get_Frame;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property AutomaticallyUpdate: WordBool read Get_AutomaticallyUpdate write Set_AutomaticallyUpdate;
    property ListTemplate: ListTemplate read Get_ListTemplate;
    property ListLevelNumber: Integer read Get_ListLevelNumber;
    property LanguageIDFarEast: WdLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property Hidden: WordBool read Get_Hidden write Set_Hidden;
    property NoProofing: Integer read Get_NoProofing write Set_NoProofing;
  end;

// *********************************************************************//
// DispIntf:  StyleDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092C-0000-0000-C000-000000000046}
// *********************************************************************//
  StyleDisp = dispinterface
    ['{0002092C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property NameLocal: WideString dispid 0;
    function BaseStyle: OleVariant; dispid 1;
    property Description: WideString readonly dispid 2;
    property Type_: WdStyleType readonly dispid 3;
    property BuiltIn: WordBool readonly dispid 4;
    function NextParagraphStyle: OleVariant; dispid 5;
    property InUse: WordBool readonly dispid 6;
    property Shading: Shading readonly dispid 7;
    property Borders: Borders dispid 8;
    property ParagraphFormat: ParagraphFormat dispid 9;
    property Font: Font dispid 10;
    property Frame: Frame readonly dispid 11;
    property LanguageID: WdLanguageID dispid 12;
    property AutomaticallyUpdate: WordBool dispid 13;
    property ListTemplate: ListTemplate readonly dispid 14;
    property ListLevelNumber: Integer readonly dispid 15;
    property LanguageIDFarEast: WdLanguageID dispid 16;
    property Hidden: WordBool dispid 17;
    procedure Delete; dispid 100;
    procedure LinkToListTemplate(const ListTemplate: ListTemplate; var ListLevelNumber: OleVariant); dispid 101;
    property NoProofing: Integer dispid 18;
  end;

// *********************************************************************//
// Interface: Frames
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092B-0000-0000-C000-000000000046}
// *********************************************************************//
  Frames = interface(IDispatch)
    ['{0002092B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Frame; safecall;
    function Add(const Range: Range): Frame; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  FramesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092B-0000-0000-C000-000000000046}
// *********************************************************************//
  FramesDisp = dispinterface
    ['{0002092B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): Frame; dispid 0;
    function Add(const Range: Range): Frame; dispid 100;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: Frame
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092A-0000-0000-C000-000000000046}
// *********************************************************************//
  Frame = interface(IDispatch)
    ['{0002092A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_HeightRule: WdFrameSizeRule; safecall;
    procedure Set_HeightRule(prop: WdFrameSizeRule); safecall;
    function Get_WidthRule: WdFrameSizeRule; safecall;
    procedure Set_WidthRule(prop: WdFrameSizeRule); safecall;
    function Get_HorizontalDistanceFromText: Single; safecall;
    procedure Set_HorizontalDistanceFromText(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HorizontalPosition: Single; safecall;
    procedure Set_HorizontalPosition(prop: Single); safecall;
    function Get_LockAnchor: WordBool; safecall;
    procedure Set_LockAnchor(prop: WordBool); safecall;
    function Get_RelativeHorizontalPosition: WdRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WdRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WdRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WdRelativeVerticalPosition); safecall;
    function Get_VerticalDistanceFromText: Single; safecall;
    procedure Set_VerticalDistanceFromText(prop: Single); safecall;
    function Get_VerticalPosition: Single; safecall;
    procedure Set_VerticalPosition(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_TextWrap: WordBool; safecall;
    procedure Set_TextWrap(prop: WordBool); safecall;
    function Get_Shading: Shading; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Range: Range; safecall;
    procedure Delete; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property HeightRule: WdFrameSizeRule read Get_HeightRule write Set_HeightRule;
    property WidthRule: WdFrameSizeRule read Get_WidthRule write Set_WidthRule;
    property HorizontalDistanceFromText: Single read Get_HorizontalDistanceFromText write Set_HorizontalDistanceFromText;
    property Height: Single read Get_Height write Set_Height;
    property HorizontalPosition: Single read Get_HorizontalPosition write Set_HorizontalPosition;
    property LockAnchor: WordBool read Get_LockAnchor write Set_LockAnchor;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WdRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property VerticalDistanceFromText: Single read Get_VerticalDistanceFromText write Set_VerticalDistanceFromText;
    property VerticalPosition: Single read Get_VerticalPosition write Set_VerticalPosition;
    property Width: Single read Get_Width write Set_Width;
    property TextWrap: WordBool read Get_TextWrap write Set_TextWrap;
    property Shading: Shading read Get_Shading;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Range: Range read Get_Range;
  end;

// *********************************************************************//
// DispIntf:  FrameDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092A-0000-0000-C000-000000000046}
// *********************************************************************//
  FrameDisp = dispinterface
    ['{0002092A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property HeightRule: WdFrameSizeRule dispid 1;
    property WidthRule: WdFrameSizeRule dispid 2;
    property HorizontalDistanceFromText: Single dispid 3;
    property Height: Single dispid 4;
    property HorizontalPosition: Single dispid 5;
    property LockAnchor: WordBool dispid 6;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition dispid 7;
    property RelativeVerticalPosition: WdRelativeVerticalPosition dispid 8;
    property VerticalDistanceFromText: Single dispid 9;
    property VerticalPosition: Single dispid 10;
    property Width: Single dispid 11;
    property TextWrap: WordBool dispid 12;
    property Shading: Shading readonly dispid 13;
    property Borders: Borders dispid 1100;
    property Range: Range readonly dispid 15;
    procedure Delete; dispid 100;
    procedure Select; dispid 65535;
    procedure Copy; dispid 101;
    procedure Cut; dispid 102;
  end;

// *********************************************************************//
// Interface: FormFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020929-0000-0000-C000-000000000046}
// *********************************************************************//
  FormFields = interface(IDispatch)
    ['{00020929-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get_Shaded: WordBool; safecall;
    procedure Set_Shaded(prop: WordBool); safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): FormField; safecall;
    function Add(const Range: Range; Type_: WdFieldType): FormField; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property Shaded: WordBool read Get_Shaded write Set_Shaded;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  FormFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020929-0000-0000-C000-000000000046}
// *********************************************************************//
  FormFieldsDisp = dispinterface
    ['{00020929-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property Shaded: WordBool dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): FormField; dispid 0;
    function Add(const Range: Range; Type_: WdFieldType): FormField; dispid 101;
  end;

// *********************************************************************//
// Interface: FormField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020928-0000-0000-C000-000000000046}
// *********************************************************************//
  FormField = interface(IDispatch)
    ['{00020928-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Type_: WdFieldType; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_EntryMacro: WideString; safecall;
    procedure Set_EntryMacro(const prop: WideString); safecall;
    function Get_ExitMacro: WideString; safecall;
    procedure Set_ExitMacro(const prop: WideString); safecall;
    function Get_OwnHelp: WordBool; safecall;
    procedure Set_OwnHelp(prop: WordBool); safecall;
    function Get_OwnStatus: WordBool; safecall;
    procedure Set_OwnStatus(prop: WordBool); safecall;
    function Get_HelpText: WideString; safecall;
    procedure Set_HelpText(const prop: WideString); safecall;
    function Get_StatusText: WideString; safecall;
    procedure Set_StatusText(const prop: WideString); safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(prop: WordBool); safecall;
    function Get_Result_: WideString; safecall;
    procedure Set_Result_(const prop: WideString); safecall;
    function Get_TextInput: TextInput; safecall;
    function Get_CheckBox: CheckBox; safecall;
    function Get_DropDown: DropDown; safecall;
    function Get_Next: FormField; safecall;
    function Get_Previous: FormField; safecall;
    function Get_CalculateOnExit: WordBool; safecall;
    procedure Set_CalculateOnExit(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Type_: WdFieldType read Get_Type_;
    property Name: WideString read Get_Name write Set_Name;
    property EntryMacro: WideString read Get_EntryMacro write Set_EntryMacro;
    property ExitMacro: WideString read Get_ExitMacro write Set_ExitMacro;
    property OwnHelp: WordBool read Get_OwnHelp write Set_OwnHelp;
    property OwnStatus: WordBool read Get_OwnStatus write Set_OwnStatus;
    property HelpText: WideString read Get_HelpText write Set_HelpText;
    property StatusText: WideString read Get_StatusText write Set_StatusText;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Result_: WideString read Get_Result_ write Set_Result_;
    property TextInput: TextInput read Get_TextInput;
    property CheckBox: CheckBox read Get_CheckBox;
    property DropDown: DropDown read Get_DropDown;
    property Next: FormField read Get_Next;
    property Previous: FormField read Get_Previous;
    property CalculateOnExit: WordBool read Get_CalculateOnExit write Set_CalculateOnExit;
    property Range: Range read Get_Range;
  end;

// *********************************************************************//
// DispIntf:  FormFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020928-0000-0000-C000-000000000046}
// *********************************************************************//
  FormFieldDisp = dispinterface
    ['{00020928-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Type_: WdFieldType readonly dispid 0;
    property Name: WideString dispid 2;
    property EntryMacro: WideString dispid 3;
    property ExitMacro: WideString dispid 4;
    property OwnHelp: WordBool dispid 5;
    property OwnStatus: WordBool dispid 6;
    property HelpText: WideString dispid 7;
    property StatusText: WideString dispid 8;
    property Enabled: WordBool dispid 9;
    property Result_: WideString dispid 10;
    property TextInput: TextInput readonly dispid 11;
    property CheckBox: CheckBox readonly dispid 12;
    property DropDown: DropDown readonly dispid 13;
    property Next: FormField readonly dispid 14;
    property Previous: FormField readonly dispid 15;
    property CalculateOnExit: WordBool dispid 16;
    property Range: Range readonly dispid 17;
    procedure Select; dispid 65535;
    procedure Copy; dispid 101;
    procedure Cut; dispid 102;
    procedure Delete; dispid 103;
  end;

// *********************************************************************//
// Interface: TextInput
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020927-0000-0000-C000-000000000046}
// *********************************************************************//
  TextInput = interface(IDispatch)
    ['{00020927-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Valid: WordBool; safecall;
    function Get_Default: WideString; safecall;
    procedure Set_Default(const prop: WideString); safecall;
    function Get_Type_: WdTextFormFieldType; safecall;
    function Get_Format: WideString; safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    procedure Clear; safecall;
    procedure EditType(Type_: WdTextFormFieldType; var Default: OleVariant; var Format: OleVariant; 
                       var Enabled: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Valid: WordBool read Get_Valid;
    property Default: WideString read Get_Default write Set_Default;
    property Type_: WdTextFormFieldType read Get_Type_;
    property Format: WideString read Get_Format;
    property Width: Integer read Get_Width write Set_Width;
  end;

// *********************************************************************//
// DispIntf:  TextInputDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020927-0000-0000-C000-000000000046}
// *********************************************************************//
  TextInputDisp = dispinterface
    ['{00020927-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Valid: WordBool readonly dispid 0;
    property Default: WideString dispid 1;
    property Type_: WdTextFormFieldType readonly dispid 2;
    property Format: WideString readonly dispid 3;
    property Width: Integer dispid 4;
    procedure Clear; dispid 101;
    procedure EditType(Type_: WdTextFormFieldType; var Default: OleVariant; var Format: OleVariant; 
                       var Enabled: OleVariant); dispid 102;
  end;

// *********************************************************************//
// Interface: CheckBox
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020926-0000-0000-C000-000000000046}
// *********************************************************************//
  CheckBox = interface(IDispatch)
    ['{00020926-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Valid: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    procedure Set_AutoSize(prop: WordBool); safecall;
    function Get_Size: Single; safecall;
    procedure Set_Size(prop: Single); safecall;
    function Get_Default: WordBool; safecall;
    procedure Set_Default(prop: WordBool); safecall;
    function Get_Value: WordBool; safecall;
    procedure Set_Value(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Valid: WordBool read Get_Valid;
    property AutoSize: WordBool read Get_AutoSize write Set_AutoSize;
    property Size: Single read Get_Size write Set_Size;
    property Default: WordBool read Get_Default write Set_Default;
    property Value: WordBool read Get_Value write Set_Value;
  end;

// *********************************************************************//
// DispIntf:  CheckBoxDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020926-0000-0000-C000-000000000046}
// *********************************************************************//
  CheckBoxDisp = dispinterface
    ['{00020926-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Valid: WordBool readonly dispid 0;
    property AutoSize: WordBool dispid 1;
    property Size: Single dispid 2;
    property Default: WordBool dispid 3;
    property Value: WordBool dispid 4;
  end;

// *********************************************************************//
// Interface: DropDown
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020925-0000-0000-C000-000000000046}
// *********************************************************************//
  DropDown = interface(IDispatch)
    ['{00020925-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Valid: WordBool; safecall;
    function Get_Default: Integer; safecall;
    procedure Set_Default(prop: Integer); safecall;
    function Get_Value: Integer; safecall;
    procedure Set_Value(prop: Integer); safecall;
    function Get_ListEntries: ListEntries; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Valid: WordBool read Get_Valid;
    property Default: Integer read Get_Default write Set_Default;
    property Value: Integer read Get_Value write Set_Value;
    property ListEntries: ListEntries read Get_ListEntries;
  end;

// *********************************************************************//
// DispIntf:  DropDownDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020925-0000-0000-C000-000000000046}
// *********************************************************************//
  DropDownDisp = dispinterface
    ['{00020925-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Valid: WordBool readonly dispid 0;
    property Default: Integer dispid 1;
    property Value: Integer dispid 2;
    property ListEntries: ListEntries readonly dispid 3;
  end;

// *********************************************************************//
// Interface: ListEntries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020924-0000-0000-C000-000000000046}
// *********************************************************************//
  ListEntries = interface(IDispatch)
    ['{00020924-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): ListEntry; safecall;
    function Add(const Name: WideString; var Index: OleVariant): ListEntry; safecall;
    procedure Clear; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListEntriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020924-0000-0000-C000-000000000046}
// *********************************************************************//
  ListEntriesDisp = dispinterface
    ['{00020924-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): ListEntry; dispid 0;
    function Add(const Name: WideString; var Index: OleVariant): ListEntry; dispid 101;
    procedure Clear; dispid 102;
  end;

// *********************************************************************//
// Interface: ListEntry
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020923-0000-0000-C000-000000000046}
// *********************************************************************//
  ListEntry = interface(IDispatch)
    ['{00020923-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name write Set_Name;
  end;

// *********************************************************************//
// DispIntf:  ListEntryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020923-0000-0000-C000-000000000046}
// *********************************************************************//
  ListEntryDisp = dispinterface
    ['{00020923-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString dispid 2;
    procedure Delete; dispid 11;
  end;

// *********************************************************************//
// Interface: TablesOfFigures
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020922-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfFigures = interface(IDispatch)
    ['{00020922-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Format: WdTofFormat; safecall;
    procedure Set_Format(prop: WdTofFormat); safecall;
    function Item(Index: Integer): TableOfFigures; safecall;
    function AddOld(const Range: Range; var Caption: OleVariant; var IncludeLabel: OleVariant; 
                    var UseHeadingStyles: OleVariant; var UpperHeadingLevel: OleVariant; 
                    var LowerHeadingLevel: OleVariant; var UseFields: OleVariant; 
                    var TableID: OleVariant; var RightAlignPageNumbers: OleVariant; 
                    var IncludePageNumbers: OleVariant; var AddedStyles: OleVariant): TableOfFigures; safecall;
    function MarkEntry(const Range: Range; var Entry: OleVariant; var EntryAutoText: OleVariant; 
                       var TableID: OleVariant; var Level: OleVariant): Field; safecall;
    function Add(const Range: Range; var Caption: OleVariant; var IncludeLabel: OleVariant; 
                 var UseHeadingStyles: OleVariant; var UpperHeadingLevel: OleVariant; 
                 var LowerHeadingLevel: OleVariant; var UseFields: OleVariant; 
                 var TableID: OleVariant; var RightAlignPageNumbers: OleVariant; 
                 var IncludePageNumbers: OleVariant; var AddedStyles: OleVariant; 
                 var UseHyperlinks: OleVariant; var HidePageNumbersInWeb: OleVariant): TableOfFigures; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Format: WdTofFormat read Get_Format write Set_Format;
  end;

// *********************************************************************//
// DispIntf:  TablesOfFiguresDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020922-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfFiguresDisp = dispinterface
    ['{00020922-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Format: WdTofFormat dispid 2;
    function Item(Index: Integer): TableOfFigures; dispid 0;
    function AddOld(const Range: Range; var Caption: OleVariant; var IncludeLabel: OleVariant; 
                    var UseHeadingStyles: OleVariant; var UpperHeadingLevel: OleVariant; 
                    var LowerHeadingLevel: OleVariant; var UseFields: OleVariant; 
                    var TableID: OleVariant; var RightAlignPageNumbers: OleVariant; 
                    var IncludePageNumbers: OleVariant; var AddedStyles: OleVariant): TableOfFigures; dispid 100;
    function MarkEntry(const Range: Range; var Entry: OleVariant; var EntryAutoText: OleVariant; 
                       var TableID: OleVariant; var Level: OleVariant): Field; dispid 101;
    function Add(const Range: Range; var Caption: OleVariant; var IncludeLabel: OleVariant; 
                 var UseHeadingStyles: OleVariant; var UpperHeadingLevel: OleVariant; 
                 var LowerHeadingLevel: OleVariant; var UseFields: OleVariant; 
                 var TableID: OleVariant; var RightAlignPageNumbers: OleVariant; 
                 var IncludePageNumbers: OleVariant; var AddedStyles: OleVariant; 
                 var UseHyperlinks: OleVariant; var HidePageNumbersInWeb: OleVariant): TableOfFigures; dispid 444;
  end;

// *********************************************************************//
// Interface: TableOfFigures
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020921-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfFigures = interface(IDispatch)
    ['{00020921-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const prop: WideString); safecall;
    function Get_IncludeLabel: WordBool; safecall;
    procedure Set_IncludeLabel(prop: WordBool); safecall;
    function Get_RightAlignPageNumbers: WordBool; safecall;
    procedure Set_RightAlignPageNumbers(prop: WordBool); safecall;
    function Get_UseHeadingStyles: WordBool; safecall;
    procedure Set_UseHeadingStyles(prop: WordBool); safecall;
    function Get_LowerHeadingLevel: Integer; safecall;
    procedure Set_LowerHeadingLevel(prop: Integer); safecall;
    function Get_UpperHeadingLevel: Integer; safecall;
    procedure Set_UpperHeadingLevel(prop: Integer); safecall;
    function Get_IncludePageNumbers: WordBool; safecall;
    procedure Set_IncludePageNumbers(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    function Get_UseFields: WordBool; safecall;
    procedure Set_UseFields(prop: WordBool); safecall;
    function Get_TableID: WideString; safecall;
    procedure Set_TableID(const prop: WideString); safecall;
    function Get_HeadingStyles: HeadingStyles; safecall;
    function Get_TabLeader: WdTabLeader; safecall;
    procedure Set_TabLeader(prop: WdTabLeader); safecall;
    procedure Delete; safecall;
    procedure UpdatePageNumbers; safecall;
    procedure Update; safecall;
    function Get_UseHyperlinks: WordBool; safecall;
    procedure Set_UseHyperlinks(prop: WordBool); safecall;
    function Get_HidePageNumbersInWeb: WordBool; safecall;
    procedure Set_HidePageNumbersInWeb(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Caption: WideString read Get_Caption write Set_Caption;
    property IncludeLabel: WordBool read Get_IncludeLabel write Set_IncludeLabel;
    property RightAlignPageNumbers: WordBool read Get_RightAlignPageNumbers write Set_RightAlignPageNumbers;
    property UseHeadingStyles: WordBool read Get_UseHeadingStyles write Set_UseHeadingStyles;
    property LowerHeadingLevel: Integer read Get_LowerHeadingLevel write Set_LowerHeadingLevel;
    property UpperHeadingLevel: Integer read Get_UpperHeadingLevel write Set_UpperHeadingLevel;
    property IncludePageNumbers: WordBool read Get_IncludePageNumbers write Set_IncludePageNumbers;
    property Range: Range read Get_Range;
    property UseFields: WordBool read Get_UseFields write Set_UseFields;
    property TableID: WideString read Get_TableID write Set_TableID;
    property HeadingStyles: HeadingStyles read Get_HeadingStyles;
    property TabLeader: WdTabLeader read Get_TabLeader write Set_TabLeader;
    property UseHyperlinks: WordBool read Get_UseHyperlinks write Set_UseHyperlinks;
    property HidePageNumbersInWeb: WordBool read Get_HidePageNumbersInWeb write Set_HidePageNumbersInWeb;
  end;

// *********************************************************************//
// DispIntf:  TableOfFiguresDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020921-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfFiguresDisp = dispinterface
    ['{00020921-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Caption: WideString dispid 1;
    property IncludeLabel: WordBool dispid 2;
    property RightAlignPageNumbers: WordBool dispid 3;
    property UseHeadingStyles: WordBool dispid 4;
    property LowerHeadingLevel: Integer dispid 5;
    property UpperHeadingLevel: Integer dispid 6;
    property IncludePageNumbers: WordBool dispid 7;
    property Range: Range readonly dispid 8;
    property UseFields: WordBool dispid 9;
    property TableID: WideString dispid 10;
    property HeadingStyles: HeadingStyles readonly dispid 11;
    property TabLeader: WdTabLeader dispid 12;
    procedure Delete; dispid 100;
    procedure UpdatePageNumbers; dispid 101;
    procedure Update; dispid 102;
    property UseHyperlinks: WordBool dispid 13;
    property HidePageNumbersInWeb: WordBool dispid 14;
  end;

// *********************************************************************//
// Interface: MailMerge
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020920-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMerge = interface(IDispatch)
    ['{00020920-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_MainDocumentType: WdMailMergeMainDocType; safecall;
    procedure Set_MainDocumentType(prop: WdMailMergeMainDocType); safecall;
    function Get_State: WdMailMergeState; safecall;
    function Get_Destination: WdMailMergeDestination; safecall;
    procedure Set_Destination(prop: WdMailMergeDestination); safecall;
    function Get_DataSource: MailMergeDataSource; safecall;
    function Get_Fields: MailMergeFields; safecall;
    function Get_ViewMailMergeFieldCodes: Integer; safecall;
    procedure Set_ViewMailMergeFieldCodes(prop: Integer); safecall;
    function Get_SuppressBlankLines: WordBool; safecall;
    procedure Set_SuppressBlankLines(prop: WordBool); safecall;
    function Get_MailAsAttachment: WordBool; safecall;
    procedure Set_MailAsAttachment(prop: WordBool); safecall;
    function Get_MailAddressFieldName: WideString; safecall;
    procedure Set_MailAddressFieldName(const prop: WideString); safecall;
    function Get_MailSubject: WideString; safecall;
    procedure Set_MailSubject(const prop: WideString); safecall;
    procedure CreateDataSource(var Name: OleVariant; var PasswordDocument: OleVariant; 
                               var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant; 
                               var MSQuery: OleVariant; var SQLStatement: OleVariant; 
                               var SQLStatement1: OleVariant; var Connection: OleVariant; 
                               var LinkToSource: OleVariant); safecall;
    procedure CreateHeaderSource(const Name: WideString; var PasswordDocument: OleVariant; 
                                 var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant); safecall;
    procedure OpenDataSource(const Name: WideString; var Format: OleVariant; 
                             var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                             var LinkToSource: OleVariant; var AddToRecentFiles: OleVariant; 
                             var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                             var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                             var WritePasswordTemplate: OleVariant; var Connection: OleVariant; 
                             var SQLStatement: OleVariant; var SQLStatement1: OleVariant); safecall;
    procedure OpenHeaderSource(const Name: WideString; var Format: OleVariant; 
                               var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                               var AddToRecentFiles: OleVariant; var PasswordDocument: OleVariant; 
                               var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                               var WritePasswordDocument: OleVariant; 
                               var WritePasswordTemplate: OleVariant); safecall;
    procedure Execute(var Pause: OleVariant); safecall;
    procedure Check; safecall;
    procedure EditDataSource; safecall;
    procedure EditHeaderSource; safecall;
    procedure EditMainDocument; safecall;
    procedure UseAddressBook(const Type_: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property MainDocumentType: WdMailMergeMainDocType read Get_MainDocumentType write Set_MainDocumentType;
    property State: WdMailMergeState read Get_State;
    property Destination: WdMailMergeDestination read Get_Destination write Set_Destination;
    property DataSource: MailMergeDataSource read Get_DataSource;
    property Fields: MailMergeFields read Get_Fields;
    property ViewMailMergeFieldCodes: Integer read Get_ViewMailMergeFieldCodes write Set_ViewMailMergeFieldCodes;
    property SuppressBlankLines: WordBool read Get_SuppressBlankLines write Set_SuppressBlankLines;
    property MailAsAttachment: WordBool read Get_MailAsAttachment write Set_MailAsAttachment;
    property MailAddressFieldName: WideString read Get_MailAddressFieldName write Set_MailAddressFieldName;
    property MailSubject: WideString read Get_MailSubject write Set_MailSubject;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020920-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDisp = dispinterface
    ['{00020920-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property MainDocumentType: WdMailMergeMainDocType dispid 1;
    property State: WdMailMergeState readonly dispid 2;
    property Destination: WdMailMergeDestination dispid 3;
    property DataSource: MailMergeDataSource readonly dispid 4;
    property Fields: MailMergeFields readonly dispid 5;
    property ViewMailMergeFieldCodes: Integer dispid 6;
    property SuppressBlankLines: WordBool dispid 7;
    property MailAsAttachment: WordBool dispid 8;
    property MailAddressFieldName: WideString dispid 9;
    property MailSubject: WideString dispid 10;
    procedure CreateDataSource(var Name: OleVariant; var PasswordDocument: OleVariant; 
                               var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant; 
                               var MSQuery: OleVariant; var SQLStatement: OleVariant; 
                               var SQLStatement1: OleVariant; var Connection: OleVariant; 
                               var LinkToSource: OleVariant); dispid 101;
    procedure CreateHeaderSource(const Name: WideString; var PasswordDocument: OleVariant; 
                                 var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant); dispid 102;
    procedure OpenDataSource(const Name: WideString; var Format: OleVariant; 
                             var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                             var LinkToSource: OleVariant; var AddToRecentFiles: OleVariant; 
                             var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                             var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                             var WritePasswordTemplate: OleVariant; var Connection: OleVariant; 
                             var SQLStatement: OleVariant; var SQLStatement1: OleVariant); dispid 103;
    procedure OpenHeaderSource(const Name: WideString; var Format: OleVariant; 
                               var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                               var AddToRecentFiles: OleVariant; var PasswordDocument: OleVariant; 
                               var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                               var WritePasswordDocument: OleVariant; 
                               var WritePasswordTemplate: OleVariant); dispid 104;
    procedure Execute(var Pause: OleVariant); dispid 105;
    procedure Check; dispid 106;
    procedure EditDataSource; dispid 107;
    procedure EditHeaderSource; dispid 108;
    procedure EditMainDocument; dispid 109;
    procedure UseAddressBook(const Type_: WideString); dispid 111;
  end;

// *********************************************************************//
// Interface: MailMergeFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091F-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFields = interface(IDispatch)
    ['{0002091F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): MailMergeField; safecall;
    function Add(const Range: Range; const Name: WideString): MailMergeField; safecall;
    function AddAsk(const Range: Range; const Name: WideString; var Prompt: OleVariant; 
                    var DefaultAskText: OleVariant; var AskOnce: OleVariant): MailMergeField; safecall;
    function AddFillIn(const Range: Range; var Prompt: OleVariant; 
                       var DefaultFillInText: OleVariant; var AskOnce: OleVariant): MailMergeField; safecall;
    function AddIf(const Range: Range; const MergeField: WideString; 
                   Comparison: WdMailMergeComparison; var CompareTo: OleVariant; 
                   var TrueAutoText: OleVariant; var TrueText: OleVariant; 
                   var FalseAutoText: OleVariant; var FalseText: OleVariant): MailMergeField; safecall;
    function AddMergeRec(const Range: Range): MailMergeField; safecall;
    function AddMergeSeq(const Range: Range): MailMergeField; safecall;
    function AddNext(const Range: Range): MailMergeField; safecall;
    function AddNextIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WdMailMergeComparison; var CompareTo: OleVariant): MailMergeField; safecall;
    function AddSet(const Range: Range; const Name: WideString; var ValueText: OleVariant; 
                    var ValueAutoText: OleVariant): MailMergeField; safecall;
    function AddSkipIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WdMailMergeComparison; var CompareTo: OleVariant): MailMergeField; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091F-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFieldsDisp = dispinterface
    ['{0002091F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): MailMergeField; dispid 0;
    function Add(const Range: Range; const Name: WideString): MailMergeField; dispid 101;
    function AddAsk(const Range: Range; const Name: WideString; var Prompt: OleVariant; 
                    var DefaultAskText: OleVariant; var AskOnce: OleVariant): MailMergeField; dispid 102;
    function AddFillIn(const Range: Range; var Prompt: OleVariant; 
                       var DefaultFillInText: OleVariant; var AskOnce: OleVariant): MailMergeField; dispid 103;
    function AddIf(const Range: Range; const MergeField: WideString; 
                   Comparison: WdMailMergeComparison; var CompareTo: OleVariant; 
                   var TrueAutoText: OleVariant; var TrueText: OleVariant; 
                   var FalseAutoText: OleVariant; var FalseText: OleVariant): MailMergeField; dispid 104;
    function AddMergeRec(const Range: Range): MailMergeField; dispid 105;
    function AddMergeSeq(const Range: Range): MailMergeField; dispid 106;
    function AddNext(const Range: Range): MailMergeField; dispid 107;
    function AddNextIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WdMailMergeComparison; var CompareTo: OleVariant): MailMergeField; dispid 108;
    function AddSet(const Range: Range; const Name: WideString; var ValueText: OleVariant; 
                    var ValueAutoText: OleVariant): MailMergeField; dispid 109;
    function AddSkipIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WdMailMergeComparison; var CompareTo: OleVariant): MailMergeField; dispid 110;
  end;

// *********************************************************************//
// Interface: MailMergeField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091E-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeField = interface(IDispatch)
    ['{0002091E-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Type_: WdFieldType; safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(prop: WordBool); safecall;
    function Get_Code: Range; safecall;
    procedure Set_Code(const prop: Range); safecall;
    function Get_Next: MailMergeField; safecall;
    function Get_Previous: MailMergeField; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Type_: WdFieldType read Get_Type_;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property Code: Range read Get_Code write Set_Code;
    property Next: MailMergeField read Get_Next;
    property Previous: MailMergeField read Get_Previous;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091E-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFieldDisp = dispinterface
    ['{0002091E-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Type_: WdFieldType readonly dispid 0;
    property Locked: WordBool dispid 3;
    property Code: Range dispid 5;
    property Next: MailMergeField readonly dispid 8;
    property Previous: MailMergeField readonly dispid 9;
    procedure Select; dispid 65535;
    procedure Copy; dispid 105;
    procedure Cut; dispid 106;
    procedure Delete; dispid 107;
  end;

// *********************************************************************//
// Interface: MailMergeDataSource
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091D-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDataSource = interface(IDispatch)
    ['{0002091D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_HeaderSourceName: WideString; safecall;
    function Get_Type_: WdMailMergeDataSource; safecall;
    function Get_HeaderSourceType: WdMailMergeDataSource; safecall;
    function Get_ConnectString: WideString; safecall;
    function Get_QueryString: WideString; safecall;
    procedure Set_QueryString(const prop: WideString); safecall;
    function Get_ActiveRecord: WdMailMergeActiveRecord; safecall;
    procedure Set_ActiveRecord(prop: WdMailMergeActiveRecord); safecall;
    function Get_FirstRecord: Integer; safecall;
    procedure Set_FirstRecord(prop: Integer); safecall;
    function Get_LastRecord: Integer; safecall;
    procedure Set_LastRecord(prop: Integer); safecall;
    function Get_FieldNames: MailMergeFieldNames; safecall;
    function Get_DataFields: MailMergeDataFields; safecall;
    function FindRecord(const FindText: WideString; const Field: WideString): WordBool; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property HeaderSourceName: WideString read Get_HeaderSourceName;
    property Type_: WdMailMergeDataSource read Get_Type_;
    property HeaderSourceType: WdMailMergeDataSource read Get_HeaderSourceType;
    property ConnectString: WideString read Get_ConnectString;
    property QueryString: WideString read Get_QueryString write Set_QueryString;
    property ActiveRecord: WdMailMergeActiveRecord read Get_ActiveRecord write Set_ActiveRecord;
    property FirstRecord: Integer read Get_FirstRecord write Set_FirstRecord;
    property LastRecord: Integer read Get_LastRecord write Set_LastRecord;
    property FieldNames: MailMergeFieldNames read Get_FieldNames;
    property DataFields: MailMergeDataFields read Get_DataFields;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDataSourceDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091D-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDataSourceDisp = dispinterface
    ['{0002091D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 1;
    property HeaderSourceName: WideString readonly dispid 2;
    property Type_: WdMailMergeDataSource readonly dispid 3;
    property HeaderSourceType: WdMailMergeDataSource readonly dispid 4;
    property ConnectString: WideString readonly dispid 5;
    property QueryString: WideString dispid 6;
    property ActiveRecord: WdMailMergeActiveRecord dispid 7;
    property FirstRecord: Integer dispid 8;
    property LastRecord: Integer dispid 9;
    property FieldNames: MailMergeFieldNames readonly dispid 10;
    property DataFields: MailMergeDataFields readonly dispid 11;
    function FindRecord(const FindText: WideString; const Field: WideString): WordBool; dispid 101;
  end;

// *********************************************************************//
// Interface: MailMergeFieldNames
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091C-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFieldNames = interface(IDispatch)
    ['{0002091C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): MailMergeFieldName; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldNamesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091C-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFieldNamesDisp = dispinterface
    ['{0002091C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): MailMergeFieldName; dispid 0;
  end;

// *********************************************************************//
// Interface: MailMergeFieldName
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091B-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFieldName = interface(IDispatch)
    ['{0002091B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldNameDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091B-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeFieldNameDisp = dispinterface
    ['{0002091B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Index: Integer readonly dispid 1;
  end;

// *********************************************************************//
// Interface: MailMergeDataFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091A-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDataFields = interface(IDispatch)
    ['{0002091A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): MailMergeDataField; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDataFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091A-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDataFieldsDisp = dispinterface
    ['{0002091A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): MailMergeDataField; dispid 0;
  end;

// *********************************************************************//
// Interface: MailMergeDataField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020919-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDataField = interface(IDispatch)
    ['{00020919-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Value: WideString; safecall;
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Value: WideString read Get_Value;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDataFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020919-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMergeDataFieldDisp = dispinterface
    ['{00020919-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Value: WideString readonly dispid 0;
    property Name: WideString readonly dispid 1;
    property Index: Integer readonly dispid 2;
  end;

// *********************************************************************//
// Interface: Envelope
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020918-0000-0000-C000-000000000046}
// *********************************************************************//
  Envelope = interface(IDispatch)
    ['{00020918-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Address: Range; safecall;
    function Get_ReturnAddress: Range; safecall;
    function Get_DefaultPrintBarCode: WordBool; safecall;
    procedure Set_DefaultPrintBarCode(prop: WordBool); safecall;
    function Get_DefaultPrintFIMA: WordBool; safecall;
    procedure Set_DefaultPrintFIMA(prop: WordBool); safecall;
    function Get_DefaultHeight: Single; safecall;
    procedure Set_DefaultHeight(prop: Single); safecall;
    function Get_DefaultWidth: Single; safecall;
    procedure Set_DefaultWidth(prop: Single); safecall;
    function Get_DefaultSize: WideString; safecall;
    procedure Set_DefaultSize(const prop: WideString); safecall;
    function Get_DefaultOmitReturnAddress: WordBool; safecall;
    procedure Set_DefaultOmitReturnAddress(prop: WordBool); safecall;
    function Get_FeedSource: WdPaperTray; safecall;
    procedure Set_FeedSource(prop: WdPaperTray); safecall;
    function Get_AddressFromLeft: Single; safecall;
    procedure Set_AddressFromLeft(prop: Single); safecall;
    function Get_AddressFromTop: Single; safecall;
    procedure Set_AddressFromTop(prop: Single); safecall;
    function Get_ReturnAddressFromLeft: Single; safecall;
    procedure Set_ReturnAddressFromLeft(prop: Single); safecall;
    function Get_ReturnAddressFromTop: Single; safecall;
    procedure Set_ReturnAddressFromTop(prop: Single); safecall;
    function Get_AddressStyle: Style; safecall;
    function Get_ReturnAddressStyle: Style; safecall;
    function Get_DefaultOrientation: WdEnvelopeOrientation; safecall;
    procedure Set_DefaultOrientation(prop: WdEnvelopeOrientation); safecall;
    function Get_DefaultFaceUp: WordBool; safecall;
    procedure Set_DefaultFaceUp(prop: WordBool); safecall;
    procedure Insert(var ExtractAddress: OleVariant; var Address: OleVariant; 
                     var AutoText: OleVariant; var OmitReturnAddress: OleVariant; 
                     var ReturnAddress: OleVariant; var ReturnAutoText: OleVariant; 
                     var PrintBarCode: OleVariant; var PrintFIMA: OleVariant; var Size: OleVariant; 
                     var Height: OleVariant; var Width: OleVariant; var FeedSource: OleVariant; 
                     var AddressFromLeft: OleVariant; var AddressFromTop: OleVariant; 
                     var ReturnAddressFromLeft: OleVariant; var ReturnAddressFromTop: OleVariant; 
                     var DefaultFaceUp: OleVariant; var DefaultOrientation: OleVariant); safecall;
    procedure PrintOut(var ExtractAddress: OleVariant; var Address: OleVariant; 
                       var AutoText: OleVariant; var OmitReturnAddress: OleVariant; 
                       var ReturnAddress: OleVariant; var ReturnAutoText: OleVariant; 
                       var PrintBarCode: OleVariant; var PrintFIMA: OleVariant; 
                       var Size: OleVariant; var Height: OleVariant; var Width: OleVariant; 
                       var FeedSource: OleVariant; var AddressFromLeft: OleVariant; 
                       var AddressFromTop: OleVariant; var ReturnAddressFromLeft: OleVariant; 
                       var ReturnAddressFromTop: OleVariant; var DefaultFaceUp: OleVariant; 
                       var DefaultOrientation: OleVariant); safecall;
    procedure UpdateDocument; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Address: Range read Get_Address;
    property ReturnAddress: Range read Get_ReturnAddress;
    property DefaultPrintBarCode: WordBool read Get_DefaultPrintBarCode write Set_DefaultPrintBarCode;
    property DefaultPrintFIMA: WordBool read Get_DefaultPrintFIMA write Set_DefaultPrintFIMA;
    property DefaultHeight: Single read Get_DefaultHeight write Set_DefaultHeight;
    property DefaultWidth: Single read Get_DefaultWidth write Set_DefaultWidth;
    property DefaultSize: WideString read Get_DefaultSize write Set_DefaultSize;
    property DefaultOmitReturnAddress: WordBool read Get_DefaultOmitReturnAddress write Set_DefaultOmitReturnAddress;
    property FeedSource: WdPaperTray read Get_FeedSource write Set_FeedSource;
    property AddressFromLeft: Single read Get_AddressFromLeft write Set_AddressFromLeft;
    property AddressFromTop: Single read Get_AddressFromTop write Set_AddressFromTop;
    property ReturnAddressFromLeft: Single read Get_ReturnAddressFromLeft write Set_ReturnAddressFromLeft;
    property ReturnAddressFromTop: Single read Get_ReturnAddressFromTop write Set_ReturnAddressFromTop;
    property AddressStyle: Style read Get_AddressStyle;
    property ReturnAddressStyle: Style read Get_ReturnAddressStyle;
    property DefaultOrientation: WdEnvelopeOrientation read Get_DefaultOrientation write Set_DefaultOrientation;
    property DefaultFaceUp: WordBool read Get_DefaultFaceUp write Set_DefaultFaceUp;
  end;

// *********************************************************************//
// DispIntf:  EnvelopeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020918-0000-0000-C000-000000000046}
// *********************************************************************//
  EnvelopeDisp = dispinterface
    ['{00020918-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Address: Range readonly dispid 1;
    property ReturnAddress: Range readonly dispid 2;
    property DefaultPrintBarCode: WordBool dispid 4;
    property DefaultPrintFIMA: WordBool dispid 5;
    property DefaultHeight: Single dispid 6;
    property DefaultWidth: Single dispid 7;
    property DefaultSize: WideString dispid 0;
    property DefaultOmitReturnAddress: WordBool dispid 9;
    property FeedSource: WdPaperTray dispid 12;
    property AddressFromLeft: Single dispid 13;
    property AddressFromTop: Single dispid 14;
    property ReturnAddressFromLeft: Single dispid 15;
    property ReturnAddressFromTop: Single dispid 16;
    property AddressStyle: Style readonly dispid 17;
    property ReturnAddressStyle: Style readonly dispid 18;
    property DefaultOrientation: WdEnvelopeOrientation dispid 19;
    property DefaultFaceUp: WordBool dispid 20;
    procedure Insert(var ExtractAddress: OleVariant; var Address: OleVariant; 
                     var AutoText: OleVariant; var OmitReturnAddress: OleVariant; 
                     var ReturnAddress: OleVariant; var ReturnAutoText: OleVariant; 
                     var PrintBarCode: OleVariant; var PrintFIMA: OleVariant; var Size: OleVariant; 
                     var Height: OleVariant; var Width: OleVariant; var FeedSource: OleVariant; 
                     var AddressFromLeft: OleVariant; var AddressFromTop: OleVariant; 
                     var ReturnAddressFromLeft: OleVariant; var ReturnAddressFromTop: OleVariant; 
                     var DefaultFaceUp: OleVariant; var DefaultOrientation: OleVariant); dispid 101;
    procedure PrintOut(var ExtractAddress: OleVariant; var Address: OleVariant; 
                       var AutoText: OleVariant; var OmitReturnAddress: OleVariant; 
                       var ReturnAddress: OleVariant; var ReturnAutoText: OleVariant; 
                       var PrintBarCode: OleVariant; var PrintFIMA: OleVariant; 
                       var Size: OleVariant; var Height: OleVariant; var Width: OleVariant; 
                       var FeedSource: OleVariant; var AddressFromLeft: OleVariant; 
                       var AddressFromTop: OleVariant; var ReturnAddressFromLeft: OleVariant; 
                       var ReturnAddressFromTop: OleVariant; var DefaultFaceUp: OleVariant; 
                       var DefaultOrientation: OleVariant); dispid 102;
    procedure UpdateDocument; dispid 103;
  end;

// *********************************************************************//
// Interface: MailingLabel
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020917-0000-0000-C000-000000000046}
// *********************************************************************//
  MailingLabel = interface(IDispatch)
    ['{00020917-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_DefaultPrintBarCode: WordBool; safecall;
    procedure Set_DefaultPrintBarCode(prop: WordBool); safecall;
    function Get_DefaultLaserTray: WdPaperTray; safecall;
    procedure Set_DefaultLaserTray(prop: WdPaperTray); safecall;
    function Get_CustomLabels: CustomLabels; safecall;
    function Get_DefaultLabelName: WideString; safecall;
    procedure Set_DefaultLabelName(const prop: WideString); safecall;
    function CreateNewDocument(var Name: OleVariant; var Address: OleVariant; 
                               var AutoText: OleVariant; var ExtractAddress: OleVariant; 
                               var LaserTray: OleVariant): Document; safecall;
    procedure PrintOut(var Name: OleVariant; var Address: OleVariant; 
                       var ExtractAddress: OleVariant; var LaserTray: OleVariant; 
                       var SingleLabel: OleVariant; var Row: OleVariant; var Column: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property DefaultPrintBarCode: WordBool read Get_DefaultPrintBarCode write Set_DefaultPrintBarCode;
    property DefaultLaserTray: WdPaperTray read Get_DefaultLaserTray write Set_DefaultLaserTray;
    property CustomLabels: CustomLabels read Get_CustomLabels;
    property DefaultLabelName: WideString read Get_DefaultLabelName write Set_DefaultLabelName;
  end;

// *********************************************************************//
// DispIntf:  MailingLabelDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020917-0000-0000-C000-000000000046}
// *********************************************************************//
  MailingLabelDisp = dispinterface
    ['{00020917-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property DefaultPrintBarCode: WordBool dispid 2;
    property DefaultLaserTray: WdPaperTray dispid 4;
    property CustomLabels: CustomLabels readonly dispid 8;
    property DefaultLabelName: WideString dispid 9;
    function CreateNewDocument(var Name: OleVariant; var Address: OleVariant; 
                               var AutoText: OleVariant; var ExtractAddress: OleVariant; 
                               var LaserTray: OleVariant): Document; dispid 101;
    procedure PrintOut(var Name: OleVariant; var Address: OleVariant; 
                       var ExtractAddress: OleVariant; var LaserTray: OleVariant; 
                       var SingleLabel: OleVariant; var Row: OleVariant; var Column: OleVariant); dispid 102;
  end;

// *********************************************************************//
// Interface: CustomLabels
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020916-0000-0000-C000-000000000046}
// *********************************************************************//
  CustomLabels = interface(IDispatch)
    ['{00020916-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): CustomLabel; safecall;
    function Add(const Name: WideString; var DotMatrix: OleVariant): CustomLabel; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  CustomLabelsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020916-0000-0000-C000-000000000046}
// *********************************************************************//
  CustomLabelsDisp = dispinterface
    ['{00020916-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): CustomLabel; dispid 0;
    function Add(const Name: WideString; var DotMatrix: OleVariant): CustomLabel; dispid 101;
  end;

// *********************************************************************//
// Interface: CustomLabel
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020915-0000-0000-C000-000000000046}
// *********************************************************************//
  CustomLabel = interface(IDispatch)
    ['{00020915-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_TopMargin: Single; safecall;
    procedure Set_TopMargin(prop: Single); safecall;
    function Get_SideMargin: Single; safecall;
    procedure Set_SideMargin(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_VerticalPitch: Single; safecall;
    procedure Set_VerticalPitch(prop: Single); safecall;
    function Get_HorizontalPitch: Single; safecall;
    procedure Set_HorizontalPitch(prop: Single); safecall;
    function Get_NumberAcross: Integer; safecall;
    procedure Set_NumberAcross(prop: Integer); safecall;
    function Get_NumberDown: Integer; safecall;
    procedure Set_NumberDown(prop: Integer); safecall;
    function Get_DotMatrix: WordBool; safecall;
    function Get_PageSize: WdCustomLabelPageSize; safecall;
    procedure Set_PageSize(prop: WdCustomLabelPageSize); safecall;
    function Get_Valid: WordBool; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name write Set_Name;
    property TopMargin: Single read Get_TopMargin write Set_TopMargin;
    property SideMargin: Single read Get_SideMargin write Set_SideMargin;
    property Height: Single read Get_Height write Set_Height;
    property Width: Single read Get_Width write Set_Width;
    property VerticalPitch: Single read Get_VerticalPitch write Set_VerticalPitch;
    property HorizontalPitch: Single read Get_HorizontalPitch write Set_HorizontalPitch;
    property NumberAcross: Integer read Get_NumberAcross write Set_NumberAcross;
    property NumberDown: Integer read Get_NumberDown write Set_NumberDown;
    property DotMatrix: WordBool read Get_DotMatrix;
    property PageSize: WdCustomLabelPageSize read Get_PageSize write Set_PageSize;
    property Valid: WordBool read Get_Valid;
  end;

// *********************************************************************//
// DispIntf:  CustomLabelDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020915-0000-0000-C000-000000000046}
// *********************************************************************//
  CustomLabelDisp = dispinterface
    ['{00020915-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString dispid 2;
    property TopMargin: Single dispid 3;
    property SideMargin: Single dispid 4;
    property Height: Single dispid 5;
    property Width: Single dispid 6;
    property VerticalPitch: Single dispid 7;
    property HorizontalPitch: Single dispid 8;
    property NumberAcross: Integer dispid 9;
    property NumberDown: Integer dispid 10;
    property DotMatrix: WordBool readonly dispid 11;
    property PageSize: WdCustomLabelPageSize dispid 12;
    property Valid: WordBool readonly dispid 13;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: TablesOfContents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020914-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfContents = interface(IDispatch)
    ['{00020914-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Format: WdTocFormat; safecall;
    procedure Set_Format(prop: WdTocFormat); safecall;
    function Item(Index: Integer): TableOfContents; safecall;
    function AddOld(const Range: Range; var UseHeadingStyles: OleVariant; 
                    var UpperHeadingLevel: OleVariant; var LowerHeadingLevel: OleVariant; 
                    var UseFields: OleVariant; var TableID: OleVariant; 
                    var RightAlignPageNumbers: OleVariant; var IncludePageNumbers: OleVariant; 
                    var AddedStyles: OleVariant): TableOfContents; safecall;
    function MarkEntry(const Range: Range; var Entry: OleVariant; var EntryAutoText: OleVariant; 
                       var TableID: OleVariant; var Level: OleVariant): Field; safecall;
    function Add(const Range: Range; var UseHeadingStyles: OleVariant; 
                 var UpperHeadingLevel: OleVariant; var LowerHeadingLevel: OleVariant; 
                 var UseFields: OleVariant; var TableID: OleVariant; 
                 var RightAlignPageNumbers: OleVariant; var IncludePageNumbers: OleVariant; 
                 var AddedStyles: OleVariant; var UseHyperlinks: OleVariant; 
                 var HidePageNumbersInWeb: OleVariant): TableOfContents; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Format: WdTocFormat read Get_Format write Set_Format;
  end;

// *********************************************************************//
// DispIntf:  TablesOfContentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020914-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfContentsDisp = dispinterface
    ['{00020914-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Format: WdTocFormat dispid 2;
    function Item(Index: Integer): TableOfContents; dispid 0;
    function AddOld(const Range: Range; var UseHeadingStyles: OleVariant; 
                    var UpperHeadingLevel: OleVariant; var LowerHeadingLevel: OleVariant; 
                    var UseFields: OleVariant; var TableID: OleVariant; 
                    var RightAlignPageNumbers: OleVariant; var IncludePageNumbers: OleVariant; 
                    var AddedStyles: OleVariant): TableOfContents; dispid 100;
    function MarkEntry(const Range: Range; var Entry: OleVariant; var EntryAutoText: OleVariant; 
                       var TableID: OleVariant; var Level: OleVariant): Field; dispid 101;
    function Add(const Range: Range; var UseHeadingStyles: OleVariant; 
                 var UpperHeadingLevel: OleVariant; var LowerHeadingLevel: OleVariant; 
                 var UseFields: OleVariant; var TableID: OleVariant; 
                 var RightAlignPageNumbers: OleVariant; var IncludePageNumbers: OleVariant; 
                 var AddedStyles: OleVariant; var UseHyperlinks: OleVariant; 
                 var HidePageNumbersInWeb: OleVariant): TableOfContents; dispid 102;
  end;

// *********************************************************************//
// Interface: TableOfContents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020913-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfContents = interface(IDispatch)
    ['{00020913-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_UseHeadingStyles: WordBool; safecall;
    procedure Set_UseHeadingStyles(prop: WordBool); safecall;
    function Get_UseFields: WordBool; safecall;
    procedure Set_UseFields(prop: WordBool); safecall;
    function Get_UpperHeadingLevel: Integer; safecall;
    procedure Set_UpperHeadingLevel(prop: Integer); safecall;
    function Get_LowerHeadingLevel: Integer; safecall;
    procedure Set_LowerHeadingLevel(prop: Integer); safecall;
    function Get_TableID: WideString; safecall;
    procedure Set_TableID(const prop: WideString); safecall;
    function Get_HeadingStyles: HeadingStyles; safecall;
    function Get_RightAlignPageNumbers: WordBool; safecall;
    procedure Set_RightAlignPageNumbers(prop: WordBool); safecall;
    function Get_IncludePageNumbers: WordBool; safecall;
    procedure Set_IncludePageNumbers(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    function Get_TabLeader: WdTabLeader; safecall;
    procedure Set_TabLeader(prop: WdTabLeader); safecall;
    procedure Delete; safecall;
    procedure UpdatePageNumbers; safecall;
    procedure Update; safecall;
    function Get_UseHyperlinks: WordBool; safecall;
    procedure Set_UseHyperlinks(prop: WordBool); safecall;
    function Get_HidePageNumbersInWeb: WordBool; safecall;
    procedure Set_HidePageNumbersInWeb(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property UseHeadingStyles: WordBool read Get_UseHeadingStyles write Set_UseHeadingStyles;
    property UseFields: WordBool read Get_UseFields write Set_UseFields;
    property UpperHeadingLevel: Integer read Get_UpperHeadingLevel write Set_UpperHeadingLevel;
    property LowerHeadingLevel: Integer read Get_LowerHeadingLevel write Set_LowerHeadingLevel;
    property TableID: WideString read Get_TableID write Set_TableID;
    property HeadingStyles: HeadingStyles read Get_HeadingStyles;
    property RightAlignPageNumbers: WordBool read Get_RightAlignPageNumbers write Set_RightAlignPageNumbers;
    property IncludePageNumbers: WordBool read Get_IncludePageNumbers write Set_IncludePageNumbers;
    property Range: Range read Get_Range;
    property TabLeader: WdTabLeader read Get_TabLeader write Set_TabLeader;
    property UseHyperlinks: WordBool read Get_UseHyperlinks write Set_UseHyperlinks;
    property HidePageNumbersInWeb: WordBool read Get_HidePageNumbersInWeb write Set_HidePageNumbersInWeb;
  end;

// *********************************************************************//
// DispIntf:  TableOfContentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020913-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfContentsDisp = dispinterface
    ['{00020913-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property UseHeadingStyles: WordBool dispid 1;
    property UseFields: WordBool dispid 2;
    property UpperHeadingLevel: Integer dispid 3;
    property LowerHeadingLevel: Integer dispid 4;
    property TableID: WideString dispid 5;
    property HeadingStyles: HeadingStyles readonly dispid 6;
    property RightAlignPageNumbers: WordBool dispid 7;
    property IncludePageNumbers: WordBool dispid 8;
    property Range: Range readonly dispid 9;
    property TabLeader: WdTabLeader dispid 10;
    procedure Delete; dispid 100;
    procedure UpdatePageNumbers; dispid 101;
    procedure Update; dispid 102;
    property UseHyperlinks: WordBool dispid 11;
    property HidePageNumbersInWeb: WordBool dispid 12;
  end;

// *********************************************************************//
// Interface: TablesOfAuthorities
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020912-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfAuthorities = interface(IDispatch)
    ['{00020912-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Format: WdToaFormat; safecall;
    procedure Set_Format(prop: WdToaFormat); safecall;
    function Item(Index: Integer): TableOfAuthorities; safecall;
    function Add(const Range: Range; var Category: OleVariant; var Bookmark: OleVariant; 
                 var Passim: OleVariant; var KeepEntryFormatting: OleVariant; 
                 var Separator: OleVariant; var IncludeSequenceName: OleVariant; 
                 var EntrySeparator: OleVariant; var PageRangeSeparator: OleVariant; 
                 var IncludeCategoryHeader: OleVariant; var PageNumberSeparator: OleVariant): TableOfAuthorities; safecall;
    procedure NextCitation(const ShortCitation: WideString); safecall;
    function MarkCitation(const Range: Range; const ShortCitation: WideString; 
                          var LongCitation: OleVariant; var LongCitationAutoText: OleVariant; 
                          var Category: OleVariant): Field; safecall;
    procedure MarkAllCitations(const ShortCitation: WideString; var LongCitation: OleVariant; 
                               var LongCitationAutoText: OleVariant; var Category: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Format: WdToaFormat read Get_Format write Set_Format;
  end;

// *********************************************************************//
// DispIntf:  TablesOfAuthoritiesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020912-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfAuthoritiesDisp = dispinterface
    ['{00020912-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Format: WdToaFormat dispid 2;
    function Item(Index: Integer): TableOfAuthorities; dispid 0;
    function Add(const Range: Range; var Category: OleVariant; var Bookmark: OleVariant; 
                 var Passim: OleVariant; var KeepEntryFormatting: OleVariant; 
                 var Separator: OleVariant; var IncludeSequenceName: OleVariant; 
                 var EntrySeparator: OleVariant; var PageRangeSeparator: OleVariant; 
                 var IncludeCategoryHeader: OleVariant; var PageNumberSeparator: OleVariant): TableOfAuthorities; dispid 100;
    procedure NextCitation(const ShortCitation: WideString); dispid 103;
    function MarkCitation(const Range: Range; const ShortCitation: WideString; 
                          var LongCitation: OleVariant; var LongCitationAutoText: OleVariant; 
                          var Category: OleVariant): Field; dispid 101;
    procedure MarkAllCitations(const ShortCitation: WideString; var LongCitation: OleVariant; 
                               var LongCitationAutoText: OleVariant; var Category: OleVariant); dispid 102;
  end;

// *********************************************************************//
// Interface: TableOfAuthorities
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020911-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfAuthorities = interface(IDispatch)
    ['{00020911-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Passim: WordBool; safecall;
    procedure Set_Passim(prop: WordBool); safecall;
    function Get_KeepEntryFormatting: WordBool; safecall;
    procedure Set_KeepEntryFormatting(prop: WordBool); safecall;
    function Get_Category: Integer; safecall;
    procedure Set_Category(prop: Integer); safecall;
    function Get_Bookmark: WideString; safecall;
    procedure Set_Bookmark(const prop: WideString); safecall;
    function Get_Separator: WideString; safecall;
    procedure Set_Separator(const prop: WideString); safecall;
    function Get_IncludeSequenceName: WideString; safecall;
    procedure Set_IncludeSequenceName(const prop: WideString); safecall;
    function Get_EntrySeparator: WideString; safecall;
    procedure Set_EntrySeparator(const prop: WideString); safecall;
    function Get_PageRangeSeparator: WideString; safecall;
    procedure Set_PageRangeSeparator(const prop: WideString); safecall;
    function Get_IncludeCategoryHeader: WordBool; safecall;
    procedure Set_IncludeCategoryHeader(prop: WordBool); safecall;
    function Get_PageNumberSeparator: WideString; safecall;
    procedure Set_PageNumberSeparator(const prop: WideString); safecall;
    function Get_Range: Range; safecall;
    function Get_TabLeader: WdTabLeader; safecall;
    procedure Set_TabLeader(prop: WdTabLeader); safecall;
    procedure Delete; safecall;
    procedure Update; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Passim: WordBool read Get_Passim write Set_Passim;
    property KeepEntryFormatting: WordBool read Get_KeepEntryFormatting write Set_KeepEntryFormatting;
    property Category: Integer read Get_Category write Set_Category;
    property Bookmark: WideString read Get_Bookmark write Set_Bookmark;
    property Separator: WideString read Get_Separator write Set_Separator;
    property IncludeSequenceName: WideString read Get_IncludeSequenceName write Set_IncludeSequenceName;
    property EntrySeparator: WideString read Get_EntrySeparator write Set_EntrySeparator;
    property PageRangeSeparator: WideString read Get_PageRangeSeparator write Set_PageRangeSeparator;
    property IncludeCategoryHeader: WordBool read Get_IncludeCategoryHeader write Set_IncludeCategoryHeader;
    property PageNumberSeparator: WideString read Get_PageNumberSeparator write Set_PageNumberSeparator;
    property Range: Range read Get_Range;
    property TabLeader: WdTabLeader read Get_TabLeader write Set_TabLeader;
  end;

// *********************************************************************//
// DispIntf:  TableOfAuthoritiesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020911-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfAuthoritiesDisp = dispinterface
    ['{00020911-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Passim: WordBool dispid 1;
    property KeepEntryFormatting: WordBool dispid 2;
    property Category: Integer dispid 3;
    property Bookmark: WideString dispid 4;
    property Separator: WideString dispid 5;
    property IncludeSequenceName: WideString dispid 6;
    property EntrySeparator: WideString dispid 7;
    property PageRangeSeparator: WideString dispid 8;
    property IncludeCategoryHeader: WordBool dispid 9;
    property PageNumberSeparator: WideString dispid 10;
    property Range: Range readonly dispid 11;
    property TabLeader: WdTabLeader dispid 12;
    procedure Delete; dispid 100;
    procedure Update; dispid 101;
  end;

// *********************************************************************//
// Interface: Dialogs
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020910-0000-0000-C000-000000000046}
// *********************************************************************//
  Dialogs = interface(IDispatch)
    ['{00020910-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: WdWordDialog): Dialog; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  DialogsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020910-0000-0000-C000-000000000046}
// *********************************************************************//
  DialogsDisp = dispinterface
    ['{00020910-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: WdWordDialog): Dialog; dispid 0;
  end;

// *********************************************************************//
// Interface: Dialog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {000209B8-0000-0000-C000-000000000046}
// *********************************************************************//
  Dialog = interface(IDispatch)
    ['{000209B8-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_DefaultTab: WdWordDialogTab; safecall;
    procedure Set_DefaultTab(prop: WdWordDialogTab); safecall;
    function Get_Type_: WdWordDialog; safecall;
    function Show(var TimeOut: OleVariant): Integer; safecall;
    function Display(var TimeOut: OleVariant): Integer; safecall;
    procedure Execute; safecall;
    procedure Update; safecall;
    function Get_CommandName: WideString; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property DefaultTab: WdWordDialogTab read Get_DefaultTab write Set_DefaultTab;
    property Type_: WdWordDialog read Get_Type_;
    property CommandName: WideString read Get_CommandName;
  end;

// *********************************************************************//
// DispIntf:  DialogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {000209B8-0000-0000-C000-000000000046}
// *********************************************************************//
  DialogDisp = dispinterface
    ['{000209B8-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 32003;
    property Creator: Integer readonly dispid 32004;
    property Parent: IDispatch readonly dispid 32005;
    property DefaultTab: WdWordDialogTab dispid 32002;
    property Type_: WdWordDialog readonly dispid 0;
    function Show(var TimeOut: OleVariant): Integer; dispid 336;
    function Display(var TimeOut: OleVariant): Integer; dispid 338;
    procedure Execute; dispid 32001;
    procedure Update; dispid 302;
    property CommandName: WideString readonly dispid 32006;
  end;

// *********************************************************************//
// Interface: PageSetup
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020971-0000-0000-C000-000000000046}
// *********************************************************************//
  PageSetup = interface(IDispatch)
    ['{00020971-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_TopMargin: Single; safecall;
    procedure Set_TopMargin(prop: Single); safecall;
    function Get_BottomMargin: Single; safecall;
    procedure Set_BottomMargin(prop: Single); safecall;
    function Get_LeftMargin: Single; safecall;
    procedure Set_LeftMargin(prop: Single); safecall;
    function Get_RightMargin: Single; safecall;
    procedure Set_RightMargin(prop: Single); safecall;
    function Get_Gutter: Single; safecall;
    procedure Set_Gutter(prop: Single); safecall;
    function Get_PageWidth: Single; safecall;
    procedure Set_PageWidth(prop: Single); safecall;
    function Get_PageHeight: Single; safecall;
    procedure Set_PageHeight(prop: Single); safecall;
    function Get_Orientation: WdOrientation; safecall;
    procedure Set_Orientation(prop: WdOrientation); safecall;
    function Get_FirstPageTray: WdPaperTray; safecall;
    procedure Set_FirstPageTray(prop: WdPaperTray); safecall;
    function Get_OtherPagesTray: WdPaperTray; safecall;
    procedure Set_OtherPagesTray(prop: WdPaperTray); safecall;
    function Get_VerticalAlignment: WdVerticalAlignment; safecall;
    procedure Set_VerticalAlignment(prop: WdVerticalAlignment); safecall;
    function Get_MirrorMargins: Integer; safecall;
    procedure Set_MirrorMargins(prop: Integer); safecall;
    function Get_HeaderDistance: Single; safecall;
    procedure Set_HeaderDistance(prop: Single); safecall;
    function Get_FooterDistance: Single; safecall;
    procedure Set_FooterDistance(prop: Single); safecall;
    function Get_SectionStart: WdSectionStart; safecall;
    procedure Set_SectionStart(prop: WdSectionStart); safecall;
    function Get_OddAndEvenPagesHeaderFooter: Integer; safecall;
    procedure Set_OddAndEvenPagesHeaderFooter(prop: Integer); safecall;
    function Get_DifferentFirstPageHeaderFooter: Integer; safecall;
    procedure Set_DifferentFirstPageHeaderFooter(prop: Integer); safecall;
    function Get_SuppressEndnotes: Integer; safecall;
    procedure Set_SuppressEndnotes(prop: Integer); safecall;
    function Get_LineNumbering: LineNumbering; safecall;
    procedure Set_LineNumbering(const prop: LineNumbering); safecall;
    function Get_TextColumns: TextColumns; safecall;
    procedure Set_TextColumns(const prop: TextColumns); safecall;
    function Get_PaperSize: WdPaperSize; safecall;
    procedure Set_PaperSize(prop: WdPaperSize); safecall;
    function Get_TwoPagesOnOne: WordBool; safecall;
    procedure Set_TwoPagesOnOne(prop: WordBool); safecall;
    function Get_GutterOnTop: WordBool; safecall;
    procedure Set_GutterOnTop(prop: WordBool); safecall;
    function Get_CharsLine: Single; safecall;
    procedure Set_CharsLine(prop: Single); safecall;
    function Get_LinesPage: Single; safecall;
    procedure Set_LinesPage(prop: Single); safecall;
    function Get_ShowGrid: WordBool; safecall;
    procedure Set_ShowGrid(prop: WordBool); safecall;
    procedure TogglePortrait; safecall;
    procedure SetAsTemplateDefault; safecall;
    function Get_GutterStyle: WdGutterStyleOld; safecall;
    procedure Set_GutterStyle(prop: WdGutterStyleOld); safecall;
    function Get_SectionDirection: WdSectionDirection; safecall;
    procedure Set_SectionDirection(prop: WdSectionDirection); safecall;
    function Get_LayoutMode: WdLayoutMode; safecall;
    procedure Set_LayoutMode(prop: WdLayoutMode); safecall;
    function Get_GutterPos: WdGutterStyle; safecall;
    procedure Set_GutterPos(prop: WdGutterStyle); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property TopMargin: Single read Get_TopMargin write Set_TopMargin;
    property BottomMargin: Single read Get_BottomMargin write Set_BottomMargin;
    property LeftMargin: Single read Get_LeftMargin write Set_LeftMargin;
    property RightMargin: Single read Get_RightMargin write Set_RightMargin;
    property Gutter: Single read Get_Gutter write Set_Gutter;
    property PageWidth: Single read Get_PageWidth write Set_PageWidth;
    property PageHeight: Single read Get_PageHeight write Set_PageHeight;
    property Orientation: WdOrientation read Get_Orientation write Set_Orientation;
    property FirstPageTray: WdPaperTray read Get_FirstPageTray write Set_FirstPageTray;
    property OtherPagesTray: WdPaperTray read Get_OtherPagesTray write Set_OtherPagesTray;
    property VerticalAlignment: WdVerticalAlignment read Get_VerticalAlignment write Set_VerticalAlignment;
    property MirrorMargins: Integer read Get_MirrorMargins write Set_MirrorMargins;
    property HeaderDistance: Single read Get_HeaderDistance write Set_HeaderDistance;
    property FooterDistance: Single read Get_FooterDistance write Set_FooterDistance;
    property SectionStart: WdSectionStart read Get_SectionStart write Set_SectionStart;
    property OddAndEvenPagesHeaderFooter: Integer read Get_OddAndEvenPagesHeaderFooter write Set_OddAndEvenPagesHeaderFooter;
    property DifferentFirstPageHeaderFooter: Integer read Get_DifferentFirstPageHeaderFooter write Set_DifferentFirstPageHeaderFooter;
    property SuppressEndnotes: Integer read Get_SuppressEndnotes write Set_SuppressEndnotes;
    property LineNumbering: LineNumbering read Get_LineNumbering write Set_LineNumbering;
    property TextColumns: TextColumns read Get_TextColumns write Set_TextColumns;
    property PaperSize: WdPaperSize read Get_PaperSize write Set_PaperSize;
    property TwoPagesOnOne: WordBool read Get_TwoPagesOnOne write Set_TwoPagesOnOne;
    property GutterOnTop: WordBool read Get_GutterOnTop write Set_GutterOnTop;
    property CharsLine: Single read Get_CharsLine write Set_CharsLine;
    property LinesPage: Single read Get_LinesPage write Set_LinesPage;
    property ShowGrid: WordBool read Get_ShowGrid write Set_ShowGrid;
    property GutterStyle: WdGutterStyleOld read Get_GutterStyle write Set_GutterStyle;
    property SectionDirection: WdSectionDirection read Get_SectionDirection write Set_SectionDirection;
    property LayoutMode: WdLayoutMode read Get_LayoutMode write Set_LayoutMode;
    property GutterPos: WdGutterStyle read Get_GutterPos write Set_GutterPos;
  end;

// *********************************************************************//
// DispIntf:  PageSetupDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020971-0000-0000-C000-000000000046}
// *********************************************************************//
  PageSetupDisp = dispinterface
    ['{00020971-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property TopMargin: Single dispid 100;
    property BottomMargin: Single dispid 101;
    property LeftMargin: Single dispid 102;
    property RightMargin: Single dispid 103;
    property Gutter: Single dispid 104;
    property PageWidth: Single dispid 105;
    property PageHeight: Single dispid 106;
    property Orientation: WdOrientation dispid 107;
    property FirstPageTray: WdPaperTray dispid 108;
    property OtherPagesTray: WdPaperTray dispid 109;
    property VerticalAlignment: WdVerticalAlignment dispid 110;
    property MirrorMargins: Integer dispid 111;
    property HeaderDistance: Single dispid 112;
    property FooterDistance: Single dispid 113;
    property SectionStart: WdSectionStart dispid 114;
    property OddAndEvenPagesHeaderFooter: Integer dispid 115;
    property DifferentFirstPageHeaderFooter: Integer dispid 116;
    property SuppressEndnotes: Integer dispid 117;
    property LineNumbering: LineNumbering dispid 118;
    property TextColumns: TextColumns dispid 119;
    property PaperSize: WdPaperSize dispid 120;
    property TwoPagesOnOne: WordBool dispid 121;
    property GutterOnTop: WordBool dispid 122;
    property CharsLine: Single dispid 123;
    property LinesPage: Single dispid 124;
    property ShowGrid: WordBool dispid 128;
    procedure TogglePortrait; dispid 201;
    procedure SetAsTemplateDefault; dispid 202;
    property GutterStyle: WdGutterStyleOld dispid 129;
    property SectionDirection: WdSectionDirection dispid 130;
    property LayoutMode: WdLayoutMode dispid 131;
    property GutterPos: WdGutterStyle dispid 1222;
  end;

// *********************************************************************//
// Interface: LineNumbering
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020972-0000-0000-C000-000000000046}
// *********************************************************************//
  LineNumbering = interface(IDispatch)
    ['{00020972-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_RestartMode: WdNumberingRule; safecall;
    procedure Set_RestartMode(prop: WdNumberingRule); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_DistanceFromText: Single; safecall;
    procedure Set_DistanceFromText(prop: Single); safecall;
    function Get_CountBy: Integer; safecall;
    procedure Set_CountBy(prop: Integer); safecall;
    function Get_Active: Integer; safecall;
    procedure Set_Active(prop: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property RestartMode: WdNumberingRule read Get_RestartMode write Set_RestartMode;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property DistanceFromText: Single read Get_DistanceFromText write Set_DistanceFromText;
    property CountBy: Integer read Get_CountBy write Set_CountBy;
    property Active: Integer read Get_Active write Set_Active;
  end;

// *********************************************************************//
// DispIntf:  LineNumberingDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020972-0000-0000-C000-000000000046}
// *********************************************************************//
  LineNumberingDisp = dispinterface
    ['{00020972-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property RestartMode: WdNumberingRule dispid 100;
    property StartingNumber: Integer dispid 101;
    property DistanceFromText: Single dispid 102;
    property CountBy: Integer dispid 103;
    property Active: Integer dispid 104;
  end;

// *********************************************************************//
// Interface: TextColumns
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020973-0000-0000-C000-000000000046}
// *********************************************************************//
  TextColumns = interface(IDispatch)
    ['{00020973-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_EvenlySpaced: Integer; safecall;
    procedure Set_EvenlySpaced(prop: Integer); safecall;
    function Get_LineBetween: Integer; safecall;
    procedure Set_LineBetween(prop: Integer); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Spacing: Single; safecall;
    procedure Set_Spacing(prop: Single); safecall;
    function Item(Index: Integer): TextColumn; safecall;
    function Add(var Width: OleVariant; var Spacing: OleVariant; var EvenlySpaced: OleVariant): TextColumn; safecall;
    procedure SetCount(NumColumns: Integer); safecall;
    function Get_FlowDirection: WdFlowDirection; safecall;
    procedure Set_FlowDirection(prop: WdFlowDirection); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property EvenlySpaced: Integer read Get_EvenlySpaced write Set_EvenlySpaced;
    property LineBetween: Integer read Get_LineBetween write Set_LineBetween;
    property Width: Single read Get_Width write Set_Width;
    property Spacing: Single read Get_Spacing write Set_Spacing;
    property FlowDirection: WdFlowDirection read Get_FlowDirection write Set_FlowDirection;
  end;

// *********************************************************************//
// DispIntf:  TextColumnsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020973-0000-0000-C000-000000000046}
// *********************************************************************//
  TextColumnsDisp = dispinterface
    ['{00020973-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property EvenlySpaced: Integer dispid 100;
    property LineBetween: Integer dispid 101;
    property Width: Single dispid 102;
    property Spacing: Single dispid 103;
    function Item(Index: Integer): TextColumn; dispid 0;
    function Add(var Width: OleVariant; var Spacing: OleVariant; var EvenlySpaced: OleVariant): TextColumn; dispid 201;
    procedure SetCount(NumColumns: Integer); dispid 202;
    property FlowDirection: WdFlowDirection dispid 104;
  end;

// *********************************************************************//
// Interface: TextColumn
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020974-0000-0000-C000-000000000046}
// *********************************************************************//
  TextColumn = interface(IDispatch)
    ['{00020974-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Width: Single read Get_Width write Set_Width;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
  end;

// *********************************************************************//
// DispIntf:  TextColumnDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020974-0000-0000-C000-000000000046}
// *********************************************************************//
  TextColumnDisp = dispinterface
    ['{00020974-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Width: Single dispid 100;
    property SpaceAfter: Single dispid 101;
  end;

// *********************************************************************//
// Interface: Selection
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020975-0000-0000-C000-000000000046}
// *********************************************************************//
  Selection = interface(IDispatch)
    ['{00020975-0000-0000-C000-000000000046}']
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_FormattedText: Range; safecall;
    procedure Set_FormattedText(const prop: Range); safecall;
    function Get_Start: Integer; safecall;
    procedure Set_Start(prop: Integer); safecall;
    function Get_End_: Integer; safecall;
    procedure Set_End_(prop: Integer); safecall;
    function Get_Font: Font; safecall;
    procedure Set_Font(const prop: Font); safecall;
    function Get_Type_: WdSelectionType; safecall;
    function Get_StoryType: WdStoryType; safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Tables: Tables; safecall;
    function Get_Words: Words; safecall;
    function Get_Sentences: Sentences; safecall;
    function Get_Characters: Characters; safecall;
    function Get_Footnotes: Footnotes; safecall;
    function Get_Endnotes: Endnotes; safecall;
    function Get_Comments: Comments; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Sections: Sections; safecall;
    function Get_Paragraphs: Paragraphs; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_Fields: Fields; safecall;
    function Get_FormFields: FormFields; safecall;
    function Get_Frames: Frames; safecall;
    function Get_ParagraphFormat: ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: ParagraphFormat); safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_Bookmarks: Bookmarks; safecall;
    function Get_StoryLength: Integer; safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_LanguageIDFarEast: WdLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WdLanguageID); safecall;
    function Get_LanguageIDOther: WdLanguageID; safecall;
    procedure Set_LanguageIDOther(prop: WdLanguageID); safecall;
    function Get_Hyperlinks: Hyperlinks; safecall;
    function Get_Columns: Columns; safecall;
    function Get_Rows: Rows; safecall;
    function Get_HeaderFooter: HeaderFooter; safecall;
    function Get_IsEndOfRowMark: WordBool; safecall;
    function Get_BookmarkID: Integer; safecall;
    function Get_PreviousBookmarkID: Integer; safecall;
    function Get_Find: Find; safecall;
    function Get_Range: Range; safecall;
    function Get_Information(Type_: WdInformation): OleVariant; safecall;
    function Get_Flags: WdSelectionFlags; safecall;
    procedure Set_Flags(prop: WdSelectionFlags); safecall;
    function Get_Active: WordBool; safecall;
    function Get_StartIsActive: WordBool; safecall;
    procedure Set_StartIsActive(prop: WordBool); safecall;
    function Get_IPAtEndOfLine: WordBool; safecall;
    function Get_ExtendMode: WordBool; safecall;
    procedure Set_ExtendMode(prop: WordBool); safecall;
    function Get_ColumnSelectMode: WordBool; safecall;
    procedure Set_ColumnSelectMode(prop: WordBool); safecall;
    function Get_Orientation: WdTextOrientation; safecall;
    procedure Set_Orientation(prop: WdTextOrientation); safecall;
    function Get_InlineShapes: InlineShapes; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Document: Document; safecall;
    function Get_ShapeRange: ShapeRange; safecall;
    procedure Select; safecall;
    procedure SetRange(Start: Integer; End_: Integer); safecall;
    procedure Collapse(var Direction: OleVariant); safecall;
    procedure InsertBefore(const Text: WideString); safecall;
    procedure InsertAfter(const Text: WideString); safecall;
    function Next(var Unit_: OleVariant; var Count: OleVariant): Range; safecall;
    function Previous(var Unit_: OleVariant; var Count: OleVariant): Range; safecall;
    function StartOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; safecall;
    function EndOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; safecall;
    function Move(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveStart(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveEnd(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveWhile(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveStartWhile(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveEndWhile(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveUntil(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveStartUntil(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    function MoveEndUntil(var Cset: OleVariant; var Count: OleVariant): Integer; safecall;
    procedure Cut; safecall;
    procedure Copy; safecall;
    procedure Paste; safecall;
    procedure InsertBreak(var Type_: OleVariant); safecall;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); safecall;
    function InStory(const Range: Range): WordBool; safecall;
    function InRange(const Range: Range): WordBool; safecall;
    function Delete(var Unit_: OleVariant; var Count: OleVariant): Integer; safecall;
    function Expand(var Unit_: OleVariant): Integer; safecall;
    procedure InsertParagraph; safecall;
    procedure InsertParagraphAfter; safecall;
    function ConvertToTableOld(var Separator: OleVariant; var NumRows: OleVariant; 
                               var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                               var Format: OleVariant; var ApplyBorders: OleVariant; 
                               var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                               var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                               var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                               var ApplyLastColumn: OleVariant; var AutoFit: OleVariant): Table; safecall;
    procedure InsertDateTimeOld(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                                var InsertAsFullWidth: OleVariant); safecall;
    procedure InsertSymbol(CharacterNumber: Integer; var Font: OleVariant; var Unicode: OleVariant; 
                           var Bias: OleVariant); safecall;
    procedure InsertCrossReference(var ReferenceType: OleVariant; ReferenceKind: WdReferenceKind; 
                                   var ReferenceItem: OleVariant; 
                                   var InsertAsHyperlink: OleVariant; 
                                   var IncludePosition: OleVariant); safecall;
    procedure InsertCaption(var Label_: OleVariant; var Title: OleVariant; 
                            var TitleAutoText: OleVariant; var Position: OleVariant); safecall;
    procedure CopyAsPicture; safecall;
    procedure SortOld(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                      var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                      var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                      var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                      var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                      var SortColumn: OleVariant; var Separator: OleVariant; 
                      var CaseSensitive: OleVariant; var LanguageID: OleVariant); safecall;
    procedure SortAscending; safecall;
    procedure SortDescending; safecall;
    function IsEqual(const Range: Range): WordBool; safecall;
    function Calculate: Single; safecall;
    function GoTo_(var What: OleVariant; var Which: OleVariant; var Count: OleVariant; 
                   var Name: OleVariant): Range; safecall;
    function GoToNext(What: WdGoToItem): Range; safecall;
    function GoToPrevious(What: WdGoToItem): Range; safecall;
    procedure PasteSpecial(var IconIndex: OleVariant; var Link: OleVariant; 
                           var Placement: OleVariant; var DisplayAsIcon: OleVariant; 
                           var DataType: OleVariant; var IconFileName: OleVariant; 
                           var IconLabel: OleVariant); safecall;
    function PreviousField: Field; safecall;
    function NextField: Field; safecall;
    procedure InsertParagraphBefore; safecall;
    procedure InsertCells(var ShiftCells: OleVariant); safecall;
    procedure Extend(var Character: OleVariant); safecall;
    procedure Shrink; safecall;
    function MoveLeft(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; safecall;
    function MoveRight(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; safecall;
    function MoveUp(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; safecall;
    function MoveDown(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; safecall;
    function HomeKey(var Unit_: OleVariant; var Extend: OleVariant): Integer; safecall;
    function EndKey(var Unit_: OleVariant; var Extend: OleVariant): Integer; safecall;
    procedure EscapeKey; safecall;
    procedure TypeText(const Text: WideString); safecall;
    procedure CopyFormat; safecall;
    procedure PasteFormat; safecall;
    procedure TypeParagraph; safecall;
    procedure TypeBackspace; safecall;
    procedure NextSubdocument; safecall;
    procedure PreviousSubdocument; safecall;
    procedure SelectColumn; safecall;
    procedure SelectCurrentFont; safecall;
    procedure SelectCurrentAlignment; safecall;
    procedure SelectCurrentSpacing; safecall;
    procedure SelectCurrentIndent; safecall;
    procedure SelectCurrentTabs; safecall;
    procedure SelectCurrentColor; safecall;
    procedure CreateTextbox; safecall;
    procedure WholeStory; safecall;
    procedure SelectRow; safecall;
    procedure SplitTable; safecall;
    procedure InsertRows(var NumRows: OleVariant); safecall;
    procedure InsertColumns; safecall;
    procedure InsertFormula(var Formula: OleVariant; var NumberFormat: OleVariant); safecall;
    function NextRevision(var Wrap: OleVariant): Revision; safecall;
    function PreviousRevision(var Wrap: OleVariant): Revision; safecall;
    procedure PasteAsNestedTable; safecall;
    function CreateAutoTextEntry(const Name: WideString; const StyleName: WideString): AutoTextEntry; safecall;
    procedure DetectLanguage; safecall;
    procedure SelectCell; safecall;
    procedure InsertRowsBelow(var NumRows: OleVariant); safecall;
    procedure InsertColumnsRight; safecall;
    procedure InsertRowsAbove(var NumRows: OleVariant); safecall;
    procedure RtlRun; safecall;
    procedure LtrRun; safecall;
    procedure BoldRun; safecall;
    procedure ItalicRun; safecall;
    procedure RtlPara; safecall;
    procedure LtrPara; safecall;
    procedure InsertDateTime(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                             var InsertAsFullWidth: OleVariant; var DateLanguage: OleVariant; 
                             var CalendarType: OleVariant); safecall;
    procedure Sort(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                   var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                   var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                   var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                   var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                   var SortColumn: OleVariant; var Separator: OleVariant; 
                   var CaseSensitive: OleVariant; var BidiSort: OleVariant; 
                   var IgnoreThe: OleVariant; var IgnoreKashida: OleVariant; 
                   var IgnoreDiacritics: OleVariant; var IgnoreHe: OleVariant; 
                   var LanguageID: OleVariant); safecall;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; safecall;
    function Get_NoProofing: Integer; safecall;
    procedure Set_NoProofing(prop: Integer); safecall;
    function Get_TopLevelTables: Tables; safecall;
    function Get_LanguageDetected: WordBool; safecall;
    procedure Set_LanguageDetected(prop: WordBool); safecall;
    function Get_FitTextWidth: Single; safecall;
    procedure Set_FitTextWidth(prop: Single); safecall;
    property Text: WideString read Get_Text write Set_Text;
    property FormattedText: Range read Get_FormattedText write Set_FormattedText;
    property Start: Integer read Get_Start write Set_Start;
    property End_: Integer read Get_End_ write Set_End_;
    property Font: Font read Get_Font write Set_Font;
    property Type_: WdSelectionType read Get_Type_;
    property StoryType: WdStoryType read Get_StoryType;
    property Tables: Tables read Get_Tables;
    property Words: Words read Get_Words;
    property Sentences: Sentences read Get_Sentences;
    property Characters: Characters read Get_Characters;
    property Footnotes: Footnotes read Get_Footnotes;
    property Endnotes: Endnotes read Get_Endnotes;
    property Comments: Comments read Get_Comments;
    property Cells: Cells read Get_Cells;
    property Sections: Sections read Get_Sections;
    property Paragraphs: Paragraphs read Get_Paragraphs;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property Fields: Fields read Get_Fields;
    property FormFields: FormFields read Get_FormFields;
    property Frames: Frames read Get_Frames;
    property ParagraphFormat: ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property Bookmarks: Bookmarks read Get_Bookmarks;
    property StoryLength: Integer read Get_StoryLength;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property LanguageIDFarEast: WdLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property LanguageIDOther: WdLanguageID read Get_LanguageIDOther write Set_LanguageIDOther;
    property Hyperlinks: Hyperlinks read Get_Hyperlinks;
    property Columns: Columns read Get_Columns;
    property Rows: Rows read Get_Rows;
    property HeaderFooter: HeaderFooter read Get_HeaderFooter;
    property IsEndOfRowMark: WordBool read Get_IsEndOfRowMark;
    property BookmarkID: Integer read Get_BookmarkID;
    property PreviousBookmarkID: Integer read Get_PreviousBookmarkID;
    property Find: Find read Get_Find;
    property Range: Range read Get_Range;
    property Information[Type_: WdInformation]: OleVariant read Get_Information;
    property Flags: WdSelectionFlags read Get_Flags write Set_Flags;
    property Active: WordBool read Get_Active;
    property StartIsActive: WordBool read Get_StartIsActive write Set_StartIsActive;
    property IPAtEndOfLine: WordBool read Get_IPAtEndOfLine;
    property ExtendMode: WordBool read Get_ExtendMode write Set_ExtendMode;
    property ColumnSelectMode: WordBool read Get_ColumnSelectMode write Set_ColumnSelectMode;
    property Orientation: WdTextOrientation read Get_Orientation write Set_Orientation;
    property InlineShapes: InlineShapes read Get_InlineShapes;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Document: Document read Get_Document;
    property ShapeRange: ShapeRange read Get_ShapeRange;
    property NoProofing: Integer read Get_NoProofing write Set_NoProofing;
    property TopLevelTables: Tables read Get_TopLevelTables;
    property LanguageDetected: WordBool read Get_LanguageDetected write Set_LanguageDetected;
    property FitTextWidth: Single read Get_FitTextWidth write Set_FitTextWidth;
  end;

// *********************************************************************//
// DispIntf:  SelectionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020975-0000-0000-C000-000000000046}
// *********************************************************************//
  SelectionDisp = dispinterface
    ['{00020975-0000-0000-C000-000000000046}']
    property Text: WideString dispid 0;
    property FormattedText: Range dispid 2;
    property Start: Integer dispid 3;
    property End_: Integer dispid 4;
    property Font: Font dispid 5;
    property Type_: WdSelectionType readonly dispid 6;
    property StoryType: WdStoryType readonly dispid 7;
    function Style: OleVariant; dispid 8;
    property Tables: Tables readonly dispid 50;
    property Words: Words readonly dispid 51;
    property Sentences: Sentences readonly dispid 52;
    property Characters: Characters readonly dispid 53;
    property Footnotes: Footnotes readonly dispid 54;
    property Endnotes: Endnotes readonly dispid 55;
    property Comments: Comments readonly dispid 56;
    property Cells: Cells readonly dispid 57;
    property Sections: Sections readonly dispid 58;
    property Paragraphs: Paragraphs readonly dispid 59;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 61;
    property Fields: Fields readonly dispid 64;
    property FormFields: FormFields readonly dispid 65;
    property Frames: Frames readonly dispid 66;
    property ParagraphFormat: ParagraphFormat dispid 1102;
    property PageSetup: PageSetup dispid 1101;
    property Bookmarks: Bookmarks readonly dispid 75;
    property StoryLength: Integer readonly dispid 152;
    property LanguageID: WdLanguageID dispid 153;
    property LanguageIDFarEast: WdLanguageID dispid 154;
    property LanguageIDOther: WdLanguageID dispid 155;
    property Hyperlinks: Hyperlinks readonly dispid 156;
    property Columns: Columns readonly dispid 302;
    property Rows: Rows readonly dispid 303;
    property HeaderFooter: HeaderFooter readonly dispid 306;
    property IsEndOfRowMark: WordBool readonly dispid 307;
    property BookmarkID: Integer readonly dispid 308;
    property PreviousBookmarkID: Integer readonly dispid 309;
    property Find: Find readonly dispid 262;
    property Range: Range readonly dispid 400;
    property Information[Type_: WdInformation]: OleVariant readonly dispid 401;
    property Flags: WdSelectionFlags dispid 402;
    property Active: WordBool readonly dispid 403;
    property StartIsActive: WordBool dispid 404;
    property IPAtEndOfLine: WordBool readonly dispid 405;
    property ExtendMode: WordBool dispid 406;
    property ColumnSelectMode: WordBool dispid 407;
    property Orientation: WdTextOrientation dispid 410;
    property InlineShapes: InlineShapes readonly dispid 411;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Document: Document readonly dispid 1003;
    property ShapeRange: ShapeRange readonly dispid 1004;
    procedure Select; dispid 65535;
    procedure SetRange(Start: Integer; End_: Integer); dispid 100;
    procedure Collapse(var Direction: OleVariant); dispid 101;
    procedure InsertBefore(const Text: WideString); dispid 102;
    procedure InsertAfter(const Text: WideString); dispid 104;
    function Next(var Unit_: OleVariant; var Count: OleVariant): Range; dispid 105;
    function Previous(var Unit_: OleVariant; var Count: OleVariant): Range; dispid 106;
    function StartOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; dispid 107;
    function EndOf(var Unit_: OleVariant; var Extend: OleVariant): Integer; dispid 108;
    function Move(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 109;
    function MoveStart(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 110;
    function MoveEnd(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 111;
    function MoveWhile(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 112;
    function MoveStartWhile(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 113;
    function MoveEndWhile(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 114;
    function MoveUntil(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 115;
    function MoveStartUntil(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 116;
    function MoveEndUntil(var Cset: OleVariant; var Count: OleVariant): Integer; dispid 117;
    procedure Cut; dispid 119;
    procedure Copy; dispid 120;
    procedure Paste; dispid 121;
    procedure InsertBreak(var Type_: OleVariant); dispid 122;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); dispid 123;
    function InStory(const Range: Range): WordBool; dispid 125;
    function InRange(const Range: Range): WordBool; dispid 126;
    function Delete(var Unit_: OleVariant; var Count: OleVariant): Integer; dispid 127;
    function Expand(var Unit_: OleVariant): Integer; dispid 129;
    procedure InsertParagraph; dispid 160;
    procedure InsertParagraphAfter; dispid 161;
    function ConvertToTableOld(var Separator: OleVariant; var NumRows: OleVariant; 
                               var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                               var Format: OleVariant; var ApplyBorders: OleVariant; 
                               var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                               var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                               var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                               var ApplyLastColumn: OleVariant; var AutoFit: OleVariant): Table; dispid 162;
    procedure InsertDateTimeOld(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                                var InsertAsFullWidth: OleVariant); dispid 163;
    procedure InsertSymbol(CharacterNumber: Integer; var Font: OleVariant; var Unicode: OleVariant; 
                           var Bias: OleVariant); dispid 164;
    procedure InsertCrossReference(var ReferenceType: OleVariant; ReferenceKind: WdReferenceKind; 
                                   var ReferenceItem: OleVariant; 
                                   var InsertAsHyperlink: OleVariant; 
                                   var IncludePosition: OleVariant); dispid 165;
    procedure InsertCaption(var Label_: OleVariant; var Title: OleVariant; 
                            var TitleAutoText: OleVariant; var Position: OleVariant); dispid 166;
    procedure CopyAsPicture; dispid 167;
    procedure SortOld(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                      var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                      var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                      var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                      var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                      var SortColumn: OleVariant; var Separator: OleVariant; 
                      var CaseSensitive: OleVariant; var LanguageID: OleVariant); dispid 168;
    procedure SortAscending; dispid 169;
    procedure SortDescending; dispid 170;
    function IsEqual(const Range: Range): WordBool; dispid 171;
    function Calculate: Single; dispid 172;
    function GoTo_(var What: OleVariant; var Which: OleVariant; var Count: OleVariant; 
                   var Name: OleVariant): Range; dispid 173;
    function GoToNext(What: WdGoToItem): Range; dispid 174;
    function GoToPrevious(What: WdGoToItem): Range; dispid 175;
    procedure PasteSpecial(var IconIndex: OleVariant; var Link: OleVariant; 
                           var Placement: OleVariant; var DisplayAsIcon: OleVariant; 
                           var DataType: OleVariant; var IconFileName: OleVariant; 
                           var IconLabel: OleVariant); dispid 176;
    function PreviousField: Field; dispid 177;
    function NextField: Field; dispid 178;
    procedure InsertParagraphBefore; dispid 212;
    procedure InsertCells(var ShiftCells: OleVariant); dispid 214;
    procedure Extend(var Character: OleVariant); dispid 300;
    procedure Shrink; dispid 301;
    function MoveLeft(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; dispid 500;
    function MoveRight(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; dispid 501;
    function MoveUp(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; dispid 502;
    function MoveDown(var Unit_: OleVariant; var Count: OleVariant; var Extend: OleVariant): Integer; dispid 503;
    function HomeKey(var Unit_: OleVariant; var Extend: OleVariant): Integer; dispid 504;
    function EndKey(var Unit_: OleVariant; var Extend: OleVariant): Integer; dispid 505;
    procedure EscapeKey; dispid 506;
    procedure TypeText(const Text: WideString); dispid 507;
    procedure CopyFormat; dispid 509;
    procedure PasteFormat; dispid 510;
    procedure TypeParagraph; dispid 512;
    procedure TypeBackspace; dispid 513;
    procedure NextSubdocument; dispid 514;
    procedure PreviousSubdocument; dispid 515;
    procedure SelectColumn; dispid 516;
    procedure SelectCurrentFont; dispid 517;
    procedure SelectCurrentAlignment; dispid 518;
    procedure SelectCurrentSpacing; dispid 519;
    procedure SelectCurrentIndent; dispid 520;
    procedure SelectCurrentTabs; dispid 521;
    procedure SelectCurrentColor; dispid 522;
    procedure CreateTextbox; dispid 523;
    procedure WholeStory; dispid 524;
    procedure SelectRow; dispid 525;
    procedure SplitTable; dispid 526;
    procedure InsertRows(var NumRows: OleVariant); dispid 528;
    procedure InsertColumns; dispid 529;
    procedure InsertFormula(var Formula: OleVariant; var NumberFormat: OleVariant); dispid 530;
    function NextRevision(var Wrap: OleVariant): Revision; dispid 531;
    function PreviousRevision(var Wrap: OleVariant): Revision; dispid 532;
    procedure PasteAsNestedTable; dispid 533;
    function CreateAutoTextEntry(const Name: WideString; const StyleName: WideString): AutoTextEntry; dispid 534;
    procedure DetectLanguage; dispid 535;
    procedure SelectCell; dispid 536;
    procedure InsertRowsBelow(var NumRows: OleVariant); dispid 537;
    procedure InsertColumnsRight; dispid 538;
    procedure InsertRowsAbove(var NumRows: OleVariant); dispid 539;
    procedure RtlRun; dispid 600;
    procedure LtrRun; dispid 601;
    procedure BoldRun; dispid 602;
    procedure ItalicRun; dispid 603;
    procedure RtlPara; dispid 605;
    procedure LtrPara; dispid 606;
    procedure InsertDateTime(var DateTimeFormat: OleVariant; var InsertAsField: OleVariant; 
                             var InsertAsFullWidth: OleVariant; var DateLanguage: OleVariant; 
                             var CalendarType: OleVariant); dispid 444;
    procedure Sort(var ExcludeHeader: OleVariant; var FieldNumber: OleVariant; 
                   var SortFieldType: OleVariant; var SortOrder: OleVariant; 
                   var FieldNumber2: OleVariant; var SortFieldType2: OleVariant; 
                   var SortOrder2: OleVariant; var FieldNumber3: OleVariant; 
                   var SortFieldType3: OleVariant; var SortOrder3: OleVariant; 
                   var SortColumn: OleVariant; var Separator: OleVariant; 
                   var CaseSensitive: OleVariant; var BidiSort: OleVariant; 
                   var IgnoreThe: OleVariant; var IgnoreKashida: OleVariant; 
                   var IgnoreDiacritics: OleVariant; var IgnoreHe: OleVariant; 
                   var LanguageID: OleVariant); dispid 445;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; dispid 457;
    property NoProofing: Integer dispid 1005;
    property TopLevelTables: Tables readonly dispid 1006;
    property LanguageDetected: WordBool dispid 1007;
    property FitTextWidth: Single dispid 1008;
  end;

// *********************************************************************//
// Interface: TablesOfAuthoritiesCategories
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020976-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfAuthoritiesCategories = interface(IDispatch)
    ['{00020976-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): TableOfAuthoritiesCategory; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  TablesOfAuthoritiesCategoriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020976-0000-0000-C000-000000000046}
// *********************************************************************//
  TablesOfAuthoritiesCategoriesDisp = dispinterface
    ['{00020976-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): TableOfAuthoritiesCategory; dispid 0;
  end;

// *********************************************************************//
// Interface: TableOfAuthoritiesCategory
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020977-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfAuthoritiesCategory = interface(IDispatch)
    ['{00020977-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Index: Integer; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name write Set_Name;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  TableOfAuthoritiesCategoryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020977-0000-0000-C000-000000000046}
// *********************************************************************//
  TableOfAuthoritiesCategoryDisp = dispinterface
    ['{00020977-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString dispid 0;
    property Index: Integer readonly dispid 1;
  end;

// *********************************************************************//
// Interface: CaptionLabels
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020978-0000-0000-C000-000000000046}
// *********************************************************************//
  CaptionLabels = interface(IDispatch)
    ['{00020978-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): CaptionLabel; safecall;
    function Add(const Name: WideString): CaptionLabel; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  CaptionLabelsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020978-0000-0000-C000-000000000046}
// *********************************************************************//
  CaptionLabelsDisp = dispinterface
    ['{00020978-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): CaptionLabel; dispid 0;
    function Add(const Name: WideString): CaptionLabel; dispid 100;
  end;

// *********************************************************************//
// Interface: CaptionLabel
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020979-0000-0000-C000-000000000046}
// *********************************************************************//
  CaptionLabel = interface(IDispatch)
    ['{00020979-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_BuiltIn: WordBool; safecall;
    function Get_ID: WdCaptionLabelID; safecall;
    function Get_IncludeChapterNumber: WordBool; safecall;
    procedure Set_IncludeChapterNumber(prop: WordBool); safecall;
    function Get_NumberStyle: WdCaptionNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WdCaptionNumberStyle); safecall;
    function Get_ChapterStyleLevel: Integer; safecall;
    procedure Set_ChapterStyleLevel(prop: Integer); safecall;
    function Get_Separator: WdSeparatorType; safecall;
    procedure Set_Separator(prop: WdSeparatorType); safecall;
    function Get_Position: WdCaptionPosition; safecall;
    procedure Set_Position(prop: WdCaptionPosition); safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property BuiltIn: WordBool read Get_BuiltIn;
    property ID: WdCaptionLabelID read Get_ID;
    property IncludeChapterNumber: WordBool read Get_IncludeChapterNumber write Set_IncludeChapterNumber;
    property NumberStyle: WdCaptionNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property ChapterStyleLevel: Integer read Get_ChapterStyleLevel write Set_ChapterStyleLevel;
    property Separator: WdSeparatorType read Get_Separator write Set_Separator;
    property Position: WdCaptionPosition read Get_Position write Set_Position;
  end;

// *********************************************************************//
// DispIntf:  CaptionLabelDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020979-0000-0000-C000-000000000046}
// *********************************************************************//
  CaptionLabelDisp = dispinterface
    ['{00020979-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property BuiltIn: WordBool readonly dispid 1;
    property ID: WdCaptionLabelID readonly dispid 2;
    property IncludeChapterNumber: WordBool dispid 3;
    property NumberStyle: WdCaptionNumberStyle dispid 4;
    property ChapterStyleLevel: Integer dispid 5;
    property Separator: WdSeparatorType dispid 6;
    property Position: WdCaptionPosition dispid 7;
    procedure Delete; dispid 100;
  end;

// *********************************************************************//
// Interface: AutoCaptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097A-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCaptions = interface(IDispatch)
    ['{0002097A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): AutoCaption; safecall;
    procedure CancelAutoInsert; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AutoCaptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097A-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCaptionsDisp = dispinterface
    ['{0002097A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): AutoCaption; dispid 0;
    procedure CancelAutoInsert; dispid 100;
  end;

// *********************************************************************//
// Interface: AutoCaption
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097B-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCaption = interface(IDispatch)
    ['{0002097B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_AutoInsert: WordBool; safecall;
    procedure Set_AutoInsert(prop: WordBool); safecall;
    function Get_Index: Integer; safecall;
    function Get_CaptionLabel: OleVariant; safecall;
    procedure Set_CaptionLabel(var prop: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property AutoInsert: WordBool read Get_AutoInsert write Set_AutoInsert;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  AutoCaptionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097B-0000-0000-C000-000000000046}
// *********************************************************************//
  AutoCaptionDisp = dispinterface
    ['{0002097B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property AutoInsert: WordBool dispid 1;
    property Index: Integer readonly dispid 2;
    function CaptionLabel: OleVariant; dispid 3;
  end;

// *********************************************************************//
// Interface: Indexes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097C-0000-0000-C000-000000000046}
// *********************************************************************//
  Indexes = interface(IDispatch)
    ['{0002097C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Format: WdIndexFormat; safecall;
    procedure Set_Format(prop: WdIndexFormat); safecall;
    function Item(Index: Integer): Index; safecall;
    function AddOld(const Range: Range; var HeadingSeparator: OleVariant; 
                    var RightAlignPageNumbers: OleVariant; var Type_: OleVariant; 
                    var NumberOfColumns: OleVariant; var AccentedLetters: OleVariant): Index; safecall;
    function MarkEntry(const Range: Range; var Entry: OleVariant; var EntryAutoText: OleVariant; 
                       var CrossReference: OleVariant; var CrossReferenceAutoText: OleVariant; 
                       var BookmarkName: OleVariant; var Bold: OleVariant; var Italic: OleVariant; 
                       var Reading: OleVariant): Field; safecall;
    procedure MarkAllEntries(const Range: Range; var Entry: OleVariant; 
                             var EntryAutoText: OleVariant; var CrossReference: OleVariant; 
                             var CrossReferenceAutoText: OleVariant; var BookmarkName: OleVariant; 
                             var Bold: OleVariant; var Italic: OleVariant); safecall;
    procedure AutoMarkEntries(const ConcordanceFileName: WideString); safecall;
    function Add(const Range: Range; var HeadingSeparator: OleVariant; 
                 var RightAlignPageNumbers: OleVariant; var Type_: OleVariant; 
                 var NumberOfColumns: OleVariant; var AccentedLetters: OleVariant; 
                 var SortBy: OleVariant; var IndexLanguage: OleVariant): Index; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Format: WdIndexFormat read Get_Format write Set_Format;
  end;

// *********************************************************************//
// DispIntf:  IndexesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097C-0000-0000-C000-000000000046}
// *********************************************************************//
  IndexesDisp = dispinterface
    ['{0002097C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Format: WdIndexFormat dispid 2;
    function Item(Index: Integer): Index; dispid 0;
    function AddOld(const Range: Range; var HeadingSeparator: OleVariant; 
                    var RightAlignPageNumbers: OleVariant; var Type_: OleVariant; 
                    var NumberOfColumns: OleVariant; var AccentedLetters: OleVariant): Index; dispid 100;
    function MarkEntry(const Range: Range; var Entry: OleVariant; var EntryAutoText: OleVariant; 
                       var CrossReference: OleVariant; var CrossReferenceAutoText: OleVariant; 
                       var BookmarkName: OleVariant; var Bold: OleVariant; var Italic: OleVariant; 
                       var Reading: OleVariant): Field; dispid 101;
    procedure MarkAllEntries(const Range: Range; var Entry: OleVariant; 
                             var EntryAutoText: OleVariant; var CrossReference: OleVariant; 
                             var CrossReferenceAutoText: OleVariant; var BookmarkName: OleVariant; 
                             var Bold: OleVariant; var Italic: OleVariant); dispid 102;
    procedure AutoMarkEntries(const ConcordanceFileName: WideString); dispid 103;
    function Add(const Range: Range; var HeadingSeparator: OleVariant; 
                 var RightAlignPageNumbers: OleVariant; var Type_: OleVariant; 
                 var NumberOfColumns: OleVariant; var AccentedLetters: OleVariant; 
                 var SortBy: OleVariant; var IndexLanguage: OleVariant): Index; dispid 104;
  end;

// *********************************************************************//
// Interface: Index
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097D-0000-0000-C000-000000000046}
// *********************************************************************//
  Index = interface(IDispatch)
    ['{0002097D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_HeadingSeparator: WdHeadingSeparator; safecall;
    procedure Set_HeadingSeparator(prop: WdHeadingSeparator); safecall;
    function Get_RightAlignPageNumbers: WordBool; safecall;
    procedure Set_RightAlignPageNumbers(prop: WordBool); safecall;
    function Get_Type_: WdIndexType; safecall;
    procedure Set_Type_(prop: WdIndexType); safecall;
    function Get_NumberOfColumns: Integer; safecall;
    procedure Set_NumberOfColumns(prop: Integer); safecall;
    function Get_Range: Range; safecall;
    function Get_TabLeader: WdTabLeader; safecall;
    procedure Set_TabLeader(prop: WdTabLeader); safecall;
    function Get_AccentedLetters: WordBool; safecall;
    procedure Set_AccentedLetters(prop: WordBool); safecall;
    function Get_SortBy: WdIndexSortBy; safecall;
    procedure Set_SortBy(prop: WdIndexSortBy); safecall;
    function Get_Filter: WdIndexFilter; safecall;
    procedure Set_Filter(prop: WdIndexFilter); safecall;
    procedure Delete; safecall;
    procedure Update; safecall;
    function Get_IndexLanguage: WdLanguageID; safecall;
    procedure Set_IndexLanguage(prop: WdLanguageID); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property HeadingSeparator: WdHeadingSeparator read Get_HeadingSeparator write Set_HeadingSeparator;
    property RightAlignPageNumbers: WordBool read Get_RightAlignPageNumbers write Set_RightAlignPageNumbers;
    property Type_: WdIndexType read Get_Type_ write Set_Type_;
    property NumberOfColumns: Integer read Get_NumberOfColumns write Set_NumberOfColumns;
    property Range: Range read Get_Range;
    property TabLeader: WdTabLeader read Get_TabLeader write Set_TabLeader;
    property AccentedLetters: WordBool read Get_AccentedLetters write Set_AccentedLetters;
    property SortBy: WdIndexSortBy read Get_SortBy write Set_SortBy;
    property Filter: WdIndexFilter read Get_Filter write Set_Filter;
    property IndexLanguage: WdLanguageID read Get_IndexLanguage write Set_IndexLanguage;
  end;

// *********************************************************************//
// DispIntf:  IndexDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097D-0000-0000-C000-000000000046}
// *********************************************************************//
  IndexDisp = dispinterface
    ['{0002097D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property HeadingSeparator: WdHeadingSeparator dispid 1;
    property RightAlignPageNumbers: WordBool dispid 2;
    property Type_: WdIndexType dispid 3;
    property NumberOfColumns: Integer dispid 4;
    property Range: Range readonly dispid 5;
    property TabLeader: WdTabLeader dispid 6;
    property AccentedLetters: WordBool dispid 7;
    property SortBy: WdIndexSortBy dispid 8;
    property Filter: WdIndexFilter dispid 9;
    procedure Delete; dispid 100;
    procedure Update; dispid 102;
    property IndexLanguage: WdLanguageID dispid 10;
  end;

// *********************************************************************//
// Interface: AddIn
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097E-0000-0000-C000-000000000046}
// *********************************************************************//
  AddIn = interface(IDispatch)
    ['{0002097E-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    function Get_Path: WideString; safecall;
    function Get_Installed: WordBool; safecall;
    procedure Set_Installed(prop: WordBool); safecall;
    function Get_Compiled: WordBool; safecall;
    function Get_Autoload: WordBool; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
    property Path: WideString read Get_Path;
    property Installed: WordBool read Get_Installed write Set_Installed;
    property Compiled: WordBool read Get_Compiled;
    property Autoload: WordBool read Get_Autoload;
  end;

// *********************************************************************//
// DispIntf:  AddInDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097E-0000-0000-C000-000000000046}
// *********************************************************************//
  AddInDisp = dispinterface
    ['{0002097E-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Index: Integer readonly dispid 1;
    property Path: WideString readonly dispid 3;
    property Installed: WordBool dispid 4;
    property Compiled: WordBool readonly dispid 5;
    property Autoload: WordBool readonly dispid 6;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: AddIns
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097F-0000-0000-C000-000000000046}
// *********************************************************************//
  AddIns = interface(IDispatch)
    ['{0002097F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): AddIn; safecall;
    function Add(const FileName: WideString; var Install: OleVariant): AddIn; safecall;
    procedure Unload(RemoveFromList: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AddInsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097F-0000-0000-C000-000000000046}
// *********************************************************************//
  AddInsDisp = dispinterface
    ['{0002097F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): AddIn; dispid 0;
    function Add(const FileName: WideString; var Install: OleVariant): AddIn; dispid 2;
    procedure Unload(RemoveFromList: WordBool); dispid 3;
  end;

// *********************************************************************//
// Interface: Revisions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020980-0000-0000-C000-000000000046}
// *********************************************************************//
  Revisions = interface(IDispatch)
    ['{00020980-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Revision; safecall;
    procedure AcceptAll; safecall;
    procedure RejectAll; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  RevisionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020980-0000-0000-C000-000000000046}
// *********************************************************************//
  RevisionsDisp = dispinterface
    ['{00020980-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 5;
    function Item(Index: Integer): Revision; dispid 0;
    procedure AcceptAll; dispid 101;
    procedure RejectAll; dispid 102;
  end;

// *********************************************************************//
// Interface: Revision
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020981-0000-0000-C000-000000000046}
// *********************************************************************//
  Revision = interface(IDispatch)
    ['{00020981-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Author: WideString; safecall;
    function Get_Date: TDateTime; safecall;
    function Get_Range: Range; safecall;
    function Get_Type_: WdRevisionType; safecall;
    function Get_Index: Integer; safecall;
    procedure Accept; safecall;
    procedure Reject; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Author: WideString read Get_Author;
    property Date: TDateTime read Get_Date;
    property Range: Range read Get_Range;
    property Type_: WdRevisionType read Get_Type_;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  RevisionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020981-0000-0000-C000-000000000046}
// *********************************************************************//
  RevisionDisp = dispinterface
    ['{00020981-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Author: WideString readonly dispid 1;
    property Date: TDateTime readonly dispid 2;
    property Range: Range readonly dispid 3;
    property Type_: WdRevisionType readonly dispid 4;
    property Index: Integer readonly dispid 5;
    procedure Accept; dispid 101;
    procedure Reject; dispid 102;
  end;

// *********************************************************************//
// Interface: Task
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020982-0000-0000-C000-000000000046}
// *********************************************************************//
  Task = interface(IDispatch)
    ['{00020982-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(prop: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(prop: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(prop: Integer); safecall;
    function Get_WindowState: WdWindowState; safecall;
    procedure Set_WindowState(prop: WdWindowState); safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    procedure Activate(var Wait: OleVariant); safecall;
    procedure Close; safecall;
    procedure Move(Left: Integer; Top: Integer); safecall;
    procedure Resize(Width: Integer; Height: Integer); safecall;
    procedure SendWindowMessage(Message: Integer; wParam: Integer; lParam: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property WindowState: WdWindowState read Get_WindowState write Set_WindowState;
    property Visible: WordBool read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  TaskDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020982-0000-0000-C000-000000000046}
// *********************************************************************//
  TaskDisp = dispinterface
    ['{00020982-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Left: Integer dispid 1;
    property Top: Integer dispid 2;
    property Width: Integer dispid 3;
    property Height: Integer dispid 4;
    property WindowState: WdWindowState dispid 5;
    property Visible: WordBool dispid 6;
    procedure Activate(var Wait: OleVariant); dispid 10;
    procedure Close; dispid 11;
    procedure Move(Left: Integer; Top: Integer); dispid 12;
    procedure Resize(Width: Integer; Height: Integer); dispid 13;
    procedure SendWindowMessage(Message: Integer; wParam: Integer; lParam: Integer); dispid 14;
  end;

// *********************************************************************//
// Interface: Tasks
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020983-0000-0000-C000-000000000046}
// *********************************************************************//
  Tasks = interface(IDispatch)
    ['{00020983-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): Task; safecall;
    function Exists(const Name: WideString): WordBool; safecall;
    procedure ExitWindows_; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  TasksDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020983-0000-0000-C000-000000000046}
// *********************************************************************//
  TasksDisp = dispinterface
    ['{00020983-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): Task; dispid 0;
    function Exists(const Name: WideString): WordBool; dispid 2;
    procedure ExitWindows_; dispid 3;
  end;

// *********************************************************************//
// Interface: HeadersFooters
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020984-0000-0000-C000-000000000046}
// *********************************************************************//
  HeadersFooters = interface(IDispatch)
    ['{00020984-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: WdHeaderFooterIndex): HeaderFooter; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  HeadersFootersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020984-0000-0000-C000-000000000046}
// *********************************************************************//
  HeadersFootersDisp = dispinterface
    ['{00020984-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: WdHeaderFooterIndex): HeaderFooter; dispid 0;
  end;

// *********************************************************************//
// Interface: HeaderFooter
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020985-0000-0000-C000-000000000046}
// *********************************************************************//
  HeaderFooter = interface(IDispatch)
    ['{00020985-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Range: Range; safecall;
    function Get_Index: WdHeaderFooterIndex; safecall;
    function Get_IsHeader: WordBool; safecall;
    function Get_Exists: WordBool; safecall;
    procedure Set_Exists(prop: WordBool); safecall;
    function Get_PageNumbers: PageNumbers; safecall;
    function Get_LinkToPrevious: WordBool; safecall;
    procedure Set_LinkToPrevious(prop: WordBool); safecall;
    function Get_Shapes: Shapes; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Range: Range read Get_Range;
    property Index: WdHeaderFooterIndex read Get_Index;
    property IsHeader: WordBool read Get_IsHeader;
    property Exists: WordBool read Get_Exists write Set_Exists;
    property PageNumbers: PageNumbers read Get_PageNumbers;
    property LinkToPrevious: WordBool read Get_LinkToPrevious write Set_LinkToPrevious;
    property Shapes: Shapes read Get_Shapes;
  end;

// *********************************************************************//
// DispIntf:  HeaderFooterDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020985-0000-0000-C000-000000000046}
// *********************************************************************//
  HeaderFooterDisp = dispinterface
    ['{00020985-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Range: Range readonly dispid 0;
    property Index: WdHeaderFooterIndex readonly dispid 2;
    property IsHeader: WordBool readonly dispid 3;
    property Exists: WordBool dispid 4;
    property PageNumbers: PageNumbers readonly dispid 5;
    property LinkToPrevious: WordBool dispid 6;
    property Shapes: Shapes readonly dispid 7;
  end;

// *********************************************************************//
// Interface: PageNumbers
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020986-0000-0000-C000-000000000046}
// *********************************************************************//
  PageNumbers = interface(IDispatch)
    ['{00020986-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_NumberStyle: WdPageNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WdPageNumberStyle); safecall;
    function Get_IncludeChapterNumber: WordBool; safecall;
    procedure Set_IncludeChapterNumber(prop: WordBool); safecall;
    function Get_HeadingLevelForChapter: Integer; safecall;
    procedure Set_HeadingLevelForChapter(prop: Integer); safecall;
    function Get_ChapterPageSeparator: WdSeparatorType; safecall;
    procedure Set_ChapterPageSeparator(prop: WdSeparatorType); safecall;
    function Get_RestartNumberingAtSection: WordBool; safecall;
    procedure Set_RestartNumberingAtSection(prop: WordBool); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_ShowFirstPageNumber: WordBool; safecall;
    procedure Set_ShowFirstPageNumber(prop: WordBool); safecall;
    function Item(Index: Integer): PageNumber; safecall;
    function Add(var PageNumberAlignment: OleVariant; var FirstPage: OleVariant): PageNumber; safecall;
    function Get_DoubleQuote: WordBool; safecall;
    procedure Set_DoubleQuote(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property NumberStyle: WdPageNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property IncludeChapterNumber: WordBool read Get_IncludeChapterNumber write Set_IncludeChapterNumber;
    property HeadingLevelForChapter: Integer read Get_HeadingLevelForChapter write Set_HeadingLevelForChapter;
    property ChapterPageSeparator: WdSeparatorType read Get_ChapterPageSeparator write Set_ChapterPageSeparator;
    property RestartNumberingAtSection: WordBool read Get_RestartNumberingAtSection write Set_RestartNumberingAtSection;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property ShowFirstPageNumber: WordBool read Get_ShowFirstPageNumber write Set_ShowFirstPageNumber;
    property DoubleQuote: WordBool read Get_DoubleQuote write Set_DoubleQuote;
  end;

// *********************************************************************//
// DispIntf:  PageNumbersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020986-0000-0000-C000-000000000046}
// *********************************************************************//
  PageNumbersDisp = dispinterface
    ['{00020986-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property NumberStyle: WdPageNumberStyle dispid 2;
    property IncludeChapterNumber: WordBool dispid 3;
    property HeadingLevelForChapter: Integer dispid 4;
    property ChapterPageSeparator: WdSeparatorType dispid 5;
    property RestartNumberingAtSection: WordBool dispid 6;
    property StartingNumber: Integer dispid 7;
    property ShowFirstPageNumber: WordBool dispid 8;
    function Item(Index: Integer): PageNumber; dispid 0;
    function Add(var PageNumberAlignment: OleVariant; var FirstPage: OleVariant): PageNumber; dispid 101;
    property DoubleQuote: WordBool dispid 10;
  end;

// *********************************************************************//
// Interface: PageNumber
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020987-0000-0000-C000-000000000046}
// *********************************************************************//
  PageNumber = interface(IDispatch)
    ['{00020987-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Alignment: WdPageNumberAlignment; safecall;
    procedure Set_Alignment(prop: WdPageNumberAlignment); safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Alignment: WdPageNumberAlignment read Get_Alignment write Set_Alignment;
  end;

// *********************************************************************//
// DispIntf:  PageNumberDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020987-0000-0000-C000-000000000046}
// *********************************************************************//
  PageNumberDisp = dispinterface
    ['{00020987-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Alignment: WdPageNumberAlignment dispid 3;
    procedure Select; dispid 65535;
    procedure Copy; dispid 101;
    procedure Cut; dispid 102;
    procedure Delete; dispid 103;
  end;

// *********************************************************************//
// Interface: Subdocuments
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020988-0000-0000-C000-000000000046}
// *********************************************************************//
  Subdocuments = interface(IDispatch)
    ['{00020988-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Expanded: WordBool; safecall;
    procedure Set_Expanded(prop: WordBool); safecall;
    function Item(Index: Integer): Subdocument; safecall;
    function AddFromFile(var Name: OleVariant; var ConfirmConversions: OleVariant; 
                         var ReadOnly: OleVariant; var PasswordDocument: OleVariant; 
                         var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                         var WritePasswordDocument: OleVariant; 
                         var WritePasswordTemplate: OleVariant): Subdocument; safecall;
    function AddFromRange(const Range: Range): Subdocument; safecall;
    procedure Merge(var FirstSubdocument: OleVariant; var LastSubdocument: OleVariant); safecall;
    procedure Delete; safecall;
    procedure Select; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Expanded: WordBool read Get_Expanded write Set_Expanded;
  end;

// *********************************************************************//
// DispIntf:  SubdocumentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020988-0000-0000-C000-000000000046}
// *********************************************************************//
  SubdocumentsDisp = dispinterface
    ['{00020988-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    property Expanded: WordBool dispid 2;
    function Item(Index: Integer): Subdocument; dispid 0;
    function AddFromFile(var Name: OleVariant; var ConfirmConversions: OleVariant; 
                         var ReadOnly: OleVariant; var PasswordDocument: OleVariant; 
                         var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                         var WritePasswordDocument: OleVariant; 
                         var WritePasswordTemplate: OleVariant): Subdocument; dispid 100;
    function AddFromRange(const Range: Range): Subdocument; dispid 101;
    procedure Merge(var FirstSubdocument: OleVariant; var LastSubdocument: OleVariant); dispid 102;
    procedure Delete; dispid 103;
    procedure Select; dispid 104;
  end;

// *********************************************************************//
// Interface: Subdocument
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020989-0000-0000-C000-000000000046}
// *********************************************************************//
  Subdocument = interface(IDispatch)
    ['{00020989-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    function Get_Name: WideString; safecall;
    function Get_Path: WideString; safecall;
    function Get_HasFile: WordBool; safecall;
    function Get_Level: Integer; safecall;
    procedure Delete; safecall;
    procedure Split(const Range: Range); safecall;
    function Open: Document; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property Range: Range read Get_Range;
    property Name: WideString read Get_Name;
    property Path: WideString read Get_Path;
    property HasFile: WordBool read Get_HasFile;
    property Level: Integer read Get_Level;
  end;

// *********************************************************************//
// DispIntf:  SubdocumentDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020989-0000-0000-C000-000000000046}
// *********************************************************************//
  SubdocumentDisp = dispinterface
    ['{00020989-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Locked: WordBool dispid 1;
    property Range: Range readonly dispid 2;
    property Name: WideString readonly dispid 3;
    property Path: WideString readonly dispid 4;
    property HasFile: WordBool readonly dispid 5;
    property Level: Integer readonly dispid 6;
    procedure Delete; dispid 100;
    procedure Split(const Range: Range); dispid 101;
    function Open: Document; dispid 102;
  end;

// *********************************************************************//
// Interface: HeadingStyles
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098A-0000-0000-C000-000000000046}
// *********************************************************************//
  HeadingStyles = interface(IDispatch)
    ['{0002098A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): HeadingStyle; safecall;
    function Add(var Style: OleVariant; Level: Smallint): HeadingStyle; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  HeadingStylesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098A-0000-0000-C000-000000000046}
// *********************************************************************//
  HeadingStylesDisp = dispinterface
    ['{0002098A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): HeadingStyle; dispid 0;
    function Add(var Style: OleVariant; Level: Smallint): HeadingStyle; dispid 100;
  end;

// *********************************************************************//
// Interface: HeadingStyle
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098B-0000-0000-C000-000000000046}
// *********************************************************************//
  HeadingStyle = interface(IDispatch)
    ['{0002098B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Level: Smallint; safecall;
    procedure Set_Level(prop: Smallint); safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Level: Smallint read Get_Level write Set_Level;
  end;

// *********************************************************************//
// DispIntf:  HeadingStyleDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098B-0000-0000-C000-000000000046}
// *********************************************************************//
  HeadingStyleDisp = dispinterface
    ['{0002098B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Style: OleVariant; dispid 0;
    property Level: Smallint dispid 2;
    procedure Delete; dispid 100;
  end;

// *********************************************************************//
// Interface: StoryRanges
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098C-0000-0000-C000-000000000046}
// *********************************************************************//
  StoryRanges = interface(IDispatch)
    ['{0002098C-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: WdStoryType): Range; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  StoryRangesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098C-0000-0000-C000-000000000046}
// *********************************************************************//
  StoryRangesDisp = dispinterface
    ['{0002098C-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: WdStoryType): Range; dispid 0;
  end;

// *********************************************************************//
// Interface: ListLevel
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098D-0000-0000-C000-000000000046}
// *********************************************************************//
  ListLevel = interface(IDispatch)
    ['{0002098D-0000-0000-C000-000000000046}']
    function Get_Index: Integer; safecall;
    function Get_NumberFormat: WideString; safecall;
    procedure Set_NumberFormat(const prop: WideString); safecall;
    function Get_TrailingCharacter: WdTrailingCharacter; safecall;
    procedure Set_TrailingCharacter(prop: WdTrailingCharacter); safecall;
    function Get_NumberStyle: WdListNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WdListNumberStyle); safecall;
    function Get_NumberPosition: Single; safecall;
    procedure Set_NumberPosition(prop: Single); safecall;
    function Get_Alignment: WdListLevelAlignment; safecall;
    procedure Set_Alignment(prop: WdListLevelAlignment); safecall;
    function Get_TextPosition: Single; safecall;
    procedure Set_TextPosition(prop: Single); safecall;
    function Get_TabPosition: Single; safecall;
    procedure Set_TabPosition(prop: Single); safecall;
    function Get_ResetOnHigherOld: WordBool; safecall;
    procedure Set_ResetOnHigherOld(prop: WordBool); safecall;
    function Get_StartAt: Integer; safecall;
    procedure Set_StartAt(prop: Integer); safecall;
    function Get_LinkedStyle: WideString; safecall;
    procedure Set_LinkedStyle(const prop: WideString); safecall;
    function Get_Font: Font; safecall;
    procedure Set_Font(const prop: Font); safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ResetOnHigher: Integer; safecall;
    procedure Set_ResetOnHigher(prop: Integer); safecall;
    property Index: Integer read Get_Index;
    property NumberFormat: WideString read Get_NumberFormat write Set_NumberFormat;
    property TrailingCharacter: WdTrailingCharacter read Get_TrailingCharacter write Set_TrailingCharacter;
    property NumberStyle: WdListNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property NumberPosition: Single read Get_NumberPosition write Set_NumberPosition;
    property Alignment: WdListLevelAlignment read Get_Alignment write Set_Alignment;
    property TextPosition: Single read Get_TextPosition write Set_TextPosition;
    property TabPosition: Single read Get_TabPosition write Set_TabPosition;
    property ResetOnHigherOld: WordBool read Get_ResetOnHigherOld write Set_ResetOnHigherOld;
    property StartAt: Integer read Get_StartAt write Set_StartAt;
    property LinkedStyle: WideString read Get_LinkedStyle write Set_LinkedStyle;
    property Font: Font read Get_Font write Set_Font;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ResetOnHigher: Integer read Get_ResetOnHigher write Set_ResetOnHigher;
  end;

// *********************************************************************//
// DispIntf:  ListLevelDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098D-0000-0000-C000-000000000046}
// *********************************************************************//
  ListLevelDisp = dispinterface
    ['{0002098D-0000-0000-C000-000000000046}']
    property Index: Integer readonly dispid 1;
    property NumberFormat: WideString dispid 2;
    property TrailingCharacter: WdTrailingCharacter dispid 3;
    property NumberStyle: WdListNumberStyle dispid 4;
    property NumberPosition: Single dispid 5;
    property Alignment: WdListLevelAlignment dispid 6;
    property TextPosition: Single dispid 7;
    property TabPosition: Single dispid 8;
    property ResetOnHigherOld: WordBool dispid 9;
    property StartAt: Integer dispid 10;
    property LinkedStyle: WideString dispid 11;
    property Font: Font dispid 12;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property ResetOnHigher: Integer dispid 13;
  end;

// *********************************************************************//
// Interface: ListLevels
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098E-0000-0000-C000-000000000046}
// *********************************************************************//
  ListLevels = interface(IDispatch)
    ['{0002098E-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): ListLevel; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListLevelsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098E-0000-0000-C000-000000000046}
// *********************************************************************//
  ListLevelsDisp = dispinterface
    ['{0002098E-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): ListLevel; dispid 0;
  end;

// *********************************************************************//
// Interface: ListTemplate
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098F-0000-0000-C000-000000000046}
// *********************************************************************//
  ListTemplate = interface(IDispatch)
    ['{0002098F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_OutlineNumbered: WordBool; safecall;
    procedure Set_OutlineNumbered(prop: WordBool); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_ListLevels: ListLevels; safecall;
    function Convert(var Level: OleVariant): ListTemplate; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property OutlineNumbered: WordBool read Get_OutlineNumbered write Set_OutlineNumbered;
    property Name: WideString read Get_Name write Set_Name;
    property ListLevels: ListLevels read Get_ListLevels;
  end;

// *********************************************************************//
// DispIntf:  ListTemplateDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098F-0000-0000-C000-000000000046}
// *********************************************************************//
  ListTemplateDisp = dispinterface
    ['{0002098F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property OutlineNumbered: WordBool dispid 1;
    property Name: WideString dispid 3;
    property ListLevels: ListLevels readonly dispid 2;
    function Convert(var Level: OleVariant): ListTemplate; dispid 101;
  end;

// *********************************************************************//
// Interface: ListTemplates
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020990-0000-0000-C000-000000000046}
// *********************************************************************//
  ListTemplates = interface(IDispatch)
    ['{00020990-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(var Index: OleVariant): ListTemplate; safecall;
    function Add(var OutlineNumbered: OleVariant; var Name: OleVariant): ListTemplate; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListTemplatesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020990-0000-0000-C000-000000000046}
// *********************************************************************//
  ListTemplatesDisp = dispinterface
    ['{00020990-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(var Index: OleVariant): ListTemplate; dispid 0;
    function Add(var OutlineNumbered: OleVariant; var Name: OleVariant): ListTemplate; dispid 100;
  end;

// *********************************************************************//
// Interface: ListParagraphs
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020991-0000-0000-C000-000000000046}
// *********************************************************************//
  ListParagraphs = interface(IDispatch)
    ['{00020991-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): Paragraph; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListParagraphsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020991-0000-0000-C000-000000000046}
// *********************************************************************//
  ListParagraphsDisp = dispinterface
    ['{00020991-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): Paragraph; dispid 0;
  end;

// *********************************************************************//
// Interface: List
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020992-0000-0000-C000-000000000046}
// *********************************************************************//
  List = interface(IDispatch)
    ['{00020992-0000-0000-C000-000000000046}']
    function Get_Range: Range; safecall;
    function Get_ListParagraphs: ListParagraphs; safecall;
    function Get_SingleListTemplate: WordBool; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    procedure ConvertNumbersToText(var NumberType: OleVariant); safecall;
    procedure RemoveNumbers(var NumberType: OleVariant); safecall;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; safecall;
    procedure ApplyListTemplateOld(const ListTemplate: ListTemplate; 
                                   var ContinuePreviousList: OleVariant); safecall;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WdContinue; safecall;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                var ContinuePreviousList: OleVariant; 
                                var DefaultListBehavior: OleVariant); safecall;
    property Range: Range read Get_Range;
    property ListParagraphs: ListParagraphs read Get_ListParagraphs;
    property SingleListTemplate: WordBool read Get_SingleListTemplate;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020992-0000-0000-C000-000000000046}
// *********************************************************************//
  ListDisp = dispinterface
    ['{00020992-0000-0000-C000-000000000046}']
    property Range: Range readonly dispid 1;
    property ListParagraphs: ListParagraphs readonly dispid 2;
    property SingleListTemplate: WordBool readonly dispid 3;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure ConvertNumbersToText(var NumberType: OleVariant); dispid 101;
    procedure RemoveNumbers(var NumberType: OleVariant); dispid 102;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; dispid 103;
    procedure ApplyListTemplateOld(const ListTemplate: ListTemplate; 
                                   var ContinuePreviousList: OleVariant); dispid 104;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WdContinue; dispid 105;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                var ContinuePreviousList: OleVariant; 
                                var DefaultListBehavior: OleVariant); dispid 106;
  end;

// *********************************************************************//
// Interface: Lists
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020993-0000-0000-C000-000000000046}
// *********************************************************************//
  Lists = interface(IDispatch)
    ['{00020993-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: Integer): List; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020993-0000-0000-C000-000000000046}
// *********************************************************************//
  ListsDisp = dispinterface
    ['{00020993-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: Integer): List; dispid 0;
  end;

// *********************************************************************//
// Interface: ListGallery
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020994-0000-0000-C000-000000000046}
// *********************************************************************//
  ListGallery = interface(IDispatch)
    ['{00020994-0000-0000-C000-000000000046}']
    function Get_ListTemplates: ListTemplates; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Modified(Index: Integer): WordBool; safecall;
    procedure Reset(Index: Integer); safecall;
    property ListTemplates: ListTemplates read Get_ListTemplates;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Modified[Index: Integer]: WordBool read Get_Modified;
  end;

// *********************************************************************//
// DispIntf:  ListGalleryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020994-0000-0000-C000-000000000046}
// *********************************************************************//
  ListGalleryDisp = dispinterface
    ['{00020994-0000-0000-C000-000000000046}']
    property ListTemplates: ListTemplates readonly dispid 1;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Modified[Index: Integer]: WordBool readonly dispid 101;
    procedure Reset(Index: Integer); dispid 100;
  end;

// *********************************************************************//
// Interface: ListGalleries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020995-0000-0000-C000-000000000046}
// *********************************************************************//
  ListGalleries = interface(IDispatch)
    ['{00020995-0000-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: WdListGalleryType): ListGallery; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ListGalleriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020995-0000-0000-C000-000000000046}
// *********************************************************************//
  ListGalleriesDisp = dispinterface
    ['{00020995-0000-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: WdListGalleryType): ListGallery; dispid 0;
  end;

// *********************************************************************//
// Interface: KeyBindings
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020996-0000-0000-C000-000000000046}
// *********************************************************************//
  KeyBindings = interface(IDispatch)
    ['{00020996-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Context: IDispatch; safecall;
    function Item(Index: Integer): KeyBinding; safecall;
    function Add(KeyCategory: WdKeyCategory; const Command: WideString; KeyCode: Integer; 
                 var KeyCode2: OleVariant; var CommandParameter: OleVariant): KeyBinding; safecall;
    procedure ClearAll; safecall;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Context: IDispatch read Get_Context;
  end;

// *********************************************************************//
// DispIntf:  KeyBindingsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020996-0000-0000-C000-000000000046}
// *********************************************************************//
  KeyBindingsDisp = dispinterface
    ['{00020996-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Context: IDispatch readonly dispid 10;
    function Item(Index: Integer): KeyBinding; dispid 0;
    function Add(KeyCategory: WdKeyCategory; const Command: WideString; KeyCode: Integer; 
                 var KeyCode2: OleVariant; var CommandParameter: OleVariant): KeyBinding; dispid 101;
    procedure ClearAll; dispid 102;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; dispid 110;
  end;

// *********************************************************************//
// Interface: KeysBoundTo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020997-0000-0000-C000-000000000046}
// *********************************************************************//
  KeysBoundTo = interface(IDispatch)
    ['{00020997-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_KeyCategory: WdKeyCategory; safecall;
    function Get_Command: WideString; safecall;
    function Get_CommandParameter: WideString; safecall;
    function Get_Context: IDispatch; safecall;
    function Item(Index: Integer): KeyBinding; safecall;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property KeyCategory: WdKeyCategory read Get_KeyCategory;
    property Command: WideString read Get_Command;
    property CommandParameter: WideString read Get_CommandParameter;
    property Context: IDispatch read Get_Context;
  end;

// *********************************************************************//
// DispIntf:  KeysBoundToDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020997-0000-0000-C000-000000000046}
// *********************************************************************//
  KeysBoundToDisp = dispinterface
    ['{00020997-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property KeyCategory: WdKeyCategory readonly dispid 3;
    property Command: WideString readonly dispid 4;
    property CommandParameter: WideString readonly dispid 5;
    property Context: IDispatch readonly dispid 10;
    function Item(Index: Integer): KeyBinding; dispid 0;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; dispid 1;
  end;

// *********************************************************************//
// Interface: KeyBinding
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020998-0000-0000-C000-000000000046}
// *********************************************************************//
  KeyBinding = interface(IDispatch)
    ['{00020998-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Command: WideString; safecall;
    function Get_KeyString: WideString; safecall;
    function Get_Protected_: WordBool; safecall;
    function Get_KeyCategory: WdKeyCategory; safecall;
    function Get_KeyCode: Integer; safecall;
    function Get_KeyCode2: Integer; safecall;
    function Get_CommandParameter: WideString; safecall;
    function Get_Context: IDispatch; safecall;
    procedure Clear; safecall;
    procedure Disable; safecall;
    procedure Execute; safecall;
    procedure Rebind(KeyCategory: WdKeyCategory; const Command: WideString; 
                     var CommandParameter: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Command: WideString read Get_Command;
    property KeyString: WideString read Get_KeyString;
    property Protected_: WordBool read Get_Protected_;
    property KeyCategory: WdKeyCategory read Get_KeyCategory;
    property KeyCode: Integer read Get_KeyCode;
    property KeyCode2: Integer read Get_KeyCode2;
    property CommandParameter: WideString read Get_CommandParameter;
    property Context: IDispatch read Get_Context;
  end;

// *********************************************************************//
// DispIntf:  KeyBindingDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020998-0000-0000-C000-000000000046}
// *********************************************************************//
  KeyBindingDisp = dispinterface
    ['{00020998-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Command: WideString readonly dispid 1;
    property KeyString: WideString readonly dispid 2;
    property Protected_: WordBool readonly dispid 3;
    property KeyCategory: WdKeyCategory readonly dispid 4;
    property KeyCode: Integer readonly dispid 6;
    property KeyCode2: Integer readonly dispid 7;
    property CommandParameter: WideString readonly dispid 8;
    property Context: IDispatch readonly dispid 10;
    procedure Clear; dispid 101;
    procedure Disable; dispid 102;
    procedure Execute; dispid 103;
    procedure Rebind(KeyCategory: WdKeyCategory; const Command: WideString; 
                     var CommandParameter: OleVariant); dispid 104;
  end;

// *********************************************************************//
// Interface: FileConverter
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020999-0000-0000-C000-000000000046}
// *********************************************************************//
  FileConverter = interface(IDispatch)
    ['{00020999-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_FormatName: WideString; safecall;
    function Get_ClassName: WideString; safecall;
    function Get_SaveFormat: Integer; safecall;
    function Get_OpenFormat: Integer; safecall;
    function Get_CanSave: WordBool; safecall;
    function Get_CanOpen: WordBool; safecall;
    function Get_Path: WideString; safecall;
    function Get_Name: WideString; safecall;
    function Get_Extensions: WideString; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property FormatName: WideString read Get_FormatName;
    property ClassName: WideString read Get_ClassName;
    property SaveFormat: Integer read Get_SaveFormat;
    property OpenFormat: Integer read Get_OpenFormat;
    property CanSave: WordBool read Get_CanSave;
    property CanOpen: WordBool read Get_CanOpen;
    property Path: WideString read Get_Path;
    property Name: WideString read Get_Name;
    property Extensions: WideString read Get_Extensions;
  end;

// *********************************************************************//
// DispIntf:  FileConverterDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020999-0000-0000-C000-000000000046}
// *********************************************************************//
  FileConverterDisp = dispinterface
    ['{00020999-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property FormatName: WideString readonly dispid 0;
    property ClassName: WideString readonly dispid 1;
    property SaveFormat: Integer readonly dispid 2;
    property OpenFormat: Integer readonly dispid 3;
    property CanSave: WordBool readonly dispid 4;
    property CanOpen: WordBool readonly dispid 5;
    property Path: WideString readonly dispid 6;
    property Name: WideString readonly dispid 7;
    property Extensions: WideString readonly dispid 8;
  end;

// *********************************************************************//
// Interface: FileConverters
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099A-0000-0000-C000-000000000046}
// *********************************************************************//
  FileConverters = interface(IDispatch)
    ['{0002099A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_ConvertMacWordChevrons: WdChevronConvertRule; safecall;
    procedure Set_ConvertMacWordChevrons(prop: WdChevronConvertRule); safecall;
    function Item(var Index: OleVariant): FileConverter; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
    property ConvertMacWordChevrons: WdChevronConvertRule read Get_ConvertMacWordChevrons write Set_ConvertMacWordChevrons;
  end;

// *********************************************************************//
// DispIntf:  FileConvertersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099A-0000-0000-C000-000000000046}
// *********************************************************************//
  FileConvertersDisp = dispinterface
    ['{0002099A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    property ConvertMacWordChevrons: WdChevronConvertRule dispid 2;
    function Item(var Index: OleVariant): FileConverter; dispid 0;
  end;

// *********************************************************************//
// Interface: SynonymInfo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099B-0000-0000-C000-000000000046}
// *********************************************************************//
  SynonymInfo = interface(IDispatch)
    ['{0002099B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Word: WideString; safecall;
    function Get_Found: WordBool; safecall;
    function Get_MeaningCount: Integer; safecall;
    function Get_MeaningList: OleVariant; safecall;
    function Get_PartOfSpeechList: OleVariant; safecall;
    function Get_SynonymList(var Meaning: OleVariant): OleVariant; safecall;
    function Get_AntonymList: OleVariant; safecall;
    function Get_RelatedExpressionList: OleVariant; safecall;
    function Get_RelatedWordList: OleVariant; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _Word: WideString read Get_Word;
    property Found: WordBool read Get_Found;
    property MeaningCount: Integer read Get_MeaningCount;
    property MeaningList: OleVariant read Get_MeaningList;
    property PartOfSpeechList: OleVariant read Get_PartOfSpeechList;
    property SynonymList[var Meaning: OleVariant]: OleVariant read Get_SynonymList;
    property AntonymList: OleVariant read Get_AntonymList;
    property RelatedExpressionList: OleVariant read Get_RelatedExpressionList;
    property RelatedWordList: OleVariant read Get_RelatedWordList;
  end;

// *********************************************************************//
// DispIntf:  SynonymInfoDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099B-0000-0000-C000-000000000046}
// *********************************************************************//
  SynonymInfoDisp = dispinterface
    ['{0002099B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _Word: WideString readonly dispid 1;
    property Found: WordBool readonly dispid 2;
    property MeaningCount: Integer readonly dispid 3;
    property MeaningList: OleVariant readonly dispid 4;
    property PartOfSpeechList: OleVariant readonly dispid 5;
    property SynonymList[var Meaning: OleVariant]: OleVariant readonly dispid 7;
    property AntonymList: OleVariant readonly dispid 8;
    property RelatedExpressionList: OleVariant readonly dispid 9;
    property RelatedWordList: OleVariant readonly dispid 10;
  end;

// *********************************************************************//
// Interface: Hyperlinks
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099C-0000-0000-C000-000000000046}
// *********************************************************************//
  Hyperlinks = interface(IDispatch)
    ['{0002099C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): Hyperlink; safecall;
    function _Add(const Anchor: IDispatch; var Address: OleVariant; var SubAddress: OleVariant): Hyperlink; safecall;
    function Add(const Anchor: IDispatch; var Address: OleVariant; var SubAddress: OleVariant; 
                 var ScreenTip: OleVariant; var TextToDisplay: OleVariant; var Target: OleVariant): Hyperlink; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  HyperlinksDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099C-0000-0000-C000-000000000046}
// *********************************************************************//
  HyperlinksDisp = dispinterface
    ['{0002099C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): Hyperlink; dispid 0;
    function _Add(const Anchor: IDispatch; var Address: OleVariant; var SubAddress: OleVariant): Hyperlink; dispid 100;
    function Add(const Anchor: IDispatch; var Address: OleVariant; var SubAddress: OleVariant; 
                 var ScreenTip: OleVariant; var TextToDisplay: OleVariant; var Target: OleVariant): Hyperlink; dispid 101;
  end;

// *********************************************************************//
// Interface: Hyperlink
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099D-0000-0000-C000-000000000046}
// *********************************************************************//
  Hyperlink = interface(IDispatch)
    ['{0002099D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_AddressOld: WideString; safecall;
    function Get_Type_: MsoHyperlinkType; safecall;
    function Get_Range: Range; safecall;
    function Get_Shape: Shape; safecall;
    function Get_SubAddressOld: WideString; safecall;
    function Get_ExtraInfoRequired: WordBool; safecall;
    procedure Delete; safecall;
    procedure Follow(var NewWindow: OleVariant; var AddHistory: OleVariant; 
                     var ExtraInfo: OleVariant; var Method: OleVariant; var HeaderInfo: OleVariant); safecall;
    procedure AddToFavorites; safecall;
    procedure CreateNewDocument(const FileName: WideString; EditNow: WordBool; Overwrite: WordBool); safecall;
    function Get_Address: WideString; safecall;
    procedure Set_Address(const prop: WideString); safecall;
    function Get_SubAddress: WideString; safecall;
    procedure Set_SubAddress(const prop: WideString); safecall;
    function Get_EmailSubject: WideString; safecall;
    procedure Set_EmailSubject(const prop: WideString); safecall;
    function Get_ScreenTip: WideString; safecall;
    procedure Set_ScreenTip(const prop: WideString); safecall;
    function Get_TextToDisplay: WideString; safecall;
    procedure Set_TextToDisplay(const prop: WideString); safecall;
    function Get_Target: WideString; safecall;
    procedure Set_Target(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property AddressOld: WideString read Get_AddressOld;
    property Type_: MsoHyperlinkType read Get_Type_;
    property Range: Range read Get_Range;
    property Shape: Shape read Get_Shape;
    property SubAddressOld: WideString read Get_SubAddressOld;
    property ExtraInfoRequired: WordBool read Get_ExtraInfoRequired;
    property Address: WideString read Get_Address write Set_Address;
    property SubAddress: WideString read Get_SubAddress write Set_SubAddress;
    property EmailSubject: WideString read Get_EmailSubject write Set_EmailSubject;
    property ScreenTip: WideString read Get_ScreenTip write Set_ScreenTip;
    property TextToDisplay: WideString read Get_TextToDisplay write Set_TextToDisplay;
    property Target: WideString read Get_Target write Set_Target;
  end;

// *********************************************************************//
// DispIntf:  HyperlinkDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099D-0000-0000-C000-000000000046}
// *********************************************************************//
  HyperlinkDisp = dispinterface
    ['{0002099D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 1003;
    property AddressOld: WideString readonly dispid 1004;
    property Type_: MsoHyperlinkType readonly dispid 1005;
    property Range: Range readonly dispid 1006;
    property Shape: Shape readonly dispid 1007;
    property SubAddressOld: WideString readonly dispid 1008;
    property ExtraInfoRequired: WordBool readonly dispid 1009;
    procedure Delete; dispid 103;
    procedure Follow(var NewWindow: OleVariant; var AddHistory: OleVariant; 
                     var ExtraInfo: OleVariant; var Method: OleVariant; var HeaderInfo: OleVariant); dispid 104;
    procedure AddToFavorites; dispid 105;
    procedure CreateNewDocument(const FileName: WideString; EditNow: WordBool; Overwrite: WordBool); dispid 106;
    property Address: WideString dispid 1100;
    property SubAddress: WideString dispid 1101;
    property EmailSubject: WideString dispid 1010;
    property ScreenTip: WideString dispid 1011;
    property TextToDisplay: WideString dispid 1012;
    property Target: WideString dispid 1013;
  end;

// *********************************************************************//
// Interface: Shapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099F-0000-0000-C000-000000000046}
// *********************************************************************//
  Shapes = interface(IDispatch)
    ['{0002099F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): Shape; safecall;
    function AddCallout(Type_: MsoCalloutType; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; safecall;
    function AddConnector(Type_: MsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                          EndY: Single): Shape; safecall;
    function AddCurve(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddLabel(Orientation: MsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                      Height: Single; var Anchor: OleVariant): Shape; safecall;
    function AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single; 
                     var Anchor: OleVariant): Shape; safecall;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Left: OleVariant; 
                        var Top: OleVariant; var Width: OleVariant; var Height: OleVariant; 
                        var Anchor: OleVariant): Shape; safecall;
    function AddPolyline(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddShape(Type_: Integer; Left: Single; Top: Single; Width: Single; Height: Single; 
                      var Anchor: OleVariant): Shape; safecall;
    function AddTextEffect(PresetTextEffect: MsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: MsoTriState; 
                           FontItalic: MsoTriState; Left: Single; Top: Single; 
                           var Anchor: OleVariant): Shape; safecall;
    function AddTextbox(Orientation: MsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; safecall;
    function BuildFreeform(EditingType: MsoEditingType; X1: Single; Y1: Single): FreeformBuilder; safecall;
    function Range(var Index: OleVariant): ShapeRange; safecall;
    procedure SelectAll; safecall;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                          var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddOLEControl(var ClassType: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                           var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  ShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099F-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapesDisp = dispinterface
    ['{0002099F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 8000;
    property Creator: Integer readonly dispid 8001;
    property Parent: IDispatch readonly dispid 1;
    property Count: Integer readonly dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): Shape; dispid 0;
    function AddCallout(Type_: MsoCalloutType; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; dispid 10;
    function AddConnector(Type_: MsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                          EndY: Single): Shape; dispid 11;
    function AddCurve(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; dispid 12;
    function AddLabel(Orientation: MsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                      Height: Single; var Anchor: OleVariant): Shape; dispid 13;
    function AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single; 
                     var Anchor: OleVariant): Shape; dispid 14;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Left: OleVariant; 
                        var Top: OleVariant; var Width: OleVariant; var Height: OleVariant; 
                        var Anchor: OleVariant): Shape; dispid 15;
    function AddPolyline(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; dispid 16;
    function AddShape(Type_: Integer; Left: Single; Top: Single; Width: Single; Height: Single; 
                      var Anchor: OleVariant): Shape; dispid 17;
    function AddTextEffect(PresetTextEffect: MsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: MsoTriState; 
                           FontItalic: MsoTriState; Left: Single; Top: Single; 
                           var Anchor: OleVariant): Shape; dispid 18;
    function AddTextbox(Orientation: MsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; dispid 19;
    function BuildFreeform(EditingType: MsoEditingType; X1: Single; Y1: Single): FreeformBuilder; dispid 20;
    function Range(var Index: OleVariant): ShapeRange; dispid 21;
    procedure SelectAll; dispid 22;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                          var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; dispid 24;
    function AddOLEControl(var ClassType: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                           var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; dispid 102;
  end;

// *********************************************************************//
// Interface: ShapeRange
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B5-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeRange = interface(IDispatch)
    ['{000209B5-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Adjustments: Adjustments; safecall;
    function Get_AutoShapeType: MsoAutoShapeType; safecall;
    procedure Set_AutoShapeType(prop: MsoAutoShapeType); safecall;
    function Get_Callout: CalloutFormat; safecall;
    function Get_ConnectionSiteCount: Integer; safecall;
    function Get_Connector: MsoTriState; safecall;
    function Get_ConnectorFormat: ConnectorFormat; safecall;
    function Get_Fill: FillFormat; safecall;
    function Get_GroupItems: GroupShapes; safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HorizontalFlip: MsoTriState; safecall;
    function Get_Left: Single; safecall;
    procedure Set_Left(prop: Single); safecall;
    function Get_Line: LineFormat; safecall;
    function Get_LockAspectRatio: MsoTriState; safecall;
    procedure Set_LockAspectRatio(prop: MsoTriState); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Nodes: ShapeNodes; safecall;
    function Get_Rotation: Single; safecall;
    procedure Set_Rotation(prop: Single); safecall;
    function Get_PictureFormat: PictureFormat; safecall;
    function Get_Shadow: ShadowFormat; safecall;
    function Get_TextEffect: TextEffectFormat; safecall;
    function Get_TextFrame: TextFrame; safecall;
    function Get_ThreeD: ThreeDFormat; safecall;
    function Get_Top: Single; safecall;
    procedure Set_Top(prop: Single); safecall;
    function Get_Type_: MsoShapeType; safecall;
    function Get_VerticalFlip: MsoTriState; safecall;
    function Get_Vertices: OleVariant; safecall;
    function Get_Visible: MsoTriState; safecall;
    procedure Set_Visible(prop: MsoTriState); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_ZOrderPosition: Integer; safecall;
    function Get_Hyperlink: Hyperlink; safecall;
    function Get_RelativeHorizontalPosition: WdRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WdRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WdRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WdRelativeVerticalPosition); safecall;
    function Get_LockAnchor: Integer; safecall;
    procedure Set_LockAnchor(prop: Integer); safecall;
    function Get_WrapFormat: WrapFormat; safecall;
    function Get_Anchor: Range; safecall;
    function Item(var Index: OleVariant): Shape; safecall;
    procedure Align(Align: MsoAlignCmd; RelativeTo: Integer); safecall;
    procedure Apply; safecall;
    procedure Delete; safecall;
    procedure Distribute(Distribute: MsoDistributeCmd; RelativeTo: Integer); safecall;
    function Duplicate: ShapeRange; safecall;
    procedure Flip(FlipCmd: MsoFlipCmd); safecall;
    procedure IncrementLeft(Increment: Single); safecall;
    procedure IncrementRotation(Increment: Single); safecall;
    procedure IncrementTop(Increment: Single); safecall;
    function Group: Shape; safecall;
    procedure PickUp; safecall;
    function Regroup: Shape; safecall;
    procedure RerouteConnections; safecall;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); safecall;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); safecall;
    procedure Select(var Replace: OleVariant); safecall;
    procedure SetShapesDefaultProperties; safecall;
    function Ungroup: ShapeRange; safecall;
    procedure ZOrder(ZOrderCmd: MsoZOrderCmd); safecall;
    function ConvertToFrame: Frame; safecall;
    function ConvertToInlineShape: InlineShape; safecall;
    procedure Activate; safecall;
    function Get_AlternativeText: WideString; safecall;
    procedure Set_AlternativeText(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Adjustments: Adjustments read Get_Adjustments;
    property AutoShapeType: MsoAutoShapeType read Get_AutoShapeType write Set_AutoShapeType;
    property Callout: CalloutFormat read Get_Callout;
    property ConnectionSiteCount: Integer read Get_ConnectionSiteCount;
    property Connector: MsoTriState read Get_Connector;
    property ConnectorFormat: ConnectorFormat read Get_ConnectorFormat;
    property Fill: FillFormat read Get_Fill;
    property GroupItems: GroupShapes read Get_GroupItems;
    property Height: Single read Get_Height write Set_Height;
    property HorizontalFlip: MsoTriState read Get_HorizontalFlip;
    property Left: Single read Get_Left write Set_Left;
    property Line: LineFormat read Get_Line;
    property LockAspectRatio: MsoTriState read Get_LockAspectRatio write Set_LockAspectRatio;
    property Name: WideString read Get_Name write Set_Name;
    property Nodes: ShapeNodes read Get_Nodes;
    property Rotation: Single read Get_Rotation write Set_Rotation;
    property PictureFormat: PictureFormat read Get_PictureFormat;
    property Shadow: ShadowFormat read Get_Shadow;
    property TextEffect: TextEffectFormat read Get_TextEffect;
    property TextFrame: TextFrame read Get_TextFrame;
    property ThreeD: ThreeDFormat read Get_ThreeD;
    property Top: Single read Get_Top write Set_Top;
    property Type_: MsoShapeType read Get_Type_;
    property VerticalFlip: MsoTriState read Get_VerticalFlip;
    property Vertices: OleVariant read Get_Vertices;
    property Visible: MsoTriState read Get_Visible write Set_Visible;
    property Width: Single read Get_Width write Set_Width;
    property ZOrderPosition: Integer read Get_ZOrderPosition;
    property Hyperlink: Hyperlink read Get_Hyperlink;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WdRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property LockAnchor: Integer read Get_LockAnchor write Set_LockAnchor;
    property WrapFormat: WrapFormat read Get_WrapFormat;
    property Anchor: Range read Get_Anchor;
    property AlternativeText: WideString read Get_AlternativeText write Set_AlternativeText;
  end;

// *********************************************************************//
// DispIntf:  ShapeRangeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B5-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeRangeDisp = dispinterface
    ['{000209B5-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 8000;
    property Creator: Integer readonly dispid 8001;
    property Parent: IDispatch readonly dispid 1;
    property Count: Integer readonly dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    property Adjustments: Adjustments readonly dispid 100;
    property AutoShapeType: MsoAutoShapeType dispid 101;
    property Callout: CalloutFormat readonly dispid 103;
    property ConnectionSiteCount: Integer readonly dispid 104;
    property Connector: MsoTriState readonly dispid 105;
    property ConnectorFormat: ConnectorFormat readonly dispid 106;
    property Fill: FillFormat readonly dispid 107;
    property GroupItems: GroupShapes readonly dispid 108;
    property Height: Single dispid 109;
    property HorizontalFlip: MsoTriState readonly dispid 110;
    property Left: Single dispid 111;
    property Line: LineFormat readonly dispid 112;
    property LockAspectRatio: MsoTriState dispid 113;
    property Name: WideString dispid 115;
    property Nodes: ShapeNodes readonly dispid 116;
    property Rotation: Single dispid 117;
    property PictureFormat: PictureFormat readonly dispid 118;
    property Shadow: ShadowFormat readonly dispid 119;
    property TextEffect: TextEffectFormat readonly dispid 120;
    property TextFrame: TextFrame readonly dispid 121;
    property ThreeD: ThreeDFormat readonly dispid 122;
    property Top: Single dispid 123;
    property Type_: MsoShapeType readonly dispid 124;
    property VerticalFlip: MsoTriState readonly dispid 125;
    property Vertices: OleVariant readonly dispid 126;
    property Visible: MsoTriState dispid 127;
    property Width: Single dispid 128;
    property ZOrderPosition: Integer readonly dispid 129;
    property Hyperlink: Hyperlink readonly dispid 1001;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition dispid 300;
    property RelativeVerticalPosition: WdRelativeVerticalPosition dispid 301;
    property LockAnchor: Integer dispid 302;
    property WrapFormat: WrapFormat readonly dispid 303;
    property Anchor: Range readonly dispid 304;
    function Item(var Index: OleVariant): Shape; dispid 0;
    procedure Align(Align: MsoAlignCmd; RelativeTo: Integer); dispid 10;
    procedure Apply; dispid 11;
    procedure Delete; dispid 12;
    procedure Distribute(Distribute: MsoDistributeCmd; RelativeTo: Integer); dispid 13;
    function Duplicate: ShapeRange; dispid 14;
    procedure Flip(FlipCmd: MsoFlipCmd); dispid 15;
    procedure IncrementLeft(Increment: Single); dispid 16;
    procedure IncrementRotation(Increment: Single); dispid 17;
    procedure IncrementTop(Increment: Single); dispid 18;
    function Group: Shape; dispid 19;
    procedure PickUp; dispid 20;
    function Regroup: Shape; dispid 21;
    procedure RerouteConnections; dispid 22;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); dispid 23;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); dispid 24;
    procedure Select(var Replace: OleVariant); dispid 25;
    procedure SetShapesDefaultProperties; dispid 26;
    function Ungroup: ShapeRange; dispid 27;
    procedure ZOrder(ZOrderCmd: MsoZOrderCmd); dispid 28;
    function ConvertToFrame: Frame; dispid 29;
    function ConvertToInlineShape: InlineShape; dispid 30;
    procedure Activate; dispid 50;
    property AlternativeText: WideString dispid 131;
  end;

// *********************************************************************//
// Interface: GroupShapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B6-0000-0000-C000-000000000046}
// *********************************************************************//
  GroupShapes = interface(IDispatch)
    ['{000209B6-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): Shape; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  GroupShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B6-0000-0000-C000-000000000046}
// *********************************************************************//
  GroupShapesDisp = dispinterface
    ['{000209B6-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 8000;
    property Creator: Integer readonly dispid 8001;
    property Parent: IDispatch readonly dispid 1;
    property Count: Integer readonly dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): Shape; dispid 0;
  end;

// *********************************************************************//
// Interface: Shape
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A0-0000-0000-C000-000000000046}
// *********************************************************************//
  Shape = interface(IDispatch)
    ['{000209A0-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Adjustments: Adjustments; safecall;
    function Get_AutoShapeType: MsoAutoShapeType; safecall;
    procedure Set_AutoShapeType(prop: MsoAutoShapeType); safecall;
    function Get_Callout: CalloutFormat; safecall;
    function Get_ConnectionSiteCount: Integer; safecall;
    function Get_Connector: MsoTriState; safecall;
    function Get_ConnectorFormat: ConnectorFormat; safecall;
    function Get_Fill: FillFormat; safecall;
    function Get_GroupItems: GroupShapes; safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HorizontalFlip: MsoTriState; safecall;
    function Get_Left: Single; safecall;
    procedure Set_Left(prop: Single); safecall;
    function Get_Line: LineFormat; safecall;
    function Get_LockAspectRatio: MsoTriState; safecall;
    procedure Set_LockAspectRatio(prop: MsoTriState); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Nodes: ShapeNodes; safecall;
    function Get_Rotation: Single; safecall;
    procedure Set_Rotation(prop: Single); safecall;
    function Get_PictureFormat: PictureFormat; safecall;
    function Get_Shadow: ShadowFormat; safecall;
    function Get_TextEffect: TextEffectFormat; safecall;
    function Get_TextFrame: TextFrame; safecall;
    function Get_ThreeD: ThreeDFormat; safecall;
    function Get_Top: Single; safecall;
    procedure Set_Top(prop: Single); safecall;
    function Get_Type_: MsoShapeType; safecall;
    function Get_VerticalFlip: MsoTriState; safecall;
    function Get_Vertices: OleVariant; safecall;
    function Get_Visible: MsoTriState; safecall;
    procedure Set_Visible(prop: MsoTriState); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_ZOrderPosition: Integer; safecall;
    function Get_Hyperlink: Hyperlink; safecall;
    function Get_RelativeHorizontalPosition: WdRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WdRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WdRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WdRelativeVerticalPosition); safecall;
    function Get_LockAnchor: Integer; safecall;
    procedure Set_LockAnchor(prop: Integer); safecall;
    function Get_WrapFormat: WrapFormat; safecall;
    function Get_OLEFormat: OLEFormat; safecall;
    function Get_Anchor: Range; safecall;
    function Get_LinkFormat: LinkFormat; safecall;
    procedure Apply; safecall;
    procedure Delete; safecall;
    function Duplicate: Shape; safecall;
    procedure Flip(FlipCmd: MsoFlipCmd); safecall;
    procedure IncrementLeft(Increment: Single); safecall;
    procedure IncrementRotation(Increment: Single); safecall;
    procedure IncrementTop(Increment: Single); safecall;
    procedure PickUp; safecall;
    procedure RerouteConnections; safecall;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); safecall;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); safecall;
    procedure Select(var Replace: OleVariant); safecall;
    procedure SetShapesDefaultProperties; safecall;
    function Ungroup: ShapeRange; safecall;
    procedure ZOrder(ZOrderCmd: MsoZOrderCmd); safecall;
    function ConvertToInlineShape: InlineShape; safecall;
    function ConvertToFrame: Frame; safecall;
    procedure Activate; safecall;
    function Get_AlternativeText: WideString; safecall;
    procedure Set_AlternativeText(const prop: WideString); safecall;
    function Get_Script: Script; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Adjustments: Adjustments read Get_Adjustments;
    property AutoShapeType: MsoAutoShapeType read Get_AutoShapeType write Set_AutoShapeType;
    property Callout: CalloutFormat read Get_Callout;
    property ConnectionSiteCount: Integer read Get_ConnectionSiteCount;
    property Connector: MsoTriState read Get_Connector;
    property ConnectorFormat: ConnectorFormat read Get_ConnectorFormat;
    property Fill: FillFormat read Get_Fill;
    property GroupItems: GroupShapes read Get_GroupItems;
    property Height: Single read Get_Height write Set_Height;
    property HorizontalFlip: MsoTriState read Get_HorizontalFlip;
    property Left: Single read Get_Left write Set_Left;
    property Line: LineFormat read Get_Line;
    property LockAspectRatio: MsoTriState read Get_LockAspectRatio write Set_LockAspectRatio;
    property Name: WideString read Get_Name write Set_Name;
    property Nodes: ShapeNodes read Get_Nodes;
    property Rotation: Single read Get_Rotation write Set_Rotation;
    property PictureFormat: PictureFormat read Get_PictureFormat;
    property Shadow: ShadowFormat read Get_Shadow;
    property TextEffect: TextEffectFormat read Get_TextEffect;
    property TextFrame: TextFrame read Get_TextFrame;
    property ThreeD: ThreeDFormat read Get_ThreeD;
    property Top: Single read Get_Top write Set_Top;
    property Type_: MsoShapeType read Get_Type_;
    property VerticalFlip: MsoTriState read Get_VerticalFlip;
    property Vertices: OleVariant read Get_Vertices;
    property Visible: MsoTriState read Get_Visible write Set_Visible;
    property Width: Single read Get_Width write Set_Width;
    property ZOrderPosition: Integer read Get_ZOrderPosition;
    property Hyperlink: Hyperlink read Get_Hyperlink;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WdRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property LockAnchor: Integer read Get_LockAnchor write Set_LockAnchor;
    property WrapFormat: WrapFormat read Get_WrapFormat;
    property OLEFormat: OLEFormat read Get_OLEFormat;
    property Anchor: Range read Get_Anchor;
    property LinkFormat: LinkFormat read Get_LinkFormat;
    property AlternativeText: WideString read Get_AlternativeText write Set_AlternativeText;
    property Script: Script read Get_Script;
  end;

// *********************************************************************//
// DispIntf:  ShapeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A0-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeDisp = dispinterface
    ['{000209A0-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 8000;
    property Creator: Integer readonly dispid 8001;
    property Parent: IDispatch readonly dispid 1;
    property Adjustments: Adjustments readonly dispid 100;
    property AutoShapeType: MsoAutoShapeType dispid 101;
    property Callout: CalloutFormat readonly dispid 103;
    property ConnectionSiteCount: Integer readonly dispid 104;
    property Connector: MsoTriState readonly dispid 105;
    property ConnectorFormat: ConnectorFormat readonly dispid 106;
    property Fill: FillFormat readonly dispid 107;
    property GroupItems: GroupShapes readonly dispid 108;
    property Height: Single dispid 109;
    property HorizontalFlip: MsoTriState readonly dispid 110;
    property Left: Single dispid 111;
    property Line: LineFormat readonly dispid 112;
    property LockAspectRatio: MsoTriState dispid 113;
    property Name: WideString dispid 115;
    property Nodes: ShapeNodes readonly dispid 116;
    property Rotation: Single dispid 117;
    property PictureFormat: PictureFormat readonly dispid 118;
    property Shadow: ShadowFormat readonly dispid 119;
    property TextEffect: TextEffectFormat readonly dispid 120;
    property TextFrame: TextFrame readonly dispid 121;
    property ThreeD: ThreeDFormat readonly dispid 122;
    property Top: Single dispid 123;
    property Type_: MsoShapeType readonly dispid 124;
    property VerticalFlip: MsoTriState readonly dispid 125;
    property Vertices: OleVariant readonly dispid 126;
    property Visible: MsoTriState dispid 127;
    property Width: Single dispid 128;
    property ZOrderPosition: Integer readonly dispid 129;
    property Hyperlink: Hyperlink readonly dispid 1001;
    property RelativeHorizontalPosition: WdRelativeHorizontalPosition dispid 300;
    property RelativeVerticalPosition: WdRelativeVerticalPosition dispid 301;
    property LockAnchor: Integer dispid 302;
    property WrapFormat: WrapFormat readonly dispid 303;
    property OLEFormat: OLEFormat readonly dispid 500;
    property Anchor: Range readonly dispid 501;
    property LinkFormat: LinkFormat readonly dispid 502;
    procedure Apply; dispid 10;
    procedure Delete; dispid 11;
    function Duplicate: Shape; dispid 12;
    procedure Flip(FlipCmd: MsoFlipCmd); dispid 13;
    procedure IncrementLeft(Increment: Single); dispid 14;
    procedure IncrementRotation(Increment: Single); dispid 15;
    procedure IncrementTop(Increment: Single); dispid 16;
    procedure PickUp; dispid 17;
    procedure RerouteConnections; dispid 18;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); dispid 19;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: MsoTriState; Scale: MsoScaleFrom); dispid 20;
    procedure Select(var Replace: OleVariant); dispid 21;
    procedure SetShapesDefaultProperties; dispid 22;
    function Ungroup: ShapeRange; dispid 23;
    procedure ZOrder(ZOrderCmd: MsoZOrderCmd); dispid 24;
    function ConvertToInlineShape: InlineShape; dispid 25;
    function ConvertToFrame: Frame; dispid 29;
    procedure Activate; dispid 50;
    property AlternativeText: WideString dispid 131;
    property Script: Script readonly dispid 503;
  end;

// *********************************************************************//
// Interface: TextFrame
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B2-0000-0000-C000-000000000046}
// *********************************************************************//
  TextFrame = interface(IDispatch)
    ['{000209B2-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: Shape; safecall;
    function Get_MarginBottom: Single; safecall;
    procedure Set_MarginBottom(prop: Single); safecall;
    function Get_MarginLeft: Single; safecall;
    procedure Set_MarginLeft(prop: Single); safecall;
    function Get_MarginRight: Single; safecall;
    procedure Set_MarginRight(prop: Single); safecall;
    function Get_MarginTop: Single; safecall;
    procedure Set_MarginTop(prop: Single); safecall;
    function Get_Orientation: MsoTextOrientation; safecall;
    procedure Set_Orientation(prop: MsoTextOrientation); safecall;
    function Get_TextRange: Range; safecall;
    function Get_ContainingRange: Range; safecall;
    function Get_Next: TextFrame; safecall;
    procedure Set_Next(const prop: TextFrame); safecall;
    function Get_Previous: TextFrame; safecall;
    procedure Set_Previous(const prop: TextFrame); safecall;
    function Get_Overflowing: WordBool; safecall;
    function Get_HasText: Integer; safecall;
    procedure BreakForwardLink; safecall;
    function ValidLinkTarget(const TargetTextFrame: TextFrame): WordBool; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: Shape read Get_Parent;
    property MarginBottom: Single read Get_MarginBottom write Set_MarginBottom;
    property MarginLeft: Single read Get_MarginLeft write Set_MarginLeft;
    property MarginRight: Single read Get_MarginRight write Set_MarginRight;
    property MarginTop: Single read Get_MarginTop write Set_MarginTop;
    property Orientation: MsoTextOrientation read Get_Orientation write Set_Orientation;
    property TextRange: Range read Get_TextRange;
    property ContainingRange: Range read Get_ContainingRange;
    property Next: TextFrame read Get_Next write Set_Next;
    property Previous: TextFrame read Get_Previous write Set_Previous;
    property Overflowing: WordBool read Get_Overflowing;
    property HasText: Integer read Get_HasText;
  end;

// *********************************************************************//
// DispIntf:  TextFrameDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B2-0000-0000-C000-000000000046}
// *********************************************************************//
  TextFrameDisp = dispinterface
    ['{000209B2-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 8000;
    property Creator: Integer readonly dispid 8001;
    property Parent: Shape readonly dispid 1;
    property MarginBottom: Single dispid 100;
    property MarginLeft: Single dispid 101;
    property MarginRight: Single dispid 102;
    property MarginTop: Single dispid 103;
    property Orientation: MsoTextOrientation dispid 104;
    property TextRange: Range readonly dispid 1001;
    property ContainingRange: Range readonly dispid 1002;
    property Next: TextFrame dispid 5001;
    property Previous: TextFrame dispid 5002;
    property Overflowing: WordBool readonly dispid 5003;
    property HasText: Integer readonly dispid 5008;
    procedure BreakForwardLink; dispid 5004;
    function ValidLinkTarget(const TargetTextFrame: TextFrame): WordBool; dispid 5006;
  end;

// *********************************************************************//
// Interface: _LetterContent
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A1-0000-0000-C000-000000000046}
// *********************************************************************//
  _LetterContent = interface(IDispatch)
    ['{000209A1-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Duplicate: LetterContent; safecall;
    function Get_DateFormat: WideString; safecall;
    procedure Set_DateFormat(const prop: WideString); safecall;
    function Get_IncludeHeaderFooter: WordBool; safecall;
    procedure Set_IncludeHeaderFooter(prop: WordBool); safecall;
    function Get_PageDesign: WideString; safecall;
    procedure Set_PageDesign(const prop: WideString); safecall;
    function Get_LetterStyle: WdLetterStyle; safecall;
    procedure Set_LetterStyle(prop: WdLetterStyle); safecall;
    function Get_Letterhead: WordBool; safecall;
    procedure Set_Letterhead(prop: WordBool); safecall;
    function Get_LetterheadLocation: WdLetterheadLocation; safecall;
    procedure Set_LetterheadLocation(prop: WdLetterheadLocation); safecall;
    function Get_LetterheadSize: Single; safecall;
    procedure Set_LetterheadSize(prop: Single); safecall;
    function Get_RecipientName: WideString; safecall;
    procedure Set_RecipientName(const prop: WideString); safecall;
    function Get_RecipientAddress: WideString; safecall;
    procedure Set_RecipientAddress(const prop: WideString); safecall;
    function Get_Salutation: WideString; safecall;
    procedure Set_Salutation(const prop: WideString); safecall;
    function Get_SalutationType: WdSalutationType; safecall;
    procedure Set_SalutationType(prop: WdSalutationType); safecall;
    function Get_RecipientReference: WideString; safecall;
    procedure Set_RecipientReference(const prop: WideString); safecall;
    function Get_MailingInstructions: WideString; safecall;
    procedure Set_MailingInstructions(const prop: WideString); safecall;
    function Get_AttentionLine: WideString; safecall;
    procedure Set_AttentionLine(const prop: WideString); safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const prop: WideString); safecall;
    function Get_EnclosureNumber: Integer; safecall;
    procedure Set_EnclosureNumber(prop: Integer); safecall;
    function Get_CCList: WideString; safecall;
    procedure Set_CCList(const prop: WideString); safecall;
    function Get_ReturnAddress: WideString; safecall;
    procedure Set_ReturnAddress(const prop: WideString); safecall;
    function Get_SenderName: WideString; safecall;
    procedure Set_SenderName(const prop: WideString); safecall;
    function Get_Closing: WideString; safecall;
    procedure Set_Closing(const prop: WideString); safecall;
    function Get_SenderCompany: WideString; safecall;
    procedure Set_SenderCompany(const prop: WideString); safecall;
    function Get_SenderJobTitle: WideString; safecall;
    procedure Set_SenderJobTitle(const prop: WideString); safecall;
    function Get_SenderInitials: WideString; safecall;
    procedure Set_SenderInitials(const prop: WideString); safecall;
    function Get_InfoBlock: WordBool; safecall;
    procedure Set_InfoBlock(prop: WordBool); safecall;
    function Get_RecipientCode: WideString; safecall;
    procedure Set_RecipientCode(const prop: WideString); safecall;
    function Get_RecipientGender: WdSalutationGender; safecall;
    procedure Set_RecipientGender(prop: WdSalutationGender); safecall;
    function Get_ReturnAddressShortForm: WideString; safecall;
    procedure Set_ReturnAddressShortForm(const prop: WideString); safecall;
    function Get_SenderCity: WideString; safecall;
    procedure Set_SenderCity(const prop: WideString); safecall;
    function Get_SenderCode: WideString; safecall;
    procedure Set_SenderCode(const prop: WideString); safecall;
    function Get_SenderGender: WdSalutationGender; safecall;
    procedure Set_SenderGender(prop: WdSalutationGender); safecall;
    function Get_SenderReference: WideString; safecall;
    procedure Set_SenderReference(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Duplicate: LetterContent read Get_Duplicate;
    property DateFormat: WideString read Get_DateFormat write Set_DateFormat;
    property IncludeHeaderFooter: WordBool read Get_IncludeHeaderFooter write Set_IncludeHeaderFooter;
    property PageDesign: WideString read Get_PageDesign write Set_PageDesign;
    property LetterStyle: WdLetterStyle read Get_LetterStyle write Set_LetterStyle;
    property Letterhead: WordBool read Get_Letterhead write Set_Letterhead;
    property LetterheadLocation: WdLetterheadLocation read Get_LetterheadLocation write Set_LetterheadLocation;
    property LetterheadSize: Single read Get_LetterheadSize write Set_LetterheadSize;
    property RecipientName: WideString read Get_RecipientName write Set_RecipientName;
    property RecipientAddress: WideString read Get_RecipientAddress write Set_RecipientAddress;
    property Salutation: WideString read Get_Salutation write Set_Salutation;
    property SalutationType: WdSalutationType read Get_SalutationType write Set_SalutationType;
    property RecipientReference: WideString read Get_RecipientReference write Set_RecipientReference;
    property MailingInstructions: WideString read Get_MailingInstructions write Set_MailingInstructions;
    property AttentionLine: WideString read Get_AttentionLine write Set_AttentionLine;
    property Subject: WideString read Get_Subject write Set_Subject;
    property EnclosureNumber: Integer read Get_EnclosureNumber write Set_EnclosureNumber;
    property CCList: WideString read Get_CCList write Set_CCList;
    property ReturnAddress: WideString read Get_ReturnAddress write Set_ReturnAddress;
    property SenderName: WideString read Get_SenderName write Set_SenderName;
    property Closing: WideString read Get_Closing write Set_Closing;
    property SenderCompany: WideString read Get_SenderCompany write Set_SenderCompany;
    property SenderJobTitle: WideString read Get_SenderJobTitle write Set_SenderJobTitle;
    property SenderInitials: WideString read Get_SenderInitials write Set_SenderInitials;
    property InfoBlock: WordBool read Get_InfoBlock write Set_InfoBlock;
    property RecipientCode: WideString read Get_RecipientCode write Set_RecipientCode;
    property RecipientGender: WdSalutationGender read Get_RecipientGender write Set_RecipientGender;
    property ReturnAddressShortForm: WideString read Get_ReturnAddressShortForm write Set_ReturnAddressShortForm;
    property SenderCity: WideString read Get_SenderCity write Set_SenderCity;
    property SenderCode: WideString read Get_SenderCode write Set_SenderCode;
    property SenderGender: WdSalutationGender read Get_SenderGender write Set_SenderGender;
    property SenderReference: WideString read Get_SenderReference write Set_SenderReference;
  end;

// *********************************************************************//
// DispIntf:  _LetterContentDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A1-0000-0000-C000-000000000046}
// *********************************************************************//
  _LetterContentDisp = dispinterface
    ['{000209A1-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Duplicate: LetterContent readonly dispid 10;
    property DateFormat: WideString dispid 101;
    property IncludeHeaderFooter: WordBool dispid 102;
    property PageDesign: WideString dispid 103;
    property LetterStyle: WdLetterStyle dispid 104;
    property Letterhead: WordBool dispid 105;
    property LetterheadLocation: WdLetterheadLocation dispid 106;
    property LetterheadSize: Single dispid 107;
    property RecipientName: WideString dispid 108;
    property RecipientAddress: WideString dispid 109;
    property Salutation: WideString dispid 110;
    property SalutationType: WdSalutationType dispid 111;
    property RecipientReference: WideString dispid 112;
    property MailingInstructions: WideString dispid 114;
    property AttentionLine: WideString dispid 115;
    property Subject: WideString dispid 116;
    property EnclosureNumber: Integer dispid 117;
    property CCList: WideString dispid 118;
    property ReturnAddress: WideString dispid 119;
    property SenderName: WideString dispid 120;
    property Closing: WideString dispid 121;
    property SenderCompany: WideString dispid 122;
    property SenderJobTitle: WideString dispid 123;
    property SenderInitials: WideString dispid 124;
    property InfoBlock: WordBool dispid 125;
    property RecipientCode: WideString dispid 126;
    property RecipientGender: WdSalutationGender dispid 127;
    property ReturnAddressShortForm: WideString dispid 128;
    property SenderCity: WideString dispid 129;
    property SenderCode: WideString dispid 130;
    property SenderGender: WdSalutationGender dispid 131;
    property SenderReference: WideString dispid 132;
  end;

// *********************************************************************//
// Interface: View
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A5-0000-0000-C000-000000000046}
// *********************************************************************//
  View = interface(IDispatch)
    ['{000209A5-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Type_: WdViewType; safecall;
    procedure Set_Type_(prop: WdViewType); safecall;
    function Get_FullScreen: WordBool; safecall;
    procedure Set_FullScreen(prop: WordBool); safecall;
    function Get_Draft: WordBool; safecall;
    procedure Set_Draft(prop: WordBool); safecall;
    function Get_ShowAll: WordBool; safecall;
    procedure Set_ShowAll(prop: WordBool); safecall;
    function Get_ShowFieldCodes: WordBool; safecall;
    procedure Set_ShowFieldCodes(prop: WordBool); safecall;
    function Get_MailMergeDataView: WordBool; safecall;
    procedure Set_MailMergeDataView(prop: WordBool); safecall;
    function Get_Magnifier: WordBool; safecall;
    procedure Set_Magnifier(prop: WordBool); safecall;
    function Get_ShowFirstLineOnly: WordBool; safecall;
    procedure Set_ShowFirstLineOnly(prop: WordBool); safecall;
    function Get_ShowFormat: WordBool; safecall;
    procedure Set_ShowFormat(prop: WordBool); safecall;
    function Get_Zoom: Zoom; safecall;
    function Get_ShowObjectAnchors: WordBool; safecall;
    procedure Set_ShowObjectAnchors(prop: WordBool); safecall;
    function Get_ShowTextBoundaries: WordBool; safecall;
    procedure Set_ShowTextBoundaries(prop: WordBool); safecall;
    function Get_ShowHighlight: WordBool; safecall;
    procedure Set_ShowHighlight(prop: WordBool); safecall;
    function Get_ShowDrawings: WordBool; safecall;
    procedure Set_ShowDrawings(prop: WordBool); safecall;
    function Get_ShowTabs: WordBool; safecall;
    procedure Set_ShowTabs(prop: WordBool); safecall;
    function Get_ShowSpaces: WordBool; safecall;
    procedure Set_ShowSpaces(prop: WordBool); safecall;
    function Get_ShowParagraphs: WordBool; safecall;
    procedure Set_ShowParagraphs(prop: WordBool); safecall;
    function Get_ShowHyphens: WordBool; safecall;
    procedure Set_ShowHyphens(prop: WordBool); safecall;
    function Get_ShowHiddenText: WordBool; safecall;
    procedure Set_ShowHiddenText(prop: WordBool); safecall;
    function Get_WrapToWindow: WordBool; safecall;
    procedure Set_WrapToWindow(prop: WordBool); safecall;
    function Get_ShowPicturePlaceHolders: WordBool; safecall;
    procedure Set_ShowPicturePlaceHolders(prop: WordBool); safecall;
    function Get_ShowBookmarks: WordBool; safecall;
    procedure Set_ShowBookmarks(prop: WordBool); safecall;
    function Get_FieldShading: WdFieldShading; safecall;
    procedure Set_FieldShading(prop: WdFieldShading); safecall;
    function Get_ShowAnimation: WordBool; safecall;
    procedure Set_ShowAnimation(prop: WordBool); safecall;
    function Get_TableGridlines: WordBool; safecall;
    procedure Set_TableGridlines(prop: WordBool); safecall;
    function Get_EnlargeFontsLessThan: Integer; safecall;
    procedure Set_EnlargeFontsLessThan(prop: Integer); safecall;
    function Get_ShowMainTextLayer: WordBool; safecall;
    procedure Set_ShowMainTextLayer(prop: WordBool); safecall;
    function Get_SeekView: WdSeekView; safecall;
    procedure Set_SeekView(prop: WdSeekView); safecall;
    function Get_SplitSpecial: WdSpecialPane; safecall;
    procedure Set_SplitSpecial(prop: WdSpecialPane); safecall;
    function Get_BrowseToWindow: Integer; safecall;
    procedure Set_BrowseToWindow(prop: Integer); safecall;
    function Get_ShowOptionalBreaks: WordBool; safecall;
    procedure Set_ShowOptionalBreaks(prop: WordBool); safecall;
    procedure CollapseOutline(var Range: OleVariant); safecall;
    procedure ExpandOutline(var Range: OleVariant); safecall;
    procedure ShowAllHeadings; safecall;
    procedure ShowHeading(Level: Integer); safecall;
    procedure PreviousHeaderFooter; safecall;
    procedure NextHeaderFooter; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Type_: WdViewType read Get_Type_ write Set_Type_;
    property FullScreen: WordBool read Get_FullScreen write Set_FullScreen;
    property Draft: WordBool read Get_Draft write Set_Draft;
    property ShowAll: WordBool read Get_ShowAll write Set_ShowAll;
    property ShowFieldCodes: WordBool read Get_ShowFieldCodes write Set_ShowFieldCodes;
    property MailMergeDataView: WordBool read Get_MailMergeDataView write Set_MailMergeDataView;
    property Magnifier: WordBool read Get_Magnifier write Set_Magnifier;
    property ShowFirstLineOnly: WordBool read Get_ShowFirstLineOnly write Set_ShowFirstLineOnly;
    property ShowFormat: WordBool read Get_ShowFormat write Set_ShowFormat;
    property Zoom: Zoom read Get_Zoom;
    property ShowObjectAnchors: WordBool read Get_ShowObjectAnchors write Set_ShowObjectAnchors;
    property ShowTextBoundaries: WordBool read Get_ShowTextBoundaries write Set_ShowTextBoundaries;
    property ShowHighlight: WordBool read Get_ShowHighlight write Set_ShowHighlight;
    property ShowDrawings: WordBool read Get_ShowDrawings write Set_ShowDrawings;
    property ShowTabs: WordBool read Get_ShowTabs write Set_ShowTabs;
    property ShowSpaces: WordBool read Get_ShowSpaces write Set_ShowSpaces;
    property ShowParagraphs: WordBool read Get_ShowParagraphs write Set_ShowParagraphs;
    property ShowHyphens: WordBool read Get_ShowHyphens write Set_ShowHyphens;
    property ShowHiddenText: WordBool read Get_ShowHiddenText write Set_ShowHiddenText;
    property WrapToWindow: WordBool read Get_WrapToWindow write Set_WrapToWindow;
    property ShowPicturePlaceHolders: WordBool read Get_ShowPicturePlaceHolders write Set_ShowPicturePlaceHolders;
    property ShowBookmarks: WordBool read Get_ShowBookmarks write Set_ShowBookmarks;
    property FieldShading: WdFieldShading read Get_FieldShading write Set_FieldShading;
    property ShowAnimation: WordBool read Get_ShowAnimation write Set_ShowAnimation;
    property TableGridlines: WordBool read Get_TableGridlines write Set_TableGridlines;
    property EnlargeFontsLessThan: Integer read Get_EnlargeFontsLessThan write Set_EnlargeFontsLessThan;
    property ShowMainTextLayer: WordBool read Get_ShowMainTextLayer write Set_ShowMainTextLayer;
    property SeekView: WdSeekView read Get_SeekView write Set_SeekView;
    property SplitSpecial: WdSpecialPane read Get_SplitSpecial write Set_SplitSpecial;
    property BrowseToWindow: Integer read Get_BrowseToWindow write Set_BrowseToWindow;
    property ShowOptionalBreaks: WordBool read Get_ShowOptionalBreaks write Set_ShowOptionalBreaks;
  end;

// *********************************************************************//
// DispIntf:  ViewDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A5-0000-0000-C000-000000000046}
// *********************************************************************//
  ViewDisp = dispinterface
    ['{000209A5-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Type_: WdViewType dispid 0;
    property FullScreen: WordBool dispid 1;
    property Draft: WordBool dispid 2;
    property ShowAll: WordBool dispid 3;
    property ShowFieldCodes: WordBool dispid 4;
    property MailMergeDataView: WordBool dispid 5;
    property Magnifier: WordBool dispid 7;
    property ShowFirstLineOnly: WordBool dispid 8;
    property ShowFormat: WordBool dispid 9;
    property Zoom: Zoom readonly dispid 10;
    property ShowObjectAnchors: WordBool dispid 11;
    property ShowTextBoundaries: WordBool dispid 12;
    property ShowHighlight: WordBool dispid 13;
    property ShowDrawings: WordBool dispid 14;
    property ShowTabs: WordBool dispid 15;
    property ShowSpaces: WordBool dispid 16;
    property ShowParagraphs: WordBool dispid 17;
    property ShowHyphens: WordBool dispid 18;
    property ShowHiddenText: WordBool dispid 19;
    property WrapToWindow: WordBool dispid 20;
    property ShowPicturePlaceHolders: WordBool dispid 21;
    property ShowBookmarks: WordBool dispid 22;
    property FieldShading: WdFieldShading dispid 23;
    property ShowAnimation: WordBool dispid 24;
    property TableGridlines: WordBool dispid 25;
    property EnlargeFontsLessThan: Integer dispid 26;
    property ShowMainTextLayer: WordBool dispid 27;
    property SeekView: WdSeekView dispid 28;
    property SplitSpecial: WdSpecialPane dispid 29;
    property BrowseToWindow: Integer dispid 30;
    property ShowOptionalBreaks: WordBool dispid 31;
    procedure CollapseOutline(var Range: OleVariant); dispid 101;
    procedure ExpandOutline(var Range: OleVariant); dispid 102;
    procedure ShowAllHeadings; dispid 103;
    procedure ShowHeading(Level: Integer); dispid 104;
    procedure PreviousHeaderFooter; dispid 105;
    procedure NextHeaderFooter; dispid 106;
  end;

// *********************************************************************//
// Interface: Zoom
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A6-0000-0000-C000-000000000046}
// *********************************************************************//
  Zoom = interface(IDispatch)
    ['{000209A6-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Percentage: Integer; safecall;
    procedure Set_Percentage(prop: Integer); safecall;
    function Get_PageFit: WdPageFit; safecall;
    procedure Set_PageFit(prop: WdPageFit); safecall;
    function Get_PageRows: Integer; safecall;
    procedure Set_PageRows(prop: Integer); safecall;
    function Get_PageColumns: Integer; safecall;
    procedure Set_PageColumns(prop: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Percentage: Integer read Get_Percentage write Set_Percentage;
    property PageFit: WdPageFit read Get_PageFit write Set_PageFit;
    property PageRows: Integer read Get_PageRows write Set_PageRows;
    property PageColumns: Integer read Get_PageColumns write Set_PageColumns;
  end;

// *********************************************************************//
// DispIntf:  ZoomDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A6-0000-0000-C000-000000000046}
// *********************************************************************//
  ZoomDisp = dispinterface
    ['{000209A6-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Percentage: Integer dispid 0;
    property PageFit: WdPageFit dispid 1;
    property PageRows: Integer dispid 2;
    property PageColumns: Integer dispid 3;
  end;

// *********************************************************************//
// Interface: Zooms
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A7-0000-0000-C000-000000000046}
// *********************************************************************//
  Zooms = interface(IDispatch)
    ['{000209A7-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Item(Index: WdViewType): Zoom; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  ZoomsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A7-0000-0000-C000-000000000046}
// *********************************************************************//
  ZoomsDisp = dispinterface
    ['{000209A7-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function Item(Index: WdViewType): Zoom; dispid 0;
  end;

// *********************************************************************//
// Interface: InlineShape
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A8-0000-0000-C000-000000000046}
// *********************************************************************//
  InlineShape = interface(IDispatch)
    ['{000209A8-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Range: Range; safecall;
    function Get_LinkFormat: LinkFormat; safecall;
    function Get_Field: Field; safecall;
    function Get_OLEFormat: OLEFormat; safecall;
    function Get_Type_: WdInlineShapeType; safecall;
    function Get_Hyperlink: Hyperlink; safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_ScaleHeight: Single; safecall;
    procedure Set_ScaleHeight(prop: Single); safecall;
    function Get_ScaleWidth: Single; safecall;
    procedure Set_ScaleWidth(prop: Single); safecall;
    function Get_LockAspectRatio: MsoTriState; safecall;
    procedure Set_LockAspectRatio(prop: MsoTriState); safecall;
    function Get_Line: LineFormat; safecall;
    function Get_Fill: FillFormat; safecall;
    function Get_PictureFormat: PictureFormat; safecall;
    procedure Set_PictureFormat(const prop: PictureFormat); safecall;
    procedure Activate; safecall;
    procedure Reset; safecall;
    procedure Delete; safecall;
    procedure Select; safecall;
    function ConvertToShape: Shape; safecall;
    function Get_HorizontalLineFormat: HorizontalLineFormat; safecall;
    function Get_Script: Script; safecall;
    function Get_OWSAnchor: Integer; safecall;
    function Get_TextEffect: TextEffectFormat; safecall;
    procedure Set_TextEffect(const prop: TextEffectFormat); safecall;
    function Get_AlternativeText: WideString; safecall;
    procedure Set_AlternativeText(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Range: Range read Get_Range;
    property LinkFormat: LinkFormat read Get_LinkFormat;
    property Field: Field read Get_Field;
    property OLEFormat: OLEFormat read Get_OLEFormat;
    property Type_: WdInlineShapeType read Get_Type_;
    property Hyperlink: Hyperlink read Get_Hyperlink;
    property Height: Single read Get_Height write Set_Height;
    property Width: Single read Get_Width write Set_Width;
    property ScaleHeight: Single read Get_ScaleHeight write Set_ScaleHeight;
    property ScaleWidth: Single read Get_ScaleWidth write Set_ScaleWidth;
    property LockAspectRatio: MsoTriState read Get_LockAspectRatio write Set_LockAspectRatio;
    property Line: LineFormat read Get_Line;
    property Fill: FillFormat read Get_Fill;
    property PictureFormat: PictureFormat read Get_PictureFormat write Set_PictureFormat;
    property HorizontalLineFormat: HorizontalLineFormat read Get_HorizontalLineFormat;
    property Script: Script read Get_Script;
    property OWSAnchor: Integer read Get_OWSAnchor;
    property TextEffect: TextEffectFormat read Get_TextEffect write Set_TextEffect;
    property AlternativeText: WideString read Get_AlternativeText write Set_AlternativeText;
  end;

// *********************************************************************//
// DispIntf:  InlineShapeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A8-0000-0000-C000-000000000046}
// *********************************************************************//
  InlineShapeDisp = dispinterface
    ['{000209A8-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Borders: Borders dispid 1100;
    property Range: Range readonly dispid 2;
    property LinkFormat: LinkFormat readonly dispid 3;
    property Field: Field readonly dispid 4;
    property OLEFormat: OLEFormat readonly dispid 5;
    property Type_: WdInlineShapeType readonly dispid 6;
    property Hyperlink: Hyperlink readonly dispid 7;
    property Height: Single dispid 8;
    property Width: Single dispid 9;
    property ScaleHeight: Single dispid 10;
    property ScaleWidth: Single dispid 11;
    property LockAspectRatio: MsoTriState dispid 12;
    property Line: LineFormat readonly dispid 112;
    property Fill: FillFormat readonly dispid 107;
    property PictureFormat: PictureFormat dispid 118;
    procedure Activate; dispid 100;
    procedure Reset; dispid 101;
    procedure Delete; dispid 102;
    procedure Select; dispid 65535;
    function ConvertToShape: Shape; dispid 104;
    property HorizontalLineFormat: HorizontalLineFormat readonly dispid 119;
    property Script: Script readonly dispid 122;
    property OWSAnchor: Integer readonly dispid 130;
    property TextEffect: TextEffectFormat dispid 120;
    property AlternativeText: WideString dispid 131;
  end;

// *********************************************************************//
// Interface: InlineShapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A9-0000-0000-C000-000000000046}
// *********************************************************************//
  InlineShapes = interface(IDispatch)
    ['{000209A9-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(Index: Integer): InlineShape; safecall;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Range: OleVariant): InlineShape; safecall;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Range: OleVariant): InlineShape; safecall;
    function AddOLEControl(var ClassType: OleVariant; var Range: OleVariant): InlineShape; safecall;
    function New(const Range: Range): InlineShape; safecall;
    function AddHorizontalLine(const FileName: WideString; var Range: OleVariant): InlineShape; safecall;
    function AddHorizontalLineStandard(var Range: OleVariant): InlineShape; safecall;
    function AddPictureBullet(const FileName: WideString; var Range: OleVariant): InlineShape; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  InlineShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A9-0000-0000-C000-000000000046}
// *********************************************************************//
  InlineShapesDisp = dispinterface
    ['{000209A9-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(Index: Integer): InlineShape; dispid 0;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Range: OleVariant): InlineShape; dispid 100;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Range: OleVariant): InlineShape; dispid 24;
    function AddOLEControl(var ClassType: OleVariant; var Range: OleVariant): InlineShape; dispid 102;
    function New(const Range: Range): InlineShape; dispid 200;
    function AddHorizontalLine(const FileName: WideString; var Range: OleVariant): InlineShape; dispid 104;
    function AddHorizontalLineStandard(var Range: OleVariant): InlineShape; dispid 105;
    function AddPictureBullet(const FileName: WideString; var Range: OleVariant): InlineShape; dispid 106;
  end;

// *********************************************************************//
// Interface: SpellingSuggestions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AA-0000-0000-C000-000000000046}
// *********************************************************************//
  SpellingSuggestions = interface(IDispatch)
    ['{000209AA-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_SpellingErrorType: WdSpellingErrorType; safecall;
    function Item(Index: Integer): SpellingSuggestion; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property SpellingErrorType: WdSpellingErrorType read Get_SpellingErrorType;
  end;

// *********************************************************************//
// DispIntf:  SpellingSuggestionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AA-0000-0000-C000-000000000046}
// *********************************************************************//
  SpellingSuggestionsDisp = dispinterface
    ['{000209AA-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property SpellingErrorType: WdSpellingErrorType readonly dispid 2;
    function Item(Index: Integer): SpellingSuggestion; dispid 0;
  end;

// *********************************************************************//
// Interface: SpellingSuggestion
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AB-0000-0000-C000-000000000046}
// *********************************************************************//
  SpellingSuggestion = interface(IDispatch)
    ['{000209AB-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  SpellingSuggestionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AB-0000-0000-C000-000000000046}
// *********************************************************************//
  SpellingSuggestionDisp = dispinterface
    ['{000209AB-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
  end;

// *********************************************************************//
// Interface: Dictionaries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AC-0000-0000-C000-000000000046}
// *********************************************************************//
  Dictionaries = interface(IDispatch)
    ['{000209AC-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Maximum: Integer; safecall;
    function Get_ActiveCustomDictionary: Dictionary; safecall;
    procedure Set_ActiveCustomDictionary(const prop: Dictionary); safecall;
    function Item(var Index: OleVariant): Dictionary; safecall;
    function Add(const FileName: WideString): Dictionary; safecall;
    procedure ClearAll; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Maximum: Integer read Get_Maximum;
    property ActiveCustomDictionary: Dictionary read Get_ActiveCustomDictionary write Set_ActiveCustomDictionary;
  end;

// *********************************************************************//
// DispIntf:  DictionariesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AC-0000-0000-C000-000000000046}
// *********************************************************************//
  DictionariesDisp = dispinterface
    ['{000209AC-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Maximum: Integer readonly dispid 2;
    property ActiveCustomDictionary: Dictionary dispid 3;
    function Item(var Index: OleVariant): Dictionary; dispid 0;
    function Add(const FileName: WideString): Dictionary; dispid 101;
    procedure ClearAll; dispid 102;
  end;

// *********************************************************************//
// Interface: HangulHanjaConversionDictionaries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E0-0000-0000-C000-000000000046}
// *********************************************************************//
  HangulHanjaConversionDictionaries = interface(IDispatch)
    ['{000209E0-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Maximum: Integer; safecall;
    function Get_ActiveCustomDictionary: Dictionary; safecall;
    procedure Set_ActiveCustomDictionary(const prop: Dictionary); safecall;
    function Get_BuiltinDictionary: Dictionary; safecall;
    function Item(var Index: OleVariant): Dictionary; safecall;
    function Add(const FileName: WideString): Dictionary; safecall;
    procedure ClearAll; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Maximum: Integer read Get_Maximum;
    property ActiveCustomDictionary: Dictionary read Get_ActiveCustomDictionary write Set_ActiveCustomDictionary;
    property BuiltinDictionary: Dictionary read Get_BuiltinDictionary;
  end;

// *********************************************************************//
// DispIntf:  HangulHanjaConversionDictionariesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E0-0000-0000-C000-000000000046}
// *********************************************************************//
  HangulHanjaConversionDictionariesDisp = dispinterface
    ['{000209E0-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Maximum: Integer readonly dispid 2;
    property ActiveCustomDictionary: Dictionary dispid 3;
    property BuiltinDictionary: Dictionary readonly dispid 4;
    function Item(var Index: OleVariant): Dictionary; dispid 0;
    function Add(const FileName: WideString): Dictionary; dispid 101;
    procedure ClearAll; dispid 102;
  end;

// *********************************************************************//
// Interface: Dictionary
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AD-0000-0000-C000-000000000046}
// *********************************************************************//
  Dictionary = interface(IDispatch)
    ['{000209AD-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Path: WideString; safecall;
    function Get_LanguageID: WdLanguageID; safecall;
    procedure Set_LanguageID(prop: WdLanguageID); safecall;
    function Get_ReadOnly: WordBool; safecall;
    function Get_Type_: WdDictionaryType; safecall;
    function Get_LanguageSpecific: WordBool; safecall;
    procedure Set_LanguageSpecific(prop: WordBool); safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Path: WideString read Get_Path;
    property LanguageID: WdLanguageID read Get_LanguageID write Set_LanguageID;
    property ReadOnly: WordBool read Get_ReadOnly;
    property Type_: WdDictionaryType read Get_Type_;
    property LanguageSpecific: WordBool read Get_LanguageSpecific write Set_LanguageSpecific;
  end;

// *********************************************************************//
// DispIntf:  DictionaryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AD-0000-0000-C000-000000000046}
// *********************************************************************//
  DictionaryDisp = dispinterface
    ['{000209AD-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Path: WideString readonly dispid 1;
    property LanguageID: WdLanguageID dispid 2;
    property ReadOnly: WordBool readonly dispid 3;
    property Type_: WdDictionaryType readonly dispid 4;
    property LanguageSpecific: WordBool dispid 5;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: ReadabilityStatistics
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AE-0000-0000-C000-000000000046}
// *********************************************************************//
  ReadabilityStatistics = interface(IDispatch)
    ['{000209AE-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): ReadabilityStatistic; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ReadabilityStatisticsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AE-0000-0000-C000-000000000046}
// *********************************************************************//
  ReadabilityStatisticsDisp = dispinterface
    ['{000209AE-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): ReadabilityStatistic; dispid 0;
  end;

// *********************************************************************//
// Interface: ReadabilityStatistic
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AF-0000-0000-C000-000000000046}
// *********************************************************************//
  ReadabilityStatistic = interface(IDispatch)
    ['{000209AF-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Value: Single; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Name: WideString read Get_Name;
    property Value: Single read Get_Value;
  end;

// *********************************************************************//
// DispIntf:  ReadabilityStatisticDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AF-0000-0000-C000-000000000046}
// *********************************************************************//
  ReadabilityStatisticDisp = dispinterface
    ['{000209AF-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Name: WideString readonly dispid 0;
    property Value: Single readonly dispid 1;
  end;

// *********************************************************************//
// Interface: Versions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B3-0000-0000-C000-000000000046}
// *********************************************************************//
  Versions = interface(IDispatch)
    ['{000209B3-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_AutoVersion: WdAutoVersions; safecall;
    procedure Set_AutoVersion(prop: WdAutoVersions); safecall;
    function Item(Index: Integer): Version; safecall;
    procedure Save(var Comment: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property AutoVersion: WdAutoVersions read Get_AutoVersion write Set_AutoVersion;
  end;

// *********************************************************************//
// DispIntf:  VersionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B3-0000-0000-C000-000000000046}
// *********************************************************************//
  VersionsDisp = dispinterface
    ['{000209B3-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property AutoVersion: WdAutoVersions dispid 3;
    function Item(Index: Integer): Version; dispid 0;
    procedure Save(var Comment: OleVariant); dispid 11;
  end;

// *********************************************************************//
// Interface: Version
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B4-0000-0000-C000-000000000046}
// *********************************************************************//
  Version = interface(IDispatch)
    ['{000209B4-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_SavedBy: WideString; safecall;
    function Get_Comment: WideString; safecall;
    function Get_Date: TDateTime; safecall;
    function Get_Index: Integer; safecall;
    procedure OpenOld; safecall;
    procedure Delete; safecall;
    function Open: Document; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property SavedBy: WideString read Get_SavedBy;
    property Comment: WideString read Get_Comment;
    property Date: TDateTime read Get_Date;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  VersionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B4-0000-0000-C000-000000000046}
// *********************************************************************//
  VersionDisp = dispinterface
    ['{000209B4-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property SavedBy: WideString readonly dispid 1003;
    property Comment: WideString readonly dispid 1004;
    property Date: TDateTime readonly dispid 1005;
    property Index: Integer readonly dispid 2;
    procedure OpenOld; dispid 101;
    procedure Delete; dispid 102;
    function Open: Document; dispid 103;
  end;

// *********************************************************************//
// Interface: Options
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B7-0000-0000-C000-000000000046}
// *********************************************************************//
  Options = interface(IDispatch)
    ['{000209B7-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_AllowAccentedUppercase: WordBool; safecall;
    procedure Set_AllowAccentedUppercase(prop: WordBool); safecall;
    function Get_WPHelp: WordBool; safecall;
    procedure Set_WPHelp(prop: WordBool); safecall;
    function Get_WPDocNavKeys: WordBool; safecall;
    procedure Set_WPDocNavKeys(prop: WordBool); safecall;
    function Get_Pagination: WordBool; safecall;
    procedure Set_Pagination(prop: WordBool); safecall;
    function Get_BlueScreen: WordBool; safecall;
    procedure Set_BlueScreen(prop: WordBool); safecall;
    function Get_EnableSound: WordBool; safecall;
    procedure Set_EnableSound(prop: WordBool); safecall;
    function Get_ConfirmConversions: WordBool; safecall;
    procedure Set_ConfirmConversions(prop: WordBool); safecall;
    function Get_UpdateLinksAtOpen: WordBool; safecall;
    procedure Set_UpdateLinksAtOpen(prop: WordBool); safecall;
    function Get_SendMailAttach: WordBool; safecall;
    procedure Set_SendMailAttach(prop: WordBool); safecall;
    function Get_MeasurementUnit: WdMeasurementUnits; safecall;
    procedure Set_MeasurementUnit(prop: WdMeasurementUnits); safecall;
    function Get_ButtonFieldClicks: Integer; safecall;
    procedure Set_ButtonFieldClicks(prop: Integer); safecall;
    function Get_ShortMenuNames: WordBool; safecall;
    procedure Set_ShortMenuNames(prop: WordBool); safecall;
    function Get_RTFInClipboard: WordBool; safecall;
    procedure Set_RTFInClipboard(prop: WordBool); safecall;
    function Get_UpdateFieldsAtPrint: WordBool; safecall;
    procedure Set_UpdateFieldsAtPrint(prop: WordBool); safecall;
    function Get_PrintProperties: WordBool; safecall;
    procedure Set_PrintProperties(prop: WordBool); safecall;
    function Get_PrintFieldCodes: WordBool; safecall;
    procedure Set_PrintFieldCodes(prop: WordBool); safecall;
    function Get_PrintComments: WordBool; safecall;
    procedure Set_PrintComments(prop: WordBool); safecall;
    function Get_PrintHiddenText: WordBool; safecall;
    procedure Set_PrintHiddenText(prop: WordBool); safecall;
    function Get_EnvelopeFeederInstalled: WordBool; safecall;
    function Get_UpdateLinksAtPrint: WordBool; safecall;
    procedure Set_UpdateLinksAtPrint(prop: WordBool); safecall;
    function Get_PrintBackground: WordBool; safecall;
    procedure Set_PrintBackground(prop: WordBool); safecall;
    function Get_PrintDrawingObjects: WordBool; safecall;
    procedure Set_PrintDrawingObjects(prop: WordBool); safecall;
    function Get_DefaultTray: WideString; safecall;
    procedure Set_DefaultTray(const prop: WideString); safecall;
    function Get_DefaultTrayID: Integer; safecall;
    procedure Set_DefaultTrayID(prop: Integer); safecall;
    function Get_CreateBackup: WordBool; safecall;
    procedure Set_CreateBackup(prop: WordBool); safecall;
    function Get_AllowFastSave: WordBool; safecall;
    procedure Set_AllowFastSave(prop: WordBool); safecall;
    function Get_SavePropertiesPrompt: WordBool; safecall;
    procedure Set_SavePropertiesPrompt(prop: WordBool); safecall;
    function Get_SaveNormalPrompt: WordBool; safecall;
    procedure Set_SaveNormalPrompt(prop: WordBool); safecall;
    function Get_SaveInterval: Integer; safecall;
    procedure Set_SaveInterval(prop: Integer); safecall;
    function Get_BackgroundSave: WordBool; safecall;
    procedure Set_BackgroundSave(prop: WordBool); safecall;
    function Get_InsertedTextMark: WdInsertedTextMark; safecall;
    procedure Set_InsertedTextMark(prop: WdInsertedTextMark); safecall;
    function Get_DeletedTextMark: WdDeletedTextMark; safecall;
    procedure Set_DeletedTextMark(prop: WdDeletedTextMark); safecall;
    function Get_RevisedLinesMark: WdRevisedLinesMark; safecall;
    procedure Set_RevisedLinesMark(prop: WdRevisedLinesMark); safecall;
    function Get_InsertedTextColor: WdColorIndex; safecall;
    procedure Set_InsertedTextColor(prop: WdColorIndex); safecall;
    function Get_DeletedTextColor: WdColorIndex; safecall;
    procedure Set_DeletedTextColor(prop: WdColorIndex); safecall;
    function Get_RevisedLinesColor: WdColorIndex; safecall;
    procedure Set_RevisedLinesColor(prop: WdColorIndex); safecall;
    function Get_DefaultFilePath(Path: WdDefaultFilePath): WideString; safecall;
    procedure Set_DefaultFilePath(Path: WdDefaultFilePath; const prop: WideString); safecall;
    function Get_Overtype: WordBool; safecall;
    procedure Set_Overtype(prop: WordBool); safecall;
    function Get_ReplaceSelection: WordBool; safecall;
    procedure Set_ReplaceSelection(prop: WordBool); safecall;
    function Get_AllowDragAndDrop: WordBool; safecall;
    procedure Set_AllowDragAndDrop(prop: WordBool); safecall;
    function Get_AutoWordSelection: WordBool; safecall;
    procedure Set_AutoWordSelection(prop: WordBool); safecall;
    function Get_INSKeyForPaste: WordBool; safecall;
    procedure Set_INSKeyForPaste(prop: WordBool); safecall;
    function Get_SmartCutPaste: WordBool; safecall;
    procedure Set_SmartCutPaste(prop: WordBool); safecall;
    function Get_TabIndentKey: WordBool; safecall;
    procedure Set_TabIndentKey(prop: WordBool); safecall;
    function Get_PictureEditor: WideString; safecall;
    procedure Set_PictureEditor(const prop: WideString); safecall;
    function Get_AnimateScreenMovements: WordBool; safecall;
    procedure Set_AnimateScreenMovements(prop: WordBool); safecall;
    function Get_VirusProtection: WordBool; safecall;
    procedure Set_VirusProtection(prop: WordBool); safecall;
    function Get_RevisedPropertiesMark: WdRevisedPropertiesMark; safecall;
    procedure Set_RevisedPropertiesMark(prop: WdRevisedPropertiesMark); safecall;
    function Get_RevisedPropertiesColor: WdColorIndex; safecall;
    procedure Set_RevisedPropertiesColor(prop: WdColorIndex); safecall;
    function Get_SnapToGrid: WordBool; safecall;
    procedure Set_SnapToGrid(prop: WordBool); safecall;
    function Get_SnapToShapes: WordBool; safecall;
    procedure Set_SnapToShapes(prop: WordBool); safecall;
    function Get_GridDistanceHorizontal: Single; safecall;
    procedure Set_GridDistanceHorizontal(prop: Single); safecall;
    function Get_GridDistanceVertical: Single; safecall;
    procedure Set_GridDistanceVertical(prop: Single); safecall;
    function Get_GridOriginHorizontal: Single; safecall;
    procedure Set_GridOriginHorizontal(prop: Single); safecall;
    function Get_GridOriginVertical: Single; safecall;
    procedure Set_GridOriginVertical(prop: Single); safecall;
    function Get_InlineConversion: WordBool; safecall;
    procedure Set_InlineConversion(prop: WordBool); safecall;
    function Get_IMEAutomaticControl: WordBool; safecall;
    procedure Set_IMEAutomaticControl(prop: WordBool); safecall;
    function Get_AutoFormatApplyHeadings: WordBool; safecall;
    procedure Set_AutoFormatApplyHeadings(prop: WordBool); safecall;
    function Get_AutoFormatApplyLists: WordBool; safecall;
    procedure Set_AutoFormatApplyLists(prop: WordBool); safecall;
    function Get_AutoFormatApplyBulletedLists: WordBool; safecall;
    procedure Set_AutoFormatApplyBulletedLists(prop: WordBool); safecall;
    function Get_AutoFormatApplyOtherParas: WordBool; safecall;
    procedure Set_AutoFormatApplyOtherParas(prop: WordBool); safecall;
    function Get_AutoFormatReplaceQuotes: WordBool; safecall;
    procedure Set_AutoFormatReplaceQuotes(prop: WordBool); safecall;
    function Get_AutoFormatReplaceSymbols: WordBool; safecall;
    procedure Set_AutoFormatReplaceSymbols(prop: WordBool); safecall;
    function Get_AutoFormatReplaceOrdinals: WordBool; safecall;
    procedure Set_AutoFormatReplaceOrdinals(prop: WordBool); safecall;
    function Get_AutoFormatReplaceFractions: WordBool; safecall;
    procedure Set_AutoFormatReplaceFractions(prop: WordBool); safecall;
    function Get_AutoFormatReplacePlainTextEmphasis: WordBool; safecall;
    procedure Set_AutoFormatReplacePlainTextEmphasis(prop: WordBool); safecall;
    function Get_AutoFormatPreserveStyles: WordBool; safecall;
    procedure Set_AutoFormatPreserveStyles(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyHeadings: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyHeadings(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyBorders: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyBorders(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyBulletedLists: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyBulletedLists(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyNumberedLists: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyNumberedLists(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceQuotes: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceQuotes(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceSymbols: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceSymbols(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceOrdinals: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceOrdinals(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceFractions: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceFractions(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplacePlainTextEmphasis: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplacePlainTextEmphasis(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeFormatListItemBeginning: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeFormatListItemBeginning(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeDefineStyles: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeDefineStyles(prop: WordBool); safecall;
    function Get_AutoFormatPlainTextWordMail: WordBool; safecall;
    procedure Set_AutoFormatPlainTextWordMail(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceHyperlinks: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceHyperlinks(prop: WordBool); safecall;
    function Get_AutoFormatReplaceHyperlinks: WordBool; safecall;
    procedure Set_AutoFormatReplaceHyperlinks(prop: WordBool); safecall;
    function Get_DefaultHighlightColorIndex: WdColorIndex; safecall;
    procedure Set_DefaultHighlightColorIndex(prop: WdColorIndex); safecall;
    function Get_DefaultBorderLineStyle: WdLineStyle; safecall;
    procedure Set_DefaultBorderLineStyle(prop: WdLineStyle); safecall;
    function Get_CheckSpellingAsYouType: WordBool; safecall;
    procedure Set_CheckSpellingAsYouType(prop: WordBool); safecall;
    function Get_CheckGrammarAsYouType: WordBool; safecall;
    procedure Set_CheckGrammarAsYouType(prop: WordBool); safecall;
    function Get_IgnoreInternetAndFileAddresses: WordBool; safecall;
    procedure Set_IgnoreInternetAndFileAddresses(prop: WordBool); safecall;
    function Get_ShowReadabilityStatistics: WordBool; safecall;
    procedure Set_ShowReadabilityStatistics(prop: WordBool); safecall;
    function Get_IgnoreUppercase: WordBool; safecall;
    procedure Set_IgnoreUppercase(prop: WordBool); safecall;
    function Get_IgnoreMixedDigits: WordBool; safecall;
    procedure Set_IgnoreMixedDigits(prop: WordBool); safecall;
    function Get_SuggestFromMainDictionaryOnly: WordBool; safecall;
    procedure Set_SuggestFromMainDictionaryOnly(prop: WordBool); safecall;
    function Get_SuggestSpellingCorrections: WordBool; safecall;
    procedure Set_SuggestSpellingCorrections(prop: WordBool); safecall;
    function Get_DefaultBorderLineWidth: WdLineWidth; safecall;
    procedure Set_DefaultBorderLineWidth(prop: WdLineWidth); safecall;
    function Get_CheckGrammarWithSpelling: WordBool; safecall;
    procedure Set_CheckGrammarWithSpelling(prop: WordBool); safecall;
    function Get_DefaultOpenFormat: WdOpenFormat; safecall;
    procedure Set_DefaultOpenFormat(prop: WdOpenFormat); safecall;
    function Get_PrintDraft: WordBool; safecall;
    procedure Set_PrintDraft(prop: WordBool); safecall;
    function Get_PrintReverse: WordBool; safecall;
    procedure Set_PrintReverse(prop: WordBool); safecall;
    function Get_MapPaperSize: WordBool; safecall;
    procedure Set_MapPaperSize(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyTables: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyTables(prop: WordBool); safecall;
    function Get_AutoFormatApplyFirstIndents: WordBool; safecall;
    procedure Set_AutoFormatApplyFirstIndents(prop: WordBool); safecall;
    function Get_AutoFormatMatchParentheses: WordBool; safecall;
    procedure Set_AutoFormatMatchParentheses(prop: WordBool); safecall;
    function Get_AutoFormatReplaceFarEastDashes: WordBool; safecall;
    procedure Set_AutoFormatReplaceFarEastDashes(prop: WordBool); safecall;
    function Get_AutoFormatDeleteAutoSpaces: WordBool; safecall;
    procedure Set_AutoFormatDeleteAutoSpaces(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyFirstIndents: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyFirstIndents(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyDates: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyDates(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyClosings: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyClosings(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeMatchParentheses: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeMatchParentheses(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceFarEastDashes: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceFarEastDashes(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeDeleteAutoSpaces: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeDeleteAutoSpaces(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeInsertClosings: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeInsertClosings(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeAutoLetterWizard: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeAutoLetterWizard(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeInsertOvers: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeInsertOvers(prop: WordBool); safecall;
    function Get_DisplayGridLines: WordBool; safecall;
    procedure Set_DisplayGridLines(prop: WordBool); safecall;
    function Get_MatchFuzzyCase: WordBool; safecall;
    procedure Set_MatchFuzzyCase(prop: WordBool); safecall;
    function Get_MatchFuzzyByte: WordBool; safecall;
    procedure Set_MatchFuzzyByte(prop: WordBool); safecall;
    function Get_MatchFuzzyHiragana: WordBool; safecall;
    procedure Set_MatchFuzzyHiragana(prop: WordBool); safecall;
    function Get_MatchFuzzySmallKana: WordBool; safecall;
    procedure Set_MatchFuzzySmallKana(prop: WordBool); safecall;
    function Get_MatchFuzzyDash: WordBool; safecall;
    procedure Set_MatchFuzzyDash(prop: WordBool); safecall;
    function Get_MatchFuzzyIterationMark: WordBool; safecall;
    procedure Set_MatchFuzzyIterationMark(prop: WordBool); safecall;
    function Get_MatchFuzzyKanji: WordBool; safecall;
    procedure Set_MatchFuzzyKanji(prop: WordBool); safecall;
    function Get_MatchFuzzyOldKana: WordBool; safecall;
    procedure Set_MatchFuzzyOldKana(prop: WordBool); safecall;
    function Get_MatchFuzzyProlongedSoundMark: WordBool; safecall;
    procedure Set_MatchFuzzyProlongedSoundMark(prop: WordBool); safecall;
    function Get_MatchFuzzyDZ: WordBool; safecall;
    procedure Set_MatchFuzzyDZ(prop: WordBool); safecall;
    function Get_MatchFuzzyBV: WordBool; safecall;
    procedure Set_MatchFuzzyBV(prop: WordBool); safecall;
    function Get_MatchFuzzyTC: WordBool; safecall;
    procedure Set_MatchFuzzyTC(prop: WordBool); safecall;
    function Get_MatchFuzzyHF: WordBool; safecall;
    procedure Set_MatchFuzzyHF(prop: WordBool); safecall;
    function Get_MatchFuzzyZJ: WordBool; safecall;
    procedure Set_MatchFuzzyZJ(prop: WordBool); safecall;
    function Get_MatchFuzzyAY: WordBool; safecall;
    procedure Set_MatchFuzzyAY(prop: WordBool); safecall;
    function Get_MatchFuzzyKiKu: WordBool; safecall;
    procedure Set_MatchFuzzyKiKu(prop: WordBool); safecall;
    function Get_MatchFuzzyPunctuation: WordBool; safecall;
    procedure Set_MatchFuzzyPunctuation(prop: WordBool); safecall;
    function Get_MatchFuzzySpace: WordBool; safecall;
    procedure Set_MatchFuzzySpace(prop: WordBool); safecall;
    function Get_ApplyFarEastFontsToAscii: WordBool; safecall;
    procedure Set_ApplyFarEastFontsToAscii(prop: WordBool); safecall;
    function Get_ConvertHighAnsiToFarEast: WordBool; safecall;
    procedure Set_ConvertHighAnsiToFarEast(prop: WordBool); safecall;
    function Get_PrintOddPagesInAscendingOrder: WordBool; safecall;
    procedure Set_PrintOddPagesInAscendingOrder(prop: WordBool); safecall;
    function Get_PrintEvenPagesInAscendingOrder: WordBool; safecall;
    procedure Set_PrintEvenPagesInAscendingOrder(prop: WordBool); safecall;
    function Get_DefaultBorderColorIndex: WdColorIndex; safecall;
    procedure Set_DefaultBorderColorIndex(prop: WdColorIndex); safecall;
    function Get_EnableMisusedWordsDictionary: WordBool; safecall;
    procedure Set_EnableMisusedWordsDictionary(prop: WordBool); safecall;
    function Get_AllowCombinedAuxiliaryForms: WordBool; safecall;
    procedure Set_AllowCombinedAuxiliaryForms(prop: WordBool); safecall;
    function Get_HangulHanjaFastConversion: WordBool; safecall;
    procedure Set_HangulHanjaFastConversion(prop: WordBool); safecall;
    function Get_CheckHangulEndings: WordBool; safecall;
    procedure Set_CheckHangulEndings(prop: WordBool); safecall;
    function Get_EnableHangulHanjaRecentOrdering: WordBool; safecall;
    procedure Set_EnableHangulHanjaRecentOrdering(prop: WordBool); safecall;
    function Get_MultipleWordConversionsMode: WdMultipleWordConversionsMode; safecall;
    procedure Set_MultipleWordConversionsMode(prop: WdMultipleWordConversionsMode); safecall;
    procedure SetWPHelpOptions(var CommandKeyHelp: OleVariant; var DocNavigationKeys: OleVariant; 
                               var MouseSimulation: OleVariant; var DemoGuidance: OleVariant; 
                               var DemoSpeed: OleVariant; var HelpType: OleVariant); safecall;
    function Get_DefaultBorderColor: WdColor; safecall;
    procedure Set_DefaultBorderColor(prop: WdColor); safecall;
    function Get_AllowPixelUnits: WordBool; safecall;
    procedure Set_AllowPixelUnits(prop: WordBool); safecall;
    function Get_UseCharacterUnit: WordBool; safecall;
    procedure Set_UseCharacterUnit(prop: WordBool); safecall;
    function Get_AllowCompoundNounProcessing: WordBool; safecall;
    procedure Set_AllowCompoundNounProcessing(prop: WordBool); safecall;
    function Get_AutoKeyboardSwitching: WordBool; safecall;
    procedure Set_AutoKeyboardSwitching(prop: WordBool); safecall;
    function Get_DocumentViewDirection: WdDocumentViewDirection; safecall;
    procedure Set_DocumentViewDirection(prop: WdDocumentViewDirection); safecall;
    function Get_ArabicNumeral: WdArabicNumeral; safecall;
    procedure Set_ArabicNumeral(prop: WdArabicNumeral); safecall;
    function Get_MonthNames: WdMonthNames; safecall;
    procedure Set_MonthNames(prop: WdMonthNames); safecall;
    function Get_CursorMovement: WdCursorMovement; safecall;
    procedure Set_CursorMovement(prop: WdCursorMovement); safecall;
    function Get_VisualSelection: WdVisualSelection; safecall;
    procedure Set_VisualSelection(prop: WdVisualSelection); safecall;
    function Get_ShowDiacritics: WordBool; safecall;
    procedure Set_ShowDiacritics(prop: WordBool); safecall;
    function Get_ShowControlCharacters: WordBool; safecall;
    procedure Set_ShowControlCharacters(prop: WordBool); safecall;
    function Get_AddControlCharacters: WordBool; safecall;
    procedure Set_AddControlCharacters(prop: WordBool); safecall;
    function Get_AddBiDirectionalMarksWhenSavingTextFile: WordBool; safecall;
    procedure Set_AddBiDirectionalMarksWhenSavingTextFile(prop: WordBool); safecall;
    function Get_StrictInitialAlefHamza: WordBool; safecall;
    procedure Set_StrictInitialAlefHamza(prop: WordBool); safecall;
    function Get_StrictFinalYaa: WordBool; safecall;
    procedure Set_StrictFinalYaa(prop: WordBool); safecall;
    function Get_HebrewMode: WdHebSpellStart; safecall;
    procedure Set_HebrewMode(prop: WdHebSpellStart); safecall;
    function Get_ArabicMode: WdAraSpeller; safecall;
    procedure Set_ArabicMode(prop: WdAraSpeller); safecall;
    function Get_AllowClickAndTypeMouse: WordBool; safecall;
    procedure Set_AllowClickAndTypeMouse(prop: WordBool); safecall;
    function Get_UseGermanSpellingReform: WordBool; safecall;
    procedure Set_UseGermanSpellingReform(prop: WordBool); safecall;
    function Get_InterpretHighAnsi: WdHighAnsiText; safecall;
    procedure Set_InterpretHighAnsi(prop: WdHighAnsiText); safecall;
    function Get_AddHebDoubleQuote: WordBool; safecall;
    procedure Set_AddHebDoubleQuote(prop: WordBool); safecall;
    function Get_UseDiffDiacColor: WordBool; safecall;
    procedure Set_UseDiffDiacColor(prop: WordBool); safecall;
    function Get_DiacriticColorVal: WdColor; safecall;
    procedure Set_DiacriticColorVal(prop: WdColor); safecall;
    function Get_OptimizeForWord97byDefault: WordBool; safecall;
    procedure Set_OptimizeForWord97byDefault(prop: WordBool); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property AllowAccentedUppercase: WordBool read Get_AllowAccentedUppercase write Set_AllowAccentedUppercase;
    property WPHelp: WordBool read Get_WPHelp write Set_WPHelp;
    property WPDocNavKeys: WordBool read Get_WPDocNavKeys write Set_WPDocNavKeys;
    property Pagination: WordBool read Get_Pagination write Set_Pagination;
    property BlueScreen: WordBool read Get_BlueScreen write Set_BlueScreen;
    property EnableSound: WordBool read Get_EnableSound write Set_EnableSound;
    property ConfirmConversions: WordBool read Get_ConfirmConversions write Set_ConfirmConversions;
    property UpdateLinksAtOpen: WordBool read Get_UpdateLinksAtOpen write Set_UpdateLinksAtOpen;
    property SendMailAttach: WordBool read Get_SendMailAttach write Set_SendMailAttach;
    property MeasurementUnit: WdMeasurementUnits read Get_MeasurementUnit write Set_MeasurementUnit;
    property ButtonFieldClicks: Integer read Get_ButtonFieldClicks write Set_ButtonFieldClicks;
    property ShortMenuNames: WordBool read Get_ShortMenuNames write Set_ShortMenuNames;
    property RTFInClipboard: WordBool read Get_RTFInClipboard write Set_RTFInClipboard;
    property UpdateFieldsAtPrint: WordBool read Get_UpdateFieldsAtPrint write Set_UpdateFieldsAtPrint;
    property PrintProperties: WordBool read Get_PrintProperties write Set_PrintProperties;
    property PrintFieldCodes: WordBool read Get_PrintFieldCodes write Set_PrintFieldCodes;
    property PrintComments: WordBool read Get_PrintComments write Set_PrintComments;
    property PrintHiddenText: WordBool read Get_PrintHiddenText write Set_PrintHiddenText;
    property EnvelopeFeederInstalled: WordBool read Get_EnvelopeFeederInstalled;
    property UpdateLinksAtPrint: WordBool read Get_UpdateLinksAtPrint write Set_UpdateLinksAtPrint;
    property PrintBackground: WordBool read Get_PrintBackground write Set_PrintBackground;
    property PrintDrawingObjects: WordBool read Get_PrintDrawingObjects write Set_PrintDrawingObjects;
    property DefaultTray: WideString read Get_DefaultTray write Set_DefaultTray;
    property DefaultTrayID: Integer read Get_DefaultTrayID write Set_DefaultTrayID;
    property CreateBackup: WordBool read Get_CreateBackup write Set_CreateBackup;
    property AllowFastSave: WordBool read Get_AllowFastSave write Set_AllowFastSave;
    property SavePropertiesPrompt: WordBool read Get_SavePropertiesPrompt write Set_SavePropertiesPrompt;
    property SaveNormalPrompt: WordBool read Get_SaveNormalPrompt write Set_SaveNormalPrompt;
    property SaveInterval: Integer read Get_SaveInterval write Set_SaveInterval;
    property BackgroundSave: WordBool read Get_BackgroundSave write Set_BackgroundSave;
    property InsertedTextMark: WdInsertedTextMark read Get_InsertedTextMark write Set_InsertedTextMark;
    property DeletedTextMark: WdDeletedTextMark read Get_DeletedTextMark write Set_DeletedTextMark;
    property RevisedLinesMark: WdRevisedLinesMark read Get_RevisedLinesMark write Set_RevisedLinesMark;
    property InsertedTextColor: WdColorIndex read Get_InsertedTextColor write Set_InsertedTextColor;
    property DeletedTextColor: WdColorIndex read Get_DeletedTextColor write Set_DeletedTextColor;
    property RevisedLinesColor: WdColorIndex read Get_RevisedLinesColor write Set_RevisedLinesColor;
    property DefaultFilePath[Path: WdDefaultFilePath]: WideString read Get_DefaultFilePath write Set_DefaultFilePath;
    property Overtype: WordBool read Get_Overtype write Set_Overtype;
    property ReplaceSelection: WordBool read Get_ReplaceSelection write Set_ReplaceSelection;
    property AllowDragAndDrop: WordBool read Get_AllowDragAndDrop write Set_AllowDragAndDrop;
    property AutoWordSelection: WordBool read Get_AutoWordSelection write Set_AutoWordSelection;
    property INSKeyForPaste: WordBool read Get_INSKeyForPaste write Set_INSKeyForPaste;
    property SmartCutPaste: WordBool read Get_SmartCutPaste write Set_SmartCutPaste;
    property TabIndentKey: WordBool read Get_TabIndentKey write Set_TabIndentKey;
    property PictureEditor: WideString read Get_PictureEditor write Set_PictureEditor;
    property AnimateScreenMovements: WordBool read Get_AnimateScreenMovements write Set_AnimateScreenMovements;
    property VirusProtection: WordBool read Get_VirusProtection write Set_VirusProtection;
    property RevisedPropertiesMark: WdRevisedPropertiesMark read Get_RevisedPropertiesMark write Set_RevisedPropertiesMark;
    property RevisedPropertiesColor: WdColorIndex read Get_RevisedPropertiesColor write Set_RevisedPropertiesColor;
    property SnapToGrid: WordBool read Get_SnapToGrid write Set_SnapToGrid;
    property SnapToShapes: WordBool read Get_SnapToShapes write Set_SnapToShapes;
    property GridDistanceHorizontal: Single read Get_GridDistanceHorizontal write Set_GridDistanceHorizontal;
    property GridDistanceVertical: Single read Get_GridDistanceVertical write Set_GridDistanceVertical;
    property GridOriginHorizontal: Single read Get_GridOriginHorizontal write Set_GridOriginHorizontal;
    property GridOriginVertical: Single read Get_GridOriginVertical write Set_GridOriginVertical;
    property InlineConversion: WordBool read Get_InlineConversion write Set_InlineConversion;
    property IMEAutomaticControl: WordBool read Get_IMEAutomaticControl write Set_IMEAutomaticControl;
    property AutoFormatApplyHeadings: WordBool read Get_AutoFormatApplyHeadings write Set_AutoFormatApplyHeadings;
    property AutoFormatApplyLists: WordBool read Get_AutoFormatApplyLists write Set_AutoFormatApplyLists;
    property AutoFormatApplyBulletedLists: WordBool read Get_AutoFormatApplyBulletedLists write Set_AutoFormatApplyBulletedLists;
    property AutoFormatApplyOtherParas: WordBool read Get_AutoFormatApplyOtherParas write Set_AutoFormatApplyOtherParas;
    property AutoFormatReplaceQuotes: WordBool read Get_AutoFormatReplaceQuotes write Set_AutoFormatReplaceQuotes;
    property AutoFormatReplaceSymbols: WordBool read Get_AutoFormatReplaceSymbols write Set_AutoFormatReplaceSymbols;
    property AutoFormatReplaceOrdinals: WordBool read Get_AutoFormatReplaceOrdinals write Set_AutoFormatReplaceOrdinals;
    property AutoFormatReplaceFractions: WordBool read Get_AutoFormatReplaceFractions write Set_AutoFormatReplaceFractions;
    property AutoFormatReplacePlainTextEmphasis: WordBool read Get_AutoFormatReplacePlainTextEmphasis write Set_AutoFormatReplacePlainTextEmphasis;
    property AutoFormatPreserveStyles: WordBool read Get_AutoFormatPreserveStyles write Set_AutoFormatPreserveStyles;
    property AutoFormatAsYouTypeApplyHeadings: WordBool read Get_AutoFormatAsYouTypeApplyHeadings write Set_AutoFormatAsYouTypeApplyHeadings;
    property AutoFormatAsYouTypeApplyBorders: WordBool read Get_AutoFormatAsYouTypeApplyBorders write Set_AutoFormatAsYouTypeApplyBorders;
    property AutoFormatAsYouTypeApplyBulletedLists: WordBool read Get_AutoFormatAsYouTypeApplyBulletedLists write Set_AutoFormatAsYouTypeApplyBulletedLists;
    property AutoFormatAsYouTypeApplyNumberedLists: WordBool read Get_AutoFormatAsYouTypeApplyNumberedLists write Set_AutoFormatAsYouTypeApplyNumberedLists;
    property AutoFormatAsYouTypeReplaceQuotes: WordBool read Get_AutoFormatAsYouTypeReplaceQuotes write Set_AutoFormatAsYouTypeReplaceQuotes;
    property AutoFormatAsYouTypeReplaceSymbols: WordBool read Get_AutoFormatAsYouTypeReplaceSymbols write Set_AutoFormatAsYouTypeReplaceSymbols;
    property AutoFormatAsYouTypeReplaceOrdinals: WordBool read Get_AutoFormatAsYouTypeReplaceOrdinals write Set_AutoFormatAsYouTypeReplaceOrdinals;
    property AutoFormatAsYouTypeReplaceFractions: WordBool read Get_AutoFormatAsYouTypeReplaceFractions write Set_AutoFormatAsYouTypeReplaceFractions;
    property AutoFormatAsYouTypeReplacePlainTextEmphasis: WordBool read Get_AutoFormatAsYouTypeReplacePlainTextEmphasis write Set_AutoFormatAsYouTypeReplacePlainTextEmphasis;
    property AutoFormatAsYouTypeFormatListItemBeginning: WordBool read Get_AutoFormatAsYouTypeFormatListItemBeginning write Set_AutoFormatAsYouTypeFormatListItemBeginning;
    property AutoFormatAsYouTypeDefineStyles: WordBool read Get_AutoFormatAsYouTypeDefineStyles write Set_AutoFormatAsYouTypeDefineStyles;
    property AutoFormatPlainTextWordMail: WordBool read Get_AutoFormatPlainTextWordMail write Set_AutoFormatPlainTextWordMail;
    property AutoFormatAsYouTypeReplaceHyperlinks: WordBool read Get_AutoFormatAsYouTypeReplaceHyperlinks write Set_AutoFormatAsYouTypeReplaceHyperlinks;
    property AutoFormatReplaceHyperlinks: WordBool read Get_AutoFormatReplaceHyperlinks write Set_AutoFormatReplaceHyperlinks;
    property DefaultHighlightColorIndex: WdColorIndex read Get_DefaultHighlightColorIndex write Set_DefaultHighlightColorIndex;
    property DefaultBorderLineStyle: WdLineStyle read Get_DefaultBorderLineStyle write Set_DefaultBorderLineStyle;
    property CheckSpellingAsYouType: WordBool read Get_CheckSpellingAsYouType write Set_CheckSpellingAsYouType;
    property CheckGrammarAsYouType: WordBool read Get_CheckGrammarAsYouType write Set_CheckGrammarAsYouType;
    property IgnoreInternetAndFileAddresses: WordBool read Get_IgnoreInternetAndFileAddresses write Set_IgnoreInternetAndFileAddresses;
    property ShowReadabilityStatistics: WordBool read Get_ShowReadabilityStatistics write Set_ShowReadabilityStatistics;
    property IgnoreUppercase: WordBool read Get_IgnoreUppercase write Set_IgnoreUppercase;
    property IgnoreMixedDigits: WordBool read Get_IgnoreMixedDigits write Set_IgnoreMixedDigits;
    property SuggestFromMainDictionaryOnly: WordBool read Get_SuggestFromMainDictionaryOnly write Set_SuggestFromMainDictionaryOnly;
    property SuggestSpellingCorrections: WordBool read Get_SuggestSpellingCorrections write Set_SuggestSpellingCorrections;
    property DefaultBorderLineWidth: WdLineWidth read Get_DefaultBorderLineWidth write Set_DefaultBorderLineWidth;
    property CheckGrammarWithSpelling: WordBool read Get_CheckGrammarWithSpelling write Set_CheckGrammarWithSpelling;
    property DefaultOpenFormat: WdOpenFormat read Get_DefaultOpenFormat write Set_DefaultOpenFormat;
    property PrintDraft: WordBool read Get_PrintDraft write Set_PrintDraft;
    property PrintReverse: WordBool read Get_PrintReverse write Set_PrintReverse;
    property MapPaperSize: WordBool read Get_MapPaperSize write Set_MapPaperSize;
    property AutoFormatAsYouTypeApplyTables: WordBool read Get_AutoFormatAsYouTypeApplyTables write Set_AutoFormatAsYouTypeApplyTables;
    property AutoFormatApplyFirstIndents: WordBool read Get_AutoFormatApplyFirstIndents write Set_AutoFormatApplyFirstIndents;
    property AutoFormatMatchParentheses: WordBool read Get_AutoFormatMatchParentheses write Set_AutoFormatMatchParentheses;
    property AutoFormatReplaceFarEastDashes: WordBool read Get_AutoFormatReplaceFarEastDashes write Set_AutoFormatReplaceFarEastDashes;
    property AutoFormatDeleteAutoSpaces: WordBool read Get_AutoFormatDeleteAutoSpaces write Set_AutoFormatDeleteAutoSpaces;
    property AutoFormatAsYouTypeApplyFirstIndents: WordBool read Get_AutoFormatAsYouTypeApplyFirstIndents write Set_AutoFormatAsYouTypeApplyFirstIndents;
    property AutoFormatAsYouTypeApplyDates: WordBool read Get_AutoFormatAsYouTypeApplyDates write Set_AutoFormatAsYouTypeApplyDates;
    property AutoFormatAsYouTypeApplyClosings: WordBool read Get_AutoFormatAsYouTypeApplyClosings write Set_AutoFormatAsYouTypeApplyClosings;
    property AutoFormatAsYouTypeMatchParentheses: WordBool read Get_AutoFormatAsYouTypeMatchParentheses write Set_AutoFormatAsYouTypeMatchParentheses;
    property AutoFormatAsYouTypeReplaceFarEastDashes: WordBool read Get_AutoFormatAsYouTypeReplaceFarEastDashes write Set_AutoFormatAsYouTypeReplaceFarEastDashes;
    property AutoFormatAsYouTypeDeleteAutoSpaces: WordBool read Get_AutoFormatAsYouTypeDeleteAutoSpaces write Set_AutoFormatAsYouTypeDeleteAutoSpaces;
    property AutoFormatAsYouTypeInsertClosings: WordBool read Get_AutoFormatAsYouTypeInsertClosings write Set_AutoFormatAsYouTypeInsertClosings;
    property AutoFormatAsYouTypeAutoLetterWizard: WordBool read Get_AutoFormatAsYouTypeAutoLetterWizard write Set_AutoFormatAsYouTypeAutoLetterWizard;
    property AutoFormatAsYouTypeInsertOvers: WordBool read Get_AutoFormatAsYouTypeInsertOvers write Set_AutoFormatAsYouTypeInsertOvers;
    property DisplayGridLines: WordBool read Get_DisplayGridLines write Set_DisplayGridLines;
    property MatchFuzzyCase: WordBool read Get_MatchFuzzyCase write Set_MatchFuzzyCase;
    property MatchFuzzyByte: WordBool read Get_MatchFuzzyByte write Set_MatchFuzzyByte;
    property MatchFuzzyHiragana: WordBool read Get_MatchFuzzyHiragana write Set_MatchFuzzyHiragana;
    property MatchFuzzySmallKana: WordBool read Get_MatchFuzzySmallKana write Set_MatchFuzzySmallKana;
    property MatchFuzzyDash: WordBool read Get_MatchFuzzyDash write Set_MatchFuzzyDash;
    property MatchFuzzyIterationMark: WordBool read Get_MatchFuzzyIterationMark write Set_MatchFuzzyIterationMark;
    property MatchFuzzyKanji: WordBool read Get_MatchFuzzyKanji write Set_MatchFuzzyKanji;
    property MatchFuzzyOldKana: WordBool read Get_MatchFuzzyOldKana write Set_MatchFuzzyOldKana;
    property MatchFuzzyProlongedSoundMark: WordBool read Get_MatchFuzzyProlongedSoundMark write Set_MatchFuzzyProlongedSoundMark;
    property MatchFuzzyDZ: WordBool read Get_MatchFuzzyDZ write Set_MatchFuzzyDZ;
    property MatchFuzzyBV: WordBool read Get_MatchFuzzyBV write Set_MatchFuzzyBV;
    property MatchFuzzyTC: WordBool read Get_MatchFuzzyTC write Set_MatchFuzzyTC;
    property MatchFuzzyHF: WordBool read Get_MatchFuzzyHF write Set_MatchFuzzyHF;
    property MatchFuzzyZJ: WordBool read Get_MatchFuzzyZJ write Set_MatchFuzzyZJ;
    property MatchFuzzyAY: WordBool read Get_MatchFuzzyAY write Set_MatchFuzzyAY;
    property MatchFuzzyKiKu: WordBool read Get_MatchFuzzyKiKu write Set_MatchFuzzyKiKu;
    property MatchFuzzyPunctuation: WordBool read Get_MatchFuzzyPunctuation write Set_MatchFuzzyPunctuation;
    property MatchFuzzySpace: WordBool read Get_MatchFuzzySpace write Set_MatchFuzzySpace;
    property ApplyFarEastFontsToAscii: WordBool read Get_ApplyFarEastFontsToAscii write Set_ApplyFarEastFontsToAscii;
    property ConvertHighAnsiToFarEast: WordBool read Get_ConvertHighAnsiToFarEast write Set_ConvertHighAnsiToFarEast;
    property PrintOddPagesInAscendingOrder: WordBool read Get_PrintOddPagesInAscendingOrder write Set_PrintOddPagesInAscendingOrder;
    property PrintEvenPagesInAscendingOrder: WordBool read Get_PrintEvenPagesInAscendingOrder write Set_PrintEvenPagesInAscendingOrder;
    property DefaultBorderColorIndex: WdColorIndex read Get_DefaultBorderColorIndex write Set_DefaultBorderColorIndex;
    property EnableMisusedWordsDictionary: WordBool read Get_EnableMisusedWordsDictionary write Set_EnableMisusedWordsDictionary;
    property AllowCombinedAuxiliaryForms: WordBool read Get_AllowCombinedAuxiliaryForms write Set_AllowCombinedAuxiliaryForms;
    property HangulHanjaFastConversion: WordBool read Get_HangulHanjaFastConversion write Set_HangulHanjaFastConversion;
    property CheckHangulEndings: WordBool read Get_CheckHangulEndings write Set_CheckHangulEndings;
    property EnableHangulHanjaRecentOrdering: WordBool read Get_EnableHangulHanjaRecentOrdering write Set_EnableHangulHanjaRecentOrdering;
    property MultipleWordConversionsMode: WdMultipleWordConversionsMode read Get_MultipleWordConversionsMode write Set_MultipleWordConversionsMode;
    property DefaultBorderColor: WdColor read Get_DefaultBorderColor write Set_DefaultBorderColor;
    property AllowPixelUnits: WordBool read Get_AllowPixelUnits write Set_AllowPixelUnits;
    property UseCharacterUnit: WordBool read Get_UseCharacterUnit write Set_UseCharacterUnit;
    property AllowCompoundNounProcessing: WordBool read Get_AllowCompoundNounProcessing write Set_AllowCompoundNounProcessing;
    property AutoKeyboardSwitching: WordBool read Get_AutoKeyboardSwitching write Set_AutoKeyboardSwitching;
    property DocumentViewDirection: WdDocumentViewDirection read Get_DocumentViewDirection write Set_DocumentViewDirection;
    property ArabicNumeral: WdArabicNumeral read Get_ArabicNumeral write Set_ArabicNumeral;
    property MonthNames: WdMonthNames read Get_MonthNames write Set_MonthNames;
    property CursorMovement: WdCursorMovement read Get_CursorMovement write Set_CursorMovement;
    property VisualSelection: WdVisualSelection read Get_VisualSelection write Set_VisualSelection;
    property ShowDiacritics: WordBool read Get_ShowDiacritics write Set_ShowDiacritics;
    property ShowControlCharacters: WordBool read Get_ShowControlCharacters write Set_ShowControlCharacters;
    property AddControlCharacters: WordBool read Get_AddControlCharacters write Set_AddControlCharacters;
    property AddBiDirectionalMarksWhenSavingTextFile: WordBool read Get_AddBiDirectionalMarksWhenSavingTextFile write Set_AddBiDirectionalMarksWhenSavingTextFile;
    property StrictInitialAlefHamza: WordBool read Get_StrictInitialAlefHamza write Set_StrictInitialAlefHamza;
    property StrictFinalYaa: WordBool read Get_StrictFinalYaa write Set_StrictFinalYaa;
    property HebrewMode: WdHebSpellStart read Get_HebrewMode write Set_HebrewMode;
    property ArabicMode: WdAraSpeller read Get_ArabicMode write Set_ArabicMode;
    property AllowClickAndTypeMouse: WordBool read Get_AllowClickAndTypeMouse write Set_AllowClickAndTypeMouse;
    property UseGermanSpellingReform: WordBool read Get_UseGermanSpellingReform write Set_UseGermanSpellingReform;
    property InterpretHighAnsi: WdHighAnsiText read Get_InterpretHighAnsi write Set_InterpretHighAnsi;
    property AddHebDoubleQuote: WordBool read Get_AddHebDoubleQuote write Set_AddHebDoubleQuote;
    property UseDiffDiacColor: WordBool read Get_UseDiffDiacColor write Set_UseDiffDiacColor;
    property DiacriticColorVal: WdColor read Get_DiacriticColorVal write Set_DiacriticColorVal;
    property OptimizeForWord97byDefault: WordBool read Get_OptimizeForWord97byDefault write Set_OptimizeForWord97byDefault;
  end;

// *********************************************************************//
// DispIntf:  OptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B7-0000-0000-C000-000000000046}
// *********************************************************************//
  OptionsDisp = dispinterface
    ['{000209B7-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property AllowAccentedUppercase: WordBool dispid 1;
    property WPHelp: WordBool dispid 17;
    property WPDocNavKeys: WordBool dispid 18;
    property Pagination: WordBool dispid 19;
    property BlueScreen: WordBool dispid 20;
    property EnableSound: WordBool dispid 21;
    property ConfirmConversions: WordBool dispid 22;
    property UpdateLinksAtOpen: WordBool dispid 23;
    property SendMailAttach: WordBool dispid 24;
    property MeasurementUnit: WdMeasurementUnits dispid 26;
    property ButtonFieldClicks: Integer dispid 27;
    property ShortMenuNames: WordBool dispid 28;
    property RTFInClipboard: WordBool dispid 29;
    property UpdateFieldsAtPrint: WordBool dispid 30;
    property PrintProperties: WordBool dispid 31;
    property PrintFieldCodes: WordBool dispid 32;
    property PrintComments: WordBool dispid 33;
    property PrintHiddenText: WordBool dispid 34;
    property EnvelopeFeederInstalled: WordBool readonly dispid 35;
    property UpdateLinksAtPrint: WordBool dispid 36;
    property PrintBackground: WordBool dispid 37;
    property PrintDrawingObjects: WordBool dispid 38;
    property DefaultTray: WideString dispid 39;
    property DefaultTrayID: Integer dispid 40;
    property CreateBackup: WordBool dispid 41;
    property AllowFastSave: WordBool dispid 42;
    property SavePropertiesPrompt: WordBool dispid 43;
    property SaveNormalPrompt: WordBool dispid 44;
    property SaveInterval: Integer dispid 45;
    property BackgroundSave: WordBool dispid 46;
    property InsertedTextMark: WdInsertedTextMark dispid 57;
    property DeletedTextMark: WdDeletedTextMark dispid 58;
    property RevisedLinesMark: WdRevisedLinesMark dispid 59;
    property InsertedTextColor: WdColorIndex dispid 60;
    property DeletedTextColor: WdColorIndex dispid 61;
    property RevisedLinesColor: WdColorIndex dispid 62;
    property DefaultFilePath[Path: WdDefaultFilePath]: WideString dispid 65;
    property Overtype: WordBool dispid 66;
    property ReplaceSelection: WordBool dispid 67;
    property AllowDragAndDrop: WordBool dispid 68;
    property AutoWordSelection: WordBool dispid 69;
    property INSKeyForPaste: WordBool dispid 70;
    property SmartCutPaste: WordBool dispid 71;
    property TabIndentKey: WordBool dispid 72;
    property PictureEditor: WideString dispid 73;
    property AnimateScreenMovements: WordBool dispid 74;
    property VirusProtection: WordBool dispid 75;
    property RevisedPropertiesMark: WdRevisedPropertiesMark dispid 76;
    property RevisedPropertiesColor: WdColorIndex dispid 77;
    property SnapToGrid: WordBool dispid 79;
    property SnapToShapes: WordBool dispid 80;
    property GridDistanceHorizontal: Single dispid 81;
    property GridDistanceVertical: Single dispid 82;
    property GridOriginHorizontal: Single dispid 83;
    property GridOriginVertical: Single dispid 84;
    property InlineConversion: WordBool dispid 86;
    property IMEAutomaticControl: WordBool dispid 87;
    property AutoFormatApplyHeadings: WordBool dispid 250;
    property AutoFormatApplyLists: WordBool dispid 251;
    property AutoFormatApplyBulletedLists: WordBool dispid 252;
    property AutoFormatApplyOtherParas: WordBool dispid 253;
    property AutoFormatReplaceQuotes: WordBool dispid 254;
    property AutoFormatReplaceSymbols: WordBool dispid 255;
    property AutoFormatReplaceOrdinals: WordBool dispid 256;
    property AutoFormatReplaceFractions: WordBool dispid 257;
    property AutoFormatReplacePlainTextEmphasis: WordBool dispid 258;
    property AutoFormatPreserveStyles: WordBool dispid 259;
    property AutoFormatAsYouTypeApplyHeadings: WordBool dispid 260;
    property AutoFormatAsYouTypeApplyBorders: WordBool dispid 261;
    property AutoFormatAsYouTypeApplyBulletedLists: WordBool dispid 262;
    property AutoFormatAsYouTypeApplyNumberedLists: WordBool dispid 263;
    property AutoFormatAsYouTypeReplaceQuotes: WordBool dispid 264;
    property AutoFormatAsYouTypeReplaceSymbols: WordBool dispid 265;
    property AutoFormatAsYouTypeReplaceOrdinals: WordBool dispid 266;
    property AutoFormatAsYouTypeReplaceFractions: WordBool dispid 267;
    property AutoFormatAsYouTypeReplacePlainTextEmphasis: WordBool dispid 268;
    property AutoFormatAsYouTypeFormatListItemBeginning: WordBool dispid 269;
    property AutoFormatAsYouTypeDefineStyles: WordBool dispid 270;
    property AutoFormatPlainTextWordMail: WordBool dispid 271;
    property AutoFormatAsYouTypeReplaceHyperlinks: WordBool dispid 272;
    property AutoFormatReplaceHyperlinks: WordBool dispid 273;
    property DefaultHighlightColorIndex: WdColorIndex dispid 274;
    property DefaultBorderLineStyle: WdLineStyle dispid 275;
    property CheckSpellingAsYouType: WordBool dispid 276;
    property CheckGrammarAsYouType: WordBool dispid 277;
    property IgnoreInternetAndFileAddresses: WordBool dispid 278;
    property ShowReadabilityStatistics: WordBool dispid 279;
    property IgnoreUppercase: WordBool dispid 280;
    property IgnoreMixedDigits: WordBool dispid 281;
    property SuggestFromMainDictionaryOnly: WordBool dispid 282;
    property SuggestSpellingCorrections: WordBool dispid 283;
    property DefaultBorderLineWidth: WdLineWidth dispid 284;
    property CheckGrammarWithSpelling: WordBool dispid 285;
    property DefaultOpenFormat: WdOpenFormat dispid 286;
    property PrintDraft: WordBool dispid 287;
    property PrintReverse: WordBool dispid 288;
    property MapPaperSize: WordBool dispid 289;
    property AutoFormatAsYouTypeApplyTables: WordBool dispid 290;
    property AutoFormatApplyFirstIndents: WordBool dispid 291;
    property AutoFormatMatchParentheses: WordBool dispid 294;
    property AutoFormatReplaceFarEastDashes: WordBool dispid 295;
    property AutoFormatDeleteAutoSpaces: WordBool dispid 296;
    property AutoFormatAsYouTypeApplyFirstIndents: WordBool dispid 297;
    property AutoFormatAsYouTypeApplyDates: WordBool dispid 298;
    property AutoFormatAsYouTypeApplyClosings: WordBool dispid 299;
    property AutoFormatAsYouTypeMatchParentheses: WordBool dispid 300;
    property AutoFormatAsYouTypeReplaceFarEastDashes: WordBool dispid 301;
    property AutoFormatAsYouTypeDeleteAutoSpaces: WordBool dispid 302;
    property AutoFormatAsYouTypeInsertClosings: WordBool dispid 303;
    property AutoFormatAsYouTypeAutoLetterWizard: WordBool dispid 304;
    property AutoFormatAsYouTypeInsertOvers: WordBool dispid 305;
    property DisplayGridLines: WordBool dispid 306;
    property MatchFuzzyCase: WordBool dispid 309;
    property MatchFuzzyByte: WordBool dispid 310;
    property MatchFuzzyHiragana: WordBool dispid 311;
    property MatchFuzzySmallKana: WordBool dispid 312;
    property MatchFuzzyDash: WordBool dispid 313;
    property MatchFuzzyIterationMark: WordBool dispid 314;
    property MatchFuzzyKanji: WordBool dispid 315;
    property MatchFuzzyOldKana: WordBool dispid 316;
    property MatchFuzzyProlongedSoundMark: WordBool dispid 317;
    property MatchFuzzyDZ: WordBool dispid 318;
    property MatchFuzzyBV: WordBool dispid 319;
    property MatchFuzzyTC: WordBool dispid 320;
    property MatchFuzzyHF: WordBool dispid 321;
    property MatchFuzzyZJ: WordBool dispid 322;
    property MatchFuzzyAY: WordBool dispid 323;
    property MatchFuzzyKiKu: WordBool dispid 324;
    property MatchFuzzyPunctuation: WordBool dispid 325;
    property MatchFuzzySpace: WordBool dispid 326;
    property ApplyFarEastFontsToAscii: WordBool dispid 327;
    property ConvertHighAnsiToFarEast: WordBool dispid 328;
    property PrintOddPagesInAscendingOrder: WordBool dispid 330;
    property PrintEvenPagesInAscendingOrder: WordBool dispid 331;
    property DefaultBorderColorIndex: WdColorIndex dispid 337;
    property EnableMisusedWordsDictionary: WordBool dispid 338;
    property AllowCombinedAuxiliaryForms: WordBool dispid 339;
    property HangulHanjaFastConversion: WordBool dispid 340;
    property CheckHangulEndings: WordBool dispid 341;
    property EnableHangulHanjaRecentOrdering: WordBool dispid 342;
    property MultipleWordConversionsMode: WdMultipleWordConversionsMode dispid 343;
    procedure SetWPHelpOptions(var CommandKeyHelp: OleVariant; var DocNavigationKeys: OleVariant; 
                               var MouseSimulation: OleVariant; var DemoGuidance: OleVariant; 
                               var DemoSpeed: OleVariant; var HelpType: OleVariant); dispid 333;
    property DefaultBorderColor: WdColor dispid 344;
    property AllowPixelUnits: WordBool dispid 345;
    property UseCharacterUnit: WordBool dispid 346;
    property AllowCompoundNounProcessing: WordBool dispid 347;
    property AutoKeyboardSwitching: WordBool dispid 399;
    property DocumentViewDirection: WdDocumentViewDirection dispid 400;
    property ArabicNumeral: WdArabicNumeral dispid 401;
    property MonthNames: WdMonthNames dispid 402;
    property CursorMovement: WdCursorMovement dispid 403;
    property VisualSelection: WdVisualSelection dispid 404;
    property ShowDiacritics: WordBool dispid 405;
    property ShowControlCharacters: WordBool dispid 406;
    property AddControlCharacters: WordBool dispid 407;
    property AddBiDirectionalMarksWhenSavingTextFile: WordBool dispid 408;
    property StrictInitialAlefHamza: WordBool dispid 409;
    property StrictFinalYaa: WordBool dispid 410;
    property HebrewMode: WdHebSpellStart dispid 411;
    property ArabicMode: WdAraSpeller dispid 412;
    property AllowClickAndTypeMouse: WordBool dispid 413;
    property UseGermanSpellingReform: WordBool dispid 415;
    property InterpretHighAnsi: WdHighAnsiText dispid 418;
    property AddHebDoubleQuote: WordBool dispid 419;
    property UseDiffDiacColor: WordBool dispid 420;
    property DiacriticColorVal: WdColor dispid 421;
    property OptimizeForWord97byDefault: WordBool dispid 423;
  end;

// *********************************************************************//
// Interface: MailMessage
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BA-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMessage = interface(IDispatch)
    ['{000209BA-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    procedure CheckName; safecall;
    procedure Delete; safecall;
    procedure DisplayMoveDialog; safecall;
    procedure DisplayProperties; safecall;
    procedure DisplaySelectNamesDialog; safecall;
    procedure Forward; safecall;
    procedure GoToNext; safecall;
    procedure GoToPrevious; safecall;
    procedure Reply; safecall;
    procedure ReplyAll; safecall;
    procedure ToggleHeader; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  MailMessageDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BA-0000-0000-C000-000000000046}
// *********************************************************************//
  MailMessageDisp = dispinterface
    ['{000209BA-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure CheckName; dispid 334;
    procedure Delete; dispid 335;
    procedure DisplayMoveDialog; dispid 336;
    procedure DisplayProperties; dispid 337;
    procedure DisplaySelectNamesDialog; dispid 338;
    procedure Forward; dispid 339;
    procedure GoToNext; dispid 340;
    procedure GoToPrevious; dispid 341;
    procedure Reply; dispid 342;
    procedure ReplyAll; dispid 343;
    procedure ToggleHeader; dispid 344;
  end;

// *********************************************************************//
// Interface: ProofreadingErrors
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BB-0000-0000-C000-000000000046}
// *********************************************************************//
  ProofreadingErrors = interface(IDispatch)
    ['{000209BB-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Type_: WdProofreadingErrorType; safecall;
    function Item(Index: Integer): Range; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Type_: WdProofreadingErrorType read Get_Type_;
  end;

// *********************************************************************//
// DispIntf:  ProofreadingErrorsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BB-0000-0000-C000-000000000046}
// *********************************************************************//
  ProofreadingErrorsDisp = dispinterface
    ['{000209BB-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Type_: WdProofreadingErrorType readonly dispid 2;
    function Item(Index: Integer): Range; dispid 0;
  end;

// *********************************************************************//
// Interface: Mailer
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BD-0000-0000-C000-000000000046}
// *********************************************************************//
  Mailer = interface(IDispatch)
    ['{000209BD-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_BCCRecipients: OleVariant; safecall;
    procedure Set_BCCRecipients(var prop: OleVariant); safecall;
    function Get_CCRecipients: OleVariant; safecall;
    procedure Set_CCRecipients(var prop: OleVariant); safecall;
    function Get_Recipients: OleVariant; safecall;
    procedure Set_Recipients(var prop: OleVariant); safecall;
    function Get_Enclosures: OleVariant; safecall;
    procedure Set_Enclosures(var prop: OleVariant); safecall;
    function Get_Sender: WideString; safecall;
    function Get_SendDateTime: TDateTime; safecall;
    function Get_Received: WordBool; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Sender: WideString read Get_Sender;
    property SendDateTime: TDateTime read Get_SendDateTime;
    property Received: WordBool read Get_Received;
    property Subject: WideString read Get_Subject write Set_Subject;
  end;

// *********************************************************************//
// DispIntf:  MailerDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BD-0000-0000-C000-000000000046}
// *********************************************************************//
  MailerDisp = dispinterface
    ['{000209BD-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    function BCCRecipients: OleVariant; dispid 100;
    function CCRecipients: OleVariant; dispid 101;
    function Recipients: OleVariant; dispid 102;
    function Enclosures: OleVariant; dispid 103;
    property Sender: WideString readonly dispid 104;
    property SendDateTime: TDateTime readonly dispid 105;
    property Received: WordBool readonly dispid 106;
    property Subject: WideString dispid 107;
  end;

// *********************************************************************//
// Interface: WrapFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C3-0000-0000-C000-000000000046}
// *********************************************************************//
  WrapFormat = interface(IDispatch)
    ['{000209C3-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Type_: WdWrapType; safecall;
    procedure Set_Type_(prop: WdWrapType); safecall;
    function Get_Side: WdWrapSideType; safecall;
    procedure Set_Side(prop: WdWrapSideType); safecall;
    function Get_DistanceTop: Single; safecall;
    procedure Set_DistanceTop(prop: Single); safecall;
    function Get_DistanceBottom: Single; safecall;
    procedure Set_DistanceBottom(prop: Single); safecall;
    function Get_DistanceLeft: Single; safecall;
    procedure Set_DistanceLeft(prop: Single); safecall;
    function Get_DistanceRight: Single; safecall;
    procedure Set_DistanceRight(prop: Single); safecall;
    function Get_AllowOverlap: Integer; safecall;
    procedure Set_AllowOverlap(prop: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Type_: WdWrapType read Get_Type_ write Set_Type_;
    property Side: WdWrapSideType read Get_Side write Set_Side;
    property DistanceTop: Single read Get_DistanceTop write Set_DistanceTop;
    property DistanceBottom: Single read Get_DistanceBottom write Set_DistanceBottom;
    property DistanceLeft: Single read Get_DistanceLeft write Set_DistanceLeft;
    property DistanceRight: Single read Get_DistanceRight write Set_DistanceRight;
    property AllowOverlap: Integer read Get_AllowOverlap write Set_AllowOverlap;
  end;

// *********************************************************************//
// DispIntf:  WrapFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C3-0000-0000-C000-000000000046}
// *********************************************************************//
  WrapFormatDisp = dispinterface
    ['{000209C3-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Type_: WdWrapType dispid 100;
    property Side: WdWrapSideType dispid 101;
    property DistanceTop: Single dispid 102;
    property DistanceBottom: Single dispid 103;
    property DistanceLeft: Single dispid 104;
    property DistanceRight: Single dispid 105;
    property AllowOverlap: Integer dispid 106;
  end;

// *********************************************************************//
// Interface: HangulAndAlphabetExceptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D1-0000-0000-C000-000000000046}
// *********************************************************************//
  HangulAndAlphabetExceptions = interface(IDispatch)
    ['{000209D1-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): HangulAndAlphabetException; safecall;
    function Add(const Name: WideString): HangulAndAlphabetException; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  HangulAndAlphabetExceptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D1-0000-0000-C000-000000000046}
// *********************************************************************//
  HangulAndAlphabetExceptionsDisp = dispinterface
    ['{000209D1-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): HangulAndAlphabetException; dispid 0;
    function Add(const Name: WideString): HangulAndAlphabetException; dispid 101;
  end;

// *********************************************************************//
// Interface: HangulAndAlphabetException
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D2-0000-0000-C000-000000000046}
// *********************************************************************//
  HangulAndAlphabetException = interface(IDispatch)
    ['{000209D2-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  HangulAndAlphabetExceptionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D2-0000-0000-C000-000000000046}
// *********************************************************************//
  HangulAndAlphabetExceptionDisp = dispinterface
    ['{000209D2-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString readonly dispid 2;
    procedure Delete; dispid 101;
  end;

// *********************************************************************//
// Interface: Adjustments
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C4-0000-0000-C000-000000000046}
// *********************************************************************//
  Adjustments = interface(IDispatch)
    ['{000209C4-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): Single; safecall;
    procedure Set_Item(Index: Integer; prop: Single); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: Single read Get_Item write Set_Item; default;
  end;

// *********************************************************************//
// DispIntf:  AdjustmentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C4-0000-0000-C000-000000000046}
// *********************************************************************//
  AdjustmentsDisp = dispinterface
    ['{000209C4-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property Count: Integer readonly dispid 2;
    property Item[Index: Integer]: Single dispid 0; default;
  end;

// *********************************************************************//
// Interface: CalloutFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C5-0000-0000-C000-000000000046}
// *********************************************************************//
  CalloutFormat = interface(IDispatch)
    ['{000209C5-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Accent: MsoTriState; safecall;
    procedure Set_Accent(prop: MsoTriState); safecall;
    function Get_Angle: MsoCalloutAngleType; safecall;
    procedure Set_Angle(prop: MsoCalloutAngleType); safecall;
    function Get_AutoAttach: MsoTriState; safecall;
    procedure Set_AutoAttach(prop: MsoTriState); safecall;
    function Get_AutoLength: MsoTriState; safecall;
    function Get_Border: MsoTriState; safecall;
    procedure Set_Border(prop: MsoTriState); safecall;
    function Get_Drop: Single; safecall;
    function Get_DropType: MsoCalloutDropType; safecall;
    function Get_Gap: Single; safecall;
    procedure Set_Gap(prop: Single); safecall;
    function Get_Length: Single; safecall;
    function Get_Type_: MsoCalloutType; safecall;
    procedure Set_Type_(prop: MsoCalloutType); safecall;
    procedure AutomaticLength; safecall;
    procedure CustomDrop(Drop: Single); safecall;
    procedure CustomLength(Length: Single); safecall;
    procedure PresetDrop(DropType: MsoCalloutDropType); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Accent: MsoTriState read Get_Accent write Set_Accent;
    property Angle: MsoCalloutAngleType read Get_Angle write Set_Angle;
    property AutoAttach: MsoTriState read Get_AutoAttach write Set_AutoAttach;
    property AutoLength: MsoTriState read Get_AutoLength;
    property Border: MsoTriState read Get_Border write Set_Border;
    property Drop: Single read Get_Drop;
    property DropType: MsoCalloutDropType read Get_DropType;
    property Gap: Single read Get_Gap write Set_Gap;
    property Length: Single read Get_Length;
    property Type_: MsoCalloutType read Get_Type_ write Set_Type_;
  end;

// *********************************************************************//
// DispIntf:  CalloutFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C5-0000-0000-C000-000000000046}
// *********************************************************************//
  CalloutFormatDisp = dispinterface
    ['{000209C5-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property Accent: MsoTriState dispid 100;
    property Angle: MsoCalloutAngleType dispid 101;
    property AutoAttach: MsoTriState dispid 102;
    property AutoLength: MsoTriState readonly dispid 103;
    property Border: MsoTriState dispid 104;
    property Drop: Single readonly dispid 105;
    property DropType: MsoCalloutDropType readonly dispid 106;
    property Gap: Single dispid 107;
    property Length: Single readonly dispid 108;
    property Type_: MsoCalloutType dispid 109;
    procedure AutomaticLength; dispid 10;
    procedure CustomDrop(Drop: Single); dispid 11;
    procedure CustomLength(Length: Single); dispid 12;
    procedure PresetDrop(DropType: MsoCalloutDropType); dispid 13;
  end;

// *********************************************************************//
// Interface: ColorFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C6-0000-0000-C000-000000000046}
// *********************************************************************//
  ColorFormat = interface(IDispatch)
    ['{000209C6-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_RGB_: Integer; safecall;
    procedure Set_RGB_(prop: Integer); safecall;
    function Get_SchemeColor: Integer; safecall;
    procedure Set_SchemeColor(prop: Integer); safecall;
    function Get_Type_: MsoColorType; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property RGB_: Integer read Get_RGB_ write Set_RGB_;
    property SchemeColor: Integer read Get_SchemeColor write Set_SchemeColor;
    property Type_: MsoColorType read Get_Type_;
  end;

// *********************************************************************//
// DispIntf:  ColorFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C6-0000-0000-C000-000000000046}
// *********************************************************************//
  ColorFormatDisp = dispinterface
    ['{000209C6-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property RGB_: Integer dispid 0;
    property SchemeColor: Integer dispid 100;
    property Type_: MsoColorType readonly dispid 101;
  end;

// *********************************************************************//
// Interface: ConnectorFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C7-0000-0000-C000-000000000046}
// *********************************************************************//
  ConnectorFormat = interface(IDispatch)
    ['{000209C7-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_BeginConnected: MsoTriState; safecall;
    function Get_BeginConnectedShape: Shape; safecall;
    function Get_BeginConnectionSite: Integer; safecall;
    function Get_EndConnected: MsoTriState; safecall;
    function Get_EndConnectedShape: Shape; safecall;
    function Get_EndConnectionSite: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Type_: MsoConnectorType; safecall;
    procedure Set_Type_(prop: MsoConnectorType); safecall;
    procedure BeginConnect(out ConnectedShape: Shape; ConnectionSite: Integer); safecall;
    procedure BeginDisconnect; safecall;
    procedure EndConnect(out ConnectedShape: Shape; ConnectionSite: Integer); safecall;
    procedure EndDisconnect; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property BeginConnected: MsoTriState read Get_BeginConnected;
    property BeginConnectedShape: Shape read Get_BeginConnectedShape;
    property BeginConnectionSite: Integer read Get_BeginConnectionSite;
    property EndConnected: MsoTriState read Get_EndConnected;
    property EndConnectedShape: Shape read Get_EndConnectedShape;
    property EndConnectionSite: Integer read Get_EndConnectionSite;
    property Parent: IDispatch read Get_Parent;
    property Type_: MsoConnectorType read Get_Type_ write Set_Type_;
  end;

// *********************************************************************//
// DispIntf:  ConnectorFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C7-0000-0000-C000-000000000046}
// *********************************************************************//
  ConnectorFormatDisp = dispinterface
    ['{000209C7-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property BeginConnected: MsoTriState readonly dispid 100;
    property BeginConnectedShape: Shape readonly dispid 101;
    property BeginConnectionSite: Integer readonly dispid 102;
    property EndConnected: MsoTriState readonly dispid 103;
    property EndConnectedShape: Shape readonly dispid 104;
    property EndConnectionSite: Integer readonly dispid 105;
    property Parent: IDispatch readonly dispid 1;
    property Type_: MsoConnectorType dispid 106;
    procedure BeginConnect(out ConnectedShape: Shape; ConnectionSite: Integer); dispid 10;
    procedure BeginDisconnect; dispid 11;
    procedure EndConnect(out ConnectedShape: Shape; ConnectionSite: Integer); dispid 12;
    procedure EndDisconnect; dispid 13;
  end;

// *********************************************************************//
// Interface: FillFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C8-0000-0000-C000-000000000046}
// *********************************************************************//
  FillFormat = interface(IDispatch)
    ['{000209C8-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_BackColor: ColorFormat; safecall;
    function Get_ForeColor: ColorFormat; safecall;
    function Get_GradientColorType: MsoGradientColorType; safecall;
    function Get_GradientDegree: Single; safecall;
    function Get_GradientStyle: MsoGradientStyle; safecall;
    function Get_GradientVariant: Integer; safecall;
    function Get_Pattern: MsoPatternType; safecall;
    function Get_PresetGradientType: MsoPresetGradientType; safecall;
    function Get_PresetTexture: MsoPresetTexture; safecall;
    function Get_TextureName: WideString; safecall;
    function Get_TextureType: MsoTextureType; safecall;
    function Get_Transparency: Single; safecall;
    procedure Set_Transparency(prop: Single); safecall;
    function Get_Type_: MsoFillType; safecall;
    function Get_Visible: MsoTriState; safecall;
    procedure Set_Visible(prop: MsoTriState); safecall;
    procedure Background; safecall;
    procedure OneColorGradient(Style: MsoGradientStyle; Variant: Integer; Degree: Single); safecall;
    procedure Patterned(Pattern: MsoPatternType); safecall;
    procedure PresetGradient(Style: MsoGradientStyle; Variant: Integer; 
                             PresetGradientType: MsoPresetGradientType); safecall;
    procedure PresetTextured(PresetTexture: MsoPresetTexture); safecall;
    procedure Solid; safecall;
    procedure TwoColorGradient(Style: MsoGradientStyle; Variant: Integer); safecall;
    procedure UserPicture(const PictureFile: WideString); safecall;
    procedure UserTextured(const TextureFile: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property BackColor: ColorFormat read Get_BackColor;
    property ForeColor: ColorFormat read Get_ForeColor;
    property GradientColorType: MsoGradientColorType read Get_GradientColorType;
    property GradientDegree: Single read Get_GradientDegree;
    property GradientStyle: MsoGradientStyle read Get_GradientStyle;
    property GradientVariant: Integer read Get_GradientVariant;
    property Pattern: MsoPatternType read Get_Pattern;
    property PresetGradientType: MsoPresetGradientType read Get_PresetGradientType;
    property PresetTexture: MsoPresetTexture read Get_PresetTexture;
    property TextureName: WideString read Get_TextureName;
    property TextureType: MsoTextureType read Get_TextureType;
    property Transparency: Single read Get_Transparency write Set_Transparency;
    property Type_: MsoFillType read Get_Type_;
    property Visible: MsoTriState read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  FillFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C8-0000-0000-C000-000000000046}
// *********************************************************************//
  FillFormatDisp = dispinterface
    ['{000209C8-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property BackColor: ColorFormat readonly dispid 100;
    property ForeColor: ColorFormat readonly dispid 101;
    property GradientColorType: MsoGradientColorType readonly dispid 102;
    property GradientDegree: Single readonly dispid 103;
    property GradientStyle: MsoGradientStyle readonly dispid 104;
    property GradientVariant: Integer readonly dispid 105;
    property Pattern: MsoPatternType readonly dispid 106;
    property PresetGradientType: MsoPresetGradientType readonly dispid 107;
    property PresetTexture: MsoPresetTexture readonly dispid 108;
    property TextureName: WideString readonly dispid 109;
    property TextureType: MsoTextureType readonly dispid 110;
    property Transparency: Single dispid 111;
    property Type_: MsoFillType readonly dispid 112;
    property Visible: MsoTriState dispid 113;
    procedure Background; dispid 10;
    procedure OneColorGradient(Style: MsoGradientStyle; Variant: Integer; Degree: Single); dispid 11;
    procedure Patterned(Pattern: MsoPatternType); dispid 12;
    procedure PresetGradient(Style: MsoGradientStyle; Variant: Integer; 
                             PresetGradientType: MsoPresetGradientType); dispid 13;
    procedure PresetTextured(PresetTexture: MsoPresetTexture); dispid 14;
    procedure Solid; dispid 15;
    procedure TwoColorGradient(Style: MsoGradientStyle; Variant: Integer); dispid 16;
    procedure UserPicture(const PictureFile: WideString); dispid 17;
    procedure UserTextured(const TextureFile: WideString); dispid 18;
  end;

// *********************************************************************//
// Interface: FreeformBuilder
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C9-0000-0000-C000-000000000046}
// *********************************************************************//
  FreeformBuilder = interface(IDispatch)
    ['{000209C9-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    procedure AddNodes(SegmentType: MsoSegmentType; EditingType: MsoEditingType; X1: Single; 
                       Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); safecall;
    function ConvertToShape(var Anchor: OleVariant): Shape; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  FreeformBuilderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C9-0000-0000-C000-000000000046}
// *********************************************************************//
  FreeformBuilderDisp = dispinterface
    ['{000209C9-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    procedure AddNodes(SegmentType: MsoSegmentType; EditingType: MsoEditingType; X1: Single; 
                       Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); dispid 10;
    function ConvertToShape(var Anchor: OleVariant): Shape; dispid 11;
  end;

// *********************************************************************//
// Interface: LineFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CA-0000-0000-C000-000000000046}
// *********************************************************************//
  LineFormat = interface(IDispatch)
    ['{000209CA-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_BackColor: ColorFormat; safecall;
    function Get_BeginArrowheadLength: MsoArrowheadLength; safecall;
    procedure Set_BeginArrowheadLength(prop: MsoArrowheadLength); safecall;
    function Get_BeginArrowheadStyle: MsoArrowheadStyle; safecall;
    procedure Set_BeginArrowheadStyle(prop: MsoArrowheadStyle); safecall;
    function Get_BeginArrowheadWidth: MsoArrowheadWidth; safecall;
    procedure Set_BeginArrowheadWidth(prop: MsoArrowheadWidth); safecall;
    function Get_DashStyle: MsoLineDashStyle; safecall;
    procedure Set_DashStyle(prop: MsoLineDashStyle); safecall;
    function Get_EndArrowheadLength: MsoArrowheadLength; safecall;
    procedure Set_EndArrowheadLength(prop: MsoArrowheadLength); safecall;
    function Get_EndArrowheadStyle: MsoArrowheadStyle; safecall;
    procedure Set_EndArrowheadStyle(prop: MsoArrowheadStyle); safecall;
    function Get_EndArrowheadWidth: MsoArrowheadWidth; safecall;
    procedure Set_EndArrowheadWidth(prop: MsoArrowheadWidth); safecall;
    function Get_ForeColor: ColorFormat; safecall;
    function Get_Pattern: MsoPatternType; safecall;
    procedure Set_Pattern(prop: MsoPatternType); safecall;
    function Get_Style: MsoLineStyle; safecall;
    procedure Set_Style(prop: MsoLineStyle); safecall;
    function Get_Transparency: Single; safecall;
    procedure Set_Transparency(prop: Single); safecall;
    function Get_Visible: MsoTriState; safecall;
    procedure Set_Visible(prop: MsoTriState); safecall;
    function Get_Weight: Single; safecall;
    procedure Set_Weight(prop: Single); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property BackColor: ColorFormat read Get_BackColor;
    property BeginArrowheadLength: MsoArrowheadLength read Get_BeginArrowheadLength write Set_BeginArrowheadLength;
    property BeginArrowheadStyle: MsoArrowheadStyle read Get_BeginArrowheadStyle write Set_BeginArrowheadStyle;
    property BeginArrowheadWidth: MsoArrowheadWidth read Get_BeginArrowheadWidth write Set_BeginArrowheadWidth;
    property DashStyle: MsoLineDashStyle read Get_DashStyle write Set_DashStyle;
    property EndArrowheadLength: MsoArrowheadLength read Get_EndArrowheadLength write Set_EndArrowheadLength;
    property EndArrowheadStyle: MsoArrowheadStyle read Get_EndArrowheadStyle write Set_EndArrowheadStyle;
    property EndArrowheadWidth: MsoArrowheadWidth read Get_EndArrowheadWidth write Set_EndArrowheadWidth;
    property ForeColor: ColorFormat read Get_ForeColor;
    property Pattern: MsoPatternType read Get_Pattern write Set_Pattern;
    property Style: MsoLineStyle read Get_Style write Set_Style;
    property Transparency: Single read Get_Transparency write Set_Transparency;
    property Visible: MsoTriState read Get_Visible write Set_Visible;
    property Weight: Single read Get_Weight write Set_Weight;
  end;

// *********************************************************************//
// DispIntf:  LineFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CA-0000-0000-C000-000000000046}
// *********************************************************************//
  LineFormatDisp = dispinterface
    ['{000209CA-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property BackColor: ColorFormat readonly dispid 100;
    property BeginArrowheadLength: MsoArrowheadLength dispid 101;
    property BeginArrowheadStyle: MsoArrowheadStyle dispid 102;
    property BeginArrowheadWidth: MsoArrowheadWidth dispid 103;
    property DashStyle: MsoLineDashStyle dispid 104;
    property EndArrowheadLength: MsoArrowheadLength dispid 105;
    property EndArrowheadStyle: MsoArrowheadStyle dispid 106;
    property EndArrowheadWidth: MsoArrowheadWidth dispid 107;
    property ForeColor: ColorFormat readonly dispid 108;
    property Pattern: MsoPatternType dispid 109;
    property Style: MsoLineStyle dispid 110;
    property Transparency: Single dispid 111;
    property Visible: MsoTriState dispid 112;
    property Weight: Single dispid 113;
  end;

// *********************************************************************//
// Interface: PictureFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CB-0000-0000-C000-000000000046}
// *********************************************************************//
  PictureFormat = interface(IDispatch)
    ['{000209CB-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Brightness: Single; safecall;
    procedure Set_Brightness(prop: Single); safecall;
    function Get_ColorType: MsoPictureColorType; safecall;
    procedure Set_ColorType(prop: MsoPictureColorType); safecall;
    function Get_Contrast: Single; safecall;
    procedure Set_Contrast(prop: Single); safecall;
    function Get_CropBottom: Single; safecall;
    procedure Set_CropBottom(prop: Single); safecall;
    function Get_CropLeft: Single; safecall;
    procedure Set_CropLeft(prop: Single); safecall;
    function Get_CropRight: Single; safecall;
    procedure Set_CropRight(prop: Single); safecall;
    function Get_CropTop: Single; safecall;
    procedure Set_CropTop(prop: Single); safecall;
    function Get_TransparencyColor: Integer; safecall;
    procedure Set_TransparencyColor(prop: Integer); safecall;
    function Get_TransparentBackground: MsoTriState; safecall;
    procedure Set_TransparentBackground(prop: MsoTriState); safecall;
    procedure IncrementBrightness(Increment: Single); safecall;
    procedure IncrementContrast(Increment: Single); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Brightness: Single read Get_Brightness write Set_Brightness;
    property ColorType: MsoPictureColorType read Get_ColorType write Set_ColorType;
    property Contrast: Single read Get_Contrast write Set_Contrast;
    property CropBottom: Single read Get_CropBottom write Set_CropBottom;
    property CropLeft: Single read Get_CropLeft write Set_CropLeft;
    property CropRight: Single read Get_CropRight write Set_CropRight;
    property CropTop: Single read Get_CropTop write Set_CropTop;
    property TransparencyColor: Integer read Get_TransparencyColor write Set_TransparencyColor;
    property TransparentBackground: MsoTriState read Get_TransparentBackground write Set_TransparentBackground;
  end;

// *********************************************************************//
// DispIntf:  PictureFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CB-0000-0000-C000-000000000046}
// *********************************************************************//
  PictureFormatDisp = dispinterface
    ['{000209CB-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property Brightness: Single dispid 100;
    property ColorType: MsoPictureColorType dispid 101;
    property Contrast: Single dispid 102;
    property CropBottom: Single dispid 103;
    property CropLeft: Single dispid 104;
    property CropRight: Single dispid 105;
    property CropTop: Single dispid 106;
    property TransparencyColor: Integer dispid 107;
    property TransparentBackground: MsoTriState dispid 108;
    procedure IncrementBrightness(Increment: Single); dispid 10;
    procedure IncrementContrast(Increment: Single); dispid 11;
  end;

// *********************************************************************//
// Interface: ShadowFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CC-0000-0000-C000-000000000046}
// *********************************************************************//
  ShadowFormat = interface(IDispatch)
    ['{000209CC-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_ForeColor: ColorFormat; safecall;
    function Get_Obscured: MsoTriState; safecall;
    procedure Set_Obscured(prop: MsoTriState); safecall;
    function Get_OffsetX: Single; safecall;
    procedure Set_OffsetX(prop: Single); safecall;
    function Get_OffsetY: Single; safecall;
    procedure Set_OffsetY(prop: Single); safecall;
    function Get_Transparency: Single; safecall;
    procedure Set_Transparency(prop: Single); safecall;
    function Get_Type_: MsoShadowType; safecall;
    procedure Set_Type_(prop: MsoShadowType); safecall;
    function Get_Visible: MsoTriState; safecall;
    procedure Set_Visible(prop: MsoTriState); safecall;
    procedure IncrementOffsetX(Increment: Single); safecall;
    procedure IncrementOffsetY(Increment: Single); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property ForeColor: ColorFormat read Get_ForeColor;
    property Obscured: MsoTriState read Get_Obscured write Set_Obscured;
    property OffsetX: Single read Get_OffsetX write Set_OffsetX;
    property OffsetY: Single read Get_OffsetY write Set_OffsetY;
    property Transparency: Single read Get_Transparency write Set_Transparency;
    property Type_: MsoShadowType read Get_Type_ write Set_Type_;
    property Visible: MsoTriState read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  ShadowFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CC-0000-0000-C000-000000000046}
// *********************************************************************//
  ShadowFormatDisp = dispinterface
    ['{000209CC-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property ForeColor: ColorFormat readonly dispid 100;
    property Obscured: MsoTriState dispid 101;
    property OffsetX: Single dispid 102;
    property OffsetY: Single dispid 103;
    property Transparency: Single dispid 104;
    property Type_: MsoShadowType dispid 105;
    property Visible: MsoTriState dispid 106;
    procedure IncrementOffsetX(Increment: Single); dispid 10;
    procedure IncrementOffsetY(Increment: Single); dispid 11;
  end;

// *********************************************************************//
// Interface: ShapeNode
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CD-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeNode = interface(IDispatch)
    ['{000209CD-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_EditingType: MsoEditingType; safecall;
    function Get_Points: OleVariant; safecall;
    function Get_SegmentType: MsoSegmentType; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property EditingType: MsoEditingType read Get_EditingType;
    property Points: OleVariant read Get_Points;
    property SegmentType: MsoSegmentType read Get_SegmentType;
  end;

// *********************************************************************//
// DispIntf:  ShapeNodeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CD-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeNodeDisp = dispinterface
    ['{000209CD-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property EditingType: MsoEditingType readonly dispid 100;
    property Points: OleVariant readonly dispid 101;
    property SegmentType: MsoSegmentType readonly dispid 102;
  end;

// *********************************************************************//
// Interface: ShapeNodes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CE-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeNodes = interface(IDispatch)
    ['{000209CE-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    procedure Delete(Index: Integer); safecall;
    function Item(var Index: OleVariant): ShapeNode; safecall;
    procedure SetEditingType(Index: Integer; EditingType: MsoEditingType); safecall;
    procedure SetPosition(Index: Integer; X1: Single; Y1: Single); safecall;
    procedure SetSegmentType(Index: Integer; SegmentType: MsoSegmentType); safecall;
    procedure Insert(Index: Integer; SegmentType: MsoSegmentType; EditingType: MsoEditingType; 
                     X1: Single; Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  ShapeNodesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CE-0000-0000-C000-000000000046}
// *********************************************************************//
  ShapeNodesDisp = dispinterface
    ['{000209CE-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property Count: Integer readonly dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    procedure Delete(Index: Integer); dispid 11;
    function Item(var Index: OleVariant): ShapeNode; dispid 0;
    procedure SetEditingType(Index: Integer; EditingType: MsoEditingType); dispid 13;
    procedure SetPosition(Index: Integer; X1: Single; Y1: Single); dispid 14;
    procedure SetSegmentType(Index: Integer; SegmentType: MsoSegmentType); dispid 15;
    procedure Insert(Index: Integer; SegmentType: MsoSegmentType; EditingType: MsoEditingType; 
                     X1: Single; Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); dispid 12;
  end;

// *********************************************************************//
// Interface: TextEffectFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CF-0000-0000-C000-000000000046}
// *********************************************************************//
  TextEffectFormat = interface(IDispatch)
    ['{000209CF-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Alignment: MsoTextEffectAlignment; safecall;
    procedure Set_Alignment(prop: MsoTextEffectAlignment); safecall;
    function Get_FontBold: MsoTriState; safecall;
    procedure Set_FontBold(prop: MsoTriState); safecall;
    function Get_FontItalic: MsoTriState; safecall;
    procedure Set_FontItalic(prop: MsoTriState); safecall;
    function Get_FontName: WideString; safecall;
    procedure Set_FontName(const prop: WideString); safecall;
    function Get_FontSize: Single; safecall;
    procedure Set_FontSize(prop: Single); safecall;
    function Get_KernedPairs: MsoTriState; safecall;
    procedure Set_KernedPairs(prop: MsoTriState); safecall;
    function Get_NormalizedHeight: MsoTriState; safecall;
    procedure Set_NormalizedHeight(prop: MsoTriState); safecall;
    function Get_PresetShape: MsoPresetTextEffectShape; safecall;
    procedure Set_PresetShape(prop: MsoPresetTextEffectShape); safecall;
    function Get_PresetTextEffect: MsoPresetTextEffect; safecall;
    procedure Set_PresetTextEffect(prop: MsoPresetTextEffect); safecall;
    function Get_RotatedChars: MsoTriState; safecall;
    procedure Set_RotatedChars(prop: MsoTriState); safecall;
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_Tracking: Single; safecall;
    procedure Set_Tracking(prop: Single); safecall;
    procedure ToggleVerticalText; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Alignment: MsoTextEffectAlignment read Get_Alignment write Set_Alignment;
    property FontBold: MsoTriState read Get_FontBold write Set_FontBold;
    property FontItalic: MsoTriState read Get_FontItalic write Set_FontItalic;
    property FontName: WideString read Get_FontName write Set_FontName;
    property FontSize: Single read Get_FontSize write Set_FontSize;
    property KernedPairs: MsoTriState read Get_KernedPairs write Set_KernedPairs;
    property NormalizedHeight: MsoTriState read Get_NormalizedHeight write Set_NormalizedHeight;
    property PresetShape: MsoPresetTextEffectShape read Get_PresetShape write Set_PresetShape;
    property PresetTextEffect: MsoPresetTextEffect read Get_PresetTextEffect write Set_PresetTextEffect;
    property RotatedChars: MsoTriState read Get_RotatedChars write Set_RotatedChars;
    property Text: WideString read Get_Text write Set_Text;
    property Tracking: Single read Get_Tracking write Set_Tracking;
  end;

// *********************************************************************//
// DispIntf:  TextEffectFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209CF-0000-0000-C000-000000000046}
// *********************************************************************//
  TextEffectFormatDisp = dispinterface
    ['{000209CF-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property Alignment: MsoTextEffectAlignment dispid 100;
    property FontBold: MsoTriState dispid 101;
    property FontItalic: MsoTriState dispid 102;
    property FontName: WideString dispid 103;
    property FontSize: Single dispid 104;
    property KernedPairs: MsoTriState dispid 105;
    property NormalizedHeight: MsoTriState dispid 106;
    property PresetShape: MsoPresetTextEffectShape dispid 107;
    property PresetTextEffect: MsoPresetTextEffect dispid 108;
    property RotatedChars: MsoTriState dispid 109;
    property Text: WideString dispid 110;
    property Tracking: Single dispid 111;
    procedure ToggleVerticalText; dispid 10;
  end;

// *********************************************************************//
// Interface: ThreeDFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D0-0000-0000-C000-000000000046}
// *********************************************************************//
  ThreeDFormat = interface(IDispatch)
    ['{000209D0-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Depth: Single; safecall;
    procedure Set_Depth(prop: Single); safecall;
    function Get_ExtrusionColor: ColorFormat; safecall;
    function Get_ExtrusionColorType: MsoExtrusionColorType; safecall;
    procedure Set_ExtrusionColorType(prop: MsoExtrusionColorType); safecall;
    function Get_Perspective: MsoTriState; safecall;
    procedure Set_Perspective(prop: MsoTriState); safecall;
    function Get_PresetExtrusionDirection: MsoPresetExtrusionDirection; safecall;
    function Get_PresetLightingDirection: MsoPresetLightingDirection; safecall;
    procedure Set_PresetLightingDirection(prop: MsoPresetLightingDirection); safecall;
    function Get_PresetLightingSoftness: MsoPresetLightingSoftness; safecall;
    procedure Set_PresetLightingSoftness(prop: MsoPresetLightingSoftness); safecall;
    function Get_PresetMaterial: MsoPresetMaterial; safecall;
    procedure Set_PresetMaterial(prop: MsoPresetMaterial); safecall;
    function Get_PresetThreeDFormat: MsoPresetThreeDFormat; safecall;
    function Get_RotationX: Single; safecall;
    procedure Set_RotationX(prop: Single); safecall;
    function Get_RotationY: Single; safecall;
    procedure Set_RotationY(prop: Single); safecall;
    function Get_Visible: MsoTriState; safecall;
    procedure Set_Visible(prop: MsoTriState); safecall;
    procedure IncrementRotationX(Increment: Single); safecall;
    procedure IncrementRotationY(Increment: Single); safecall;
    procedure ResetRotation; safecall;
    procedure SetExtrusionDirection(PresetExtrusionDirection: MsoPresetExtrusionDirection); safecall;
    procedure SetThreeDFormat(PresetThreeDFormat: MsoPresetThreeDFormat); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Depth: Single read Get_Depth write Set_Depth;
    property ExtrusionColor: ColorFormat read Get_ExtrusionColor;
    property ExtrusionColorType: MsoExtrusionColorType read Get_ExtrusionColorType write Set_ExtrusionColorType;
    property Perspective: MsoTriState read Get_Perspective write Set_Perspective;
    property PresetExtrusionDirection: MsoPresetExtrusionDirection read Get_PresetExtrusionDirection;
    property PresetLightingDirection: MsoPresetLightingDirection read Get_PresetLightingDirection write Set_PresetLightingDirection;
    property PresetLightingSoftness: MsoPresetLightingSoftness read Get_PresetLightingSoftness write Set_PresetLightingSoftness;
    property PresetMaterial: MsoPresetMaterial read Get_PresetMaterial write Set_PresetMaterial;
    property PresetThreeDFormat: MsoPresetThreeDFormat read Get_PresetThreeDFormat;
    property RotationX: Single read Get_RotationX write Set_RotationX;
    property RotationY: Single read Get_RotationY write Set_RotationY;
    property Visible: MsoTriState read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  ThreeDFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D0-0000-0000-C000-000000000046}
// *********************************************************************//
  ThreeDFormatDisp = dispinterface
    ['{000209D0-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1;
    property Depth: Single dispid 100;
    property ExtrusionColor: ColorFormat readonly dispid 101;
    property ExtrusionColorType: MsoExtrusionColorType dispid 102;
    property Perspective: MsoTriState dispid 103;
    property PresetExtrusionDirection: MsoPresetExtrusionDirection readonly dispid 104;
    property PresetLightingDirection: MsoPresetLightingDirection dispid 105;
    property PresetLightingSoftness: MsoPresetLightingSoftness dispid 106;
    property PresetMaterial: MsoPresetMaterial dispid 107;
    property PresetThreeDFormat: MsoPresetThreeDFormat readonly dispid 108;
    property RotationX: Single dispid 109;
    property RotationY: Single dispid 110;
    property Visible: MsoTriState dispid 111;
    procedure IncrementRotationX(Increment: Single); dispid 10;
    procedure IncrementRotationY(Increment: Single); dispid 11;
    procedure ResetRotation; dispid 12;
    procedure SetExtrusionDirection(PresetExtrusionDirection: MsoPresetExtrusionDirection); dispid 14;
    procedure SetThreeDFormat(PresetThreeDFormat: MsoPresetThreeDFormat); dispid 13;
  end;

// *********************************************************************//
// DispIntf:  ApplicationEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209F7-0000-0000-C000-000000000046}
// *********************************************************************//
  ApplicationEvents = dispinterface
    ['{000209F7-0000-0000-C000-000000000046}']
  end;

// *********************************************************************//
// DispIntf:  ApplicationEvents2
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209FE-0000-0000-C000-000000000046}
// *********************************************************************//
  ApplicationEvents2 = dispinterface
    ['{000209FE-0000-0000-C000-000000000046}']
    procedure QueryInterface_(var riid: {??TGUID} OleVariant; out ppvObj: {??Pointer} OleVariant); dispid 1610612736;
    function AddRef_: UINT; dispid 1610612737;
    function Release_: UINT; dispid 1610612738;
    procedure GetTypeInfoCount_(out pctinfo: SYSUINT); dispid 1610678272;
    procedure GetTypeInfo_(itinfo: SYSUINT; lcid: UINT; out pptinfo: {??Pointer} OleVariant); dispid 1610678273;
    procedure GetIDsOfNames_(var riid: {??TGUID} OleVariant; rgszNames: {??PPShortint1} OleVariant; 
                             cNames: SYSUINT; lcid: UINT; out rgdispid: Integer); dispid 1610678274;
    procedure Invoke_(dispidMember: Integer; var riid: {??TGUID} OleVariant; lcid: UINT; 
                      wFlags: {??Word} OleVariant; var pdispparams: {??DISPPARAMS} OleVariant; 
                      out pvarResult: OleVariant; out pexcepinfo: {??EXCEPINFO} OleVariant; 
                      out puArgErr: SYSUINT); dispid 1610678275;
    procedure Startup; dispid 1;
    procedure Quit; dispid 2;
    procedure DocumentChange; dispid 3;
    procedure DocumentOpen(const Doc: Document); dispid 4;
    procedure DocumentBeforeClose(const Doc: Document; var Cancel: WordBool); dispid 6;
    procedure DocumentBeforePrint(const Doc: Document; var Cancel: WordBool); dispid 7;
    procedure DocumentBeforeSave(const Doc: Document; var SaveAsUI: WordBool; var Cancel: WordBool); dispid 8;
    procedure NewDocument(const Doc: Document); dispid 9;
    procedure WindowActivate(const Doc: Document; const Wn: Window_); dispid 10;
    procedure WindowDeactivate(const Doc: Document; const Wn: Window_); dispid 11;
    procedure WindowSelectionChange(const Sel: Selection); dispid 12;
    procedure WindowBeforeRightClick(const Sel: Selection; var Cancel: WordBool); dispid 13;
    procedure WindowBeforeDoubleClick(const Sel: Selection; var Cancel: WordBool); dispid 14;
  end;

// *********************************************************************//
// DispIntf:  DocumentEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209F6-0000-0000-C000-000000000046}
// *********************************************************************//
  DocumentEvents = dispinterface
    ['{000209F6-0000-0000-C000-000000000046}']
    procedure New; dispid 4;
    procedure Open; dispid 5;
    procedure Close; dispid 6;
  end;

// *********************************************************************//
// DispIntf:  OCXEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209F3-0000-0000-C000-000000000046}
// *********************************************************************//
  OCXEvents = dispinterface
    ['{000209F3-0000-0000-C000-000000000046}']
    procedure GotFocus; dispid -2147417888;
    procedure LostFocus; dispid -2147417887;
  end;

// *********************************************************************//
// Interface: IApplicationEvents
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209F7-0001-0000-C000-000000000046}
// *********************************************************************//
  IApplicationEvents = interface(IDispatch)
    ['{000209F7-0001-0000-C000-000000000046}']
    procedure Startup; safecall;
    procedure Quit; safecall;
    procedure DocumentChange; safecall;
  end;

// *********************************************************************//
// DispIntf:  IApplicationEventsDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209F7-0001-0000-C000-000000000046}
// *********************************************************************//
  IApplicationEventsDisp = dispinterface
    ['{000209F7-0001-0000-C000-000000000046}']
    procedure Startup; dispid 1;
    procedure Quit; dispid 2;
    procedure DocumentChange; dispid 3;
  end;

// *********************************************************************//
// Interface: IApplicationEvents2
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209FE-0001-0000-C000-000000000046}
// *********************************************************************//
  IApplicationEvents2 = interface(IDispatch)
    ['{000209FE-0001-0000-C000-000000000046}']
    procedure Startup; safecall;
    procedure Quit; safecall;
    procedure DocumentChange; safecall;
    procedure DocumentOpen(const Doc: Document); safecall;
    procedure DocumentBeforeClose(const Doc: Document; var Cancel: WordBool); safecall;
    procedure DocumentBeforePrint(const Doc: Document; var Cancel: WordBool); safecall;
    procedure DocumentBeforeSave(const Doc: Document; var SaveAsUI: WordBool; var Cancel: WordBool); safecall;
    procedure NewDocument(const Doc: Document); safecall;
    procedure WindowActivate(const Doc: Document; const Wn: Window_); safecall;
    procedure WindowDeactivate(const Doc: Document; const Wn: Window_); safecall;
    procedure WindowSelectionChange(const Sel: Selection); safecall;
    procedure WindowBeforeRightClick(const Sel: Selection; var Cancel: WordBool); safecall;
    procedure WindowBeforeDoubleClick(const Sel: Selection; var Cancel: WordBool); safecall;
  end;

// *********************************************************************//
// DispIntf:  IApplicationEvents2Disp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209FE-0001-0000-C000-000000000046}
// *********************************************************************//
  IApplicationEvents2Disp = dispinterface
    ['{000209FE-0001-0000-C000-000000000046}']
    procedure Startup; dispid 1;
    procedure Quit; dispid 2;
    procedure DocumentChange; dispid 3;
    procedure DocumentOpen(const Doc: Document); dispid 4;
    procedure DocumentBeforeClose(const Doc: Document; var Cancel: WordBool); dispid 6;
    procedure DocumentBeforePrint(const Doc: Document; var Cancel: WordBool); dispid 7;
    procedure DocumentBeforeSave(const Doc: Document; var SaveAsUI: WordBool; var Cancel: WordBool); dispid 8;
    procedure NewDocument(const Doc: Document); dispid 9;
    procedure WindowActivate(const Doc: Document; const Wn: Window_); dispid 10;
    procedure WindowDeactivate(const Doc: Document; const Wn: Window_); dispid 11;
    procedure WindowSelectionChange(const Sel: Selection); dispid 12;
    procedure WindowBeforeRightClick(const Sel: Selection; var Cancel: WordBool); dispid 13;
    procedure WindowBeforeDoubleClick(const Sel: Selection; var Cancel: WordBool); dispid 14;
  end;

// *********************************************************************//
// Interface: EmailAuthor
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D7-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailAuthor = interface(IDispatch)
    ['{000209D7-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Style: Style; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Style: Style read Get_Style;
  end;

// *********************************************************************//
// DispIntf:  EmailAuthorDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209D7-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailAuthorDisp = dispinterface
    ['{000209D7-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Style: Style readonly dispid 103;
  end;

// *********************************************************************//
// Interface: EmailOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DB-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailOptions = interface(IDispatch)
    ['{000209DB-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_UseThemeStyle: WordBool; safecall;
    procedure Set_UseThemeStyle(prop: WordBool); safecall;
    function Get_MarkCommentsWith: WideString; safecall;
    procedure Set_MarkCommentsWith(const prop: WideString); safecall;
    function Get_MarkComments: WordBool; safecall;
    procedure Set_MarkComments(prop: WordBool); safecall;
    function Get_EmailSignature: EmailSignature; safecall;
    function Get_ComposeStyle: Style; safecall;
    function Get_ReplyStyle: Style; safecall;
    function Get_ThemeName: WideString; safecall;
    procedure Set_ThemeName(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property UseThemeStyle: WordBool read Get_UseThemeStyle write Set_UseThemeStyle;
    property MarkCommentsWith: WideString read Get_MarkCommentsWith write Set_MarkCommentsWith;
    property MarkComments: WordBool read Get_MarkComments write Set_MarkComments;
    property EmailSignature: EmailSignature read Get_EmailSignature;
    property ComposeStyle: Style read Get_ComposeStyle;
    property ReplyStyle: Style read Get_ReplyStyle;
    property ThemeName: WideString read Get_ThemeName write Set_ThemeName;
  end;

// *********************************************************************//
// DispIntf:  EmailOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DB-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailOptionsDisp = dispinterface
    ['{000209DB-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 100;
    property Creator: Integer readonly dispid 101;
    property Parent: IDispatch readonly dispid 102;
    property UseThemeStyle: WordBool dispid 103;
    property MarkCommentsWith: WideString dispid 106;
    property MarkComments: WordBool dispid 107;
    property EmailSignature: EmailSignature readonly dispid 108;
    property ComposeStyle: Style readonly dispid 109;
    property ReplyStyle: Style readonly dispid 110;
    property ThemeName: WideString dispid 114;
  end;

// *********************************************************************//
// Interface: EmailSignature
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DC-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailSignature = interface(IDispatch)
    ['{000209DC-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_NewMessageSignature: WideString; safecall;
    procedure Set_NewMessageSignature(const prop: WideString); safecall;
    function Get_ReplyMessageSignature: WideString; safecall;
    procedure Set_ReplyMessageSignature(const prop: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property NewMessageSignature: WideString read Get_NewMessageSignature write Set_NewMessageSignature;
    property ReplyMessageSignature: WideString read Get_ReplyMessageSignature write Set_ReplyMessageSignature;
  end;

// *********************************************************************//
// DispIntf:  EmailSignatureDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DC-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailSignatureDisp = dispinterface
    ['{000209DC-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 100;
    property Creator: Integer readonly dispid 101;
    property Parent: IDispatch readonly dispid 102;
    property NewMessageSignature: WideString dispid 103;
    property ReplyMessageSignature: WideString dispid 104;
  end;

// *********************************************************************//
// Interface: Email
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DD-0000-0000-C000-000000000046}
// *********************************************************************//
  Email = interface(IDispatch)
    ['{000209DD-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_CurrentEmailAuthor: EmailAuthor; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property CurrentEmailAuthor: EmailAuthor read Get_CurrentEmailAuthor;
  end;

// *********************************************************************//
// DispIntf:  EmailDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DD-0000-0000-C000-000000000046}
// *********************************************************************//
  EmailDisp = dispinterface
    ['{000209DD-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 100;
    property Creator: Integer readonly dispid 101;
    property Parent: IDispatch readonly dispid 102;
    property CurrentEmailAuthor: EmailAuthor readonly dispid 105;
  end;

// *********************************************************************//
// Interface: HorizontalLineFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DE-0000-0000-C000-000000000046}
// *********************************************************************//
  HorizontalLineFormat = interface(IDispatch)
    ['{000209DE-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_PercentWidth: Single; safecall;
    procedure Set_PercentWidth(prop: Single); safecall;
    function Get_NoShade: WordBool; safecall;
    procedure Set_NoShade(prop: WordBool); safecall;
    function Get_Alignment: WdHorizontalLineAlignment; safecall;
    procedure Set_Alignment(prop: WdHorizontalLineAlignment); safecall;
    function Get_WidthType: WdHorizontalLineWidthType; safecall;
    procedure Set_WidthType(prop: WdHorizontalLineWidthType); safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property PercentWidth: Single read Get_PercentWidth write Set_PercentWidth;
    property NoShade: WordBool read Get_NoShade write Set_NoShade;
    property Alignment: WdHorizontalLineAlignment read Get_Alignment write Set_Alignment;
    property WidthType: WdHorizontalLineWidthType read Get_WidthType write Set_WidthType;
  end;

// *********************************************************************//
// DispIntf:  HorizontalLineFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DE-0000-0000-C000-000000000046}
// *********************************************************************//
  HorizontalLineFormatDisp = dispinterface
    ['{000209DE-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property PercentWidth: Single dispid 2;
    property NoShade: WordBool dispid 3;
    property Alignment: WdHorizontalLineAlignment dispid 4;
    property WidthType: WdHorizontalLineWidthType dispid 5;
  end;

// *********************************************************************//
// Interface: Frameset
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E2-0000-0000-C000-000000000046}
// *********************************************************************//
  Frameset = interface(IDispatch)
    ['{000209E2-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_ParentFrameset: Frameset; safecall;
    function Get_Type_: WdFramesetType; safecall;
    function Get_WidthType: WdFramesetSizeType; safecall;
    procedure Set_WidthType(prop: WdFramesetSizeType); safecall;
    function Get_HeightType: WdFramesetSizeType; safecall;
    procedure Set_HeightType(prop: WdFramesetSizeType); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(prop: Integer); safecall;
    function Get_ChildFramesetCount: Integer; safecall;
    function Get_ChildFramesetItem(Index: Integer): Frameset; safecall;
    function Get_FramesetBorderWidth: Single; safecall;
    procedure Set_FramesetBorderWidth(prop: Single); safecall;
    function Get_FramesetBorderColor: WdColor; safecall;
    procedure Set_FramesetBorderColor(prop: WdColor); safecall;
    function Get_FrameScrollbarType: WdScrollbarType; safecall;
    procedure Set_FrameScrollbarType(prop: WdScrollbarType); safecall;
    function Get_FrameResizable: WordBool; safecall;
    procedure Set_FrameResizable(prop: WordBool); safecall;
    function Get_FrameName: WideString; safecall;
    procedure Set_FrameName(const prop: WideString); safecall;
    function Get_FrameDisplayBorders: WordBool; safecall;
    procedure Set_FrameDisplayBorders(prop: WordBool); safecall;
    function Get_FrameDefaultURL: WideString; safecall;
    procedure Set_FrameDefaultURL(const prop: WideString); safecall;
    function Get_FrameLinkToFile: WordBool; safecall;
    procedure Set_FrameLinkToFile(prop: WordBool); safecall;
    function AddNewFrame(Where: WdFramesetNewFrameLocation): Frameset; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property ParentFrameset: Frameset read Get_ParentFrameset;
    property Type_: WdFramesetType read Get_Type_;
    property WidthType: WdFramesetSizeType read Get_WidthType write Set_WidthType;
    property HeightType: WdFramesetSizeType read Get_HeightType write Set_HeightType;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property ChildFramesetCount: Integer read Get_ChildFramesetCount;
    property ChildFramesetItem[Index: Integer]: Frameset read Get_ChildFramesetItem;
    property FramesetBorderWidth: Single read Get_FramesetBorderWidth write Set_FramesetBorderWidth;
    property FramesetBorderColor: WdColor read Get_FramesetBorderColor write Set_FramesetBorderColor;
    property FrameScrollbarType: WdScrollbarType read Get_FrameScrollbarType write Set_FrameScrollbarType;
    property FrameResizable: WordBool read Get_FrameResizable write Set_FrameResizable;
    property FrameName: WideString read Get_FrameName write Set_FrameName;
    property FrameDisplayBorders: WordBool read Get_FrameDisplayBorders write Set_FrameDisplayBorders;
    property FrameDefaultURL: WideString read Get_FrameDefaultURL write Set_FrameDefaultURL;
    property FrameLinkToFile: WordBool read Get_FrameLinkToFile write Set_FrameLinkToFile;
  end;

// *********************************************************************//
// DispIntf:  FramesetDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E2-0000-0000-C000-000000000046}
// *********************************************************************//
  FramesetDisp = dispinterface
    ['{000209E2-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property ParentFrameset: Frameset readonly dispid 1003;
    property Type_: WdFramesetType readonly dispid 0;
    property WidthType: WdFramesetSizeType dispid 1;
    property HeightType: WdFramesetSizeType dispid 2;
    property Width: Integer dispid 3;
    property Height: Integer dispid 4;
    property ChildFramesetCount: Integer readonly dispid 5;
    property ChildFramesetItem[Index: Integer]: Frameset readonly dispid 6;
    property FramesetBorderWidth: Single dispid 20;
    property FramesetBorderColor: WdColor dispid 21;
    property FrameScrollbarType: WdScrollbarType dispid 30;
    property FrameResizable: WordBool dispid 31;
    property FrameName: WideString dispid 34;
    property FrameDisplayBorders: WordBool dispid 35;
    property FrameDefaultURL: WideString dispid 36;
    property FrameLinkToFile: WordBool dispid 37;
    function AddNewFrame(Where: WdFramesetNewFrameLocation): Frameset; dispid 50;
    procedure Delete; dispid 51;
  end;

// *********************************************************************//
// Interface: DefaultWebOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E3-0000-0000-C000-000000000046}
// *********************************************************************//
  DefaultWebOptions = interface(IDispatch)
    ['{000209E3-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_OptimizeForBrowser: WordBool; safecall;
    procedure Set_OptimizeForBrowser(prop: WordBool); safecall;
    function Get_BrowserLevel: WdBrowserLevel; safecall;
    procedure Set_BrowserLevel(prop: WdBrowserLevel); safecall;
    function Get_RelyOnCSS: WordBool; safecall;
    procedure Set_RelyOnCSS(prop: WordBool); safecall;
    function Get_OrganizeInFolder: WordBool; safecall;
    procedure Set_OrganizeInFolder(prop: WordBool); safecall;
    function Get_UpdateLinksOnSave: WordBool; safecall;
    procedure Set_UpdateLinksOnSave(prop: WordBool); safecall;
    function Get_UseLongFileNames: WordBool; safecall;
    procedure Set_UseLongFileNames(prop: WordBool); safecall;
    function Get_CheckIfOfficeIsHTMLEditor: WordBool; safecall;
    procedure Set_CheckIfOfficeIsHTMLEditor(prop: WordBool); safecall;
    function Get_CheckIfWordIsDefaultHTMLEditor: WordBool; safecall;
    procedure Set_CheckIfWordIsDefaultHTMLEditor(prop: WordBool); safecall;
    function Get_RelyOnVML: WordBool; safecall;
    procedure Set_RelyOnVML(prop: WordBool); safecall;
    function Get_AllowPNG: WordBool; safecall;
    procedure Set_AllowPNG(prop: WordBool); safecall;
    function Get_ScreenSize: MsoScreenSize; safecall;
    procedure Set_ScreenSize(prop: MsoScreenSize); safecall;
    function Get_PixelsPerInch: Integer; safecall;
    procedure Set_PixelsPerInch(prop: Integer); safecall;
    function Get_Encoding: MsoEncoding; safecall;
    procedure Set_Encoding(prop: MsoEncoding); safecall;
    function Get_AlwaysSaveInDefaultEncoding: WordBool; safecall;
    procedure Set_AlwaysSaveInDefaultEncoding(prop: WordBool); safecall;
    function Get_Fonts: WebPageFonts; safecall;
    function Get_FolderSuffix: WideString; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property OptimizeForBrowser: WordBool read Get_OptimizeForBrowser write Set_OptimizeForBrowser;
    property BrowserLevel: WdBrowserLevel read Get_BrowserLevel write Set_BrowserLevel;
    property RelyOnCSS: WordBool read Get_RelyOnCSS write Set_RelyOnCSS;
    property OrganizeInFolder: WordBool read Get_OrganizeInFolder write Set_OrganizeInFolder;
    property UpdateLinksOnSave: WordBool read Get_UpdateLinksOnSave write Set_UpdateLinksOnSave;
    property UseLongFileNames: WordBool read Get_UseLongFileNames write Set_UseLongFileNames;
    property CheckIfOfficeIsHTMLEditor: WordBool read Get_CheckIfOfficeIsHTMLEditor write Set_CheckIfOfficeIsHTMLEditor;
    property CheckIfWordIsDefaultHTMLEditor: WordBool read Get_CheckIfWordIsDefaultHTMLEditor write Set_CheckIfWordIsDefaultHTMLEditor;
    property RelyOnVML: WordBool read Get_RelyOnVML write Set_RelyOnVML;
    property AllowPNG: WordBool read Get_AllowPNG write Set_AllowPNG;
    property ScreenSize: MsoScreenSize read Get_ScreenSize write Set_ScreenSize;
    property PixelsPerInch: Integer read Get_PixelsPerInch write Set_PixelsPerInch;
    property Encoding: MsoEncoding read Get_Encoding write Set_Encoding;
    property AlwaysSaveInDefaultEncoding: WordBool read Get_AlwaysSaveInDefaultEncoding write Set_AlwaysSaveInDefaultEncoding;
    property Fonts: WebPageFonts read Get_Fonts;
    property FolderSuffix: WideString read Get_FolderSuffix;
  end;

// *********************************************************************//
// DispIntf:  DefaultWebOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E3-0000-0000-C000-000000000046}
// *********************************************************************//
  DefaultWebOptionsDisp = dispinterface
    ['{000209E3-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property OptimizeForBrowser: WordBool dispid 1;
    property BrowserLevel: WdBrowserLevel dispid 2;
    property RelyOnCSS: WordBool dispid 3;
    property OrganizeInFolder: WordBool dispid 4;
    property UpdateLinksOnSave: WordBool dispid 5;
    property UseLongFileNames: WordBool dispid 6;
    property CheckIfOfficeIsHTMLEditor: WordBool dispid 7;
    property CheckIfWordIsDefaultHTMLEditor: WordBool dispid 8;
    property RelyOnVML: WordBool dispid 9;
    property AllowPNG: WordBool dispid 10;
    property ScreenSize: MsoScreenSize dispid 11;
    property PixelsPerInch: Integer dispid 12;
    property Encoding: MsoEncoding dispid 13;
    property AlwaysSaveInDefaultEncoding: WordBool dispid 14;
    property Fonts: WebPageFonts readonly dispid 15;
    property FolderSuffix: WideString readonly dispid 16;
  end;

// *********************************************************************//
// Interface: WebOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E4-0000-0000-C000-000000000046}
// *********************************************************************//
  WebOptions = interface(IDispatch)
    ['{000209E4-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_OptimizeForBrowser: WordBool; safecall;
    procedure Set_OptimizeForBrowser(prop: WordBool); safecall;
    function Get_BrowserLevel: WdBrowserLevel; safecall;
    procedure Set_BrowserLevel(prop: WdBrowserLevel); safecall;
    function Get_RelyOnCSS: WordBool; safecall;
    procedure Set_RelyOnCSS(prop: WordBool); safecall;
    function Get_OrganizeInFolder: WordBool; safecall;
    procedure Set_OrganizeInFolder(prop: WordBool); safecall;
    function Get_UseLongFileNames: WordBool; safecall;
    procedure Set_UseLongFileNames(prop: WordBool); safecall;
    function Get_RelyOnVML: WordBool; safecall;
    procedure Set_RelyOnVML(prop: WordBool); safecall;
    function Get_AllowPNG: WordBool; safecall;
    procedure Set_AllowPNG(prop: WordBool); safecall;
    function Get_ScreenSize: MsoScreenSize; safecall;
    procedure Set_ScreenSize(prop: MsoScreenSize); safecall;
    function Get_PixelsPerInch: Integer; safecall;
    procedure Set_PixelsPerInch(prop: Integer); safecall;
    function Get_Encoding: MsoEncoding; safecall;
    procedure Set_Encoding(prop: MsoEncoding); safecall;
    function Get_FolderSuffix: WideString; safecall;
    procedure UseDefaultFolderSuffix; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property OptimizeForBrowser: WordBool read Get_OptimizeForBrowser write Set_OptimizeForBrowser;
    property BrowserLevel: WdBrowserLevel read Get_BrowserLevel write Set_BrowserLevel;
    property RelyOnCSS: WordBool read Get_RelyOnCSS write Set_RelyOnCSS;
    property OrganizeInFolder: WordBool read Get_OrganizeInFolder write Set_OrganizeInFolder;
    property UseLongFileNames: WordBool read Get_UseLongFileNames write Set_UseLongFileNames;
    property RelyOnVML: WordBool read Get_RelyOnVML write Set_RelyOnVML;
    property AllowPNG: WordBool read Get_AllowPNG write Set_AllowPNG;
    property ScreenSize: MsoScreenSize read Get_ScreenSize write Set_ScreenSize;
    property PixelsPerInch: Integer read Get_PixelsPerInch write Set_PixelsPerInch;
    property Encoding: MsoEncoding read Get_Encoding write Set_Encoding;
    property FolderSuffix: WideString read Get_FolderSuffix;
  end;

// *********************************************************************//
// DispIntf:  WebOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E4-0000-0000-C000-000000000046}
// *********************************************************************//
  WebOptionsDisp = dispinterface
    ['{000209E4-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property OptimizeForBrowser: WordBool dispid 1;
    property BrowserLevel: WdBrowserLevel dispid 2;
    property RelyOnCSS: WordBool dispid 3;
    property OrganizeInFolder: WordBool dispid 4;
    property UseLongFileNames: WordBool dispid 5;
    property RelyOnVML: WordBool dispid 6;
    property AllowPNG: WordBool dispid 7;
    property ScreenSize: MsoScreenSize dispid 8;
    property PixelsPerInch: Integer dispid 9;
    property Encoding: MsoEncoding dispid 10;
    property FolderSuffix: WideString readonly dispid 11;
    procedure UseDefaultFolderSuffix; dispid 101;
  end;

// *********************************************************************//
// Interface: OtherCorrectionsExceptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DF-0000-0000-C000-000000000046}
// *********************************************************************//
  OtherCorrectionsExceptions = interface(IDispatch)
    ['{000209DF-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): OtherCorrectionsException; safecall;
    function Add(const Name: WideString): OtherCorrectionsException; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  OtherCorrectionsExceptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209DF-0000-0000-C000-000000000046}
// *********************************************************************//
  OtherCorrectionsExceptionsDisp = dispinterface
    ['{000209DF-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): OtherCorrectionsException; dispid 0;
    function Add(const Name: WideString): OtherCorrectionsException; dispid 101;
  end;

// *********************************************************************//
// Interface: OtherCorrectionsException
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E1-0000-0000-C000-000000000046}
// *********************************************************************//
  OtherCorrectionsException = interface(IDispatch)
    ['{000209E1-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  OtherCorrectionsExceptionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209E1-0000-0000-C000-000000000046}
// *********************************************************************//
  OtherCorrectionsExceptionDisp = dispinterface
    ['{000209E1-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Index: Integer readonly dispid 1;
    property Name: WideString readonly dispid 2;
    procedure Delete; dispid 101;
  end;

(*  CoGlobal = class
    class function Create: _Global;
    class function CreateRemote(const MachineName: string): _Global;
  end;
*)
  CoDocument = class
    class function Create: _Document;
    class function CreateRemote(const MachineName: string): _Document;
  end;

  CoFont = class
    class function Create: _Font;
    class function CreateRemote(const MachineName: string): _Font;
  end;

  CoParagraphFormat = class
    class function Create: _ParagraphFormat;
    class function CreateRemote(const MachineName: string): _ParagraphFormat;
  end;

  CoOLEControl = class
    class function Create: _OLEControl;
    class function CreateRemote(const MachineName: string): _OLEControl;
  end;

  CoLetterContent = class
    class function Create: _LetterContent;
    class function CreateRemote(const MachineName: string): _LetterContent;
  end;

  CoApplication_ = class
    class function Create: _Application;
    class function CreateRemote(const MachineName: string): _Application;
  end;

implementation

uses ComObj;

(*class function CoGlobal.Create: _Global;
begin
  Result := CreateComObject(CLASS_Global) as _Global;
end;

class function CoGlobal.CreateRemote(const MachineName: string): _Global;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Global) as _Global;
end;
*)
class function CoDocument.Create: _Document;
begin
  Result := CreateComObject(CLASS_Document) as _Document;
end;

class function CoDocument.CreateRemote(const MachineName: string): _Document;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Document) as _Document;
end;

class function CoFont.Create: _Font;
begin
  Result := CreateComObject(CLASS_Font) as _Font;
end;

class function CoFont.CreateRemote(const MachineName: string): _Font;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Font) as _Font;
end;

class function CoParagraphFormat.Create: _ParagraphFormat;
begin
  Result := CreateComObject(CLASS_ParagraphFormat) as _ParagraphFormat;
end;

class function CoParagraphFormat.CreateRemote(const MachineName: string): _ParagraphFormat;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ParagraphFormat) as _ParagraphFormat;
end;

class function CoOLEControl.Create: _OLEControl;
begin
  Result := CreateComObject(CLASS_OLEControl) as _OLEControl;
end;

class function CoOLEControl.CreateRemote(const MachineName: string): _OLEControl;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OLEControl) as _OLEControl;
end;

class function CoLetterContent.Create: _LetterContent;
begin
  Result := CreateComObject(CLASS_LetterContent) as _LetterContent;
end;

class function CoLetterContent.CreateRemote(const MachineName: string): _LetterContent;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_LetterContent) as _LetterContent;
end;

class function CoApplication_.Create: _Application;
begin
  Result := CreateComObject(CLASS_Application_) as _Application;
end;

class function CoApplication_.CreateRemote(const MachineName: string): _Application;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Application_) as _Application;
end;

end.
