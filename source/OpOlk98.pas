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

  {$WARN SYMBOL_PLATFORM OFF}

{$IFDEF CBuilder}
  {$Warnings Off}
{$ENDIF}

unit OpOlk98;

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
// File generated on 3/10/99 2:27:06 PM from Type Library described below.

// ************************************************************************ //
// Type Lib: C:\Program Files\MSOffice\Office\Msoutl85.olb
// IID\LCID: {00062FFF-0000-0000-C000-000000000046}\0
// Helpfile: C:\Program Files\MSOffice\Office\VBAOUTL.HLP
// HelpString: Outlook 98 Type Library
// Version:    8.5
// ************************************************************************ //

interface

uses Windows, ActiveX, Classes, Graphics, OleCtrls, StdVCL, OpShared;

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:      //
//   Type Libraries     : LIBID_xxxx                                    //
//   CoClasses          : CLASS_xxxx                                    //
//   DISPInterfaces     : DIID_xxxx                                     //
//   Non-DISP interfaces: IID_xxxx                                      //
// *********************************************************************//
const
  LIBID_Outlook: TGUID = '{00062FFF-0000-0000-C000-000000000046}';
  IID_Action: TGUID = '{00063043-0000-0000-C000-000000000046}';
  IID__Application: TGUID = '{00063001-0000-0000-C000-000000000046}';
  CLASS_Application_: TGUID = '{0006F033-0000-0000-C000-000000000046}';
  IID_NameSpace: TGUID = '{00063002-0000-0000-C000-000000000046}';
  IID_Recipient: TGUID = '{00063045-0000-0000-C000-000000000046}';
  IID_AddressEntry: TGUID = '{0006304B-0000-0000-C000-000000000046}';
  IID_AddressEntries: TGUID = '{0006304A-0000-0000-C000-000000000046}';
  IID_Folders: TGUID = '{00063040-0000-0000-C000-000000000046}';
  IID_MAPIFolder: TGUID = '{00063006-0000-0000-C000-000000000046}';
  IID_Items: TGUID = '{00063041-0000-0000-C000-000000000046}';
  IID_Explorer: TGUID = '{00063003-0000-0000-C000-000000000046}';
  IID_AddressLists: TGUID = '{00063048-0000-0000-C000-000000000046}';
  IID_AddressList: TGUID = '{00063049-0000-0000-C000-000000000046}';
  IID_Inspector: TGUID = '{00063005-0000-0000-C000-000000000046}';
  IID_Actions: TGUID = '{0006303E-0000-0000-C000-000000000046}';
  IID__AppointmentItem: TGUID = '{00063033-0000-0000-C000-000000000046}';
  IID_Attachments: TGUID = '{0006303C-0000-0000-C000-000000000046}';
  IID_Attachment: TGUID = '{00063007-0000-0000-C000-000000000046}';
  IID_FormDescription: TGUID = '{00063046-0000-0000-C000-000000000046}';
  IID_UserProperties: TGUID = '{0006303D-0000-0000-C000-000000000046}';
  IID_UserProperty: TGUID = '{00063042-0000-0000-C000-000000000046}';
  IID_Recipients: TGUID = '{0006303B-0000-0000-C000-000000000046}';
  IID_ItemEvents: TGUID = '{0006303A-0000-0000-C000-000000000046}';
  IID__MailItem: TGUID = '{00063034-0000-0000-C000-000000000046}';
  CLASS_MailItem: TGUID = '{00061033-0000-0000-C000-000000000046}';
  IID_RecurrencePattern: TGUID = '{00063044-0000-0000-C000-000000000046}';
  IID_Exceptions: TGUID = '{0006304C-0000-0000-C000-000000000046}';
  IID_Exception: TGUID = '{0006304D-0000-0000-C000-000000000046}';
  CLASS_AppointmentItem: TGUID = '{00061030-0000-0000-C000-000000000046}';
  IID__MeetingItem: TGUID = '{00063062-0000-0000-C000-000000000046}';
  CLASS_MeetingItem: TGUID = '{00061036-0000-0000-C000-000000000046}';
  IID__ContactItem: TGUID = '{00063021-0000-0000-C000-000000000046}';
  IID__DocumentItem: TGUID = '{00063020-0000-0000-C000-000000000046}';
  IID__JournalItem: TGUID = '{00063022-0000-0000-C000-000000000046}';
  IID__NoteItem: TGUID = '{00063025-0000-0000-C000-000000000046}';
  IID_Pages: TGUID = '{0006303F-0000-0000-C000-000000000046}';
  IID__PostItem: TGUID = '{00063024-0000-0000-C000-000000000046}';
  IID__RemoteItem: TGUID = '{00063023-0000-0000-C000-000000000046}';
  IID__ReportItem: TGUID = '{00063026-0000-0000-C000-000000000046}';
  IID__TaskItem: TGUID = '{00063035-0000-0000-C000-000000000046}';
  CLASS_TaskItem: TGUID = '{00061032-0000-0000-C000-000000000046}';
  IID__TaskRequestAcceptItem: TGUID = '{00063038-0000-0000-C000-000000000046}';
  IID__TaskRequestDeclineItem: TGUID = '{00063039-0000-0000-C000-000000000046}';
  IID__TaskRequestItem: TGUID = '{00063036-0000-0000-C000-000000000046}';
  IID__TaskRequestUpdateItem: TGUID = '{00063037-0000-0000-C000-000000000046}';
  CLASS_ContactItem: TGUID = '{00061031-0000-0000-C000-000000000046}';
  CLASS_DocumentItem: TGUID = '{00061061-0000-0000-C000-000000000046}';
  CLASS_JournalItem: TGUID = '{00061037-0000-0000-C000-000000000046}';
  CLASS_NoteItem: TGUID = '{00061034-0000-0000-C000-000000000046}';
  CLASS_PostItem: TGUID = '{0006103A-0000-0000-C000-000000000046}';
  CLASS_RemoteItem: TGUID = '{00061060-0000-0000-C000-000000000046}';
  CLASS_ReportItem: TGUID = '{00061035-0000-0000-C000-000000000046}';
  CLASS_TaskRequestAcceptItem: TGUID = '{00061052-0000-0000-C000-000000000046}';
  CLASS_TaskRequestDeclineItem: TGUID = '{00061053-0000-0000-C000-000000000046}';
  CLASS_TaskRequestItem: TGUID = '{00061050-0000-0000-C000-000000000046}';
  CLASS_TaskRequestUpdateItem: TGUID = '{00061051-0000-0000-C000-000000000046}';
  IID__IRecipientControl: TGUID = '{D87E7E16-6897-11CE-A6C0-00AA00608FAA}';
  DIID__DRecipientControl: TGUID = '{0006F025-0000-0000-C000-000000000046}';
  DIID__DRecipientControlEvents: TGUID = '{D87E7E17-6897-11CE-A6C0-00AA00608FAA}';
  CLASS__RecipientControl: TGUID = '{0006F023-0000-0000-C000-000000000046}';
  IID__IDocSiteControl: TGUID = '{43507DD0-811D-11CE-B565-00AA00608FAA}';
  DIID__DDocSiteControl: TGUID = '{0006F026-0000-0000-C000-000000000046}';
  DIID__DDocSiteControlEvents: TGUID = '{50BB9B50-811D-11CE-B565-00AA00608FAA}';
  CLASS__DocSiteControl: TGUID = '{0006F024-0000-0000-C000-000000000046}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                  //
// *********************************************************************//
// OlActionCopyLike constants
type
  OlActionCopyLike = TOleEnum;
const
  olReply = $00000000;
  olReplyAll = $00000001;
  olForward = $00000002;
  olReplyFolder = $00000003;
  olRespond = $00000004;

// OlActionReplyStyle constants
type
  OlActionReplyStyle = TOleEnum;
const
  olOmitOriginalText = $00000000;
  olEmbedOriginalItem = $00000001;
  olIncludeOriginalText = $00000002;
  olIndentOriginalText = $00000003;
  olLinkOriginalItem = $00000004;
  olUserPreference = $00000005;
  olReplyTickOriginalText = $000003E8;

// OlActionResponseStyle constants
type
  OlActionResponseStyle = TOleEnum;
const
  olOpen = $00000000;
  olSend = $00000001;
  olPrompt = $00000002;

// OlActionShowOn constants
type
  OlActionShowOn = TOleEnum;
const
  olDontShow = $00000000;
  olMenu = $00000001;
  olMenuAndToolbar = $00000002;

// OlAttachmentType constants
type
  OlAttachmentType = TOleEnum;
const
  olByValue = $00000001;
  olByReference = $00000004;
  olEmbeddeditem = $00000005;
  olOLE = $00000006;

// OlBusyStatus constants
type
  OlBusyStatus = TOleEnum;
const
  olFree = $00000000;
  olTentative = $00000001;
  olBusy = $00000002;
  olOutOfOffice = $00000003;

// OlDaysOfWeek constants
type
  OlDaysOfWeek = TOleEnum;
const
  olSunday = $00000001;
  olMonday = $00000002;
  olTuesday = $00000004;
  olWednesday = $00000008;
  olThursday = $00000010;
  olFriday = $00000020;
  olSaturday = $00000040;

// OlDefaultFolders constants
type
  OlDefaultFolders = TOleEnum;
const
  olFolderDeletedItems = $00000003;
  olFolderOutbox = $00000004;
  olFolderSentMail = $00000005;
  olFolderInbox = $00000006;
  olFolderCalendar = $00000009;
  olFolderContacts = $0000000A;
  olFolderJournal = $0000000B;
  olFolderNotes = $0000000C;
  olFolderTasks = $0000000D;
  olFolderDrafts = $00000010;

// OlDisplayType constants
type
  OlDisplayType = TOleEnum;
const
  olUser = $00000000;
  olDistList = $00000001;
  olForum = $00000002;
  olAgent = $00000003;
  olOrganization = $00000004;
  olPrivateDistList = $00000005;
  olRemoteUser = $00000006;

// OlEditorType constants
type
  OlEditorType = TOleEnum;
const
  olEditorText = $00000001;
  olEditorHTML = $00000002;
  olEditorRTF = $00000003;
  olEditorWord = $00000004;

// OlFlagStatus constants
type
  OlFlagStatus = TOleEnum;
const
  olNoFlag = $00000000;
  olFlagComplete = $00000001;
  olFlagMarked = $00000002;

// OlFolderDisplayMode constants
type
  OlFolderDisplayMode = TOleEnum;
const
  olFolderDisplayNormal = $00000000;
  olFolderDisplayFolderOnly = $00000001;
  olFolderDisplayNoNavigation = $00000002;

// OlFormRegistry constants
type
  OlFormRegistry = TOleEnum;
const
  olDefaultRegistry = $00000000;
  olPersonalRegistry = $00000002;
  olFolderRegistry = $00000003;
  olOrganizationRegistry = $00000004;

// OlGender constants
type
  OlGender = TOleEnum;
const
  olUnspecified = $00000000;
  olFemale = $00000001;
  olMale = $00000002;

// OlImportance constants
type
  OlImportance = TOleEnum;
const
  olImportanceLow = $00000000;
  olImportanceNormal = $00000001;
  olImportanceHigh = $00000002;

// OlInspectorClose constants
type
  OlInspectorClose = TOleEnum;
const
  olSave = $00000000;
  olDiscard = $00000001;
  olPromptForSave = $00000002;

// OlItemType constants
type
  OlItemType = TOleEnum;
const
  olMailItem = $00000000;
  olAppointmentItem = $00000001;
  olContactItem = $00000002;
  olTaskItem = $00000003;
  olJournalItem = $00000004;
  olNoteItem = $00000005;
  olPostItem = $00000006;

// OlJournalRecipientType constants
type
  OlJournalRecipientType = TOleEnum;
const
  olAssociatedContact = $00000001;

// OlMailingAddress constants
type
  OlMailingAddress = TOleEnum;
const
  olNone = $00000000;
  olHome = $00000001;
  olBusiness = $00000002;
  olOther = $00000003;

// OlMailRecipientType constants
type
  OlMailRecipientType = TOleEnum;
const
  olOriginator = $00000000;
  olTo = $00000001;
  olCC = $00000002;
  olBCC = $00000003;

// OlMeetingRecipientType constants
type
  OlMeetingRecipientType = TOleEnum;
const
  olOrganizer = $00000000;
  olRequired = $00000001;
  olOptional = $00000002;
  olResource = $00000003;

// OlMeetingResponse constants
type
  OlMeetingResponse = TOleEnum;
const
  olMeetingTentative = $00000002;
  olMeetingAccepted = $00000003;
  olMeetingDeclined = $00000004;

// OlMeetingStatus constants
type
  OlMeetingStatus = TOleEnum;
const
  olNonMeeting = $00000000;
  olMeeting = $00000001;
  olMeetingReceived = $00000003;
  olMeetingCanceled = $00000005;

// OlNetMeetingType constants
type
  OlNetMeetingType = TOleEnum;
const
  olNetMeeting = $00000000;
  olNetShow = $00000001;
  olChat = $00000002;

// OlNoteColor constants
type
  OlNoteColor = TOleEnum;
const
  olBlue = $00000000;
  olGreen = $00000001;
  olPink = $00000002;
  olYellow = $00000003;
  olWhite = $00000004;

// OlObjectClass constants
type
  OlObjectClass = TOleEnum;
const
  olApplication = $00000000;
  olNamespace = $00000001;
  olFolder = $00000002;
  olRecipient = $00000004;
  olAttachment = $00000005;
  olAddressList = $00000007;
  olAddressEntry = $00000008;
  olFolders = $0000000F;
  olItems = $00000010;
  olRecipients = $00000011;
  olAttachments = $00000012;
  olAddressLists = $00000014;
  olAddressEntries = $00000015;
  olAppointment = $0000001A;
  olMeetingRequest = $00000035;
  olMeetingCancellation = $00000036;
  olMeetingResponseNegative = $00000037;
  olMeetingResponsePositive = $00000038;
  olMeetingResponseTentative = $00000039;
  olRecurrencePattern = $0000001C;
  olExceptions = $0000001D;
  olException = $0000001E;
  olAction = $00000020;
  olActions = $00000021;
  olExplorer = $00000022;
  olInspector = $00000023;
  olPages = $00000024;
  olFormDescription = $00000025;
  olUserProperties = $00000026;
  olUserProperty = $00000027;
  olContact = $00000028;
  olDocument = $00000029;
  olJournal = $0000002A;
  olMail = $0000002B;
  olNote = $0000002C;
  olPost = $0000002D;
  olReport = $0000002E;
  olRemote = $0000002F;
  olTask = $00000030;
  olTaskRequest = $00000031;
  olTaskRequestUpdate = $00000032;
  olTaskRequestAccept = $00000033;
  olTaskRequestDecline = $00000034;

// OlRecurrenceState constants
type
  OlRecurrenceState = TOleEnum;
const
  olApptNotRecurring = $00000000;
  olApptMaster = $00000001;
  olApptOccurrence = $00000002;
  olApptException = $00000003;

// OlRecurrenceType constants
type
  OlRecurrenceType = TOleEnum;
const
  olRecursDaily = $00000000;
  olRecursWeekly = $00000001;
  olRecursMonthly = $00000002;
  olRecursMonthNth = $00000003;
  olRecursYearly = $00000005;
  olRecursYearNth = $00000006;

// OlRemoteStatus constants
type
  OlRemoteStatus = TOleEnum;
const
  olRemoteStatusNone = $00000000;
  olUnMarked = $00000001;
  olMarkedForDownload = $00000002;
  olMarkedForCopy = $00000003;
  olMarkedForDelete = $00000004;

// OlResponseStatus constants
type
  OlResponseStatus = TOleEnum;
const
  olResponseNone = $00000000;
  olResponseOrganized = $00000001;
  olResponseTentative = $00000002;
  olResponseAccepted = $00000003;
  olResponseDeclined = $00000004;
  olResponseNotResponded = $00000005;

// OlSaveAsType constants
type
  OlSaveAsType = TOleEnum;
const
  olTXT = $00000000;
  olRTF = $00000001;
  olTemplate = $00000002;
  olMSG = $00000003;
  olDoc = $00000004;
  olHTML = $00000005;
  olVCard = $00000006;
  olVCal = $00000007;

// OlSensitivity constants
type
  OlSensitivity = TOleEnum;
const
  olNormal = $00000000;
  olPersonal = $00000001;
  olPrivate = $00000002;
  olConfidential = $00000003;

// OlSortOrder constants
type
  OlSortOrder = TOleEnum;
const
  olSortNone = $00000000;
  olAscending = $00000001;
  olDescending = $00000002;

// OlTaskDelegationState constants
type
  OlTaskDelegationState = TOleEnum;
const
  olTaskNotDelegated = $00000000;
  olTaskDelegationUnknown = $00000001;
  olTaskDelegationAccepted = $00000002;
  olTaskDelegationDeclined = $00000003;

// OlTaskOwnership constants
type
  OlTaskOwnership = TOleEnum;
const
  olNewTask = $00000000;
  olDelegatedTask = $00000001;
  olOwnTask = $00000002;

// OlTaskRecipientType constants
type
  OlTaskRecipientType = TOleEnum;
const
  olUpdate = $00000002;
  olFinalStatus = $00000003;

// OlTaskResponse constants
type
  OlTaskResponse = TOleEnum;
const
  olTaskSimple = $00000000;
  olTaskAssign = $00000001;
  olTaskAccept = $00000002;
  olTaskDecline = $00000003;

// OlTaskStatus constants
type
  OlTaskStatus = TOleEnum;
const
  olTaskNotStarted = $00000000;
  olTaskInProgress = $00000001;
  olTaskComplete = $00000002;
  olTaskWaiting = $00000003;
  olTaskDeferred = $00000004;

// OlTrackingStatus constants
type
  OlTrackingStatus = TOleEnum;
const
  olTrackingNone = $00000000;
  olTrackingDelivered = $00000001;
  olTrackingNotDelivered = $00000002;
  olTrackingNotRead = $00000003;
  olTrackingRecallFailure = $00000004;
  olTrackingRecallSuccess = $00000005;
  olTrackingRead = $00000006;
  olTrackingReplied = $00000007;

// OlUserPropertyType constants
type
  OlUserPropertyType = TOleEnum;
const
  olText = $00000001;
  olNumber = $00000003;
  olDateTime = $00000005;
  olYesNo = $00000006;
  olDuration = $00000007;
  olKeywords = $0000000B;
  olPercent = $0000000C;
  olCurrency = $0000000E;
  olFormula = $00000012;
  olCombination = $00000013;

type

// *********************************************************************//
// Forward declaration of interfaces defined in Type Library            //
// *********************************************************************//
  Action = interface;
  ActionDisp = dispinterface;
  _Application = interface;
  _ApplicationDisp = dispinterface;
  NameSpace = interface;
  NameSpaceDisp = dispinterface;
  Recipient = interface;
  RecipientDisp = dispinterface;
  AddressEntry = interface;
  AddressEntryDisp = dispinterface;
  AddressEntries = interface;
  AddressEntriesDisp = dispinterface;
  Folders = interface;
  FoldersDisp = dispinterface;
  MAPIFolder = interface;
  MAPIFolderDisp = dispinterface;
  Items = interface;
  ItemsDisp = dispinterface;
  Explorer = interface;
  ExplorerDisp = dispinterface;
  AddressLists = interface;
  AddressListsDisp = dispinterface;
  AddressList = interface;
  AddressListDisp = dispinterface;
  Inspector = interface;
  InspectorDisp = dispinterface;
  Actions = interface;
  ActionsDisp = dispinterface;
  _AppointmentItem = interface;
  _AppointmentItemDisp = dispinterface;
  Attachments = interface;
  AttachmentsDisp = dispinterface;
  Attachment = interface;
  AttachmentDisp = dispinterface;
  FormDescription = interface;
  FormDescriptionDisp = dispinterface;
  UserProperties = interface;
  UserPropertiesDisp = dispinterface;
  UserProperty = interface;
  UserPropertyDisp = dispinterface;
  Recipients = interface;
  RecipientsDisp = dispinterface;
  ItemEvents = interface;
  ItemEventsDisp = dispinterface;
  _MailItem = interface;
  _MailItemDisp = dispinterface;
  RecurrencePattern = interface;
  RecurrencePatternDisp = dispinterface;
  Exceptions = interface;
  ExceptionsDisp = dispinterface;
  Exception = interface;
  ExceptionDisp = dispinterface;
  _MeetingItem = interface;
  _MeetingItemDisp = dispinterface;
  _ContactItem = interface;
  _ContactItemDisp = dispinterface;
  _DocumentItem = interface;
  _DocumentItemDisp = dispinterface;
  _JournalItem = interface;
  _JournalItemDisp = dispinterface;
  _NoteItem = interface;
  _NoteItemDisp = dispinterface;
  Pages = interface;
  PagesDisp = dispinterface;
  _PostItem = interface;
  _PostItemDisp = dispinterface;
  _RemoteItem = interface;
  _RemoteItemDisp = dispinterface;
  _ReportItem = interface;
  _ReportItemDisp = dispinterface;
  _TaskItem = interface;
  _TaskItemDisp = dispinterface;
  _TaskRequestAcceptItem = interface;
  _TaskRequestAcceptItemDisp = dispinterface;
  _TaskRequestDeclineItem = interface;
  _TaskRequestDeclineItemDisp = dispinterface;
  _TaskRequestItem = interface;
  _TaskRequestItemDisp = dispinterface;
  _TaskRequestUpdateItem = interface;
  _TaskRequestUpdateItemDisp = dispinterface;
  _IRecipientControl = interface;
  _DRecipientControl = dispinterface;
  _DRecipientControlEvents = dispinterface;
  _IDocSiteControl = interface;
  _DDocSiteControl = dispinterface;
  _DDocSiteControlEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                     //
// (NOTE: Here we map each CoClass to its Default Interface)            //
// *********************************************************************//
  Application_ = _Application;
  MailItem = _MailItem;
  AppointmentItem = _AppointmentItem;
  MeetingItem = _MeetingItem;
  TaskItem = _TaskItem;
  ContactItem = _ContactItem;
  DocumentItem = _DocumentItem;
  JournalItem = _JournalItem;
  NoteItem = _NoteItem;
  PostItem = _PostItem;
  RemoteItem = _RemoteItem;
  ReportItem = _ReportItem;
  TaskRequestAcceptItem = _TaskRequestAcceptItem;
  TaskRequestDeclineItem = _TaskRequestDeclineItem;
  TaskRequestItem = _TaskRequestItem;
  TaskRequestUpdateItem = _TaskRequestUpdateItem;
  _RecipientControl = _DRecipientControl;
  _DocSiteControl = _DDocSiteControl;


// *********************************************************************//
// Declaration of structures, unions and aliases.                       //
// *********************************************************************//
  PUserType1 = ^TGUID; {*}
  PShortint1 = ^Shortint; {*}
  PPShortint1 = ^PShortint1; {*}
  //PUserType2 = ^DISPPARAMS; {*}


// *********************************************************************//
// Interface: Action
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063043-0000-0000-C000-000000000046}
// *********************************************************************//
  Action = interface(IDispatch)
    ['{00063043-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_CopyLike: OlActionCopyLike; safecall;
    procedure Set_CopyLike(CopyLike: OlActionCopyLike); safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Enabled: WordBool); safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Name: WideString); safecall;
    function Get_Prefix: WideString; safecall;
    procedure Set_Prefix(const Prefix: WideString); safecall;
    function Get_ReplyStyle: OlActionReplyStyle; safecall;
    procedure Set_ReplyStyle(ReplyStyle: OlActionReplyStyle); safecall;
    function Get_ResponseStyle: OlActionResponseStyle; safecall;
    procedure Set_ResponseStyle(ResponseStyle: OlActionResponseStyle); safecall;
    function Get_ShowOn: OlActionShowOn; safecall;
    procedure Set_ShowOn(ShowOn: OlActionShowOn); safecall;
    procedure Delete; safecall;
    function Execute: IDispatch; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property CopyLike: OlActionCopyLike read Get_CopyLike write Set_CopyLike;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Name: WideString read Get_Name write Set_Name;
    property Prefix: WideString read Get_Prefix write Set_Prefix;
    property ReplyStyle: OlActionReplyStyle read Get_ReplyStyle write Set_ReplyStyle;
    property ResponseStyle: OlActionResponseStyle read Get_ResponseStyle write Set_ResponseStyle;
    property ShowOn: OlActionShowOn read Get_ShowOn write Set_ShowOn;
  end;

// *********************************************************************//
// DispIntf:  ActionDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063043-0000-0000-C000-000000000046}
// *********************************************************************//
  ActionDisp = dispinterface
    ['{00063043-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property CopyLike: OlActionCopyLike dispid 100;
    property Enabled: WordBool dispid 103;
    property MessageClass: WideString dispid 26;
    property Name: WideString dispid 12289;
    property Prefix: WideString dispid 61;
    property ReplyStyle: OlActionReplyStyle dispid 101;
    property ResponseStyle: OlActionResponseStyle dispid 102;
    property ShowOn: OlActionShowOn dispid 105;
    procedure Delete; dispid 108;
    function Execute: IDispatch; dispid 106;
  end;

// *********************************************************************//
// Interface: _Application
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063001-0000-0000-C000-000000000046}
// *********************************************************************//
  _Application = interface(IDispatch)
    ['{00063001-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Assistant: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Version: WideString; safecall;
    function ActiveExplorer: Explorer; safecall;
    function ActiveInspector: Inspector; safecall;
    function CreateItem(ItemType: OlItemType): IDispatch; safecall;
    function CreateItemFromTemplate(const TemplatePath: WideString; InFolder: OleVariant): IDispatch; safecall;
    function CreateObject(const ObjectName: WideString): IDispatch; safecall;
    function GetNamespace(const Type_: WideString): NameSpace; safecall;
    procedure Quit; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Assistant: IDispatch read Get_Assistant;
    property Name: WideString read Get_Name;
    property Version: WideString read Get_Version;
  end;

// *********************************************************************//
// DispIntf:  _ApplicationDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063001-0000-0000-C000-000000000046}
// *********************************************************************//
  _ApplicationDisp = dispinterface
    ['{00063001-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Assistant: IDispatch readonly dispid 276;
    property Name: WideString readonly dispid 12289;
    property Version: WideString readonly dispid 278;
    function ActiveExplorer: Explorer; dispid 273;
    function ActiveInspector: Inspector; dispid 274;
    function CreateItem(ItemType: OlItemType): IDispatch; dispid 266;
    function CreateItemFromTemplate(const TemplatePath: WideString; InFolder: OleVariant): IDispatch; dispid 267;
    function CreateObject(const ObjectName: WideString): IDispatch; dispid 277;
    function GetNamespace(const Type_: WideString): NameSpace; dispid 272;
    procedure Quit; dispid 275;
  end;

// *********************************************************************//
// Interface: NameSpace
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063002-0000-0000-C000-000000000046}
// *********************************************************************//
  NameSpace = interface(IDispatch)
    ['{00063002-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_CurrentUser: Recipient; safecall;
    function Get_Folders: Folders; safecall;
    function Get_Type_: WideString; safecall;
    function Get_AddressLists: AddressLists; safecall;
    function CreateRecipient(const RecipientName: WideString): Recipient; safecall;
    function GetDefaultFolder(FolderType: OlDefaultFolders): MAPIFolder; safecall;
    function GetFolderFromID(const EntryIDFolder: WideString; EntryIDStore: OleVariant): MAPIFolder; safecall;
    function GetItemFromID(const EntryIDItem: WideString; EntryIDStore: OleVariant): IDispatch; safecall;
    function GetRecipientFromID(const EntryID: WideString): Recipient; safecall;
    function GetSharedDefaultFolder(const Recipient: Recipient; FolderType: OlDefaultFolders): MAPIFolder; safecall;
    procedure Logoff; safecall;
    procedure Logon(Profile: OleVariant; Password: OleVariant; ShowDialog: OleVariant; 
                    NewSession: OleVariant); safecall;
    function PickFolder: MAPIFolder; safecall;
    procedure RefreshRemoteHeaders; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property CurrentUser: Recipient read Get_CurrentUser;
    property Folders: Folders read Get_Folders;
    property Type_: WideString read Get_Type_;
    property AddressLists: AddressLists read Get_AddressLists;
  end;

// *********************************************************************//
// DispIntf:  NameSpaceDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063002-0000-0000-C000-000000000046}
// *********************************************************************//
  NameSpaceDisp = dispinterface
    ['{00063002-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property CurrentUser: Recipient readonly dispid 8449;
    property Folders: Folders readonly dispid 8451;
    property Type_: WideString readonly dispid 8452;
    property AddressLists: AddressLists readonly dispid 8461;
    function CreateRecipient(const RecipientName: WideString): Recipient; dispid 8458;
    function GetDefaultFolder(FolderType: OlDefaultFolders): MAPIFolder; dispid 8459;
    function GetFolderFromID(const EntryIDFolder: WideString; EntryIDStore: OleVariant): MAPIFolder; dispid 8456;
    function GetItemFromID(const EntryIDItem: WideString; EntryIDStore: OleVariant): IDispatch; dispid 8457;
    function GetRecipientFromID(const EntryID: WideString): Recipient; dispid 8455;
    function GetSharedDefaultFolder(const Recipient: Recipient; FolderType: OlDefaultFolders): MAPIFolder; dispid 8460;
    procedure Logoff; dispid 8454;
    procedure Logon(Profile: OleVariant; Password: OleVariant; ShowDialog: OleVariant; 
                    NewSession: OleVariant); dispid 8453;
    function PickFolder: MAPIFolder; dispid 8462;
    procedure RefreshRemoteHeaders; dispid 8471;
  end;

// *********************************************************************//
// Interface: Recipient
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063045-0000-0000-C000-000000000046}
// *********************************************************************//
  Recipient = interface(IDispatch)
    ['{00063045-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Address: WideString; safecall;
    function Get_AddressEntry: AddressEntry; safecall;
    procedure Set_AddressEntry(const AddressEntry: AddressEntry); safecall;
    function Get_AutoResponse: WideString; safecall;
    procedure Set_AutoResponse(const AutoResponse: WideString); safecall;
    function Get_DisplayType: OlDisplayType; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_Index: Integer; safecall;
    function Get_MeetingResponseStatus: OlResponseStatus; safecall;
    function Get_Name: WideString; safecall;
    function Get_Resolved: WordBool; safecall;
    function Get_TrackingStatus: OlTrackingStatus; safecall;
    procedure Set_TrackingStatus(TrackingStatus: OlTrackingStatus); safecall;
    function Get_TrackingStatusTime: TDateTime; safecall;
    procedure Set_TrackingStatusTime(TrackingStatusTime: TDateTime); safecall;
    function Get_Type_: Integer; safecall;
    procedure Set_Type_(Type_: Integer); safecall;
    procedure Delete; safecall;
    function FreeBusy(Start: TDateTime; MinPerChar: Integer; CompleteFormat: OleVariant): WideString; safecall;
    function Resolve: WordBool; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Address: WideString read Get_Address;
    property AddressEntry: AddressEntry read Get_AddressEntry write Set_AddressEntry;
    property AutoResponse: WideString read Get_AutoResponse write Set_AutoResponse;
    property DisplayType: OlDisplayType read Get_DisplayType;
    property EntryID: WideString read Get_EntryID;
    property Index: Integer read Get_Index;
    property MeetingResponseStatus: OlResponseStatus read Get_MeetingResponseStatus;
    property Name: WideString read Get_Name;
    property Resolved: WordBool read Get_Resolved;
    property TrackingStatus: OlTrackingStatus read Get_TrackingStatus write Set_TrackingStatus;
    property TrackingStatusTime: TDateTime read Get_TrackingStatusTime write Set_TrackingStatusTime;
    property Type_: Integer read Get_Type_ write Set_Type_;
  end;

// *********************************************************************//
// DispIntf:  RecipientDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063045-0000-0000-C000-000000000046}
// *********************************************************************//
  RecipientDisp = dispinterface
    ['{00063045-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 109;
    property Address: WideString readonly dispid 12291;
    property AddressEntry: AddressEntry dispid 121;
    property AutoResponse: WideString dispid 106;
    property DisplayType: OlDisplayType readonly dispid 14592;
    property EntryID: WideString readonly dispid 61470;
    property Index: Integer readonly dispid 91;
    property MeetingResponseStatus: OlResponseStatus readonly dispid 102;
    property Name: WideString readonly dispid 12289;
    property Resolved: WordBool readonly dispid 100;
    property TrackingStatus: OlTrackingStatus dispid 118;
    property TrackingStatusTime: TDateTime dispid 119;
    property Type_: Integer dispid 3093;
    procedure Delete; dispid 110;
    function FreeBusy(Start: TDateTime; MinPerChar: Integer; CompleteFormat: OleVariant): WideString; dispid 111;
    function Resolve: WordBool; dispid 113;
  end;

// *********************************************************************//
// Interface: AddressEntry
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304B-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressEntry = interface(IDispatch)
    ['{0006304B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Address: WideString; safecall;
    procedure Set_Address(const Address: WideString); safecall;
    function Get_DisplayType: OlDisplayType; safecall;
    function Get_ID: WideString; safecall;
    function Get_Manager: AddressEntry; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    procedure Set_MAPIOBJECT(const MAPIOBJECT: IUnknown); safecall;
    function Get_Members: AddressEntries; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Name: WideString); safecall;
    function Get_Type_: WideString; safecall;
    procedure Set_Type_(const Type_: WideString); safecall;
    procedure Delete; safecall;
    procedure Details(HWnd: OleVariant); safecall;
    function GetFreeBusy(Start: TDateTime; MinPerChar: Integer; CompleteFormat: OleVariant): WideString; safecall;
    procedure Update(MakePermanent: OleVariant; Refresh: OleVariant); safecall;
    procedure UpdateFreeBusy; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Address: WideString read Get_Address write Set_Address;
    property DisplayType: OlDisplayType read Get_DisplayType;
    property ID: WideString read Get_ID;
    property Manager: AddressEntry read Get_Manager;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT write Set_MAPIOBJECT;
    property Members: AddressEntries read Get_Members;
    property Name: WideString read Get_Name write Set_Name;
    property Type_: WideString read Get_Type_ write Set_Type_;
  end;

// *********************************************************************//
// DispIntf:  AddressEntryDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304B-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressEntryDisp = dispinterface
    ['{0006304B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Address: WideString dispid 12291;
    property DisplayType: OlDisplayType readonly dispid 14592;
    property ID: WideString readonly dispid 61470;
    property Manager: AddressEntry readonly dispid 771;
    property MAPIOBJECT: IUnknown dispid 61696;
    property Members: AddressEntries readonly dispid 772;
    property Name: WideString dispid 12289;
    property Type_: WideString dispid 12290;
    procedure Delete; dispid 770;
    procedure Details(HWnd: OleVariant); dispid 769;
    function GetFreeBusy(Start: TDateTime; MinPerChar: Integer; CompleteFormat: OleVariant): WideString; dispid 774;
    procedure Update(MakePermanent: OleVariant; Refresh: OleVariant); dispid 768;
    procedure UpdateFreeBusy; dispid 775;
  end;

// *********************************************************************//
// Interface: AddressEntries
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304A-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressEntries = interface(IDispatch)
    ['{0006304A-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): AddressEntry; safecall;
    function Get_RawTable: IUnknown; safecall;
    function Add(const Type_: WideString; Name: OleVariant; Address: OleVariant): AddressEntry; safecall;
    function GetFirst: AddressEntry; safecall;
    function GetLast: AddressEntry; safecall;
    function GetNext: AddressEntry; safecall;
    function GetPrevious: AddressEntry; safecall;
    procedure Sort(Property_: OleVariant; Order: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property RawTable: IUnknown read Get_RawTable;
  end;

// *********************************************************************//
// DispIntf:  AddressEntriesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304A-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressEntriesDisp = dispinterface
    ['{0006304A-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): AddressEntry; dispid 81;
    property RawTable: IUnknown readonly dispid 90;
    function Add(const Type_: WideString; Name: OleVariant; Address: OleVariant): AddressEntry; dispid 95;
    function GetFirst: AddressEntry; dispid 86;
    function GetLast: AddressEntry; dispid 88;
    function GetNext: AddressEntry; dispid 87;
    function GetPrevious: AddressEntry; dispid 89;
    procedure Sort(Property_: OleVariant; Order: OleVariant); dispid 97;
  end;

// *********************************************************************//
// Interface: Folders
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063040-0000-0000-C000-000000000046}
// *********************************************************************//
  Folders = interface(IDispatch)
    ['{00063040-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): MAPIFolder; safecall;
    function Get_RawTable: IUnknown; safecall;
    function Add(const Name: WideString; Type_: OleVariant): MAPIFolder; safecall;
    function GetFirst: MAPIFolder; safecall;
    function GetLast: MAPIFolder; safecall;
    function GetNext: MAPIFolder; safecall;
    function GetPrevious: MAPIFolder; safecall;
    procedure Remove(Index: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property RawTable: IUnknown read Get_RawTable;
  end;

// *********************************************************************//
// DispIntf:  FoldersDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063040-0000-0000-C000-000000000046}
// *********************************************************************//
  FoldersDisp = dispinterface
    ['{00063040-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): MAPIFolder; dispid 81;
    property RawTable: IUnknown readonly dispid 90;
    function Add(const Name: WideString; Type_: OleVariant): MAPIFolder; dispid 95;
    function GetFirst: MAPIFolder; dispid 86;
    function GetLast: MAPIFolder; dispid 88;
    function GetNext: MAPIFolder; dispid 87;
    function GetPrevious: MAPIFolder; dispid 89;
    procedure Remove(Index: Integer); dispid 84;
  end;

// *********************************************************************//
// Interface: MAPIFolder
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063006-0000-0000-C000-000000000046}
// *********************************************************************//
  MAPIFolder = interface(IDispatch)
    ['{00063006-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_DefaultItemType: OlItemType; safecall;
    function Get_DefaultMessageClass: WideString; safecall;
    function Get_Description: WideString; safecall;
    procedure Set_Description(const Description: WideString); safecall;
    function Get_EntryID: WideString; safecall;
    function Get_Folders: Folders; safecall;
    function Get_Items: Items; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Name: WideString); safecall;
    function Get_StoreID: WideString; safecall;
    function Get_UnReadItemCount: Integer; safecall;
    function CopyTo(const DestinationFolder: MAPIFolder): MAPIFolder; safecall;
    procedure Delete; safecall;
    procedure Display; safecall;
    function GetExplorer(DisplayMode: OleVariant): Explorer; safecall;
    procedure MoveTo(const DestinationFolder: MAPIFolder); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property DefaultItemType: OlItemType read Get_DefaultItemType;
    property DefaultMessageClass: WideString read Get_DefaultMessageClass;
    property Description: WideString read Get_Description write Set_Description;
    property EntryID: WideString read Get_EntryID;
    property Folders: Folders read Get_Folders;
    property Items: Items read Get_Items;
    property Name: WideString read Get_Name write Set_Name;
    property StoreID: WideString read Get_StoreID;
    property UnReadItemCount: Integer read Get_UnReadItemCount;
  end;

// *********************************************************************//
// DispIntf:  MAPIFolderDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063006-0000-0000-C000-000000000046}
// *********************************************************************//
  MAPIFolderDisp = dispinterface
    ['{00063006-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property DefaultItemType: OlItemType readonly dispid 12550;
    property DefaultMessageClass: WideString readonly dispid 12551;
    property Description: WideString dispid 12292;
    property EntryID: WideString readonly dispid 61470;
    property Folders: Folders readonly dispid 8451;
    property Items: Items readonly dispid 12544;
    property Name: WideString dispid 12289;
    property StoreID: WideString readonly dispid 12552;
    property UnReadItemCount: Integer readonly dispid 13827;
    function CopyTo(const DestinationFolder: MAPIFolder): MAPIFolder; dispid 61490;
    procedure Delete; dispid 61509;
    procedure Display; dispid 12548;
    function GetExplorer(DisplayMode: OleVariant): Explorer; dispid 12545;
    procedure MoveTo(const DestinationFolder: MAPIFolder); dispid 61492;
  end;

// *********************************************************************//
// Interface: Items
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063041-0000-0000-C000-000000000046}
// *********************************************************************//
  Items = interface(IDispatch)
    ['{00063041-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): IDispatch; safecall;
    function Get_RawTable: IUnknown; safecall;
    function Get_IncludeRecurrences: WordBool; safecall;
    procedure Set_IncludeRecurrences(IncludeRecurrences: WordBool); safecall;
    function Add(Type_: OleVariant): IDispatch; safecall;
    function Find(const Filter: WideString): IDispatch; safecall;
    function FindNext: IDispatch; safecall;
    function GetFirst: IDispatch; safecall;
    function GetLast: IDispatch; safecall;
    function GetNext: IDispatch; safecall;
    function GetPrevious: IDispatch; safecall;
    procedure Remove(Index: Integer); safecall;
    procedure ResetColumns; safecall;
    function Restrict(const Filter: WideString): Items; safecall;
    procedure SetColumns(const Columns: WideString); safecall;
    procedure Sort(const Property_: WideString; Descending: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
    property RawTable: IUnknown read Get_RawTable;
    property IncludeRecurrences: WordBool read Get_IncludeRecurrences write Set_IncludeRecurrences;
  end;

// *********************************************************************//
// DispIntf:  ItemsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063041-0000-0000-C000-000000000046}
// *********************************************************************//
  ItemsDisp = dispinterface
    ['{00063041-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): IDispatch; dispid 81;
    property RawTable: IUnknown readonly dispid 90;
    property IncludeRecurrences: WordBool dispid 206;
    function Add(Type_: OleVariant): IDispatch; dispid 95;
    function Find(const Filter: WideString): IDispatch; dispid 98;
    function FindNext: IDispatch; dispid 99;
    function GetFirst: IDispatch; dispid 86;
    function GetLast: IDispatch; dispid 88;
    function GetNext: IDispatch; dispid 87;
    function GetPrevious: IDispatch; dispid 89;
    procedure Remove(Index: Integer); dispid 84;
    procedure ResetColumns; dispid 93;
    function Restrict(const Filter: WideString): Items; dispid 100;
    procedure SetColumns(const Columns: WideString); dispid 92;
    procedure Sort(const Property_: WideString; Descending: OleVariant); dispid 97;
  end;

// *********************************************************************//
// Interface: Explorer
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063003-0000-0000-C000-000000000046}
// *********************************************************************//
  Explorer = interface(IDispatch)
    ['{00063003-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_CommandBars: IDispatch; safecall;
    function Get_CurrentFolder: MAPIFolder; safecall;
    procedure Set_CurrentFolder(const CurrentFolder: MAPIFolder); safecall;
    procedure Close; safecall;
    procedure Display; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property CommandBars: IDispatch read Get_CommandBars;
    property CurrentFolder: MAPIFolder read Get_CurrentFolder write Set_CurrentFolder;
  end;

// *********************************************************************//
// DispIntf:  ExplorerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063003-0000-0000-C000-000000000046}
// *********************************************************************//
  ExplorerDisp = dispinterface
    ['{00063003-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property CommandBars: IDispatch readonly dispid 8448;
    property CurrentFolder: MAPIFolder dispid 8449;
    procedure Close; dispid 8451;
    procedure Display; dispid 8452;
  end;

// *********************************************************************//
// Interface: AddressLists
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063048-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressLists = interface(IDispatch)
    ['{00063048-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): AddressList; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AddressListsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063048-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressListsDisp = dispinterface
    ['{00063048-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): AddressList; dispid 81;
  end;

// *********************************************************************//
// Interface: AddressList
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063049-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressList = interface(IDispatch)
    ['{00063049-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_AddressEntries: AddressEntries; safecall;
    function Get_ID: WideString; safecall;
    function Get_Index: Integer; safecall;
    function Get_IsReadOnly: WordBool; safecall;
    function Get_Name: WideString; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property AddressEntries: AddressEntries read Get_AddressEntries;
    property ID: WideString read Get_ID;
    property Index: Integer read Get_Index;
    property IsReadOnly: WordBool read Get_IsReadOnly;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  AddressListDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063049-0000-0000-C000-000000000046}
// *********************************************************************//
  AddressListDisp = dispinterface
    ['{00063049-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property AddressEntries: AddressEntries readonly dispid 256;
    property ID: WideString readonly dispid 61470;
    property Index: Integer readonly dispid 91;
    property IsReadOnly: WordBool readonly dispid 61463;
    property Name: WideString readonly dispid 12289;
  end;

// *********************************************************************//
// Interface: Inspector
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063005-0000-0000-C000-000000000046}
// *********************************************************************//
  Inspector = interface(IDispatch)
    ['{00063005-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_CommandBars: IDispatch; safecall;
    function Get_CurrentItem: IDispatch; safecall;
    function Get_EditorType: OlEditorType; safecall;
    function Get_ModifiedFormPages: IDispatch; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    procedure Display(Modal: OleVariant); safecall;
    procedure HideFormPage(const PageName: WideString); safecall;
    function IsWordMail: WordBool; safecall;
    procedure SetCurrentFormPage(const PageName: WideString); safecall;
    procedure ShowFormPage(const PageName: WideString); safecall;
    function Get_HTMLEditor: IDispatch; safecall;
    function Get_WordEditor: IDispatch; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property CommandBars: IDispatch read Get_CommandBars;
    property CurrentItem: IDispatch read Get_CurrentItem;
    property EditorType: OlEditorType read Get_EditorType;
    property ModifiedFormPages: IDispatch read Get_ModifiedFormPages;
    property HTMLEditor: IDispatch read Get_HTMLEditor;
    property WordEditor: IDispatch read Get_WordEditor;
  end;

// *********************************************************************//
// DispIntf:  InspectorDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063005-0000-0000-C000-000000000046}
// *********************************************************************//
  InspectorDisp = dispinterface
    ['{00063005-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property CommandBars: IDispatch readonly dispid 8448;
    property CurrentItem: IDispatch readonly dispid 8450;
    property EditorType: OlEditorType readonly dispid 8464;
    property ModifiedFormPages: IDispatch readonly dispid 8454;
    procedure Close(SaveMode: OlInspectorClose); dispid 8451;
    procedure Display(Modal: OleVariant); dispid 8452;
    procedure HideFormPage(const PageName: WideString); dispid 8456;
    function IsWordMail: WordBool; dispid 8453;
    procedure SetCurrentFormPage(const PageName: WideString); dispid 8460;
    procedure ShowFormPage(const PageName: WideString); dispid 8457;
    property HTMLEditor: IDispatch readonly dispid 8462;
    property WordEditor: IDispatch readonly dispid 8463;
  end;

// *********************************************************************//
// Interface: Actions
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303E-0000-0000-C000-000000000046}
// *********************************************************************//
  Actions = interface(IDispatch)
    ['{0006303E-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): Action; safecall;
    function Add: Action; safecall;
    procedure Remove(Index: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ActionsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303E-0000-0000-C000-000000000046}
// *********************************************************************//
  ActionsDisp = dispinterface
    ['{0006303E-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): Action; dispid 81;
    function Add: Action; dispid 100;
    procedure Remove(Index: Integer); dispid 82;
  end;

// *********************************************************************//
// Interface: _AppointmentItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063033-0000-0000-C000-000000000046}
// *********************************************************************//
  _AppointmentItem = interface(IDispatch)
    ['{00063033-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_AllDayEvent: WordBool; safecall;
    procedure Set_AllDayEvent(AllDayEvent: WordBool); safecall;
    function Get_BusyStatus: OlBusyStatus; safecall;
    procedure Set_BusyStatus(BusyStatus: OlBusyStatus); safecall;
    function Get_Duration: Integer; safecall;
    procedure Set_Duration(Duration: Integer); safecall;
    function Get_End_: TDateTime; safecall;
    procedure Set_End_(End_: TDateTime); safecall;
    function Get_IsOnlineMeeting: WordBool; safecall;
    procedure Set_IsOnlineMeeting(IsOnlineMeeting: WordBool); safecall;
    function Get_IsRecurring: WordBool; safecall;
    function Get_Location: WideString; safecall;
    procedure Set_Location(const Location: WideString); safecall;
    function Get_MeetingStatus: OlMeetingStatus; safecall;
    procedure Set_MeetingStatus(MeetingStatus: OlMeetingStatus); safecall;
    function Get_NetMeetingAutoStart: WordBool; safecall;
    procedure Set_NetMeetingAutoStart(NetMeetingAutoStart: WordBool); safecall;
    function Get_NetMeetingOrganizerAlias: WideString; safecall;
    procedure Set_NetMeetingOrganizerAlias(const NetMeetingOrganizerAlias: WideString); safecall;
    function Get_NetMeetingServer: WideString; safecall;
    procedure Set_NetMeetingServer(const NetMeetingServer: WideString); safecall;
    function Get_NetMeetingType: OlNetMeetingType; safecall;
    procedure Set_NetMeetingType(NetMeetingType: OlNetMeetingType); safecall;
    function Get_OptionalAttendees: WideString; safecall;
    procedure Set_OptionalAttendees(const OptionalAttendees: WideString); safecall;
    function Get_Organizer: WideString; safecall;
    function Get_Recipients: Recipients; safecall;
    function Get_RecurrenceState: OlRecurrenceState; safecall;
    function Get_ReminderMinutesBeforeStart: Integer; safecall;
    procedure Set_ReminderMinutesBeforeStart(ReminderMinutesBeforeStart: Integer); safecall;
    function Get_ReminderOverrideDefault: WordBool; safecall;
    procedure Set_ReminderOverrideDefault(ReminderOverrideDefault: WordBool); safecall;
    function Get_ReminderPlaySound: WordBool; safecall;
    procedure Set_ReminderPlaySound(ReminderPlaySound: WordBool); safecall;
    function Get_ReminderSet: WordBool; safecall;
    procedure Set_ReminderSet(ReminderSet: WordBool); safecall;
    function Get_ReminderSoundFile: WideString; safecall;
    procedure Set_ReminderSoundFile(const ReminderSoundFile: WideString); safecall;
    function Get_ReplyTime: TDateTime; safecall;
    procedure Set_ReplyTime(ReplyTime: TDateTime); safecall;
    function Get_RequiredAttendees: WideString; safecall;
    procedure Set_RequiredAttendees(const RequiredAttendees: WideString); safecall;
    function Get_Resources: WideString; safecall;
    procedure Set_Resources(const Resources: WideString); safecall;
    function Get_ResponseRequested: WordBool; safecall;
    procedure Set_ResponseRequested(ResponseRequested: WordBool); safecall;
    function Get_ResponseStatus: OlResponseStatus; safecall;
    function Get_Start: TDateTime; safecall;
    procedure Set_Start(Start: TDateTime); safecall;
    procedure ClearRecurrencePattern; safecall;
    function ForwardAsVcal: MailItem; safecall;
    function GetRecurrencePattern: RecurrencePattern; safecall;
    function Respond(Response: OlMeetingResponse; fNoUI: OleVariant; 
                     fAdditionalTextDialog: OleVariant): MeetingItem; safecall;
    procedure Send; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property AllDayEvent: WordBool read Get_AllDayEvent write Set_AllDayEvent;
    property BusyStatus: OlBusyStatus read Get_BusyStatus write Set_BusyStatus;
    property Duration: Integer read Get_Duration write Set_Duration;
    property End_: TDateTime read Get_End_ write Set_End_;
    property IsOnlineMeeting: WordBool read Get_IsOnlineMeeting write Set_IsOnlineMeeting;
    property IsRecurring: WordBool read Get_IsRecurring;
    property Location: WideString read Get_Location write Set_Location;
    property MeetingStatus: OlMeetingStatus read Get_MeetingStatus write Set_MeetingStatus;
    property NetMeetingAutoStart: WordBool read Get_NetMeetingAutoStart write Set_NetMeetingAutoStart;
    property NetMeetingOrganizerAlias: WideString read Get_NetMeetingOrganizerAlias write Set_NetMeetingOrganizerAlias;
    property NetMeetingServer: WideString read Get_NetMeetingServer write Set_NetMeetingServer;
    property NetMeetingType: OlNetMeetingType read Get_NetMeetingType write Set_NetMeetingType;
    property OptionalAttendees: WideString read Get_OptionalAttendees write Set_OptionalAttendees;
    property Organizer: WideString read Get_Organizer;
    property Recipients: Recipients read Get_Recipients;
    property RecurrenceState: OlRecurrenceState read Get_RecurrenceState;
    property ReminderMinutesBeforeStart: Integer read Get_ReminderMinutesBeforeStart write Set_ReminderMinutesBeforeStart;
    property ReminderOverrideDefault: WordBool read Get_ReminderOverrideDefault write Set_ReminderOverrideDefault;
    property ReminderPlaySound: WordBool read Get_ReminderPlaySound write Set_ReminderPlaySound;
    property ReminderSet: WordBool read Get_ReminderSet write Set_ReminderSet;
    property ReminderSoundFile: WideString read Get_ReminderSoundFile write Set_ReminderSoundFile;
    property ReplyTime: TDateTime read Get_ReplyTime write Set_ReplyTime;
    property RequiredAttendees: WideString read Get_RequiredAttendees write Set_RequiredAttendees;
    property Resources: WideString read Get_Resources write Set_Resources;
    property ResponseRequested: WordBool read Get_ResponseRequested write Set_ResponseRequested;
    property ResponseStatus: OlResponseStatus read Get_ResponseStatus;
    property Start: TDateTime read Get_Start write Set_Start;
  end;

// *********************************************************************//
// DispIntf:  _AppointmentItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063033-0000-0000-C000-000000000046}
// *********************************************************************//
  _AppointmentItemDisp = dispinterface
    ['{00063033-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property AllDayEvent: WordBool dispid 33301;
    property BusyStatus: OlBusyStatus dispid 33285;
    property Duration: Integer dispid 33299;
    property End_: TDateTime dispid 33294;
    property IsOnlineMeeting: WordBool dispid 33344;
    property IsRecurring: WordBool readonly dispid 33315;
    property Location: WideString dispid 33288;
    property MeetingStatus: OlMeetingStatus dispid 33303;
    property NetMeetingAutoStart: WordBool dispid 33348;
    property NetMeetingOrganizerAlias: WideString dispid 33347;
    property NetMeetingServer: WideString dispid 33346;
    property NetMeetingType: OlNetMeetingType dispid 33345;
    property OptionalAttendees: WideString dispid 3587;
    property Organizer: WideString readonly dispid 66;
    property Recipients: Recipients readonly dispid 63508;
    property RecurrenceState: OlRecurrenceState readonly dispid 62789;
    property ReminderMinutesBeforeStart: Integer dispid 34049;
    property ReminderOverrideDefault: WordBool dispid 34076;
    property ReminderPlaySound: WordBool dispid 34078;
    property ReminderSet: WordBool dispid 34051;
    property ReminderSoundFile: WideString dispid 34079;
    property ReplyTime: TDateTime dispid 33312;
    property RequiredAttendees: WideString dispid 3588;
    property Resources: WideString dispid 3586;
    property ResponseRequested: WordBool dispid 99;
    property ResponseStatus: OlResponseStatus readonly dispid 33304;
    property Start: TDateTime dispid 33293;
    procedure ClearRecurrencePattern; dispid 61605;
    function ForwardAsVcal: MailItem; dispid 62791;
    function GetRecurrencePattern: RecurrencePattern; dispid 61604;
    function Respond(Response: OlMeetingResponse; fNoUI: OleVariant; 
                     fAdditionalTextDialog: OleVariant): MeetingItem; dispid 62722;
    procedure Send; dispid 61557;
  end;

// *********************************************************************//
// Interface: Attachments
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303C-0000-0000-C000-000000000046}
// *********************************************************************//
  Attachments = interface(IDispatch)
    ['{0006303C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): Attachment; safecall;
    function Add(Source: OleVariant; Type_: OleVariant; Position: OleVariant; 
                 DisplayName: OleVariant): Attachment; safecall;
    procedure Remove(Index: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AttachmentsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303C-0000-0000-C000-000000000046}
// *********************************************************************//
  AttachmentsDisp = dispinterface
    ['{0006303C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): Attachment; dispid 81;
    function Add(Source: OleVariant; Type_: OleVariant; Position: OleVariant; 
                 DisplayName: OleVariant): Attachment; dispid 101;
    procedure Remove(Index: Integer); dispid 84;
  end;

// *********************************************************************//
// Interface: Attachment
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063007-0000-0000-C000-000000000046}
// *********************************************************************//
  Attachment = interface(IDispatch)
    ['{00063007-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_DisplayName: WideString; safecall;
    procedure Set_DisplayName(const DisplayName: WideString); safecall;
    function Get_FileName: WideString; safecall;
    function Get_Index: Integer; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_PathName: WideString; safecall;
    function Get_Position: Integer; safecall;
    procedure Set_Position(Position: Integer); safecall;
    function Get_Type_: OlAttachmentType; safecall;
    procedure Delete; safecall;
    procedure SaveAsFile(const Path: WideString); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property DisplayName: WideString read Get_DisplayName write Set_DisplayName;
    property FileName: WideString read Get_FileName;
    property Index: Integer read Get_Index;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property PathName: WideString read Get_PathName;
    property Position: Integer read Get_Position write Set_Position;
    property Type_: OlAttachmentType read Get_Type_;
  end;

// *********************************************************************//
// DispIntf:  AttachmentDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063007-0000-0000-C000-000000000046}
// *********************************************************************//
  AttachmentDisp = dispinterface
    ['{00063007-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 113;
    property DisplayName: WideString dispid 12289;
    property FileName: WideString readonly dispid 14084;
    property Index: Integer readonly dispid 91;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property PathName: WideString readonly dispid 14088;
    property Position: Integer dispid 114;
    property Type_: OlAttachmentType readonly dispid 14085;
    procedure Delete; dispid 105;
    procedure SaveAsFile(const Path: WideString); dispid 104;
  end;

// *********************************************************************//
// Interface: FormDescription
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063046-0000-0000-C000-000000000046}
// *********************************************************************//
  FormDescription = interface(IDispatch)
    ['{00063046-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Category: WideString; safecall;
    procedure Set_Category(const Category: WideString); safecall;
    function Get_CategorySub: WideString; safecall;
    procedure Set_CategorySub(const CategorySub: WideString); safecall;
    function Get_Comment: WideString; safecall;
    procedure Set_Comment(const Comment: WideString); safecall;
    function Get_ContactName: WideString; safecall;
    procedure Set_ContactName(const ContactName: WideString); safecall;
    function Get_DisplayName: WideString; safecall;
    procedure Set_DisplayName(const DisplayName: WideString); safecall;
    function Get_Hidden: WordBool; safecall;
    procedure Set_Hidden(Hidden: WordBool); safecall;
    function Get_Icon: WideString; safecall;
    procedure Set_Icon(const Icon: WideString); safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(Locked: WordBool); safecall;
    function Get_MessageClass: WideString; safecall;
    function Get_MiniIcon: WideString; safecall;
    procedure Set_MiniIcon(const MiniIcon: WideString); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Name: WideString); safecall;
    function Get_Number: WideString; safecall;
    procedure Set_Number(const Number: WideString); safecall;
    function Get_OneOff: WordBool; safecall;
    procedure Set_OneOff(OneOff: WordBool); safecall;
    function Get_Password: WideString; safecall;
    procedure Set_Password(const Password: WideString); safecall;
    function Get_ScriptText: WideString; safecall;
    function Get_Template: WideString; safecall;
    procedure Set_Template(const Template: WideString); safecall;
    function Get_UseWordMail: WordBool; safecall;
    procedure Set_UseWordMail(UseWordMail: WordBool); safecall;
    function Get_Version: WideString; safecall;
    procedure Set_Version(const Version: WideString); safecall;
    procedure PublishForm(Registry: OlFormRegistry; Folder: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Category: WideString read Get_Category write Set_Category;
    property CategorySub: WideString read Get_CategorySub write Set_CategorySub;
    property Comment: WideString read Get_Comment write Set_Comment;
    property ContactName: WideString read Get_ContactName write Set_ContactName;
    property DisplayName: WideString read Get_DisplayName write Set_DisplayName;
    property Hidden: WordBool read Get_Hidden write Set_Hidden;
    property Icon: WideString read Get_Icon write Set_Icon;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property MessageClass: WideString read Get_MessageClass;
    property MiniIcon: WideString read Get_MiniIcon write Set_MiniIcon;
    property Name: WideString read Get_Name write Set_Name;
    property Number: WideString read Get_Number write Set_Number;
    property OneOff: WordBool read Get_OneOff write Set_OneOff;
    property Password: WideString read Get_Password write Set_Password;
    property ScriptText: WideString read Get_ScriptText;
    property Template: WideString read Get_Template write Set_Template;
    property UseWordMail: WordBool read Get_UseWordMail write Set_UseWordMail;
    property Version: WideString read Get_Version write Set_Version;
  end;

// *********************************************************************//
// DispIntf:  FormDescriptionDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063046-0000-0000-C000-000000000046}
// *********************************************************************//
  FormDescriptionDisp = dispinterface
    ['{00063046-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Category: WideString dispid 13060;
    property CategorySub: WideString dispid 13061;
    property Comment: WideString dispid 12292;
    property ContactName: WideString dispid 13059;
    property DisplayName: WideString dispid 12289;
    property Hidden: WordBool dispid 13063;
    property Icon: WideString dispid 4093;
    property Locked: WordBool dispid 102;
    property MessageClass: WideString readonly dispid 26;
    property MiniIcon: WideString dispid 4092;
    property Name: WideString dispid 61469;
    property Number: WideString dispid 104;
    property OneOff: WordBool dispid 101;
    property Password: WideString dispid 103;
    property ScriptText: WideString readonly dispid 109;
    property Template: WideString dispid 106;
    property UseWordMail: WordBool dispid 105;
    property Version: WideString dispid 13057;
    procedure PublishForm(Registry: OlFormRegistry; Folder: OleVariant); dispid 107;
  end;

// *********************************************************************//
// Interface: UserProperties
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303D-0000-0000-C000-000000000046}
// *********************************************************************//
  UserProperties = interface(IDispatch)
    ['{0006303D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): UserProperty; safecall;
    function Add(const Name: WideString; Type_: OlUserPropertyType; AddToFolderFields: OleVariant; 
                 DisplayFormat: OleVariant): UserProperty; safecall;
    function Find(const Name: WideString; Custom: OleVariant): UserProperty; safecall;
    procedure Remove(Index: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  UserPropertiesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303D-0000-0000-C000-000000000046}
// *********************************************************************//
  UserPropertiesDisp = dispinterface
    ['{0006303D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): UserProperty; dispid 81;
    function Add(const Name: WideString; Type_: OlUserPropertyType; AddToFolderFields: OleVariant; 
                 DisplayFormat: OleVariant): UserProperty; dispid 102;
    function Find(const Name: WideString; Custom: OleVariant): UserProperty; dispid 103;
    procedure Remove(Index: Integer); dispid 82;
  end;

// *********************************************************************//
// Interface: UserProperty
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063042-0000-0000-C000-000000000046}
// *********************************************************************//
  UserProperty = interface(IDispatch)
    ['{00063042-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Formula: WideString; safecall;
    procedure Set_Formula(const Formula: WideString); safecall;
    function Get_Name: WideString; safecall;
    function Get_Type_: OlUserPropertyType; safecall;
    function Get_ValidationFormula: WideString; safecall;
    procedure Set_ValidationFormula(const ValidationFormula: WideString); safecall;
    function Get_ValidationText: WideString; safecall;
    procedure Set_ValidationText(const ValidationText: WideString); safecall;
    function Get_Value: OleVariant; safecall;
    procedure Set_Value(Value: OleVariant); safecall;
    procedure Delete; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Formula: WideString read Get_Formula write Set_Formula;
    property Name: WideString read Get_Name;
    property Type_: OlUserPropertyType read Get_Type_;
    property ValidationFormula: WideString read Get_ValidationFormula write Set_ValidationFormula;
    property ValidationText: WideString read Get_ValidationText write Set_ValidationText;
    property Value: OleVariant read Get_Value write Set_Value;
  end;

// *********************************************************************//
// DispIntf:  UserPropertyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063042-0000-0000-C000-000000000046}
// *********************************************************************//
  UserPropertyDisp = dispinterface
    ['{00063042-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Formula: WideString dispid 103;
    property Name: WideString readonly dispid 112;
    property Type_: OlUserPropertyType readonly dispid 109;
    property ValidationFormula: WideString dispid 104;
    property ValidationText: WideString dispid 105;
    property Value: OleVariant dispid 0;
    procedure Delete; dispid 108;
  end;

// *********************************************************************//
// Interface: Recipients
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303B-0000-0000-C000-000000000046}
// *********************************************************************//
  Recipients = interface(IDispatch)
    ['{0006303B-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): Recipient; safecall;
    function Add(const Name: WideString): Recipient; safecall;
    procedure Remove(Index: Integer); safecall;
    function ResolveAll: WordBool; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  RecipientsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303B-0000-0000-C000-000000000046}
// *********************************************************************//
  RecipientsDisp = dispinterface
    ['{0006303B-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): Recipient; dispid 81;
    function Add(const Name: WideString): Recipient; dispid 111;
    procedure Remove(Index: Integer); dispid 84;
    function ResolveAll: WordBool; dispid 126;
  end;

// *********************************************************************//
// Interface: ItemEvents
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303A-0000-0000-C000-000000000046}
// *********************************************************************//
  ItemEvents = interface(IDispatch)
    ['{0006303A-0000-0000-C000-000000000046}']
    procedure CustomAction(const Action: IDispatch; const Response: IDispatch); safecall;
    procedure CustomPropertyChange(const Name: WideString); safecall;
    procedure Forward(const Forward: IDispatch); safecall;
    procedure Open; safecall;
    procedure Close; safecall;
    procedure PropertyChange(const Name: WideString); safecall;
    procedure Read; safecall;
    procedure Reply(const Response: IDispatch); safecall;
    procedure ReplyAll(const Response: IDispatch); safecall;
    procedure Send; safecall;
    procedure Write; safecall;
  end;

// *********************************************************************//
// DispIntf:  ItemEventsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303A-0000-0000-C000-000000000046}
// *********************************************************************//
  ItemEventsDisp = dispinterface
    ['{0006303A-0000-0000-C000-000000000046}']
    procedure CustomAction(const Action: IDispatch; const Response: IDispatch); dispid 61446;
    procedure CustomPropertyChange(const Name: WideString); dispid 61448;
    procedure Forward(const Forward: IDispatch); dispid 62568;
    procedure Open; dispid 61443;
    procedure Close; dispid 61444;
    procedure PropertyChange(const Name: WideString); dispid 61449;
    procedure Read; dispid 61441;
    procedure Reply(const Response: IDispatch); dispid 62566;
    procedure ReplyAll(const Response: IDispatch); dispid 62567;
    procedure Send; dispid 61445;
    procedure Write; dispid 61442;
  end;

// *********************************************************************//
// Interface: _MailItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063034-0000-0000-C000-000000000046}
// *********************************************************************//
  _MailItem = interface(IDispatch)
    ['{00063034-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_AlternateRecipientAllowed: WordBool; safecall;
    procedure Set_AlternateRecipientAllowed(AlternateRecipientAllowed: WordBool); safecall;
    function Get_AutoForwarded: WordBool; safecall;
    procedure Set_AutoForwarded(AutoForwarded: WordBool); safecall;
    function Get_BCC: WideString; safecall;
    procedure Set_BCC(const BCC: WideString); safecall;
    function Get_CC: WideString; safecall;
    procedure Set_CC(const CC: WideString); safecall;
    function Get_DeferredDeliveryTime: TDateTime; safecall;
    procedure Set_DeferredDeliveryTime(DeferredDeliveryTime: TDateTime); safecall;
    function Get_DeleteAfterSubmit: WordBool; safecall;
    procedure Set_DeleteAfterSubmit(DeleteAfterSubmit: WordBool); safecall;
    function Get_ExpiryTime: TDateTime; safecall;
    procedure Set_ExpiryTime(ExpiryTime: TDateTime); safecall;
    function Get_FlagDueBy: TDateTime; safecall;
    procedure Set_FlagDueBy(FlagDueBy: TDateTime); safecall;
    function Get_FlagRequest: WideString; safecall;
    procedure Set_FlagRequest(const FlagRequest: WideString); safecall;
    function Get_FlagStatus: OlFlagStatus; safecall;
    procedure Set_FlagStatus(FlagStatus: OlFlagStatus); safecall;
    function Get_HTMLBody: WideString; safecall;
    procedure Set_HTMLBody(const HTMLBody: WideString); safecall;
    function Get_OriginatorDeliveryReportRequested: WordBool; safecall;
    procedure Set_OriginatorDeliveryReportRequested(OriginatorDeliveryReportRequested: WordBool); safecall;
    function Get_ReadReceiptRequested: WordBool; safecall;
    procedure Set_ReadReceiptRequested(ReadReceiptRequested: WordBool); safecall;
    function Get_ReceivedByEntryID: WideString; safecall;
    function Get_ReceivedByName: WideString; safecall;
    function Get_ReceivedOnBehalfOfEntryID: WideString; safecall;
    function Get_ReceivedOnBehalfOfName: WideString; safecall;
    function Get_ReceivedTime: TDateTime; safecall;
    function Get_RecipientReassignmentProhibited: WordBool; safecall;
    procedure Set_RecipientReassignmentProhibited(RecipientReassignmentProhibited: WordBool); safecall;
    function Get_Recipients: Recipients; safecall;
    function Get_ReminderOverrideDefault: WordBool; safecall;
    procedure Set_ReminderOverrideDefault(ReminderOverrideDefault: WordBool); safecall;
    function Get_ReminderPlaySound: WordBool; safecall;
    procedure Set_ReminderPlaySound(ReminderPlaySound: WordBool); safecall;
    function Get_ReminderSet: WordBool; safecall;
    procedure Set_ReminderSet(ReminderSet: WordBool); safecall;
    function Get_ReminderSoundFile: WideString; safecall;
    procedure Set_ReminderSoundFile(const ReminderSoundFile: WideString); safecall;
    function Get_ReminderTime: TDateTime; safecall;
    procedure Set_ReminderTime(ReminderTime: TDateTime); safecall;
    function Get_RemoteStatus: OlRemoteStatus; safecall;
    procedure Set_RemoteStatus(RemoteStatus: OlRemoteStatus); safecall;
    function Get_ReplyRecipientNames: WideString; safecall;
    function Get_ReplyRecipients: Recipients; safecall;
    function Get_SaveSentMessageFolder: MAPIFolder; safecall;
    procedure Set_SaveSentMessageFolder(const SaveSentMessageFolder: MAPIFolder); safecall;
    function Get_SenderName: WideString; safecall;
    function Get_Sent: WordBool; safecall;
    function Get_SentOn: TDateTime; safecall;
    function Get_SentOnBehalfOfName: WideString; safecall;
    procedure Set_SentOnBehalfOfName(const SentOnBehalfOfName: WideString); safecall;
    function Get_Submitted: WordBool; safecall;
    function Get_To_: WideString; safecall;
    procedure Set_To_(const To_: WideString); safecall;
    function Get_VotingOptions: WideString; safecall;
    procedure Set_VotingOptions(const VotingOptions: WideString); safecall;
    function Get_VotingResponse: WideString; safecall;
    procedure Set_VotingResponse(const VotingResponse: WideString); safecall;
    procedure ClearConversationIndex; safecall;
    function Forward: MailItem; safecall;
    function Reply: MailItem; safecall;
    function ReplyAll: MailItem; safecall;
    procedure Send; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property AlternateRecipientAllowed: WordBool read Get_AlternateRecipientAllowed write Set_AlternateRecipientAllowed;
    property AutoForwarded: WordBool read Get_AutoForwarded write Set_AutoForwarded;
    property BCC: WideString read Get_BCC write Set_BCC;
    property CC: WideString read Get_CC write Set_CC;
    property DeferredDeliveryTime: TDateTime read Get_DeferredDeliveryTime write Set_DeferredDeliveryTime;
    property DeleteAfterSubmit: WordBool read Get_DeleteAfterSubmit write Set_DeleteAfterSubmit;
    property ExpiryTime: TDateTime read Get_ExpiryTime write Set_ExpiryTime;
    property FlagDueBy: TDateTime read Get_FlagDueBy write Set_FlagDueBy;
    property FlagRequest: WideString read Get_FlagRequest write Set_FlagRequest;
    property FlagStatus: OlFlagStatus read Get_FlagStatus write Set_FlagStatus;
    property HTMLBody: WideString read Get_HTMLBody write Set_HTMLBody;
    property OriginatorDeliveryReportRequested: WordBool read Get_OriginatorDeliveryReportRequested write Set_OriginatorDeliveryReportRequested;
    property ReadReceiptRequested: WordBool read Get_ReadReceiptRequested write Set_ReadReceiptRequested;
    property ReceivedByEntryID: WideString read Get_ReceivedByEntryID;
    property ReceivedByName: WideString read Get_ReceivedByName;
    property ReceivedOnBehalfOfEntryID: WideString read Get_ReceivedOnBehalfOfEntryID;
    property ReceivedOnBehalfOfName: WideString read Get_ReceivedOnBehalfOfName;
    property ReceivedTime: TDateTime read Get_ReceivedTime;
    property RecipientReassignmentProhibited: WordBool read Get_RecipientReassignmentProhibited write Set_RecipientReassignmentProhibited;
    property Recipients: Recipients read Get_Recipients;
    property ReminderOverrideDefault: WordBool read Get_ReminderOverrideDefault write Set_ReminderOverrideDefault;
    property ReminderPlaySound: WordBool read Get_ReminderPlaySound write Set_ReminderPlaySound;
    property ReminderSet: WordBool read Get_ReminderSet write Set_ReminderSet;
    property ReminderSoundFile: WideString read Get_ReminderSoundFile write Set_ReminderSoundFile;
    property ReminderTime: TDateTime read Get_ReminderTime write Set_ReminderTime;
    property RemoteStatus: OlRemoteStatus read Get_RemoteStatus write Set_RemoteStatus;
    property ReplyRecipientNames: WideString read Get_ReplyRecipientNames;
    property ReplyRecipients: Recipients read Get_ReplyRecipients;
    property SaveSentMessageFolder: MAPIFolder read Get_SaveSentMessageFolder write Set_SaveSentMessageFolder;
    property SenderName: WideString read Get_SenderName;
    property Sent: WordBool read Get_Sent;
    property SentOn: TDateTime read Get_SentOn;
    property SentOnBehalfOfName: WideString read Get_SentOnBehalfOfName write Set_SentOnBehalfOfName;
    property Submitted: WordBool read Get_Submitted;
    property To_: WideString read Get_To_ write Set_To_;
    property VotingOptions: WideString read Get_VotingOptions write Set_VotingOptions;
    property VotingResponse: WideString read Get_VotingResponse write Set_VotingResponse;
  end;

// *********************************************************************//
// DispIntf:  _MailItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063034-0000-0000-C000-000000000046}
// *********************************************************************//
  _MailItemDisp = dispinterface
    ['{00063034-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property AlternateRecipientAllowed: WordBool dispid 2;
    property AutoForwarded: WordBool dispid 5;
    property BCC: WideString dispid 3586;
    property CC: WideString dispid 3587;
    property DeferredDeliveryTime: TDateTime dispid 15;
    property DeleteAfterSubmit: WordBool dispid 3585;
    property ExpiryTime: TDateTime dispid 21;
    property FlagDueBy: TDateTime dispid 48;
    property FlagRequest: WideString dispid 34096;
    property FlagStatus: OlFlagStatus dispid 4240;
    property HTMLBody: WideString dispid 62468;
    property OriginatorDeliveryReportRequested: WordBool dispid 35;
    property ReadReceiptRequested: WordBool dispid 41;
    property ReceivedByEntryID: WideString readonly dispid 63;
    property ReceivedByName: WideString readonly dispid 64;
    property ReceivedOnBehalfOfEntryID: WideString readonly dispid 67;
    property ReceivedOnBehalfOfName: WideString readonly dispid 68;
    property ReceivedTime: TDateTime readonly dispid 3590;
    property RecipientReassignmentProhibited: WordBool dispid 43;
    property Recipients: Recipients readonly dispid 63508;
    property ReminderOverrideDefault: WordBool dispid 34076;
    property ReminderPlaySound: WordBool dispid 34078;
    property ReminderSet: WordBool dispid 34051;
    property ReminderSoundFile: WideString dispid 34079;
    property ReminderTime: TDateTime dispid 34050;
    property RemoteStatus: OlRemoteStatus dispid 34065;
    property ReplyRecipientNames: WideString readonly dispid 80;
    property ReplyRecipients: Recipients readonly dispid 61459;
    property SaveSentMessageFolder: MAPIFolder dispid 62465;
    property SenderName: WideString readonly dispid 3098;
    property Sent: WordBool readonly dispid 62466;
    property SentOn: TDateTime readonly dispid 57;
    property SentOnBehalfOfName: WideString dispid 66;
    property Submitted: WordBool readonly dispid 62467;
    property To_: WideString dispid 3588;
    property VotingOptions: WideString dispid 61467;
    property VotingResponse: WideString dispid 34084;
    procedure ClearConversationIndex; dispid 63522;
    function Forward: MailItem; dispid 63507;
    function Reply: MailItem; dispid 63504;
    function ReplyAll: MailItem; dispid 63505;
    procedure Send; dispid 61557;
  end;

// *********************************************************************//
// Interface: RecurrencePattern
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063044-0000-0000-C000-000000000046}
// *********************************************************************//
  RecurrencePattern = interface(IDispatch)
    ['{00063044-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_DayOfMonth: Integer; safecall;
    procedure Set_DayOfMonth(DayOfMonth: Integer); safecall;
    function Get_DayOfWeekMask: OlDaysOfWeek; safecall;
    procedure Set_DayOfWeekMask(DayOfWeekMask: OlDaysOfWeek); safecall;
    function Get_Duration: Integer; safecall;
    procedure Set_Duration(Duration: Integer); safecall;
    function Get_EndTime: TDateTime; safecall;
    procedure Set_EndTime(EndTime: TDateTime); safecall;
    function Get_Exceptions: Exceptions; safecall;
    function Get_Instance: Integer; safecall;
    procedure Set_Instance(Instance: Integer); safecall;
    function Get_Interval: Integer; safecall;
    procedure Set_Interval(Interval: Integer); safecall;
    function Get_MonthOfYear: Integer; safecall;
    procedure Set_MonthOfYear(MonthOfYear: Integer); safecall;
    function Get_NoEndDate: WordBool; safecall;
    procedure Set_NoEndDate(NoEndDate: WordBool); safecall;
    function Get_Occurrences: Integer; safecall;
    procedure Set_Occurrences(Occurrences: Integer); safecall;
    function Get_PatternEndDate: TDateTime; safecall;
    procedure Set_PatternEndDate(PatternEndDate: TDateTime); safecall;
    function Get_PatternStartDate: TDateTime; safecall;
    procedure Set_PatternStartDate(PatternStartDate: TDateTime); safecall;
    function Get_RecurrenceType: OlRecurrenceType; safecall;
    procedure Set_RecurrenceType(RecurrenceType: OlRecurrenceType); safecall;
    function Get_Regenerate: WordBool; safecall;
    procedure Set_Regenerate(Regenerate: WordBool); safecall;
    function Get_StartTime: TDateTime; safecall;
    procedure Set_StartTime(StartTime: TDateTime); safecall;
    function GetOccurrence(StartDate: TDateTime): AppointmentItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property DayOfMonth: Integer read Get_DayOfMonth write Set_DayOfMonth;
    property DayOfWeekMask: OlDaysOfWeek read Get_DayOfWeekMask write Set_DayOfWeekMask;
    property Duration: Integer read Get_Duration write Set_Duration;
    property EndTime: TDateTime read Get_EndTime write Set_EndTime;
    property Exceptions: Exceptions read Get_Exceptions;
    property Instance: Integer read Get_Instance write Set_Instance;
    property Interval: Integer read Get_Interval write Set_Interval;
    property MonthOfYear: Integer read Get_MonthOfYear write Set_MonthOfYear;
    property NoEndDate: WordBool read Get_NoEndDate write Set_NoEndDate;
    property Occurrences: Integer read Get_Occurrences write Set_Occurrences;
    property PatternEndDate: TDateTime read Get_PatternEndDate write Set_PatternEndDate;
    property PatternStartDate: TDateTime read Get_PatternStartDate write Set_PatternStartDate;
    property RecurrenceType: OlRecurrenceType read Get_RecurrenceType write Set_RecurrenceType;
    property Regenerate: WordBool read Get_Regenerate write Set_Regenerate;
    property StartTime: TDateTime read Get_StartTime write Set_StartTime;
  end;

// *********************************************************************//
// DispIntf:  RecurrencePatternDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063044-0000-0000-C000-000000000046}
// *********************************************************************//
  RecurrencePatternDisp = dispinterface
    ['{00063044-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property DayOfMonth: Integer dispid 4096;
    property DayOfWeekMask: OlDaysOfWeek dispid 4097;
    property Duration: Integer dispid 4109;
    property EndTime: TDateTime dispid 4108;
    property Exceptions: Exceptions readonly dispid 4110;
    property Instance: Integer dispid 4099;
    property Interval: Integer dispid 4100;
    property MonthOfYear: Integer dispid 4102;
    property NoEndDate: WordBool dispid 4107;
    property Occurrences: Integer dispid 4101;
    property PatternEndDate: TDateTime dispid 4098;
    property PatternStartDate: TDateTime dispid 4104;
    property RecurrenceType: OlRecurrenceType dispid 4103;
    property Regenerate: WordBool dispid 4106;
    property StartTime: TDateTime dispid 4105;
    function GetOccurrence(StartDate: TDateTime): AppointmentItem; dispid 4111;
  end;

// *********************************************************************//
// Interface: Exceptions
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304C-0000-0000-C000-000000000046}
// *********************************************************************//
  Exceptions = interface(IDispatch)
    ['{0006304C-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): Exception; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ExceptionsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304C-0000-0000-C000-000000000046}
// *********************************************************************//
  ExceptionsDisp = dispinterface
    ['{0006304C-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): Exception; dispid 81;
  end;

// *********************************************************************//
// Interface: Exception
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304D-0000-0000-C000-000000000046}
// *********************************************************************//
  Exception = interface(IDispatch)
    ['{0006304D-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_AppointmentItem: AppointmentItem; safecall;
    function Get_Deleted: WordBool; safecall;
    function Get_OriginalDate: TDateTime; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property AppointmentItem: AppointmentItem read Get_AppointmentItem;
    property Deleted: WordBool read Get_Deleted;
    property OriginalDate: TDateTime read Get_OriginalDate;
  end;

// *********************************************************************//
// DispIntf:  ExceptionDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006304D-0000-0000-C000-000000000046}
// *********************************************************************//
  ExceptionDisp = dispinterface
    ['{0006304D-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property AppointmentItem: AppointmentItem readonly dispid 8193;
    property Deleted: WordBool readonly dispid 8194;
    property OriginalDate: TDateTime readonly dispid 8192;
  end;

// *********************************************************************//
// Interface: _MeetingItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063062-0000-0000-C000-000000000046}
// *********************************************************************//
  _MeetingItem = interface(IDispatch)
    ['{00063062-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_AutoForwarded: WordBool; safecall;
    procedure Set_AutoForwarded(AutoForwarded: WordBool); safecall;
    function Get_DeferredDeliveryTime: TDateTime; safecall;
    procedure Set_DeferredDeliveryTime(DeferredDeliveryTime: TDateTime); safecall;
    function Get_DeleteAfterSubmit: WordBool; safecall;
    procedure Set_DeleteAfterSubmit(DeleteAfterSubmit: WordBool); safecall;
    function Get_ExpiryTime: TDateTime; safecall;
    procedure Set_ExpiryTime(ExpiryTime: TDateTime); safecall;
    function Get_FlagDueBy: TDateTime; safecall;
    procedure Set_FlagDueBy(FlagDueBy: TDateTime); safecall;
    function Get_FlagRequest: WideString; safecall;
    procedure Set_FlagRequest(const FlagRequest: WideString); safecall;
    function Get_FlagStatus: OlFlagStatus; safecall;
    procedure Set_FlagStatus(FlagStatus: OlFlagStatus); safecall;
    function Get_OriginatorDeliveryReportRequested: WordBool; safecall;
    procedure Set_OriginatorDeliveryReportRequested(OriginatorDeliveryReportRequested: WordBool); safecall;
    function Get_ReceivedTime: TDateTime; safecall;
    procedure Set_ReceivedTime(ReceivedTime: TDateTime); safecall;
    function Get_Recipients: Recipients; safecall;
    function Get_ReminderSet: WordBool; safecall;
    procedure Set_ReminderSet(ReminderSet: WordBool); safecall;
    function Get_ReminderTime: TDateTime; safecall;
    procedure Set_ReminderTime(ReminderTime: TDateTime); safecall;
    function Get_ReplyRecipients: Recipients; safecall;
    function Get_SaveSentMessageFolder: MAPIFolder; safecall;
    procedure Set_SaveSentMessageFolder(const SaveSentMessageFolder: MAPIFolder); safecall;
    function Get_SenderName: WideString; safecall;
    function Get_Sent: WordBool; safecall;
    function Get_SentOn: TDateTime; safecall;
    procedure Set_SentOn(SentOn: TDateTime); safecall;
    function Get_Submitted: WordBool; safecall;
    function Forward: MeetingItem; safecall;
    function GetAssociatedAppointment(AddToCalendar: WordBool): AppointmentItem; safecall;
    function Reply: MailItem; safecall;
    function ReplyAll: MailItem; safecall;
    procedure Send; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property AutoForwarded: WordBool read Get_AutoForwarded write Set_AutoForwarded;
    property DeferredDeliveryTime: TDateTime read Get_DeferredDeliveryTime write Set_DeferredDeliveryTime;
    property DeleteAfterSubmit: WordBool read Get_DeleteAfterSubmit write Set_DeleteAfterSubmit;
    property ExpiryTime: TDateTime read Get_ExpiryTime write Set_ExpiryTime;
    property FlagDueBy: TDateTime read Get_FlagDueBy write Set_FlagDueBy;
    property FlagRequest: WideString read Get_FlagRequest write Set_FlagRequest;
    property FlagStatus: OlFlagStatus read Get_FlagStatus write Set_FlagStatus;
    property OriginatorDeliveryReportRequested: WordBool read Get_OriginatorDeliveryReportRequested write Set_OriginatorDeliveryReportRequested;
    property ReceivedTime: TDateTime read Get_ReceivedTime write Set_ReceivedTime;
    property Recipients: Recipients read Get_Recipients;
    property ReminderSet: WordBool read Get_ReminderSet write Set_ReminderSet;
    property ReminderTime: TDateTime read Get_ReminderTime write Set_ReminderTime;
    property ReplyRecipients: Recipients read Get_ReplyRecipients;
    property SaveSentMessageFolder: MAPIFolder read Get_SaveSentMessageFolder write Set_SaveSentMessageFolder;
    property SenderName: WideString read Get_SenderName;
    property Sent: WordBool read Get_Sent;
    property SentOn: TDateTime read Get_SentOn write Set_SentOn;
    property Submitted: WordBool read Get_Submitted;
  end;

// *********************************************************************//
// DispIntf:  _MeetingItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063062-0000-0000-C000-000000000046}
// *********************************************************************//
  _MeetingItemDisp = dispinterface
    ['{00063062-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property AutoForwarded: WordBool dispid 5;
    property DeferredDeliveryTime: TDateTime dispid 15;
    property DeleteAfterSubmit: WordBool dispid 3585;
    property ExpiryTime: TDateTime dispid 21;
    property FlagDueBy: TDateTime dispid 48;
    property FlagRequest: WideString dispid 34096;
    property FlagStatus: OlFlagStatus dispid 4240;
    property OriginatorDeliveryReportRequested: WordBool dispid 35;
    property ReceivedTime: TDateTime dispid 3590;
    property Recipients: Recipients readonly dispid 63508;
    property ReminderSet: WordBool dispid 34051;
    property ReminderTime: TDateTime dispid 34050;
    property ReplyRecipients: Recipients readonly dispid 61459;
    property SaveSentMessageFolder: MAPIFolder dispid 62465;
    property SenderName: WideString readonly dispid 3098;
    property Sent: WordBool readonly dispid 62466;
    property SentOn: TDateTime dispid 57;
    property Submitted: WordBool readonly dispid 62467;
    function Forward: MeetingItem; dispid 63507;
    function GetAssociatedAppointment(AddToCalendar: WordBool): AppointmentItem; dispid 63328;
    function Reply: MailItem; dispid 63504;
    function ReplyAll: MailItem; dispid 63505;
    procedure Send; dispid 61557;
  end;

// *********************************************************************//
// Interface: _ContactItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063021-0000-0000-C000-000000000046}
// *********************************************************************//
  _ContactItem = interface(IDispatch)
    ['{00063021-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_Account: WideString; safecall;
    procedure Set_Account(const Account: WideString); safecall;
    function Get_Anniversary: TDateTime; safecall;
    procedure Set_Anniversary(Anniversary: TDateTime); safecall;
    function Get_AssistantName: WideString; safecall;
    procedure Set_AssistantName(const AssistantName: WideString); safecall;
    function Get_AssistantTelephoneNumber: WideString; safecall;
    procedure Set_AssistantTelephoneNumber(const AssistantTelephoneNumber: WideString); safecall;
    function Get_Birthday: TDateTime; safecall;
    procedure Set_Birthday(Birthday: TDateTime); safecall;
    function Get_Business2TelephoneNumber: WideString; safecall;
    procedure Set_Business2TelephoneNumber(const Business2TelephoneNumber: WideString); safecall;
    function Get_BusinessAddress: WideString; safecall;
    procedure Set_BusinessAddress(const BusinessAddress: WideString); safecall;
    function Get_BusinessAddressCity: WideString; safecall;
    procedure Set_BusinessAddressCity(const BusinessAddressCity: WideString); safecall;
    function Get_BusinessAddressCountry: WideString; safecall;
    procedure Set_BusinessAddressCountry(const BusinessAddressCountry: WideString); safecall;
    function Get_BusinessAddressPostalCode: WideString; safecall;
    procedure Set_BusinessAddressPostalCode(const BusinessAddressPostalCode: WideString); safecall;
    function Get_BusinessAddressPostOfficeBox: WideString; safecall;
    procedure Set_BusinessAddressPostOfficeBox(const BusinessAddressPostOfficeBox: WideString); safecall;
    function Get_BusinessAddressState: WideString; safecall;
    procedure Set_BusinessAddressState(const BusinessAddressState: WideString); safecall;
    function Get_BusinessAddressStreet: WideString; safecall;
    procedure Set_BusinessAddressStreet(const BusinessAddressStreet: WideString); safecall;
    function Get_BusinessFaxNumber: WideString; safecall;
    procedure Set_BusinessFaxNumber(const BusinessFaxNumber: WideString); safecall;
    function Get_BusinessHomePage: WideString; safecall;
    procedure Set_BusinessHomePage(const BusinessHomePage: WideString); safecall;
    function Get_BusinessTelephoneNumber: WideString; safecall;
    procedure Set_BusinessTelephoneNumber(const BusinessTelephoneNumber: WideString); safecall;
    function Get_CallbackTelephoneNumber: WideString; safecall;
    procedure Set_CallbackTelephoneNumber(const CallbackTelephoneNumber: WideString); safecall;
    function Get_CarTelephoneNumber: WideString; safecall;
    procedure Set_CarTelephoneNumber(const CarTelephoneNumber: WideString); safecall;
    function Get_Children: WideString; safecall;
    procedure Set_Children(const Children: WideString); safecall;
    function Get_CompanyAndFullName: WideString; safecall;
    function Get_CompanyLastFirstNoSpace: WideString; safecall;
    function Get_CompanyLastFirstSpaceOnly: WideString; safecall;
    function Get_CompanyMainTelephoneNumber: WideString; safecall;
    procedure Set_CompanyMainTelephoneNumber(const CompanyMainTelephoneNumber: WideString); safecall;
    function Get_CompanyName: WideString; safecall;
    procedure Set_CompanyName(const CompanyName: WideString); safecall;
    function Get_ComputerNetworkName: WideString; safecall;
    procedure Set_ComputerNetworkName(const ComputerNetworkName: WideString); safecall;
    function Get_CustomerID: WideString; safecall;
    procedure Set_CustomerID(const CustomerID: WideString); safecall;
    function Get_Department: WideString; safecall;
    procedure Set_Department(const Department: WideString); safecall;
    function Get_Email1Address: WideString; safecall;
    procedure Set_Email1Address(const Email1Address: WideString); safecall;
    function Get_Email1AddressType: WideString; safecall;
    procedure Set_Email1AddressType(const Email1AddressType: WideString); safecall;
    function Get_Email1DisplayName: WideString; safecall;
    function Get_Email1EntryID: WideString; safecall;
    function Get_Email2Address: WideString; safecall;
    procedure Set_Email2Address(const Email2Address: WideString); safecall;
    function Get_Email2AddressType: WideString; safecall;
    procedure Set_Email2AddressType(const Email2AddressType: WideString); safecall;
    function Get_Email2DisplayName: WideString; safecall;
    function Get_Email2EntryID: WideString; safecall;
    function Get_Email3Address: WideString; safecall;
    procedure Set_Email3Address(const Email3Address: WideString); safecall;
    function Get_Email3AddressType: WideString; safecall;
    procedure Set_Email3AddressType(const Email3AddressType: WideString); safecall;
    function Get_Email3DisplayName: WideString; safecall;
    function Get_Email3EntryID: WideString; safecall;
    function Get_FileAs: WideString; safecall;
    procedure Set_FileAs(const FileAs: WideString); safecall;
    function Get_FirstName: WideString; safecall;
    procedure Set_FirstName(const FirstName: WideString); safecall;
    function Get_FTPSite: WideString; safecall;
    procedure Set_FTPSite(const FTPSite: WideString); safecall;
    function Get_FullName: WideString; safecall;
    procedure Set_FullName(const FullName: WideString); safecall;
    function Get_FullNameAndCompany: WideString; safecall;
    function Get_Gender: OlGender; safecall;
    procedure Set_Gender(Gender: OlGender); safecall;
    function Get_GovernmentIDNumber: WideString; safecall;
    procedure Set_GovernmentIDNumber(const GovernmentIDNumber: WideString); safecall;
    function Get_Hobby: WideString; safecall;
    procedure Set_Hobby(const Hobby: WideString); safecall;
    function Get_Home2TelephoneNumber: WideString; safecall;
    procedure Set_Home2TelephoneNumber(const Home2TelephoneNumber: WideString); safecall;
    function Get_HomeAddress: WideString; safecall;
    procedure Set_HomeAddress(const HomeAddress: WideString); safecall;
    function Get_HomeAddressCity: WideString; safecall;
    procedure Set_HomeAddressCity(const HomeAddressCity: WideString); safecall;
    function Get_HomeAddressCountry: WideString; safecall;
    procedure Set_HomeAddressCountry(const HomeAddressCountry: WideString); safecall;
    function Get_HomeAddressPostalCode: WideString; safecall;
    procedure Set_HomeAddressPostalCode(const HomeAddressPostalCode: WideString); safecall;
    function Get_HomeAddressPostOfficeBox: WideString; safecall;
    procedure Set_HomeAddressPostOfficeBox(const HomeAddressPostOfficeBox: WideString); safecall;
    function Get_HomeAddressState: WideString; safecall;
    procedure Set_HomeAddressState(const HomeAddressState: WideString); safecall;
    function Get_HomeAddressStreet: WideString; safecall;
    procedure Set_HomeAddressStreet(const HomeAddressStreet: WideString); safecall;
    function Get_HomeFaxNumber: WideString; safecall;
    procedure Set_HomeFaxNumber(const HomeFaxNumber: WideString); safecall;
    function Get_HomeTelephoneNumber: WideString; safecall;
    procedure Set_HomeTelephoneNumber(const HomeTelephoneNumber: WideString); safecall;
    function Get_Initials: WideString; safecall;
    procedure Set_Initials(const Initials: WideString); safecall;
    function Get_InternetFreeBusyAddress: WideString; safecall;
    procedure Set_InternetFreeBusyAddress(const InternetFreeBusyAddress: WideString); safecall;
    function Get_ISDNNumber: WideString; safecall;
    procedure Set_ISDNNumber(const ISDNNumber: WideString); safecall;
    function Get_JobTitle: WideString; safecall;
    procedure Set_JobTitle(const JobTitle: WideString); safecall;
    function Get_Journal: WordBool; safecall;
    procedure Set_Journal(Journal: WordBool); safecall;
    function Get_Language: WideString; safecall;
    procedure Set_Language(const Language: WideString); safecall;
    function Get_LastFirstAndSuffix: WideString; safecall;
    function Get_LastFirstNoSpace: WideString; safecall;
    function Get_LastFirstNoSpaceCompany: WideString; safecall;
    function Get_LastFirstSpaceOnly: WideString; safecall;
    function Get_LastFirstSpaceOnlyCompany: WideString; safecall;
    function Get_LastName: WideString; safecall;
    procedure Set_LastName(const LastName: WideString); safecall;
    function Get_LastNameAndFirstName: WideString; safecall;
    function Get_MailingAddress: WideString; safecall;
    procedure Set_MailingAddress(const MailingAddress: WideString); safecall;
    function Get_MailingAddressCity: WideString; safecall;
    procedure Set_MailingAddressCity(const MailingAddressCity: WideString); safecall;
    function Get_MailingAddressCountry: WideString; safecall;
    procedure Set_MailingAddressCountry(const MailingAddressCountry: WideString); safecall;
    function Get_MailingAddressPostalCode: WideString; safecall;
    procedure Set_MailingAddressPostalCode(const MailingAddressPostalCode: WideString); safecall;
    function Get_MailingAddressPostOfficeBox: WideString; safecall;
    procedure Set_MailingAddressPostOfficeBox(const MailingAddressPostOfficeBox: WideString); safecall;
    function Get_MailingAddressState: WideString; safecall;
    procedure Set_MailingAddressState(const MailingAddressState: WideString); safecall;
    function Get_MailingAddressStreet: WideString; safecall;
    procedure Set_MailingAddressStreet(const MailingAddressStreet: WideString); safecall;
    function Get_ManagerName: WideString; safecall;
    procedure Set_ManagerName(const ManagerName: WideString); safecall;
    function Get_MiddleName: WideString; safecall;
    procedure Set_MiddleName(const MiddleName: WideString); safecall;
    function Get_MobileTelephoneNumber: WideString; safecall;
    procedure Set_MobileTelephoneNumber(const MobileTelephoneNumber: WideString); safecall;
    function Get_NetMeetingAlias: WideString; safecall;
    procedure Set_NetMeetingAlias(const NetMeetingAlias: WideString); safecall;
    function Get_NetMeetingServer: WideString; safecall;
    procedure Set_NetMeetingServer(const NetMeetingServer: WideString); safecall;
    function Get_NickName: WideString; safecall;
    procedure Set_NickName(const NickName: WideString); safecall;
    function Get_OfficeLocation: WideString; safecall;
    procedure Set_OfficeLocation(const OfficeLocation: WideString); safecall;
    function Get_OrganizationalIDNumber: WideString; safecall;
    procedure Set_OrganizationalIDNumber(const OrganizationalIDNumber: WideString); safecall;
    function Get_OtherAddress: WideString; safecall;
    procedure Set_OtherAddress(const OtherAddress: WideString); safecall;
    function Get_OtherAddressCity: WideString; safecall;
    procedure Set_OtherAddressCity(const OtherAddressCity: WideString); safecall;
    function Get_OtherAddressCountry: WideString; safecall;
    procedure Set_OtherAddressCountry(const OtherAddressCountry: WideString); safecall;
    function Get_OtherAddressPostalCode: WideString; safecall;
    procedure Set_OtherAddressPostalCode(const OtherAddressPostalCode: WideString); safecall;
    function Get_OtherAddressPostOfficeBox: WideString; safecall;
    procedure Set_OtherAddressPostOfficeBox(const OtherAddressPostOfficeBox: WideString); safecall;
    function Get_OtherAddressState: WideString; safecall;
    procedure Set_OtherAddressState(const OtherAddressState: WideString); safecall;
    function Get_OtherAddressStreet: WideString; safecall;
    procedure Set_OtherAddressStreet(const OtherAddressStreet: WideString); safecall;
    function Get_OtherFaxNumber: WideString; safecall;
    procedure Set_OtherFaxNumber(const OtherFaxNumber: WideString); safecall;
    function Get_OtherTelephoneNumber: WideString; safecall;
    procedure Set_OtherTelephoneNumber(const OtherTelephoneNumber: WideString); safecall;
    function Get_PagerNumber: WideString; safecall;
    procedure Set_PagerNumber(const PagerNumber: WideString); safecall;
    function Get_PersonalHomePage: WideString; safecall;
    procedure Set_PersonalHomePage(const PersonalHomePage: WideString); safecall;
    function Get_PrimaryTelephoneNumber: WideString; safecall;
    procedure Set_PrimaryTelephoneNumber(const PrimaryTelephoneNumber: WideString); safecall;
    function Get_Profession: WideString; safecall;
    procedure Set_Profession(const Profession: WideString); safecall;
    function Get_RadioTelephoneNumber: WideString; safecall;
    procedure Set_RadioTelephoneNumber(const RadioTelephoneNumber: WideString); safecall;
    function Get_ReferredBy: WideString; safecall;
    procedure Set_ReferredBy(const ReferredBy: WideString); safecall;
    function Get_SelectedMailingAddress: OlMailingAddress; safecall;
    procedure Set_SelectedMailingAddress(SelectedMailingAddress: OlMailingAddress); safecall;
    function Get_Spouse: WideString; safecall;
    procedure Set_Spouse(const Spouse: WideString); safecall;
    function Get_Suffix: WideString; safecall;
    procedure Set_Suffix(const Suffix: WideString); safecall;
    function Get_TelexNumber: WideString; safecall;
    procedure Set_TelexNumber(const TelexNumber: WideString); safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Title: WideString); safecall;
    function Get_TTYTDDTelephoneNumber: WideString; safecall;
    procedure Set_TTYTDDTelephoneNumber(const TTYTDDTelephoneNumber: WideString); safecall;
    function Get_User1: WideString; safecall;
    procedure Set_User1(const User1: WideString); safecall;
    function Get_User2: WideString; safecall;
    procedure Set_User2(const User2: WideString); safecall;
    function Get_User3: WideString; safecall;
    procedure Set_User3(const User3: WideString); safecall;
    function Get_User4: WideString; safecall;
    procedure Set_User4(const User4: WideString); safecall;
    function Get_UserCertificate: WideString; safecall;
    procedure Set_UserCertificate(const UserCertificate: WideString); safecall;
    function Get_WebPage: WideString; safecall;
    procedure Set_WebPage(const WebPage: WideString); safecall;
    function Get_YomiCompanyName: WideString; safecall;
    procedure Set_YomiCompanyName(const YomiCompanyName: WideString); safecall;
    function Get_YomiFirstName: WideString; safecall;
    procedure Set_YomiFirstName(const YomiFirstName: WideString); safecall;
    function Get_YomiLastName: WideString; safecall;
    procedure Set_YomiLastName(const YomiLastName: WideString); safecall;
    function ForwardAsVcard: MailItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property Account: WideString read Get_Account write Set_Account;
    property Anniversary: TDateTime read Get_Anniversary write Set_Anniversary;
    property AssistantName: WideString read Get_AssistantName write Set_AssistantName;
    property AssistantTelephoneNumber: WideString read Get_AssistantTelephoneNumber write Set_AssistantTelephoneNumber;
    property Birthday: TDateTime read Get_Birthday write Set_Birthday;
    property Business2TelephoneNumber: WideString read Get_Business2TelephoneNumber write Set_Business2TelephoneNumber;
    property BusinessAddress: WideString read Get_BusinessAddress write Set_BusinessAddress;
    property BusinessAddressCity: WideString read Get_BusinessAddressCity write Set_BusinessAddressCity;
    property BusinessAddressCountry: WideString read Get_BusinessAddressCountry write Set_BusinessAddressCountry;
    property BusinessAddressPostalCode: WideString read Get_BusinessAddressPostalCode write Set_BusinessAddressPostalCode;
    property BusinessAddressPostOfficeBox: WideString read Get_BusinessAddressPostOfficeBox write Set_BusinessAddressPostOfficeBox;
    property BusinessAddressState: WideString read Get_BusinessAddressState write Set_BusinessAddressState;
    property BusinessAddressStreet: WideString read Get_BusinessAddressStreet write Set_BusinessAddressStreet;
    property BusinessFaxNumber: WideString read Get_BusinessFaxNumber write Set_BusinessFaxNumber;
    property BusinessHomePage: WideString read Get_BusinessHomePage write Set_BusinessHomePage;
    property BusinessTelephoneNumber: WideString read Get_BusinessTelephoneNumber write Set_BusinessTelephoneNumber;
    property CallbackTelephoneNumber: WideString read Get_CallbackTelephoneNumber write Set_CallbackTelephoneNumber;
    property CarTelephoneNumber: WideString read Get_CarTelephoneNumber write Set_CarTelephoneNumber;
    property Children: WideString read Get_Children write Set_Children;
    property CompanyAndFullName: WideString read Get_CompanyAndFullName;
    property CompanyLastFirstNoSpace: WideString read Get_CompanyLastFirstNoSpace;
    property CompanyLastFirstSpaceOnly: WideString read Get_CompanyLastFirstSpaceOnly;
    property CompanyMainTelephoneNumber: WideString read Get_CompanyMainTelephoneNumber write Set_CompanyMainTelephoneNumber;
    property CompanyName: WideString read Get_CompanyName write Set_CompanyName;
    property ComputerNetworkName: WideString read Get_ComputerNetworkName write Set_ComputerNetworkName;
    property CustomerID: WideString read Get_CustomerID write Set_CustomerID;
    property Department: WideString read Get_Department write Set_Department;
    property Email1Address: WideString read Get_Email1Address write Set_Email1Address;
    property Email1AddressType: WideString read Get_Email1AddressType write Set_Email1AddressType;
    property Email1DisplayName: WideString read Get_Email1DisplayName;
    property Email1EntryID: WideString read Get_Email1EntryID;
    property Email2Address: WideString read Get_Email2Address write Set_Email2Address;
    property Email2AddressType: WideString read Get_Email2AddressType write Set_Email2AddressType;
    property Email2DisplayName: WideString read Get_Email2DisplayName;
    property Email2EntryID: WideString read Get_Email2EntryID;
    property Email3Address: WideString read Get_Email3Address write Set_Email3Address;
    property Email3AddressType: WideString read Get_Email3AddressType write Set_Email3AddressType;
    property Email3DisplayName: WideString read Get_Email3DisplayName;
    property Email3EntryID: WideString read Get_Email3EntryID;
    property FileAs: WideString read Get_FileAs write Set_FileAs;
    property FirstName: WideString read Get_FirstName write Set_FirstName;
    property FTPSite: WideString read Get_FTPSite write Set_FTPSite;
    property FullName: WideString read Get_FullName write Set_FullName;
    property FullNameAndCompany: WideString read Get_FullNameAndCompany;
    property Gender: OlGender read Get_Gender write Set_Gender;
    property GovernmentIDNumber: WideString read Get_GovernmentIDNumber write Set_GovernmentIDNumber;
    property Hobby: WideString read Get_Hobby write Set_Hobby;
    property Home2TelephoneNumber: WideString read Get_Home2TelephoneNumber write Set_Home2TelephoneNumber;
    property HomeAddress: WideString read Get_HomeAddress write Set_HomeAddress;
    property HomeAddressCity: WideString read Get_HomeAddressCity write Set_HomeAddressCity;
    property HomeAddressCountry: WideString read Get_HomeAddressCountry write Set_HomeAddressCountry;
    property HomeAddressPostalCode: WideString read Get_HomeAddressPostalCode write Set_HomeAddressPostalCode;
    property HomeAddressPostOfficeBox: WideString read Get_HomeAddressPostOfficeBox write Set_HomeAddressPostOfficeBox;
    property HomeAddressState: WideString read Get_HomeAddressState write Set_HomeAddressState;
    property HomeAddressStreet: WideString read Get_HomeAddressStreet write Set_HomeAddressStreet;
    property HomeFaxNumber: WideString read Get_HomeFaxNumber write Set_HomeFaxNumber;
    property HomeTelephoneNumber: WideString read Get_HomeTelephoneNumber write Set_HomeTelephoneNumber;
    property Initials: WideString read Get_Initials write Set_Initials;
    property InternetFreeBusyAddress: WideString read Get_InternetFreeBusyAddress write Set_InternetFreeBusyAddress;
    property ISDNNumber: WideString read Get_ISDNNumber write Set_ISDNNumber;
    property JobTitle: WideString read Get_JobTitle write Set_JobTitle;
    property Journal: WordBool read Get_Journal write Set_Journal;
    property Language: WideString read Get_Language write Set_Language;
    property LastFirstAndSuffix: WideString read Get_LastFirstAndSuffix;
    property LastFirstNoSpace: WideString read Get_LastFirstNoSpace;
    property LastFirstNoSpaceCompany: WideString read Get_LastFirstNoSpaceCompany;
    property LastFirstSpaceOnly: WideString read Get_LastFirstSpaceOnly;
    property LastFirstSpaceOnlyCompany: WideString read Get_LastFirstSpaceOnlyCompany;
    property LastName: WideString read Get_LastName write Set_LastName;
    property LastNameAndFirstName: WideString read Get_LastNameAndFirstName;
    property MailingAddress: WideString read Get_MailingAddress write Set_MailingAddress;
    property MailingAddressCity: WideString read Get_MailingAddressCity write Set_MailingAddressCity;
    property MailingAddressCountry: WideString read Get_MailingAddressCountry write Set_MailingAddressCountry;
    property MailingAddressPostalCode: WideString read Get_MailingAddressPostalCode write Set_MailingAddressPostalCode;
    property MailingAddressPostOfficeBox: WideString read Get_MailingAddressPostOfficeBox write Set_MailingAddressPostOfficeBox;
    property MailingAddressState: WideString read Get_MailingAddressState write Set_MailingAddressState;
    property MailingAddressStreet: WideString read Get_MailingAddressStreet write Set_MailingAddressStreet;
    property ManagerName: WideString read Get_ManagerName write Set_ManagerName;
    property MiddleName: WideString read Get_MiddleName write Set_MiddleName;
    property MobileTelephoneNumber: WideString read Get_MobileTelephoneNumber write Set_MobileTelephoneNumber;
    property NetMeetingAlias: WideString read Get_NetMeetingAlias write Set_NetMeetingAlias;
    property NetMeetingServer: WideString read Get_NetMeetingServer write Set_NetMeetingServer;
    property NickName: WideString read Get_NickName write Set_NickName;
    property OfficeLocation: WideString read Get_OfficeLocation write Set_OfficeLocation;
    property OrganizationalIDNumber: WideString read Get_OrganizationalIDNumber write Set_OrganizationalIDNumber;
    property OtherAddress: WideString read Get_OtherAddress write Set_OtherAddress;
    property OtherAddressCity: WideString read Get_OtherAddressCity write Set_OtherAddressCity;
    property OtherAddressCountry: WideString read Get_OtherAddressCountry write Set_OtherAddressCountry;
    property OtherAddressPostalCode: WideString read Get_OtherAddressPostalCode write Set_OtherAddressPostalCode;
    property OtherAddressPostOfficeBox: WideString read Get_OtherAddressPostOfficeBox write Set_OtherAddressPostOfficeBox;
    property OtherAddressState: WideString read Get_OtherAddressState write Set_OtherAddressState;
    property OtherAddressStreet: WideString read Get_OtherAddressStreet write Set_OtherAddressStreet;
    property OtherFaxNumber: WideString read Get_OtherFaxNumber write Set_OtherFaxNumber;
    property OtherTelephoneNumber: WideString read Get_OtherTelephoneNumber write Set_OtherTelephoneNumber;
    property PagerNumber: WideString read Get_PagerNumber write Set_PagerNumber;
    property PersonalHomePage: WideString read Get_PersonalHomePage write Set_PersonalHomePage;
    property PrimaryTelephoneNumber: WideString read Get_PrimaryTelephoneNumber write Set_PrimaryTelephoneNumber;
    property Profession: WideString read Get_Profession write Set_Profession;
    property RadioTelephoneNumber: WideString read Get_RadioTelephoneNumber write Set_RadioTelephoneNumber;
    property ReferredBy: WideString read Get_ReferredBy write Set_ReferredBy;
    property SelectedMailingAddress: OlMailingAddress read Get_SelectedMailingAddress write Set_SelectedMailingAddress;
    property Spouse: WideString read Get_Spouse write Set_Spouse;
    property Suffix: WideString read Get_Suffix write Set_Suffix;
    property TelexNumber: WideString read Get_TelexNumber write Set_TelexNumber;
    property Title: WideString read Get_Title write Set_Title;
    property TTYTDDTelephoneNumber: WideString read Get_TTYTDDTelephoneNumber write Set_TTYTDDTelephoneNumber;
    property User1: WideString read Get_User1 write Set_User1;
    property User2: WideString read Get_User2 write Set_User2;
    property User3: WideString read Get_User3 write Set_User3;
    property User4: WideString read Get_User4 write Set_User4;
    property UserCertificate: WideString read Get_UserCertificate write Set_UserCertificate;
    property WebPage: WideString read Get_WebPage write Set_WebPage;
    property YomiCompanyName: WideString read Get_YomiCompanyName write Set_YomiCompanyName;
    property YomiFirstName: WideString read Get_YomiFirstName write Set_YomiFirstName;
    property YomiLastName: WideString read Get_YomiLastName write Set_YomiLastName;
  end;

// *********************************************************************//
// DispIntf:  _ContactItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063021-0000-0000-C000-000000000046}
// *********************************************************************//
  _ContactItemDisp = dispinterface
    ['{00063021-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property Account: WideString dispid 14848;
    property Anniversary: TDateTime dispid 14913;
    property AssistantName: WideString dispid 14896;
    property AssistantTelephoneNumber: WideString dispid 14894;
    property Birthday: TDateTime dispid 14914;
    property Business2TelephoneNumber: WideString dispid 14875;
    property BusinessAddress: WideString dispid 32795;
    property BusinessAddressCity: WideString dispid 32838;
    property BusinessAddressCountry: WideString dispid 32841;
    property BusinessAddressPostalCode: WideString dispid 32840;
    property BusinessAddressPostOfficeBox: WideString dispid 32842;
    property BusinessAddressState: WideString dispid 32839;
    property BusinessAddressStreet: WideString dispid 32837;
    property BusinessFaxNumber: WideString dispid 14884;
    property BusinessHomePage: WideString dispid 14929;
    property BusinessTelephoneNumber: WideString dispid 14856;
    property CallbackTelephoneNumber: WideString dispid 14850;
    property CarTelephoneNumber: WideString dispid 14878;
    property Children: WideString dispid 32780;
    property CompanyAndFullName: WideString readonly dispid 32792;
    property CompanyLastFirstNoSpace: WideString readonly dispid 32818;
    property CompanyLastFirstSpaceOnly: WideString readonly dispid 32819;
    property CompanyMainTelephoneNumber: WideString dispid 14935;
    property CompanyName: WideString dispid 14870;
    property ComputerNetworkName: WideString dispid 14921;
    property CustomerID: WideString dispid 14922;
    property Department: WideString dispid 14872;
    property Email1Address: WideString dispid 32899;
    property Email1AddressType: WideString dispid 32898;
    property Email1DisplayName: WideString readonly dispid 32896;
    property Email1EntryID: WideString readonly dispid 32901;
    property Email2Address: WideString dispid 32915;
    property Email2AddressType: WideString dispid 32914;
    property Email2DisplayName: WideString readonly dispid 32912;
    property Email2EntryID: WideString readonly dispid 32917;
    property Email3Address: WideString dispid 32931;
    property Email3AddressType: WideString dispid 32930;
    property Email3DisplayName: WideString readonly dispid 32928;
    property Email3EntryID: WideString readonly dispid 32933;
    property FileAs: WideString dispid 32773;
    property FirstName: WideString dispid 14854;
    property FTPSite: WideString dispid 14924;
    property FullName: WideString dispid 12289;
    property FullNameAndCompany: WideString readonly dispid 32793;
    property Gender: OlGender dispid 14925;
    property GovernmentIDNumber: WideString dispid 14855;
    property Hobby: WideString dispid 14915;
    property Home2TelephoneNumber: WideString dispid 14895;
    property HomeAddress: WideString dispid 32794;
    property HomeAddressCity: WideString dispid 14937;
    property HomeAddressCountry: WideString dispid 14938;
    property HomeAddressPostalCode: WideString dispid 14939;
    property HomeAddressPostOfficeBox: WideString dispid 14942;
    property HomeAddressState: WideString dispid 14940;
    property HomeAddressStreet: WideString dispid 14941;
    property HomeFaxNumber: WideString dispid 14885;
    property HomeTelephoneNumber: WideString dispid 14857;
    property Initials: WideString dispid 14858;
    property InternetFreeBusyAddress: WideString dispid 32984;
    property ISDNNumber: WideString dispid 14893;
    property JobTitle: WideString dispid 14871;
    property Journal: WordBool dispid 32805;
    property Language: WideString dispid 14860;
    property LastFirstAndSuffix: WideString readonly dispid 32822;
    property LastFirstNoSpace: WideString readonly dispid 32816;
    property LastFirstNoSpaceCompany: WideString readonly dispid 32820;
    property LastFirstSpaceOnly: WideString readonly dispid 32817;
    property LastFirstSpaceOnlyCompany: WideString readonly dispid 32821;
    property LastName: WideString dispid 14865;
    property LastNameAndFirstName: WideString readonly dispid 32791;
    property MailingAddress: WideString dispid 14869;
    property MailingAddressCity: WideString dispid 14887;
    property MailingAddressCountry: WideString dispid 14886;
    property MailingAddressPostalCode: WideString dispid 14890;
    property MailingAddressPostOfficeBox: WideString dispid 14891;
    property MailingAddressState: WideString dispid 14888;
    property MailingAddressStreet: WideString dispid 14889;
    property ManagerName: WideString dispid 14926;
    property MiddleName: WideString dispid 14916;
    property MobileTelephoneNumber: WideString dispid 14876;
    property NetMeetingAlias: WideString dispid 32863;
    property NetMeetingServer: WideString dispid 32864;
    property NickName: WideString dispid 14927;
    property OfficeLocation: WideString dispid 14873;
    property OrganizationalIDNumber: WideString dispid 14864;
    property OtherAddress: WideString dispid 32796;
    property OtherAddressCity: WideString dispid 14943;
    property OtherAddressCountry: WideString dispid 14944;
    property OtherAddressPostalCode: WideString dispid 14945;
    property OtherAddressPostOfficeBox: WideString dispid 14948;
    property OtherAddressState: WideString dispid 14946;
    property OtherAddressStreet: WideString dispid 14947;
    property OtherFaxNumber: WideString dispid 14883;
    property OtherTelephoneNumber: WideString dispid 14879;
    property PagerNumber: WideString dispid 14881;
    property PersonalHomePage: WideString dispid 14928;
    property PrimaryTelephoneNumber: WideString dispid 14874;
    property Profession: WideString dispid 14918;
    property RadioTelephoneNumber: WideString dispid 14877;
    property ReferredBy: WideString dispid 14919;
    property SelectedMailingAddress: OlMailingAddress dispid 32802;
    property Spouse: WideString dispid 14920;
    property Suffix: WideString dispid 14853;
    property TelexNumber: WideString dispid 14892;
    property Title: WideString dispid 14917;
    property TTYTDDTelephoneNumber: WideString dispid 14923;
    property User1: WideString dispid 32847;
    property User2: WideString dispid 32848;
    property User3: WideString dispid 32849;
    property User4: WideString dispid 32850;
    property UserCertificate: WideString dispid 32790;
    property WebPage: WideString dispid 32811;
    property YomiCompanyName: WideString dispid 32814;
    property YomiFirstName: WideString dispid 32812;
    property YomiLastName: WideString dispid 32813;
    function ForwardAsVcard: MailItem; dispid 63649;
  end;

// *********************************************************************//
// Interface: _DocumentItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063020-0000-0000-C000-000000000046}
// *********************************************************************//
  _DocumentItem = interface(IDispatch)
    ['{00063020-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
  end;

// *********************************************************************//
// DispIntf:  _DocumentItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063020-0000-0000-C000-000000000046}
// *********************************************************************//
  _DocumentItemDisp = dispinterface
    ['{00063020-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
  end;

// *********************************************************************//
// Interface: _JournalItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063022-0000-0000-C000-000000000046}
// *********************************************************************//
  _JournalItem = interface(IDispatch)
    ['{00063022-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_ContactNames: WideString; safecall;
    procedure Set_ContactNames(const ContactNames: WideString); safecall;
    function Get_DocPosted: WordBool; safecall;
    procedure Set_DocPosted(DocPosted: WordBool); safecall;
    function Get_DocPrinted: WordBool; safecall;
    procedure Set_DocPrinted(DocPrinted: WordBool); safecall;
    function Get_DocRouted: WordBool; safecall;
    procedure Set_DocRouted(DocRouted: WordBool); safecall;
    function Get_DocSaved: WordBool; safecall;
    procedure Set_DocSaved(DocSaved: WordBool); safecall;
    function Get_Duration: Integer; safecall;
    procedure Set_Duration(Duration: Integer); safecall;
    function Get_End_: TDateTime; safecall;
    procedure Set_End_(End_: TDateTime); safecall;
    function Get_Type_: WideString; safecall;
    procedure Set_Type_(const Type_: WideString); safecall;
    function Get_Recipients: Recipients; safecall;
    function Get_Start: TDateTime; safecall;
    procedure Set_Start(Start: TDateTime); safecall;
    function Forward: MailItem; safecall;
    function Reply: MailItem; safecall;
    function ReplyAll: MailItem; safecall;
    procedure StartTimer; safecall;
    procedure StopTimer; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property ContactNames: WideString read Get_ContactNames write Set_ContactNames;
    property DocPosted: WordBool read Get_DocPosted write Set_DocPosted;
    property DocPrinted: WordBool read Get_DocPrinted write Set_DocPrinted;
    property DocRouted: WordBool read Get_DocRouted write Set_DocRouted;
    property DocSaved: WordBool read Get_DocSaved write Set_DocSaved;
    property Duration: Integer read Get_Duration write Set_Duration;
    property End_: TDateTime read Get_End_ write Set_End_;
    property Type_: WideString read Get_Type_ write Set_Type_;
    property Recipients: Recipients read Get_Recipients;
    property Start: TDateTime read Get_Start write Set_Start;
  end;

// *********************************************************************//
// DispIntf:  _JournalItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063022-0000-0000-C000-000000000046}
// *********************************************************************//
  _JournalItemDisp = dispinterface
    ['{00063022-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property ContactNames: WideString dispid 3588;
    property DocPosted: WordBool dispid 34577;
    property DocPrinted: WordBool dispid 34574;
    property DocRouted: WordBool dispid 34576;
    property DocSaved: WordBool dispid 34575;
    property Duration: Integer dispid 34567;
    property End_: TDateTime dispid 34568;
    property Type_: WideString dispid 34560;
    property Recipients: Recipients readonly dispid 63508;
    property Start: TDateTime dispid 34566;
    function Forward: MailItem; dispid 63507;
    function Reply: MailItem; dispid 63504;
    function ReplyAll: MailItem; dispid 63505;
    procedure StartTimer; dispid 63269;
    procedure StopTimer; dispid 63270;
  end;

// *********************************************************************//
// Interface: _NoteItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063025-0000-0000-C000-000000000046}
// *********************************************************************//
  _NoteItem = interface(IDispatch)
    ['{00063025-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Color: OlNoteColor; safecall;
    procedure Set_Color(Color: OlNoteColor); safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(Height: Integer); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(Left: Integer); safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(Top: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(Width: Integer); safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Color: OlNoteColor read Get_Color write Set_Color;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property GetInspector: Inspector read Get_GetInspector;
    property Height: Integer read Get_Height write Set_Height;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property Left: Integer read Get_Left write Set_Left;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Saved: WordBool read Get_Saved;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
  end;

// *********************************************************************//
// DispIntf:  _NoteItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063025-0000-0000-C000-000000000046}
// *********************************************************************//
  _NoteItemDisp = dispinterface
    ['{00063025-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Color: OlNoteColor dispid 35584;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property GetInspector: Inspector readonly dispid 61502;
    property Height: Integer dispid 35587;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property Left: Integer dispid 35588;
    property MessageClass: WideString dispid 26;
    property Saved: WordBool readonly dispid 61603;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString readonly dispid 63392;
    property Top: Integer dispid 35589;
    property Width: Integer dispid 35586;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
  end;

// *********************************************************************//
// Interface: Pages
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303F-0000-0000-C000-000000000046}
// *********************************************************************//
  Pages = interface(IDispatch)
    ['{0006303F-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: OleVariant): IDispatch; safecall;
    function Add(Name: OleVariant): IDispatch; safecall;
    procedure Remove(Index: Integer); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  PagesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0006303F-0000-0000-C000-000000000046}
// *********************************************************************//
  PagesDisp = dispinterface
    ['{0006303F-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Count: Integer readonly dispid 80;
    function Item(Index: OleVariant): IDispatch; dispid 81;
    function Add(Name: OleVariant): IDispatch; dispid 300;
    procedure Remove(Index: Integer); dispid 301;
  end;

// *********************************************************************//
// Interface: _PostItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063024-0000-0000-C000-000000000046}
// *********************************************************************//
  _PostItem = interface(IDispatch)
    ['{00063024-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_ExpiryTime: TDateTime; safecall;
    procedure Set_ExpiryTime(ExpiryTime: TDateTime); safecall;
    function Get_HTMLBody: WideString; safecall;
    procedure Set_HTMLBody(const HTMLBody: WideString); safecall;
    function Get_ReceivedTime: TDateTime; safecall;
    function Get_SenderName: WideString; safecall;
    function Get_SentOn: TDateTime; safecall;
    procedure ClearConversationIndex; safecall;
    function Forward: MailItem; safecall;
    procedure Post; safecall;
    function Reply: MailItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property ExpiryTime: TDateTime read Get_ExpiryTime write Set_ExpiryTime;
    property HTMLBody: WideString read Get_HTMLBody write Set_HTMLBody;
    property ReceivedTime: TDateTime read Get_ReceivedTime;
    property SenderName: WideString read Get_SenderName;
    property SentOn: TDateTime read Get_SentOn;
  end;

// *********************************************************************//
// DispIntf:  _PostItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063024-0000-0000-C000-000000000046}
// *********************************************************************//
  _PostItemDisp = dispinterface
    ['{00063024-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property ExpiryTime: TDateTime dispid 21;
    property HTMLBody: WideString dispid 62468;
    property ReceivedTime: TDateTime readonly dispid 3590;
    property SenderName: WideString readonly dispid 3098;
    property SentOn: TDateTime readonly dispid 57;
    procedure ClearConversationIndex; dispid 63522;
    function Forward: MailItem; dispid 63507;
    procedure Post; dispid 61557;
    function Reply: MailItem; dispid 63504;
  end;

// *********************************************************************//
// Interface: _RemoteItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063023-0000-0000-C000-000000000046}
// *********************************************************************//
  _RemoteItem = interface(IDispatch)
    ['{00063023-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_HasAttachment: WordBool; safecall;
    function Get_RemoteMessageClass: WideString; safecall;
    function Get_TransferSize: Integer; safecall;
    function Get_TransferTime: Integer; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property HasAttachment: WordBool read Get_HasAttachment;
    property RemoteMessageClass: WideString read Get_RemoteMessageClass;
    property TransferSize: Integer read Get_TransferSize;
    property TransferTime: Integer read Get_TransferTime;
  end;

// *********************************************************************//
// DispIntf:  _RemoteItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063023-0000-0000-C000-000000000046}
// *********************************************************************//
  _RemoteItemDisp = dispinterface
    ['{00063023-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property HasAttachment: WordBool readonly dispid 36615;
    property RemoteMessageClass: WideString readonly dispid 36610;
    property TransferSize: Integer readonly dispid 36613;
    property TransferTime: Integer readonly dispid 36612;
  end;

// *********************************************************************//
// Interface: _ReportItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063026-0000-0000-C000-000000000046}
// *********************************************************************//
  _ReportItem = interface(IDispatch)
    ['{00063026-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
  end;

// *********************************************************************//
// DispIntf:  _ReportItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063026-0000-0000-C000-000000000046}
// *********************************************************************//
  _ReportItemDisp = dispinterface
    ['{00063026-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
  end;

// *********************************************************************//
// Interface: _TaskItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063035-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskItem = interface(IDispatch)
    ['{00063035-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function Get_ActualWork: Integer; safecall;
    procedure Set_ActualWork(ActualWork: Integer); safecall;
    function Get_CardData: WideString; safecall;
    procedure Set_CardData(const CardData: WideString); safecall;
    function Get_Complete: WordBool; safecall;
    procedure Set_Complete(Complete: WordBool); safecall;
    function Get_Contacts: WideString; safecall;
    procedure Set_Contacts(const Contacts: WideString); safecall;
    function Get_ContactNames: WideString; safecall;
    procedure Set_ContactNames(const ContactNames: WideString); safecall;
    function Get_DateCompleted: TDateTime; safecall;
    procedure Set_DateCompleted(DateCompleted: TDateTime); safecall;
    function Get_DelegationState: OlTaskDelegationState; safecall;
    function Get_Delegator: WideString; safecall;
    function Get_DueDate: TDateTime; safecall;
    procedure Set_DueDate(DueDate: TDateTime); safecall;
    function Get_IsRecurring: WordBool; safecall;
    function Get_Ordinal: Integer; safecall;
    procedure Set_Ordinal(Ordinal: Integer); safecall;
    function Get_Owner: WideString; safecall;
    procedure Set_Owner(const Owner: WideString); safecall;
    function Get_Ownership: OlTaskOwnership; safecall;
    function Get_PercentComplete: Integer; safecall;
    procedure Set_PercentComplete(PercentComplete: Integer); safecall;
    function Get_Recipients: Recipients; safecall;
    function Get_ReminderTime: TDateTime; safecall;
    procedure Set_ReminderTime(ReminderTime: TDateTime); safecall;
    function Get_ReminderOverrideDefault: WordBool; safecall;
    procedure Set_ReminderOverrideDefault(ReminderOverrideDefault: WordBool); safecall;
    function Get_ReminderPlaySound: WordBool; safecall;
    procedure Set_ReminderPlaySound(ReminderPlaySound: WordBool); safecall;
    function Get_ReminderSet: WordBool; safecall;
    procedure Set_ReminderSet(ReminderSet: WordBool); safecall;
    function Get_ReminderSoundFile: WideString; safecall;
    procedure Set_ReminderSoundFile(const ReminderSoundFile: WideString); safecall;
    function Get_ResponseState: OlTaskResponse; safecall;
    function Get_Role: WideString; safecall;
    procedure Set_Role(const Role: WideString); safecall;
    function Get_SchedulePlusPriority: WideString; safecall;
    procedure Set_SchedulePlusPriority(const SchedulePlusPriority: WideString); safecall;
    function Get_StartDate: TDateTime; safecall;
    procedure Set_StartDate(StartDate: TDateTime); safecall;
    function Get_Status: OlTaskStatus; safecall;
    procedure Set_Status(Status: OlTaskStatus); safecall;
    function Get_StatusOnCompletionRecipients: WideString; safecall;
    procedure Set_StatusOnCompletionRecipients(const StatusOnCompletionRecipients: WideString); safecall;
    function Get_StatusUpdateRecipients: WideString; safecall;
    procedure Set_StatusUpdateRecipients(const StatusUpdateRecipients: WideString); safecall;
    function Get_TeamTask: WordBool; safecall;
    procedure Set_TeamTask(TeamTask: WordBool); safecall;
    function Get_TotalWork: Integer; safecall;
    procedure Set_TotalWork(TotalWork: Integer); safecall;
    function Assign: TaskItem; safecall;
    procedure CancelResponseState; safecall;
    procedure ClearRecurrencePattern; safecall;
    function GetRecurrencePattern: RecurrencePattern; safecall;
    procedure MarkComplete; safecall;
    function Respond(Response: OlTaskResponse; fNoUI: OleVariant; fAdditionalTextDialog: OleVariant): TaskItem; safecall;
    procedure Send; safecall;
    function SkipRecurrence: WordBool; safecall;
    function StatusReport: IDispatch; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
    property ActualWork: Integer read Get_ActualWork write Set_ActualWork;
    property CardData: WideString read Get_CardData write Set_CardData;
    property Complete: WordBool read Get_Complete write Set_Complete;
    property Contacts: WideString read Get_Contacts write Set_Contacts;
    property ContactNames: WideString read Get_ContactNames write Set_ContactNames;
    property DateCompleted: TDateTime read Get_DateCompleted write Set_DateCompleted;
    property DelegationState: OlTaskDelegationState read Get_DelegationState;
    property Delegator: WideString read Get_Delegator;
    property DueDate: TDateTime read Get_DueDate write Set_DueDate;
    property IsRecurring: WordBool read Get_IsRecurring;
    property Ordinal: Integer read Get_Ordinal write Set_Ordinal;
    property Owner: WideString read Get_Owner write Set_Owner;
    property Ownership: OlTaskOwnership read Get_Ownership;
    property PercentComplete: Integer read Get_PercentComplete write Set_PercentComplete;
    property Recipients: Recipients read Get_Recipients;
    property ReminderTime: TDateTime read Get_ReminderTime write Set_ReminderTime;
    property ReminderOverrideDefault: WordBool read Get_ReminderOverrideDefault write Set_ReminderOverrideDefault;
    property ReminderPlaySound: WordBool read Get_ReminderPlaySound write Set_ReminderPlaySound;
    property ReminderSet: WordBool read Get_ReminderSet write Set_ReminderSet;
    property ReminderSoundFile: WideString read Get_ReminderSoundFile write Set_ReminderSoundFile;
    property ResponseState: OlTaskResponse read Get_ResponseState;
    property Role: WideString read Get_Role write Set_Role;
    property SchedulePlusPriority: WideString read Get_SchedulePlusPriority write Set_SchedulePlusPriority;
    property StartDate: TDateTime read Get_StartDate write Set_StartDate;
    property Status: OlTaskStatus read Get_Status write Set_Status;
    property StatusOnCompletionRecipients: WideString read Get_StatusOnCompletionRecipients write Set_StatusOnCompletionRecipients;
    property StatusUpdateRecipients: WideString read Get_StatusUpdateRecipients write Set_StatusUpdateRecipients;
    property TeamTask: WordBool read Get_TeamTask write Set_TeamTask;
    property TotalWork: Integer read Get_TotalWork write Set_TotalWork;
  end;

// *********************************************************************//
// DispIntf:  _TaskItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063035-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskItemDisp = dispinterface
    ['{00063035-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    property ActualWork: Integer dispid 33040;
    property CardData: WideString dispid 33067;
    property Complete: WordBool dispid 33052;
    property Contacts: WideString dispid 34106;
    property ContactNames: WideString dispid 34108;
    property DateCompleted: TDateTime dispid 33039;
    property DelegationState: OlTaskDelegationState readonly dispid 33066;
    property Delegator: WideString readonly dispid 33057;
    property DueDate: TDateTime dispid 33029;
    property IsRecurring: WordBool readonly dispid 62999;
    property Ordinal: Integer dispid 33059;
    property Owner: WideString dispid 33055;
    property Ownership: OlTaskOwnership readonly dispid 33065;
    property PercentComplete: Integer dispid 63007;
    property Recipients: Recipients readonly dispid 63508;
    property ReminderTime: TDateTime dispid 34050;
    property ReminderOverrideDefault: WordBool dispid 34076;
    property ReminderPlaySound: WordBool dispid 34078;
    property ReminderSet: WordBool dispid 34051;
    property ReminderSoundFile: WideString dispid 34079;
    property ResponseState: OlTaskResponse readonly dispid 63011;
    property Role: WideString dispid 33063;
    property SchedulePlusPriority: WideString dispid 33071;
    property StartDate: TDateTime dispid 33028;
    property Status: OlTaskStatus dispid 33025;
    property StatusOnCompletionRecipients: WideString dispid 3586;
    property StatusUpdateRecipients: WideString dispid 3587;
    property TeamTask: WordBool dispid 33027;
    property TotalWork: Integer dispid 33041;
    function Assign: TaskItem; dispid 63008;
    procedure CancelResponseState; dispid 63010;
    procedure ClearRecurrencePattern; dispid 61605;
    function GetRecurrencePattern: RecurrencePattern; dispid 61604;
    procedure MarkComplete; dispid 62989;
    function Respond(Response: OlTaskResponse; fNoUI: OleVariant; fAdditionalTextDialog: OleVariant): TaskItem; dispid 63009;
    procedure Send; dispid 61557;
    function SkipRecurrence: WordBool; dispid 63012;
    function StatusReport: IDispatch; dispid 62994;
  end;

// *********************************************************************//
// Interface: _TaskRequestAcceptItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063038-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestAcceptItem = interface(IDispatch)
    ['{00063038-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
  end;

// *********************************************************************//
// DispIntf:  _TaskRequestAcceptItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063038-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestAcceptItemDisp = dispinterface
    ['{00063038-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; dispid 61460;
  end;

// *********************************************************************//
// Interface: _TaskRequestDeclineItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063039-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestDeclineItem = interface(IDispatch)
    ['{00063039-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
  end;

// *********************************************************************//
// DispIntf:  _TaskRequestDeclineItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063039-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestDeclineItemDisp = dispinterface
    ['{00063039-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; dispid 61460;
  end;

// *********************************************************************//
// Interface: _TaskRequestItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063036-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestItem = interface(IDispatch)
    ['{00063036-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
  end;

// *********************************************************************//
// DispIntf:  _TaskRequestItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063036-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestItemDisp = dispinterface
    ['{00063036-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; dispid 61460;
  end;

// *********************************************************************//
// Interface: _TaskRequestUpdateItem
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063037-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestUpdateItem = interface(IDispatch)
    ['{00063037-0000-0000-C000-000000000046}']
    function Get_Application_: Application_; safecall;
    function Get_Class_: OlObjectClass; safecall;
    function Get_Session: NameSpace; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Actions: Actions; safecall;
    function Get_Attachments: Attachments; safecall;
    function Get_BillingInformation: WideString; safecall;
    procedure Set_BillingInformation(const BillingInformation: WideString); safecall;
    function Get_Body: WideString; safecall;
    procedure Set_Body(const Body: WideString); safecall;
    function Get_Categories: WideString; safecall;
    procedure Set_Categories(const Categories: WideString); safecall;
    function Get_Companies: WideString; safecall;
    procedure Set_Companies(const Companies: WideString); safecall;
    function Get_ConversationIndex: WideString; safecall;
    function Get_ConversationTopic: WideString; safecall;
    function Get_CreationTime: TDateTime; safecall;
    function Get_EntryID: WideString; safecall;
    function Get_FormDescription: FormDescription; safecall;
    function Get_GetInspector: Inspector; safecall;
    function Get_Importance: OlImportance; safecall;
    procedure Set_Importance(Importance: OlImportance); safecall;
    function Get_LastModificationTime: TDateTime; safecall;
    function Get_MAPIOBJECT: IUnknown; safecall;
    function Get_MessageClass: WideString; safecall;
    procedure Set_MessageClass(const MessageClass: WideString); safecall;
    function Get_Mileage: WideString; safecall;
    procedure Set_Mileage(const Mileage: WideString); safecall;
    function Get_NoAging: WordBool; safecall;
    procedure Set_NoAging(NoAging: WordBool); safecall;
    function Get_OutlookInternalVersion: Integer; safecall;
    function Get_OutlookVersion: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    function Get_Sensitivity: OlSensitivity; safecall;
    procedure Set_Sensitivity(Sensitivity: OlSensitivity); safecall;
    function Get_Size: Integer; safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const Subject: WideString); safecall;
    function Get_UnRead: WordBool; safecall;
    procedure Set_UnRead(UnRead: WordBool); safecall;
    function Get_UserProperties: UserProperties; safecall;
    procedure Close(SaveMode: OlInspectorClose); safecall;
    function Copy: IDispatch; safecall;
    procedure Delete; safecall;
    procedure Display(Modal: OleVariant); safecall;
    function Move(const DestFldr: MAPIFolder): IDispatch; safecall;
    procedure PrintOut; safecall;
    procedure Save; safecall;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); safecall;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; safecall;
    property Application_: Application_ read Get_Application_;
    property Class_: OlObjectClass read Get_Class_;
    property Session: NameSpace read Get_Session;
    property Parent: IDispatch read Get_Parent;
    property Actions: Actions read Get_Actions;
    property Attachments: Attachments read Get_Attachments;
    property BillingInformation: WideString read Get_BillingInformation write Set_BillingInformation;
    property Body: WideString read Get_Body write Set_Body;
    property Categories: WideString read Get_Categories write Set_Categories;
    property Companies: WideString read Get_Companies write Set_Companies;
    property ConversationIndex: WideString read Get_ConversationIndex;
    property ConversationTopic: WideString read Get_ConversationTopic;
    property CreationTime: TDateTime read Get_CreationTime;
    property EntryID: WideString read Get_EntryID;
    property FormDescription: FormDescription read Get_FormDescription;
    property GetInspector: Inspector read Get_GetInspector;
    property Importance: OlImportance read Get_Importance write Set_Importance;
    property LastModificationTime: TDateTime read Get_LastModificationTime;
    property MAPIOBJECT: IUnknown read Get_MAPIOBJECT;
    property MessageClass: WideString read Get_MessageClass write Set_MessageClass;
    property Mileage: WideString read Get_Mileage write Set_Mileage;
    property NoAging: WordBool read Get_NoAging write Set_NoAging;
    property OutlookInternalVersion: Integer read Get_OutlookInternalVersion;
    property OutlookVersion: WideString read Get_OutlookVersion;
    property Saved: WordBool read Get_Saved;
    property Sensitivity: OlSensitivity read Get_Sensitivity write Set_Sensitivity;
    property Size: Integer read Get_Size;
    property Subject: WideString read Get_Subject write Set_Subject;
    property UnRead: WordBool read Get_UnRead write Set_UnRead;
    property UserProperties: UserProperties read Get_UserProperties;
  end;

// *********************************************************************//
// DispIntf:  _TaskRequestUpdateItemDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00063037-0000-0000-C000-000000000046}
// *********************************************************************//
  _TaskRequestUpdateItemDisp = dispinterface
    ['{00063037-0000-0000-C000-000000000046}']
    property Application_: Application_ readonly dispid 61440;
    property Class_: OlObjectClass readonly dispid 61450;
    property Session: NameSpace readonly dispid 61451;
    property Parent: IDispatch readonly dispid 61441;
    property Actions: Actions readonly dispid 63511;
    property Attachments: Attachments readonly dispid 63509;
    property BillingInformation: WideString dispid 34101;
    property Body: WideString dispid 37120;
    property Categories: WideString dispid 36865;
    property Companies: WideString dispid 34107;
    property ConversationIndex: WideString readonly dispid 113;
    property ConversationTopic: WideString readonly dispid 112;
    property CreationTime: TDateTime readonly dispid 12295;
    property EntryID: WideString readonly dispid 61470;
    property FormDescription: FormDescription readonly dispid 61589;
    property GetInspector: Inspector readonly dispid 61502;
    property Importance: OlImportance dispid 23;
    property LastModificationTime: TDateTime readonly dispid 12296;
    property MAPIOBJECT: IUnknown readonly dispid 61696;
    property MessageClass: WideString dispid 26;
    property Mileage: WideString dispid 34100;
    property NoAging: WordBool dispid 34062;
    property OutlookInternalVersion: Integer readonly dispid 34130;
    property OutlookVersion: WideString readonly dispid 34132;
    property Saved: WordBool readonly dispid 61603;
    property Sensitivity: OlSensitivity dispid 54;
    property Size: Integer readonly dispid 3592;
    property Subject: WideString dispid 55;
    property UnRead: WordBool dispid 61468;
    property UserProperties: UserProperties readonly dispid 63510;
    procedure Close(SaveMode: OlInspectorClose); dispid 61475;
    function Copy: IDispatch; dispid 61490;
    procedure Delete; dispid 61514;
    procedure Display(Modal: OleVariant); dispid 61606;
    function Move(const DestFldr: MAPIFolder): IDispatch; dispid 61492;
    procedure PrintOut; dispid 61491;
    procedure Save; dispid 61512;
    procedure SaveAs(const Path: WideString; Type_: OleVariant); dispid 61521;
    function GetAssociatedTask(AddToTaskList: WordBool): TaskItem; dispid 61460;
  end;

// *********************************************************************//
// Interface: _IRecipientControl
// Flags:     (4112) Hidden Dispatchable
// GUID:      {D87E7E16-6897-11CE-A6C0-00AA00608FAA}
// *********************************************************************//
  _IRecipientControl = interface(IDispatch)
    ['{D87E7E16-6897-11CE-A6C0-00AA00608FAA}']
    function Get_Enabled(out Enabled: Shortint): HResult; stdcall;
    function Set_Enabled(Enabled: Shortint): HResult; stdcall;
    function Get_BackColor(out BackColor: Integer): HResult; stdcall;
    function Set_BackColor(BackColor: Integer): HResult; stdcall;
    function Get_ForeColor(out ForeColor: Integer): HResult; stdcall;
    function Set_ForeColor(ForeColor: Integer): HResult; stdcall;
    function Get_ReadOnly(out ReadOnly: Shortint): HResult; stdcall;
    function Set_ReadOnly(ReadOnly: Shortint): HResult; stdcall;
    function Get_Font(out Font: IFontDisp): HResult; stdcall;
    function Set_Font(const Font: IFontDisp): HResult; stdcall;
    function Get_SpecialEffect(out Effect: Integer): HResult; stdcall;
    function Set_SpecialEffect(Effect: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// DispIntf:  _DRecipientControl
// Flags:     (4112) Hidden Dispatchable
// GUID:      {0006F025-0000-0000-C000-000000000046}
// *********************************************************************//
  _DRecipientControl = dispinterface
    ['{0006F025-0000-0000-C000-000000000046}']
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
    property Enabled: {??Shortint} OleVariant dispid -514;
    property BackColor: Integer dispid -501;
    property ForeColor: Integer dispid -513;
    property ReadOnly: {??Shortint} OleVariant dispid -2147356664;
    property Font: IFontDisp dispid -512;
    property SpecialEffect: Integer dispid 12;
  end;

// *********************************************************************//
// DispIntf:  _DRecipientControlEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {D87E7E17-6897-11CE-A6C0-00AA00608FAA}
// *********************************************************************//
  _DRecipientControlEvents = dispinterface
    ['{D87E7E17-6897-11CE-A6C0-00AA00608FAA}']
  end;

// *********************************************************************//
// Interface: _IDocSiteControl
// Flags:     (4112) Hidden Dispatchable
// GUID:      {43507DD0-811D-11CE-B565-00AA00608FAA}
// *********************************************************************//
  _IDocSiteControl = interface(IDispatch)
    ['{43507DD0-811D-11CE-B565-00AA00608FAA}']
    function Get_ReadOnly(out ReadOnly: Shortint): HResult; stdcall;
    function Set_ReadOnly(ReadOnly: Shortint): HResult; stdcall;
  end;

// *********************************************************************//
// DispIntf:  _DDocSiteControl
// Flags:     (4112) Hidden Dispatchable
// GUID:      {0006F026-0000-0000-C000-000000000046}
// *********************************************************************//
  _DDocSiteControl = dispinterface
    ['{0006F026-0000-0000-C000-000000000046}']
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
    property ReadOnly: {??Shortint} OleVariant dispid -2147356664;
  end;

// *********************************************************************//
// DispIntf:  _DDocSiteControlEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {50BB9B50-811D-11CE-B565-00AA00608FAA}
// *********************************************************************//
  _DDocSiteControlEvents = dispinterface
    ['{50BB9B50-811D-11CE-B565-00AA00608FAA}']
  end;

  CoApplication_ = class
    class function Create: _Application;
    class function CreateRemote(const MachineName: string): _Application;
  end;

  CoMailItem = class
    class function Create: _MailItem;
    class function CreateRemote(const MachineName: string): _MailItem;
  end;

  CoAppointmentItem = class
    class function Create: _AppointmentItem;
    class function CreateRemote(const MachineName: string): _AppointmentItem;
  end;

  CoMeetingItem = class
    class function Create: _MeetingItem;
    class function CreateRemote(const MachineName: string): _MeetingItem;
  end;

  CoTaskItem = class
    class function Create: _TaskItem;
    class function CreateRemote(const MachineName: string): _TaskItem;
  end;

  CoContactItem = class
    class function Create: _ContactItem;
    class function CreateRemote(const MachineName: string): _ContactItem;
  end;

  CoDocumentItem = class
    class function Create: _DocumentItem;
    class function CreateRemote(const MachineName: string): _DocumentItem;
  end;

  CoJournalItem = class
    class function Create: _JournalItem;
    class function CreateRemote(const MachineName: string): _JournalItem;
  end;

  CoNoteItem = class
    class function Create: _NoteItem;
    class function CreateRemote(const MachineName: string): _NoteItem;
  end;

  CoPostItem = class
    class function Create: _PostItem;
    class function CreateRemote(const MachineName: string): _PostItem;
  end;

  CoRemoteItem = class
    class function Create: _RemoteItem;
    class function CreateRemote(const MachineName: string): _RemoteItem;
  end;

  CoReportItem = class
    class function Create: _ReportItem;
    class function CreateRemote(const MachineName: string): _ReportItem;
  end;

  CoTaskRequestAcceptItem = class
    class function Create: _TaskRequestAcceptItem;
    class function CreateRemote(const MachineName: string): _TaskRequestAcceptItem;
  end;

  CoTaskRequestDeclineItem = class
    class function Create: _TaskRequestDeclineItem;
    class function CreateRemote(const MachineName: string): _TaskRequestDeclineItem;
  end;

  CoTaskRequestItem = class
    class function Create: _TaskRequestItem;
    class function CreateRemote(const MachineName: string): _TaskRequestItem;
  end;

  CoTaskRequestUpdateItem = class
    class function Create: _TaskRequestUpdateItem;
    class function CreateRemote(const MachineName: string): _TaskRequestUpdateItem;
  end;

implementation

uses ComObj;

class function CoApplication_.Create: _Application;
begin
  Result := CreateComObject(CLASS_Application_) as _Application;
end;

class function CoApplication_.CreateRemote(const MachineName: string): _Application;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Application_) as _Application;
end;

class function CoMailItem.Create: _MailItem;
begin
  Result := CreateComObject(CLASS_MailItem) as _MailItem;
end;

class function CoMailItem.CreateRemote(const MachineName: string): _MailItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_MailItem) as _MailItem;
end;

class function CoAppointmentItem.Create: _AppointmentItem;
begin
  Result := CreateComObject(CLASS_AppointmentItem) as _AppointmentItem;
end;

class function CoAppointmentItem.CreateRemote(const MachineName: string): _AppointmentItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_AppointmentItem) as _AppointmentItem;
end;

class function CoMeetingItem.Create: _MeetingItem;
begin
  Result := CreateComObject(CLASS_MeetingItem) as _MeetingItem;
end;

class function CoMeetingItem.CreateRemote(const MachineName: string): _MeetingItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_MeetingItem) as _MeetingItem;
end;

class function CoTaskItem.Create: _TaskItem;
begin
  Result := CreateComObject(CLASS_TaskItem) as _TaskItem;
end;

class function CoTaskItem.CreateRemote(const MachineName: string): _TaskItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskItem) as _TaskItem;
end;

class function CoContactItem.Create: _ContactItem;
begin
  Result := CreateComObject(CLASS_ContactItem) as _ContactItem;
end;

class function CoContactItem.CreateRemote(const MachineName: string): _ContactItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ContactItem) as _ContactItem;
end;

class function CoDocumentItem.Create: _DocumentItem;
begin
  Result := CreateComObject(CLASS_DocumentItem) as _DocumentItem;
end;

class function CoDocumentItem.CreateRemote(const MachineName: string): _DocumentItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DocumentItem) as _DocumentItem;
end;

class function CoJournalItem.Create: _JournalItem;
begin
  Result := CreateComObject(CLASS_JournalItem) as _JournalItem;
end;

class function CoJournalItem.CreateRemote(const MachineName: string): _JournalItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_JournalItem) as _JournalItem;
end;

class function CoNoteItem.Create: _NoteItem;
begin
  Result := CreateComObject(CLASS_NoteItem) as _NoteItem;
end;

class function CoNoteItem.CreateRemote(const MachineName: string): _NoteItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_NoteItem) as _NoteItem;
end;

class function CoPostItem.Create: _PostItem;
begin
  Result := CreateComObject(CLASS_PostItem) as _PostItem;
end;

class function CoPostItem.CreateRemote(const MachineName: string): _PostItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PostItem) as _PostItem;
end;

class function CoRemoteItem.Create: _RemoteItem;
begin
  Result := CreateComObject(CLASS_RemoteItem) as _RemoteItem;
end;

class function CoRemoteItem.CreateRemote(const MachineName: string): _RemoteItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_RemoteItem) as _RemoteItem;
end;

class function CoReportItem.Create: _ReportItem;
begin
  Result := CreateComObject(CLASS_ReportItem) as _ReportItem;
end;

class function CoReportItem.CreateRemote(const MachineName: string): _ReportItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ReportItem) as _ReportItem;
end;

class function CoTaskRequestAcceptItem.Create: _TaskRequestAcceptItem;
begin
  Result := CreateComObject(CLASS_TaskRequestAcceptItem) as _TaskRequestAcceptItem;
end;

class function CoTaskRequestAcceptItem.CreateRemote(const MachineName: string): _TaskRequestAcceptItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskRequestAcceptItem) as _TaskRequestAcceptItem;
end;

class function CoTaskRequestDeclineItem.Create: _TaskRequestDeclineItem;
begin
  Result := CreateComObject(CLASS_TaskRequestDeclineItem) as _TaskRequestDeclineItem;
end;

class function CoTaskRequestDeclineItem.CreateRemote(const MachineName: string): _TaskRequestDeclineItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskRequestDeclineItem) as _TaskRequestDeclineItem;
end;

class function CoTaskRequestItem.Create: _TaskRequestItem;
begin
  Result := CreateComObject(CLASS_TaskRequestItem) as _TaskRequestItem;
end;

class function CoTaskRequestItem.CreateRemote(const MachineName: string): _TaskRequestItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskRequestItem) as _TaskRequestItem;
end;

class function CoTaskRequestUpdateItem.Create: _TaskRequestUpdateItem;
begin
  Result := CreateComObject(CLASS_TaskRequestUpdateItem) as _TaskRequestUpdateItem;
end;

class function CoTaskRequestUpdateItem.CreateRemote(const MachineName: string): _TaskRequestUpdateItem;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskRequestUpdateItem) as _TaskRequestUpdateItem;
end;

end.
