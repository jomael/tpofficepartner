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
USERES("P164_R35.res");
USEPACKAGE("vclx35.bpi");
USEPACKAGE("VCL35.bpi");
USEPACKAGE("VCLDB35.bpi");
USEUNIT("..\source\OpDbOfc.pas");
USEUNIT("..\source\OpDbOlk.pas");
USEUNIT("..\source\OpXLXP.pas");
USEUNIT("..\source\OpExcel.pas");
USEUNIT("..\source\OpMSO.pas");
USEUNIT("..\source\OpOfcXP.pas");
USEUNIT("..\source\OpConst.pas");
USEUNIT("..\source\OpEvents.pas");
USEUNIT("..\source\OpModels.pas");
USEUNIT("..\source\OpShared.pas");
USEUNIT("..\source\OpOlkXP.pas");
USEUNIT("..\source\OpOlk98.pas");
USEUNIT("..\source\OpOutlk.pas");
USEUNIT("..\source\OpPptXP.pas");
USEUNIT("..\source\OpVbIdXP.pas");
USEUNIT("..\source\OpWrdXP.pas");
USEUNIT("..\source\OpWord.pas");
USEUNIT("..\source\OpPower.pas");
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------
//   Package source.
//---------------------------------------------------------------------------
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void*)
{
        return 1;
}
//---------------------------------------------------------------------------
