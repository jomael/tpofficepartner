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

(* NOTES:
  This console application is used to correct the BCB6 headers that are 
  generated when the package is compiled. "CheckSafecallResult" must
  be replaced with "OleCheck". It is used by passing the path of the 
  headers as a commandline variable. IF no Variable is passed in, it
  will attempt look for the files in the current directory.*)

program OPHeaderFix;
{$APPTYPE CONSOLE}
uses
  SysUtils,
  Registry,
  Windows,
  Classes;

const
  cOldPattern = 'System::CheckSafeCallResult(hr)';
  cNewPattern = 'OleCheck(hr)';
  cComObjHeader = '#include <ComObj.hpp> //added by OPHeaderFix program on: %s';

var
  SearchDir: string;

function FindComObjHeader(FileLines : TStringList; var EndLocation : Integer): boolean;
(*returns line number of the last entry in the include list, which is where
  we will insert the #include <ComObj.hpp> UNLESS it is already there, then
  this function will return -1, also indicated if the comobj.hpp was already
  found in the list *)
var
  I, IncludesFound, CharIndex : Integer;
  EndOfIncludeList : boolean;
begin
  I := 0;
  IncludesFound := 0;
  EndOfIncludeList := False;
  result := false;
  While ((I < FileLines.Count -1) and not (EndOfIncludeList)) do
  begin
    //loop through the lines of the file until we find a #include
    CharIndex :=   pos('#include',FileLines.Strings[I]);
    if CharIndex <> 0 then //meaning we found the word '#include' in it
    begin
      IncludesFound := IncludesFound + 1;
      if pos('<ComObj.hpp>', FileLines.Strings[I]) <> 0 then
      begin
        Result := true;
      end;
    end
    else if ((IncludesFound > 0) and (CharIndex = 0)) then
    begin
      //end of the include list was found
      EndOfIncludeList := true;
      EndLocation := I;
    end;
    I := I + 1;
  end;
end;

procedure FixPath(var Path : string);
begin
  if Path[Length(path)] <> '\' then
    Path := Path + '\';
end;

function GetOfficePartnerLoc : string;
(*
var
  OPReg : TRegIniFile;
*)
begin
  //look in the registry for the location of the OfficePartner files
  if ParamCount <> 0 then
    result := ParamStr(1) //used passed in var
  else
  begin
     result :=  ExtractFilePath(ParamStr(0));
    (*
    OPReg := TRegIniFile.Create;
    try
      OpReg.RootKey := HKEY_LOCAL_MACHINE;
      //look in registry, if cannont find, use app directory
      result := OpReg.ReadString('SOFTWARE\TurboPower\OfficePartner','AppPath',ExtractFilePath(ParamStr(0)));
    finally
      OpReg.Free;
    end;
    *)
  end;
end;


procedure GetHeaderFiles(var HeaderFiles : TStringList);
var
  DirFiles : TSearchRec;
begin
  //get a list of the header files and place them into the var HeaderFiles
  //Call GetOfficePartnerLoc for the location of the files
  SearchDir := GetOfficePartnerLoc;
  FixPath(SearchDir);
  if FindFirst(SearchDir+'Op*.hpp', faAnyFile, DirFiles) = 0 then
  begin
    HeaderFiles.Add(SearchDir + DirFiles.Name);
    while FindNext(DirFiles) = 0 do
      HeaderFiles.Add(SearchDir + DirFiles.Name);
    SysUtils.FindClose(DirFiles);
  end; //if
end;

var
  HPPFiles , HeaderFile: TStringList;
  I, J  : Integer;
  TempStr : string;
  Modified : Boolean;
  FixedFileCount : Integer;
  FoundComObjHeader : boolean;
  IncludeListEndLoc : Integer;
begin
  FixedFileCount := 0;
  try
    try
      HPPFiles := TStringList.Create;
      GetHeaderFiles(HPPFiles);
      //iterate through all of the header files
      if HPPFiles.Count = 0 then
      begin
        Writeln('No office partner header files found in directory: ' + SearchDir);
      end
      else
      begin
        for I := 0 to HPPFiles.count - 1 do
        begin
          HeaderFile := TStringList.Create;
          try
            //Open the current file
            HeaderFile.LoadFromFile(HPPFiles.Strings[I]);
            IncludeListEndLoc := 1;
            FoundComObjHeader := FindComObjHeader(HeaderFile,IncludeListEndLoc);
            Modified := false; //for each file, default is NOT modified
            for J := IncludeListEndLoc to HeaderFile.Count - 1 do //read each line in it
            begin
              (*on each line do a stringreplace for System::CheckAutoResult(hr),
                to OleCheck(hr); *)
              TempStr := StringReplace(HeaderFile.strings[J],
                cOldPattern,cNewPattern, [rfReplaceAll, rfIgnoreCase]);
              if CompareStr(TempStr, HeaderFile.strings[J]) <> 0 then
              begin
                Modified := true;
                HeaderFile.strings[J] := TempStr;
              end;
            end; //for J := 0 to HeaderFile.Count - 1 do
            if Modified then
            begin
              (*if file was modified, save it and give
                user feedback*)
              if not FoundComObjHeader then
              begin
                HeaderFile.Insert(IncludeListEndLoc, Format(cComObjHeader,
                  [DateTimeToStr(now)]));
              end;
              HeaderFile.SaveToFile(HPPFiles.Strings[I]); //save the file
              Writeln('File Modified: '+HPPFiles.Strings[I]);
              FixedFileCount := FixedFileCount + 1;
            end;
          finally
            HeaderFile.Free;
          end;
        end; //for I to HPPFiles.Count
      end; //else begin
    finally
      HPPFiles.Free;
    end;
    Writeln(inttostr(FixedFileCount) + ' file(s) modified')
  except on E : Exception do
    Writeln(E.Message);
  end;
end.