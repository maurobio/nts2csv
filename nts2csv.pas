{=======================================================================}
{        NTS2CSV - Converts NTSYSpc data files to CSV format            }
{       (c) 2003-2008 Mauro J. Cavalcanti. All rights reserved.         }
{                                                                       }
{  This program has been developed using Borland Delphi 4.0             }
{  Professional Edition for Windows 95, 98, Me, NT, 2000, XP            }
{  and FreePascal 2.0 for Windows XP & GNU/Linux.                       }
{                                                                       }
{  This code may be freely copied  and modified, as long as the above   }
{  copyright notice is not removed and modifications are documented.    }
{  No price or fee must be charged for it.                              }
{                                                                       }
{  If anyone makes any significant alterations or has any bright ideas  }
{  to enhance the program then please forward them to me so I can keep  }
{  one up to date copy. Comments may be sent by e-mail to the address:  }
{  maurobio@gmail.com                                                   }
{                                                                       }
{  Disclaimer: this program is provided as is, without warranty of any  }
{  kind. Use it at your own risk.                                       }
{                                                                       }
{  REVISION HISTORY:                                                    }
{    Version 1.00, 27th Sep 03 - Initial release                        }
{    Version 1.01, 27th Jul 08 - Minor modifications                    }
{=======================================================================}

program NTS2CSV;

{$APPTYPE CONSOLE}

uses Classes, SysUtils;

var
  Outfile: TextFile;
  inf, outf: string;
  ObjList, VarList, CommList: TStringList;
  DatMat: array of array of Double;
  I, J, N, M: Word;
  Kind, Missing: Byte;
  NC: Integer;

{ strip all occurrences of a character from a string }

function StripChar (Str: string; Ch: Char): string;
var
  i: Integer;
begin
  i := 1;
  repeat
    if (Str[i] = Ch) and (Length (Str) > 0) then
      Delete (Str, i, 1)
    else
      i := Succ (i);
  until (i > Length (Str)) or (Str = '');
  StripChar := Str;
end;

function GetSubStr (aString, SepChar: string; TokenNum: Byte): string;
{
parameters: aString : the complete string
      SepChar : a single character used as separator
           between the substrings
      TokenNum: the number of the substring you want
result		: the substring or an empty string if the are less then
      'TokenNum' substrings
}
var
  Token: string;
  StrLen: Byte;
  TNum: Byte;
  TEnd: Byte;

begin
  StrLen := Length (aString);
  TNum := 1;
  TEnd := StrLen;
  while ((TNum <= TokenNum) and (TEnd <> 0)) do begin
    TEnd := Pos (SepChar, aString);
    if TEnd <> 0 then begin
      Token := Copy (aString, 1, TEnd - 1);
      Delete (aString, 1, TEnd);
      Inc (TNum);
    end
    else begin
      Token := aString;
    end;
  end;
  if TNum >= TokenNum then begin
    GetSubStr := Token;
  end
  else begin
    GetSubStr := '';
  end;
end;

function IIF (BoolVar: Boolean; IfTrue, IfFalse: string): string;
begin
  if BoolVar then
    IIF := IfTrue
  else
    IIF := IfFalse;
end;

{  Open I/O files  }

procedure OpenFiles;
var
  reply: Char;
  ok: Boolean;
begin
  repeat
    WriteLn;
    Write ('Name of NTSYSpc file: ');
    ReadLn (inf);
    if inf = '' then Halt;
    if Pos ('.', inf) = 0 then inf := inf + '.nts';
    ok := FileExists (inf);
    if not ok then begin
      Beep;
      Write ('*** ERROR: File "', inf, '" not found!');
    end;
  until ok;

  repeat
    WriteLn;
    Write ('Name for CSV file: ');
    ReadLn (outf);
    if outf = '' then outf := Copy (inf, 1, Pos ('.', inf) - 1);
    if Pos ('.', outf) = 0 then
      outf := outf + '.csv';
    ok := FileExists (outf);
    if ok then begin
      Beep;
      Write ('*** WARNING: File "', outf, '" already exists! Overwrite (Y/N)? ');
      ReadLn (reply);
      reply := UpCase (reply);
      ok := (reply = 'N');
    end;
  until not ok;
  AssignFile (Outfile, outf);
  Rewrite (Outfile);
end;

{ Check for I/O errors through the TURBO procedure IOResult }

procedure IOCHECK (var f: TextFile);
var
  IOErr: Boolean;
begin
  IOErr := (IOResult <> 0);
  if IOErr then begin
    WriteLn ('File error');
    CloseFile (f);
  end;
end;

{ Check for end of file while reading data }

procedure CheckForEOF (var f: TextFile);
begin
  if EoF (f) then begin
    WriteLn ('Not enough data in file');
    CloseFile (f);
    IOCHECK (f);
  end;
end;

{ Read a NTSYSpc file }

function ReadFile (const FileName: string): Boolean;
var
  Ch: Char;
  I, J: Word;
  inData: Double;
  Count: Integer;
  S, cp, inLabel: string;
  rowLabels, colLabels: Boolean;
  Infile: TextFile;

  procedure ParseLabel (var Infile: TextFile; var Lbl: string);
  { reads labels from the input file. Labels must be
    separated by spaces or end of line }
  var
    wordFound, endWord: Boolean;
    CH: Char;

  begin
    lbl := '';
    wordFound := False;
    endWord := False;
    repeat
      Read (Infile, CH);
      IOCheck (Infile);
      CheckForEOF (Infile);
      if CH > #32 then begin
        wordFound := True;
        lbl := Concat (lbl, CH);
      end;
      if ((CH = #32) or (CH = #13)) and wordFound then
        endWord := True;
    until wordFound and endWord;
  end;

begin
  ReadFile := False; { Preset function result }
  AssignFile (Infile, FileName);
  Reset (Infile);
  IOCHECK (Infile);

  { Go to matrix specification }
  Count := 0;
  CommList.Clear;
  while not EoF (Infile) do begin
    ReadLn (Infile, S);
    IOCHECK (Infile);
    S := Trim (S);
    if (S <> '') then begin
      Ch := S[1];
      { Skip comment character }
      if (Ch = Chr (34)) or (Ch = Chr (39)) then begin
        CommList.Add (Trim (StripChar (S, Ch)));
        Inc (Count);
      end else Break;
    end;
  end;

  { Parse matrix specification }
  cp := GetSubStr (S, ' ', 1);
  Kind := StrToInt (cp);
  if (cp <> '1') then begin
    WriteLn ('*** ERROR: Invalid matrix type. Must be "1"');
    CloseFile (Infile);
    IOCHECK (Infile);
    Result := False;
    Exit;
  end;

  { Rows }
  cp := GetSubStr (S, ' ', 2);
  if Pos ('L', UpperCase (cp)) = 0 then begin
    rowLabels := False;
    N := StrToInt (cp);
  end
  else begin
    rowlabels := True;
    N := StrToInt (StripChar (cp, 'L'));
  end;
  if not (N > 0) then begin
    WriteLn ('*** ERROR: Invalid row specification in file!');
    CloseFile (Infile);
    IOCHECK (Infile);
    Result := False;
    Exit;
  end;

  { Cols }
  cp := GetSubStr (S, ' ', 3);
  if Pos ('L', UpperCase (cp)) = 0 then begin
    colLabels := False;
    M := StrToInt (cp);
  end
  else begin
    colLabels := True;
    M := StrToInt (StripChar (cp, 'L'));
  end;
  if not (M > 0) then begin
    WriteLn ('*** ERROR: Invalid column specification in file!');
    CloseFile (Infile);
    IOCHECK (Infile);
    Result := False;
    Exit;
  end;

  { Missing values }
  cp := GetSubStr (S, ' ', 4);
  Missing := StrToIntDef (cp, 0);
  if (Missing = 0) then
    NC := Missing
  else begin
    cp := GetSubStr (S, ' ', 5);
    NC := StrToInt (cp);
  end;

  { Read row labels }
  ObjList.Clear;
  if rowLabels then begin
    for I := 0 to N - 1 do begin
      CheckForEOF (Infile);
      ParseLabel (Infile, inLabel);
      IOCHECK (Infile);
      ObjList.Add (inLabel);
    end;
  end
  else
    for I := 0 to N - 1 do
      ObjList.Add (Concat ('ROW', IntToStr (I + 1)));

  { Read column labels }
  VarList.Clear;
  if colLabels then begin
    for J := 0 to M - 1 do begin
      CheckForEOF (Infile);
      ParseLabel (Infile, inLabel);
      IOCHECK (Infile);
      VarList.Add (inLabel);
    end;
  end else
    for J := 0 to M - 1 do
      VarList.Add (Concat ('COL', IntToStr (J + 1)));

  { Declare global arrays }
  SetLength (DatMat, N, M);

  { Read data }
  for I := 0 to N - 1 do begin
    for J := 0 to M - 1 do begin
      CheckForEOF (Infile);
      Read (Infile, inData);
      IOCHECK (Infile);
      DatMat[I, J] := inData;
    end;
  end;

  CloseFile (Infile);
  IOCHECK (Infile);
  ReadFile := True;
end;

begin
  WriteLn ('Convert files in NTSYSpc format to comma-separated values (CSV)');
  WriteLn ('(c) 2003-2008 Mauro J. Cavalcanti');
  WriteLn ('Departamento de Zoologia, Universidade do Estado do Rio de Janeiro, Brasil');
  WriteLn ('E-mail: maurobio@gmail.com');
  OpenFiles;

  ObjList := TStringList.Create;
  VarList := TStringList.Create;
  CommList := TStringList.Create;

  if ReadFile (inf) then begin
    WriteLn;
    WriteLn ('Input file: ', ExpandFileName (inf));
    WriteLn ('ID records:');
    for I := 0 to CommList.Count - 1 do
      WriteLn ('"', CommList.Strings[I]);
    Write ('type=', Kind, ', size=', N, ' by ', M);
    WriteLn (IIf ((Missing = 0), ', nc=none', ', nc=' + IntToStr(NC)));

    { Write first row (variable names) }
    Write (Outfile, 'OBJ,');
    WriteLn (Outfile, VarList.CommaText);

    { Write data matrix }
    for I := 0 to N - 1 do begin
      Write (Outfile, ObjList.Strings[I], ',');
      for J := 0 to M - 1 do
        Write (Outfile, DatMat[I, J], IIf (J < M - 1, ',', ''));
      WriteLn (Outfile);
    end;
  end;

  ObjList.Free;
  VarList.Free;
  CommList.Free;

  Close (Outfile);
  WriteLn;
  WriteLn ('Output file - ');
  WriteLn ('   ', ExpandFileName (outf));
  WriteLn;
  WriteLn ('- done -');
end.
