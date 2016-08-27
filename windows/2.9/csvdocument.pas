{
  CSV Parser and Document classes.
  Version 0.2 2010-05-31

  Copyright (C) 2010 Vladimir Zhirov <vvzh.home@gmail.com>

  This library is free software; you can redistribute it and/or modify it
  under the terms of the GNU Library General Public License as published by
  the Free Software Foundation; either version 2 of the License, or (at your
  option) any later version with the following modification:

  As a special exception, the copyright holders of this library give you
  permission to link this library with independent modules to produce an
  executable, regardless of the license terms of these independent modules,and
  to copy and distribute the resulting executable under terms of your choice,
  provided that you also meet, for each linked independent module, the terms
  and conditions of the license of that module. An independent module is a
  module which is not derived from or based on this library. If you modify
  this library, you may extend this exception to your version of the library,
  but you are not obligated to do so. If you do not wish to do so, delete this
  exception statement from your version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
  FITNESS FOR A PARTICULAR PURPOSE. See the GNU Library General Public License
  for more details.

  You should have received a copy of the GNU Library General Public License
  along with this library; if not, write to the Free Software Foundation,
  Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
}

unit CsvDocument;

{$IFDEF FPC}
  {$MODE DELPHI}
{$ENDIF}

interface

uses
  Classes, SysUtils, Contnrs;

type

  { TCSVParser }

  TCSVChar = Char;
  TCSVParser = class;
  TCSVParserEvent = procedure(Parser: TCSVParser) of object;

  TCSVParser = class(TObject)
  private
    // special chars
    FQuoteChar: TCSVChar;
    FDelimiter: TCSVChar;
    // fields
    FStream: TStream;
    FStopFlag: Boolean;
    FOnNewCell: TCSVParserEvent;
    // parser state
    EOF: Boolean;
    FCurrentChar: TCSVChar;
    FCurrentRow: Integer;
    FCurrentCol: Integer;
    FCurrentCellText: string;
    // parsing routines
    procedure ClearCellText;
    procedure IncCellText;
    function  EOL: Boolean;
    procedure SkipEOL;
    procedure SkipDelimiter;
    procedure NextChar;
    procedure ParseInput;
    procedure ParseRow;
    procedure ParseCell;
    procedure DoParseValue;
    procedure ParseQuotedValue;
    procedure ParseValue;
    // other routines
    procedure DoOnNewCell;
  public
    procedure ParseStream(AStream: TStream);
    procedure ParseString(const AString: string);
    procedure Stop;
    property OnNewCell: TCSVParserEvent read FOnNewCell write FOnNewCell;
    property CurrentRow: Integer read FCurrentRow;
    property CurrentCol: Integer read FCurrentCol;
    property CurrentCellText: string read FCurrentCellText;
    property Delimiter: TCSVChar read FDelimiter write FDelimiter;
    property QuoteChar: TCSVChar read FQuoteChar write FQuoteChar;
  end;

  { TCSVList }

  TCSVDocument = class(TObject)
  private
    FRows: TObjectList;
    FDelimiter: TCSVChar;
    FQuoteChar: TCSVChar;
    FLineEnding: string;
    FParser: TCSVParser;
    FLastParsedRow: Integer;
    // helpers
    procedure ForceRowIndex(ARowIndex: Integer);
    function  CreateNewRow(const AFirstCell: string = ''): TObject;
    // property getters/setters
    function  GetCell(ACol, ARow: Integer): string;
    procedure SetCell(ACol, ARow: Integer; const AValue: string);
    function  GetCSVText: string;
    procedure SetCSVText(const AValue: string);
    function  GetColCount(ARow: Integer): Integer;
    function  GetRowCount: Integer;
    // document reading stuff
    procedure ParserOnNewCell(Parser: TCSVParser);
    procedure UpdateParserSpecialChars;
  public
    constructor Create;
    destructor  Destroy; override;
    // input/output
    procedure LoadFromFile(AFilename: string);
    procedure LoadFromStream(AStream: TStream);
    procedure SaveToFile(AFilename: string);
    procedure SaveToStream(AStream: TStream);
    // row and cell operations
    procedure AddRow(AFirstCell: string = '');
    procedure AddCell(ARow: Integer; const AValue: string = '');
    procedure InsertRow(ARow: Integer; const AFirstCell: string = '');
    procedure InsertCell(ACol, ARow: Integer; const AValue: string = '');
    procedure RemoveRow(ARow: Integer);
    procedure RemoveCell(ACol, ARow: Integer);
    function  HasRow(ARow: Integer): Boolean;
    function  HasCell(ACol, ARow: Integer): Boolean;
    // utils
    procedure CloneRow(ARow, AInsertPos: Integer);
    procedure ExchangeRows(ARow1, ARow2: Integer);
    procedure UnifyEmbeddedLineEndings;
    procedure TrimEmptyCells;
    // properties
    property Cells[ACol, ARow: Integer]: string read GetCell write SetCell;
    property RowCount: Integer read GetRowCount;
    property ColCount[ARow: Integer]: Integer read GetColCount;
    property Delimiter: TCSVChar read FDelimiter write FDelimiter;
    property QuoteChar: TCSVChar read FQuoteChar write FQuoteChar;
    property LineEnding: string read FLineEnding write FLineEnding;
    property CSVText: string read GetCSVText write SetCSVText;
  end;

function QuoteCSVString(const AValue: string; const AQuoteChar,
  ADelimiter: TCSVChar): string;

implementation

const
  CR = #13;
  LF = #10;

procedure AppendStringToStream(const AString: String; AStream: TStream);
begin
  if (AString <> '') and (Assigned(AStream)) then
    AStream.WriteBuffer(AString[1], Length(AString));
end;

function ReadStringFromStream(AStream: TStream): string;
var
  buf: string;
  sz: Integer;
begin
  if Assigned(AStream) then
  begin
    sz := AStream.Size - AStream.Position;
    SetLength(buf, sz);
    AStream.ReadBuffer(Pointer(buf)^, sz);
    Result := buf;
  end else
    Result := '';
end;

function QuoteCSVString(const AValue: string; const AQuoteChar,
  ADelimiter: TCSVChar): string;
var
  i: Integer;
  needQuotation: Boolean;
begin
  needQuotation := False;
  for i := 1 to Length(AValue) do
  begin
    if AValue[i] in [CR, LF, ADelimiter, AQuoteChar] then
    begin
      needQuotation := True;
      Break;
    end;
  end;

  Result := AValue;
  if needQuotation then
  begin
    // double existing quotes
    Result := StringReplace(Result, AQuoteChar,
      AQuoteChar + AQuoteChar, [rfReplaceAll]);
    Result := AQuoteChar + Result + AQuoteChar;
  end;
end;

{ TCSVParser }

procedure TCSVParser.ClearCellText;
begin
  FCurrentCellText := '';
end;

procedure TCSVParser.IncCellText;
begin
  { TODO: Speed this up by memory preallocation }
  FCurrentCellText := FCurrentCellText + FCurrentChar;
end;

function TCSVParser.EOL: Boolean;
begin
  Result := (FCurrentChar = CR) or (FCurrentChar = LF);
end;

procedure TCSVParser.SkipEOL;
begin
  if (FCurrentChar = CR) then
    NextChar;
  if EOF then
    Exit;
  if (FCurrentChar = LF) then
    NextChar;
end;

procedure TCSVParser.SkipDelimiter;
begin
  if FCurrentChar = FDelimiter then
    NextChar;
end;

procedure TCSVParser.NextChar;
var
  csize: Integer;
begin
  csize := SizeOf(FCurrentChar);
  if FStream.Position <= FStream.Size - csize then
    FStream.ReadBuffer(FCurrentChar, csize)
  else
  begin
    FCurrentChar := #0;
    EOF := True;
  end;
end;

procedure TCSVParser.ParseInput;
begin
  FCurrentCellText := '';
  FCurrentRow := 0;
  FCurrentCol := 0;
  EOF := False;
  NextChar;
  while not EOF do
  begin
    ParseRow;
    Inc(FCurrentRow);
    if FStopFlag then
      Exit;
  end;
end;

procedure TCSVParser.ParseRow;
begin
  FCurrentCol := 0;
  // process first cell separately because it
  // does not start with delimiter unlike other cells
  if not (EOL or EOF) then
  begin
    ParseCell;
    Inc(FCurrentCol);
    if FStopFlag then
      Exit;
  end;
  // process other cells
  while not (EOL or EOF) do
  begin
    SkipDelimiter;
    ParseCell;
    Inc(FCurrentCol);
    if FStopFlag then
      Exit;
  end;
  SkipEOL;
end;

procedure TCSVParser.ParseCell;
begin
  if FCurrentChar = FQuoteChar then
    ParseQuotedValue
  else
    ParseValue;
  DoOnNewCell;
end;

procedure TCSVParser.DoParseValue;
begin
  while (FCurrentChar <> FDelimiter) and not (EOL or EOF) do
  begin
    IncCellText;
    NextChar;
  end;
end;

procedure TCSVParser.ParseQuotedValue;
var
  quotationEnd: Boolean;
begin
  ClearCellText;
  NextChar; // skip opening quotation char
  if EOF then
    Exit;
  repeat
    // read value up to next quotation char
    while (FCurrentChar <> FQuoteChar) and not EOF do
    begin
      IncCellText;
      NextChar;
    end;
    NextChar; // skip quotation char (closing or escaping)
    if EOF then
      Exit;
    // check if it was escaping
    if FCurrentChar = FQuoteChar then
    begin
      IncCellText;
      quotationEnd := False;
      NextChar;
    end else
      quotationEnd := True;
  until quotationEnd;
  // read the rest of the value until separator or new line
  DoParseValue;
end;

procedure TCSVParser.ParseValue;
begin
  ClearCellText;
  DoParseValue;
end;

procedure TCSVParser.DoOnNewCell;
begin
  if Assigned(FOnNewCell) then
    FOnNewCell(Self);
end;

procedure TCSVParser.ParseStream(AStream: TStream);
begin
  FStream := AStream;
  FStream.Seek(0, soFromBeginning);
  ParseInput;
end;

procedure TCSVParser.ParseString(const AString: string);
var
  ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  try
    AppendStringToStream(AString, ms);
    ParseStream(ms);
  finally
    FreeAndNil(ms);
  end;
end;

procedure TCSVParser.Stop;
begin
  FStopFlag := True;
end;

//------------------------------------------------------------------------------

type
  TCSVCell = class
  public
    Value: string;
  end;

  TCSVRow = class
  private
    FCells: TObjectList;
    procedure ForceCellIndex(ACellIndex: Integer);
    function  CreateNewCell(const AValue: string): TCSVCell;
    function  GetCellValue(ACol: Integer): string;
    procedure SetCellValue(ACol: Integer; const AValue: string);
    function  GetColCount: Integer;
  public
    constructor Create;
    destructor  Destroy; override;
    // cell operations
    procedure AddCell(const AValue: string = '');
    procedure InsertCell(ACol: Integer; const AValue: string);
    procedure RemoveCell(ACol: Integer);
    function  HasCell(ACol: Integer): Boolean;
    // utilities
    function  Clone: TCSVRow;
    procedure TrimEmptyCells;
    procedure SetValuesLineEnding(const ALineEnding: string);
    // properties
    property CellValue[ACol: Integer]: string read GetCellValue write SetCellValue;
    property ColCount: Integer read GetColCount;
  end;

{ TCSVRow }

function TCSVRow.GetCellValue(ACol: Integer): string;
begin
  if HasCell(ACol) then
    Result := TCSVCell(FCells[ACol]).Value
  else
    Result := '';
end;

procedure TCSVRow.SetCellValue(ACol: Integer; const AValue: string);
begin
  ForceCellIndex(ACol);
  TCSVCell(FCells[ACol]).Value := AValue;
end;

procedure TCSVRow.ForceCellIndex(ACellIndex: Integer);
begin
  while FCells.Count <= ACellIndex do
    AddCell();
end;

function TCSVRow.CreateNewCell(const AValue: string): TCSVCell;
begin
  Result := TCSVCell.Create;
  Result.Value := AValue;
end;

function TCSVRow.GetColCount: Integer;
begin
  Result := FCells.Count;
end;

constructor TCSVRow.Create;
begin
  inherited Create;
  FCells := TObjectList.Create;
end;

destructor TCSVRow.Destroy;
begin
  FreeAndNil(FCells);
  inherited Destroy;
end;

procedure TCSVRow.AddCell(const AValue: string = '');
begin
  FCells.Add(CreateNewCell(AValue));
end;

procedure TCSVRow.InsertCell(ACol: Integer; const AValue: string);
begin
  FCells.Insert(ACol, CreateNewCell(AValue));
end;

procedure TCSVRow.RemoveCell(ACol: Integer);
begin
  if HasCell(ACol) then
    FCells.Delete(ACol);
end;

function TCSVRow.HasCell(ACol: Integer): Boolean;
begin
  Result := (ACol >= 0) and (ACol < FCells.Count);
end;

function TCSVRow.Clone: TCSVRow;
var
  i: Integer;
begin
  Result := TCSVRow.Create;
  for i := 0 to ColCount - 1 do
    Result.AddCell(CellValue[i]);
end;

procedure TCSVRow.TrimEmptyCells;
var
  i: Integer;
  maxcol: Integer;
begin
  maxcol := FCells.Count - 1;
  for i := maxcol downto 0 do
    if (TCSVCell(FCells[i]).Value = '') and (FCells.Count > 1) then
      FCells.Delete(i);
end;

procedure TCSVRow.SetValuesLineEnding(const ALineEnding: string);

  function ChangeLineEndings(const AString, ALineEnding: string): string;
  var
    i: Integer;
  begin
    for i := 1 to Length(AString) do
      if AString[i] in [CR, LF] then
      begin
        // first unify line endings to single-char value
        Result := StringReplace(AString, CR+LF, LF, [rfReplaceAll]);
        Result := StringReplace(Result, CR, LF, [rfReplaceAll]);
        // then replace this single-char value with value we need
        Result := StringReplace(Result, LF, ALineEnding, [rfReplaceAll]);
        Exit;
      end;
    Result := AString;
  end;

var
  i: Integer;
begin
  for i := 0 to FCells.Count - 1 do
    CellValue[i] := ChangeLineEndings(CellValue[i], ALineEnding);
end;

{ TCSVDocument }

procedure TCSVDocument.AddRow(AFirstCell: string = '');
begin
  FRows.Add(CreateNewRow(AFirstCell));
end;

procedure TCSVDocument.AddCell(ARow: Integer; const AValue: string = '');
begin
  ForceRowIndex(ARow);
  TCSVRow(FRows[ARow]).AddCell(AValue);
end;

procedure TCSVDocument.InsertRow(ARow: Integer; const AFirstCell: string = '');
begin
  if HasRow(ARow) then
    FRows.Insert(ARow, CreateNewRow(AFirstCell))
  else
    AddRow(AFirstCell);
end;

procedure TCSVDocument.InsertCell(ACol, ARow: Integer; const AValue: string);
begin
  ForceRowIndex(ARow);
  TCSVRow(FRows[ARow]).InsertCell(ACol, AValue);
end;

constructor TCSVDocument.Create;
begin
  inherited Create;
  FDelimiter := ';';
  FQuoteChar := '"';
  FLineEnding := sLineBreak;
  FRows := TObjectList.Create;
  FParser := TCSVParser.Create;
  UpdateParserSpecialChars;
  FParser.OnNewCell := ParserOnNewCell;
end;

destructor TCSVDocument.Destroy;
begin
  FreeAndNil(FParser);
  FreeAndNil(FRows);
  inherited Destroy;
end;

function TCSVDocument.GetCSVText: string;
var
  ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  try
    SaveToStream(ms);
    ms.Seek(0, soFromBeginning);
    Result := ReadStringFromStream(ms);
  finally
    FreeAndNil(ms);
  end;
end;

procedure TCSVDocument.SetCSVText(const AValue: string);
var
  ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  try
    AppendStringToStream(AValue, ms);
    LoadFromStream(ms);
  finally
    FreeAndNil(ms);
  end;
end;

function TCSVDocument.GetCell(ACol, ARow: Integer): string;
begin
  if HasRow(ARow) then
    Result := TCSVRow(FRows[ARow]).CellValue[ACol]
  else
    Result := '';
end;

procedure TCSVDocument.SetCell(ACol, ARow: Integer; const AValue: string);
begin
  ForceRowIndex(ARow);
  TCSVRow(FRows[ARow]).CellValue[ACol] := AValue;
end;

function TCSVDocument.GetColCount(ARow: Integer): Integer;
begin
  if HasRow(ARow) then
    Result := TCSVRow(FRows[ARow]).ColCount
  else
    Result := 0;
end;

function TCSVDocument.GetRowCount: Integer;
begin
  Result := FRows.Count;
end;

procedure TCSVDocument.ParserOnNewCell(Parser: TCSVParser);
begin
  Cells[Parser.CurrentCol, Parser.CurrentRow] := Parser.CurrentCellText;
end;

procedure TCSVDocument.UpdateParserSpecialChars;
begin
  FParser.Delimiter := FDelimiter;
  FParser.QuoteChar := FQuoteChar;
end;

procedure TCSVDocument.ForceRowIndex(ARowIndex: Integer);
begin
  while FRows.Count <= ARowIndex do
    AddRow();
end;

function TCSVDocument.CreateNewRow(const AFirstCell: string): TObject;
var
  newRow: TCSVRow;
begin
  newRow := TCSVRow.Create;
  if AFirstCell <> '' then
    newRow.AddCell(AFirstCell);
  Result := newRow;
end;

procedure TCSVDocument.LoadFromFile(AFilename: string);
var
  fs: TFileStream;
begin
  fs := TFileStream.Create(AFilename, fmOpenRead or fmShareDenyNone);
  try
    LoadFromStream(fs);
  finally
    fs.Free;
  end;
end;

procedure TCSVDocument.LoadFromStream(AStream: TStream);
begin
  FRows.Clear;
  UpdateParserSpecialChars;
  FLastParsedRow := 0;
  FParser.ParseStream(AStream);
end;

procedure TCSVDocument.SaveToFile(AFilename: string);
var
  fs: TFileStream;
begin
  fs := TFileStream.Create(AFilename, fmCreate);
  try
    SaveToStream(fs);
  finally
    fs.Free;
  end;
end;

procedure TCSVDocument.SaveToStream(AStream: TStream);
var
  i, j, maxcol: Integer;
begin
  for i := 0 to RowCount - 1 do
  begin
    maxcol := ColCount[i] - 1;
    for j := 0 to maxcol do
    begin
      AppendStringToStream(
        QuoteCSVString(Cells[j, i], FQuoteChar, FDelimiter),
        AStream);
      if j < maxcol then
        AppendStringToStream(FDelimiter, AStream);
    end;
    AppendStringToStream(FLineEnding, AStream);
  end;
end;

procedure TCSVDocument.RemoveRow(ARow: Integer);
begin
  if HasRow(ARow) then
    FRows.Delete(ARow);
end;

procedure TCSVDocument.RemoveCell(ACol, ARow: Integer);
begin
  if HasRow(ARow) then
    TCSVRow(FRows[ARow]).RemoveCell(ACol);
end;

function TCSVDocument.HasRow(ARow: Integer): Boolean;
begin
  Result := (ARow >= 0) and (ARow < FRows.Count);
end;

function TCSVDocument.HasCell(ACol, ARow: Integer): Boolean;
begin
  if HasRow(ARow) then
    Result := TCSVRow(FRows[ARow]).HasCell(ACol)
  else
    Result := False;
end;

procedure TCSVDocument.CloneRow(ARow, AInsertPos: Integer);
var
  row: TObject;
begin
  if not HasRow(ARow) then
    Exit;
  row := TCSVRow(FRows[ARow]).Clone;
  if not HasRow(AInsertPos) then
  begin
    ForceRowIndex(AInsertPos - 1);
    FRows.Add(row);
  end else
    FRows.Insert(AInsertPos, row);
end;

procedure TCSVDocument.ExchangeRows(ARow1, ARow2: Integer);
begin
  if not (HasRow(ARow1) and HasRow(ARow2)) then
    Exit;
  FRows.Exchange(ARow1, ARow2);
end;

procedure TCSVDocument.UnifyEmbeddedLineEndings;
var
  i: Integer;
begin
  for i := 0 to FRows.Count - 1 do
    TCSVRow(FRows[i]).SetValuesLineEnding(FLineEnding);
end;

procedure TCSVDocument.TrimEmptyCells;
var
  i: Integer;
begin
  for i := 0 to FRows.Count - 1 do
    TCSVRow(FRows[i]).TrimEmptyCells;
end;

end.
