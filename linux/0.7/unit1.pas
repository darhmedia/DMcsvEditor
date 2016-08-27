unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  Menus, Grids, ComCtrls, StdCtrls, LCLProc, LazHelpHTML, UTF8Process,
   clipbrd, PrintersDlgs,Variants,Printers,LCLIntf, LCLType,lconvencoding,inifiles,csvdocument;

type

  { TForm1 }

  TForm1 = class(TForm)
    FindDialog1: TFindDialog;
    ImageList1: TImageList;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem38: TMenuItem;
    MenuItem39: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem40: TMenuItem;
    MenuItem41: TMenuItem;
    MenuItem42: TMenuItem;
    MenuItem43: TMenuItem;
    MenuItem44: TMenuItem;
    MenuItem45: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem49: TMenuItem;
    MenuItem50: TMenuItem;
    MenuItem51: TMenuItem;
    MenuItem52: TMenuItem;
    saveexit: TMenuItem;
    MenuItem47: TMenuItem;
    MenuItem48: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    PrintDialog1: TPrintDialog;
    StatusBar1: TStatusBar;
    StringGrid1: TStringGrid;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    procedure FindDialog1Find(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of String);

    procedure FormShow(Sender: TObject);
    procedure MenuItem10Click(Sender: TObject);
    procedure MenuItem12Click(Sender: TObject);
    procedure MenuItem13Click(Sender: TObject);
    procedure MenuItem14Click(Sender: TObject);
    procedure MenuItem16Click(Sender: TObject);
    procedure MenuItem17Click(Sender: TObject);
    procedure MenuItem18Click(Sender: TObject);
    procedure MenuItem20Click(Sender: TObject);
    procedure MenuItem21Click(Sender: TObject);
    procedure MenuItem22Click(Sender: TObject);
    procedure MenuItem23Click(Sender: TObject);
    procedure MenuItem24Click(Sender: TObject);
    procedure MenuItem25Click(Sender: TObject);
    procedure MenuItem27Click(Sender: TObject);

    procedure MenuItem30Click(Sender: TObject);
    procedure MenuItem32Click(Sender: TObject);
    procedure MenuItem33Click(Sender: TObject);
    procedure MenuItem35Click(Sender: TObject);
    procedure MenuItem37Click(Sender: TObject);
    procedure MenuItem39Click(Sender: TObject);
    procedure MenuItem41Click(Sender: TObject);
    procedure MenuItem42Click(Sender: TObject);
    procedure MenuItem43Click(Sender: TObject);
    procedure MenuItem45Click(Sender: TObject);
    procedure MenuItem46Click(Sender: TObject);
    procedure MenuItem49Click(Sender: TObject);
    procedure MenuItem52Click(Sender: TObject);

    procedure saveexitClick(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem8Click(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; aCol, aRow: Integer;
      aRect: TRect; aState: TGridDrawState);

  private
    { private declarations }
    procedure feldolgozas(csv:string);
    procedure Sg2Csv(sep:char;ff:string);
    procedure Sg2Xml(F:string);
    procedure PrintGrid(var sGrid : TStringGrid);
    //procedure SGridToHtml(SG: TStringgrid;filenamex:string);
  public
    { public declarations }
  end; 

var
  Form1: TForm1;
  slista,tlista   : tstringlist;
  sep : char = ',';
  line1    : Tcolor = $00eeeeee;
  line2    : Tcolor = clwhite;
  labelc    : Tcolor = clBtnFace;
  sgfont   : Tcolor = clblack;
  save     : Tsavedialog;
  lng,inif : tinifile;
  //vers  : single = 0.6;
  LastRow, LastCol,eid,sid,lid,ATop, ABot : integer;
  BrowserPath, BrowserParams,parameter,filter0,filter1,filter2,filter3,filter4,filter5,filter6,filter7,filter8,filter9,filter10,saveas,assoc,assoc_txt,filename:string;
  s_true :boolean= false;
  csv1,csv2 : TCSVDocument;


 const
  conf = '.dmcsveditor';
  verz  = '0.7';

implementation

{ TForm1  }

function elvalaszto(const elsosor: string): char; // auto separator
var
sepa : array [0..6] of char;
tmpmax,max,i:integer;
c:char;
begin
max := 0;
tmpmax := 0;
sepa[0] := ',';
sepa[1] := ';';
sepa[2] := '*';
sepa[3] := '|';
sepa[4] := ':';
sepa[5] := '$';
sepa[6] := #9;

for i:=0 to high(sepa) do begin

 tmpmax := Length(elsosor) - Length(StringReplace(elsosor, sepa[i], '', [rfReplaceAll, rfIgnoreCase]));

 if max < tmpmax then begin
    max := tmpmax;
    c := sepa[i];
 end;

end;

Result := c;
end;

procedure Sortgrid(Grid : TStringGrid; SortCol:integer);
{A simple exchange sort of grid rows}
var
   i,j : integer;
   temp:tstringlist;
begin
  temp:=tstringlist.create;
  with Grid do
  for i := FixedRows to RowCount - 2 do  {because last row has no next row}
  for j:= i+1 to rowcount-1 do {from next row to end}
  if AnsiCompareText(Cells[SortCol, i], Cells[SortCol,j]) > 0
  then
  begin
    temp.assign(rows[j]);
    rows[j].assign(rows[i]);
    rows[i].assign(temp);
  end;
  temp.free;
end;

procedure tform1.Sg2Xml(F:string);
var
xml:tstringlist;
i,a:integer;
begin
xml := tstringlist.Create;

xml.Add('<?xml version="1.0" encoding="utf-8" standalone="yes"?>');
xml.Add('<!-- Creator: DMcsvEditor v'+verz+' (linux) '+datetostr(now)+' -->');
xml.Add('<grid cols="'+inttostr(stringgrid1.ColCount)+'" rows="'+inttostr(stringgrid1.RowCount)+'">');

for i:=0 to stringgrid1.ColCount - 1 do begin

  for a:=0 to stringgrid1.RowCount - 1 do begin

      if stringgrid1.Cells[i,a] <> '' then
      xml.Add(' <cell col="'+inttostr(i+1)+'" row="'+inttostr(a+1)+'">'+stringgrid1.Cells[i,a]+'</cell>')
      else
      xml.Add(' <cell col="'+inttostr(i+1)+'" row="'+inttostr(a+1)+'" />');

  end;

end;

xml.Add('</grid>');
xml.Text := UTF8Encode(xml.Text);
xml.SaveToFile(F);
xml.Free;
end;
// CSV fájl feldogozása

procedure TForm1.feldolgozas(csv:string);
var
i,a : integer;
begin

csv1 := TCSVDocument.Create;
csv1.Delimiter:= sep;
csv1.LoadFromFile(UTF8ToANSI(csv));
stringgrid1.BeginUpdate;
    try
    stringgrid1.Font.CharSet:= 4;
    stringgrid1.RowCount := csv1.RowCount;
    stringgrid1.ColCount := csv1.ColCount[0];
   for i:=0 to csv1.RowCount -1 do begin

         for a:=0 to csv1.ColCount[i]-1 do begin

                 stringgrid1.Cells[a,i] := sysToutf8(StringReplace(csv1.Cells[a,i],#13#10,' ',[rfReplaceAll]));

         end;

   end;

    finally
      stringgrid1.EndUpdate;
      stringgrid1.FixedRows:= 0;
      stringgrid1.FixedCols:= 0;
      FreeAndNil(csv1);
    end;

end;

procedure TForm1.PrintGrid(var sGrid : TStringGrid);
var
  r, c,x,y: Integer;

begin
     Printer.Title := extractfilename(filename);
     printer.BeginDoc;
     Printer.Canvas.Font.Name  := stringgrid1.Font.Name;
     Printer.Canvas.Font.Size  := stringgrid1.Font.Size;


     x:= 10;
     //y1 := 50;
     for r:=0 to sgrid.RowCount - 1 do begin

         if printer.PageHeight <= (x+250) then begin
            printer.NewPage;
            x := 10;
         end;

         if  r = 0 then begin
             if sgrid.FixedRows > 0 then  begin
             Printer.Canvas.Font.Style := [fsBold];

             end;
         end else  begin
             Printer.Canvas.Font.Style := [];
         end;

        y := 20;
        for c := 0 to sgrid.ColCount - 1 do begin

        printer.Canvas.TextOut(y,x,sGrid.Cells[c,r]);
        y :=  sgrid.ColWidths[c] * 5+y;
        end;
     x := x +80;
     end;

  Printer.EndDoc;
end;

procedure TForm1.MenuItem17Click(Sender: TObject);
begin
     if stringgrid1.AutoFillColumns then begin
        stringgrid1.AutoFillColumns := false;
     end else begin

          stringgrid1.AutoFillColumns := true;

     end;
end;

procedure TForm1.MenuItem16Click(Sender: TObject);
begin
  application.Terminate;
end;

procedure TForm1.MenuItem10Click(Sender: TObject);
var
open : topendialog;
begin
open := topendialog.Create(self);
open.Filter := filter0+' (*.csv;*.tab;*.tsv;*.txt)|*.csv;*.tab;*.tsv;*.txt|'+filter4+' (*.csv)|*.csv|'+filter5+' (*.csv)|*.csv|'+filter6+' (*.csv)|*.csv|'+filter7+' (*.csv)|*.csv|'+filter8+' (*.csv)|*.csv|'+filter9+' (*.csv)|*.csv|'+filter10+' (*.tab;*.tsv)|*.tab;*.tsv';
open.FilterIndex := lid;
if open.Execute then begin


    Filename := open.FileName;
    tlista.LoadFromFile(open.FileName);

    if open.FilterIndex = 1 then begin

      sep := elvalaszto(tlista.Strings[0]);

    end else begin

        if open.FilterIndex = 2 then
           sep := ',' ;

        if open.FilterIndex = 3 then
           sep := ';';

        if open.FilterIndex = 4 then
           sep := '|';

        if open.FilterIndex = 5 then
           sep := '*';

        if open.FilterIndex = 6 then
           sep := ':';

        if open.FilterIndex = 7 then
           sep := '$';

        if open.FilterIndex = 8 then
           sep := #9;

    end;
    lid := open.FilterIndex;
    feldolgozas(FileName);
    Statusbar1.Panels[2].Text := Filename;
    if sep = #9 then
       Statusbar1.Panels[1].Text := 'Tab'
    else
        Statusbar1.Panels[1].Text := sep;


end;

open.Free;

end;


function htmlcolor(szin: TColor): string;
var
a,b,tmp :string;
begin

  tmp := ColorToString(szin);

    if tmp[1] = '$' then begin

      delete(tmp,1,3);
      a := tmp[1] + tmp[2];
      b := tmp[5] + tmp[6];
      delete(tmp,1,2);
      delete(tmp,3,4);
      tmp := '#'+b+tmp+a;

    end;

    if tmp[1]+tmp[2] = 'cl' then begin

      tmp := stringreplace(tmp,'cl','',[rfReplaceAll]);

    end;
  result := strlower(pchar(tmp));
end;

procedure SGridToHtml(SG: TStringgrid;filenamex:string);
var
  i, p: integer;
   Text,wh: string;
   dest :tstringlist;
begin
  dest := tstringlist.Create;
  Dest.Add('<!DOCTYPE html>');
  Dest.Add('<html>');
  Dest.Add('<head>');
  Dest.Add('<meta charset="utf-8" />');
  Dest.Add('<title>DMcsvEditor v'+verz+' linux export: ' + ExtractFileName(filename) + '</title>');
  Dest.Add('<style type="text/css">');
  Dest.Add('<!--');
  Dest.Add('body{font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 12px;color: '+htmlcolor(form1.StringGrid1.Font.Color)+';}');
  Dest.Add('a,a:hover {text-decoration: underline; color: '+htmlcolor(sgfont)+';}');
  Dest.Add('.center {margin: auto;text-align: left; background: '+htmlcolor(form1.Color)+'; border: 1px solid #eee;}');
  Dest.Add('.bg1 {background-color: '+htmlcolor(form1.stringgrid1.color)+';}');
  Dest.Add('.bg1:hover {background-color: #eee;}');
  Dest.Add('.bg2 {background: '+htmlcolor(form1.stringgrid1.AlternateColor)+';}');
  Dest.Add('.bg2:hover {background-color: #eee;}');
  Dest.Add('.fix {background-color: #c0c0c0;}');
  Dest.Add('-->');
  Dest.Add('</style>');
  Dest.Add('</head>');
  Dest.Add('<body>');
  Dest.Add('<div style="text-align: center; font-size: 15px;"><strong>' + ExtractFileName(filename) + '</strong></div>');
  Dest.Add('<div style="text-align: center;">');
  Dest.Add(' <table class="center" cellpadding="1" cellspacing="1">');

  for i := 0 to SG.RowCount - 1 do
  begin

    if odd(i) then begin
       Dest.Add('  <tr class="bg2">');
    end else begin

    if (sg.FixedRows = 1) and (i = 0) then begin
        Dest.Add('  <tr class="fix">');

      end else begin
        Dest.Add('  <tr class="bg1">');
      end
    end;




    for p := 0 to SG.ColCount - 1 do
    begin

      Text := sg.Cells[p, i];
      if Text = '' then wh := '' else wh := ' width: '+inttostr(sg.ColWidths[p]+50)+'px;';
      Dest.Add('   <td style="height: 15px;'+wh+'">'+Text+'</td>');
    end;
    Dest.Add('  </tr>');

  end;
  Dest.Add('  </table>');
  Dest.Add('<br />');
  Dest.Add('<br />');
  Dest.Add('<div style="text-align: center;">Created <strong>DMcsvEditor linux v'+verz+'</strong> by: <a href="http://darhmedia.blogspot.hu/">Darh Media - Tivadar</a></div>');
  Dest.Add('</div>');
  Dest.Add('</body>');
  Dest.Add('</html>');
  dest.Text := UTF8encode(dest.Text);
   dest.SaveToFile(filenamex);
  dest.Free;
end;
procedure TForm1.MenuItem12Click(Sender: TObject);
begin
    if filename <> '' then begin
		Sg2Csv(sep,filename);
		application.MessageBox(pchar(format('File "%s" saved.',[filename])),pchar('Save'),0);
	end else begin
		MenuItem13Click(Sender);
	end;
end;

procedure TForm1.MenuItem13Click(Sender: TObject);
begin
save := tsavedialog.Create(self);
save.Filter := filter4+' (*.csv)|*.csv|'+filter5+' (*.csv)|*.csv|'+filter6+' (*.csv)|*.csv|'+filter7+' (*.csv)|*.csv|'+filter8+' (*.csv)|*.csv|'+filter9+' (*.csv)|*.csv|'+filter10+' (*.tab)|*.tab|'+filter10+' (*.tsv)|*.tsv';
save.FilterIndex := sid;
if save.Execute then begin



if save.FilterIndex = 1 then
  sep := ',';
if save.FilterIndex = 2 then
  sep := ';';
if save.FilterIndex = 3 then
  sep := '|';
if save.FilterIndex = 4 then
  sep := '*';
if save.FilterIndex = 5 then
  sep := ':';
if save.FilterIndex = 6 then
  sep := '$';
if save.FilterIndex = 7 then
  sep := #9;
if save.FilterIndex = 8 then
  sep := #9;

  sid := save.FilterIndex;

  if fileexists(save.FileName) then begin
     if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.filename])),pchar('Save as...'),1) =1 then begin
        sg2csv(sep,save.FileName);
        filename := save.FileName;

     end;
  end else begin

    if sep = #9 then begin

       if save.FilterIndex = 7 then begin

       sg2csv(sep,save.FileName+'.tab');
       filename := save.FileName+'.tab';

       end else begin

       sg2csv(sep,save.FileName+'.tsv');
       filename := save.FileName+'.tsv';

       end;
    end else begin
        sg2csv(sep,save.FileName+'.csv');
        filename := save.FileName+'.csv';
    end;

  end;
  if sep = #9 then
      Statusbar1.Panels[1].Text := 'Tab'
  else
      Statusbar1.Panels[1].Text := sep;

Statusbar1.Panels[2].Text := filename;
end;

save.Free;

end;

procedure XlsWriteCellLabel(XlsStream: TStream; const ACol, ARow: Word;
  const AValue: string);
var
  L: Word;
const
  {$J+}
  CXlsLabel: array[0..5] of Word = ($204, 0, 0, 0, 0, 0);
  {$J-}
begin
  L := Length(AValue);
  CXlsLabel[1] := 8 + L;
  CXlsLabel[2] := ARow;
  CXlsLabel[3] := ACol;
  CXlsLabel[5] := L;
  XlsStream.WriteBuffer(CXlsLabel, SizeOf(CXlsLabel));
  XlsStream.WriteBuffer(Pointer(UTF8ToCP1252(AValue))^, L);
end;


function SaveAsExcelFile(AGrid: TStringGrid; AFileName: string): Boolean;
const
  {$J+} CXlsBof: array[0..5] of Word = ($809, 8, 00, $10, 0, 0); {$J-}
  CXlsEof: array[0..1] of Word = ($0A, 00);
var
  FStream: TFileStream;
  I, J: Integer;
begin
  Result := False;

  FStream := TFileStream.Create(PChar(AFileName), fmCreate or fmOpenWrite);

  try
    CXlsBof[4] := 0;
    FStream.WriteBuffer(CXlsBof, SizeOf(CXlsBof));
    for i := 0 to AGrid.ColCount - 1 do
      for j := 0 to AGrid.RowCount - 1 do
        XlsWriteCellLabel(FStream, I, J, AGrid.cells[i, j]);
    FStream.WriteBuffer(CXlsEof, SizeOf(CXlsEof));
    Result := True;
  finally
    FStream.Free;
  end;
end;
procedure TForm1.MenuItem14Click(Sender: TObject);
begin
 save := tsavedialog.Create(self);
save.Title := 'Export';
save.Filter := filter2+' (*.xls)|*.xls|'+filter3+' (*.html)|*.html|'+filter1+' (*.xml)|*.xml';
save.FilterIndex := eid;
if save.Execute then begin

  if save.FilterIndex = 1 then begin
    if fileexists(save.FileName) then begin

      if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.filename])),pchar(saveas),1) =1 then begin

        SaveAsExcelFile(StringGrid1, save.FileName);

      end;

    end else begin

      SaveAsExcelFile(StringGrid1, save.FileName+'.xls');

    end;

  end;

  if save.FilterIndex = 2 then begin
       if fileexists(save.FileName) then begin

          if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.filename])),pchar(saveas),1) =1 then begin

           SGridToHtml(Stringgrid1,save.FileName);

          end;

       end else begin

            SGridToHtml(Stringgrid1,save.FileName+'.html');
       end;

  end;

  if save.FilterIndex = 3 then begin
       if fileexists(save.FileName) then begin

          if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.filename])),pchar(saveas),1) =1 then begin

           sg2xml(save.FileName);

          end;

       end else begin

            sg2xml(save.FileName+'.xml');
       end;

  end;
eid := save.FilterIndex;
end;

save.Free;
end;

procedure Tform1.Sg2Csv(sep:char;ff:string);
var
i,a : integer;
begin

csv2 := TCSVDocument.Create;
    csv2.Delimiter:= sep;
     csv2.QuoteChar := '"';

    for i:=0 to stringgrid1.RowCount -1 do begin

         for a:=0 to stringgrid1.ColCount -1 do begin

                csv2.Cells[a,i] := stringgrid1.Cells[a,i];
         end;

   end;

    csv2.SaveToFile(ff);
    FreeAndNil(csv2);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
tlista := tstringlist.Create;
filter0 := 'Automatic recognition';
filter1 := 'Xml format';
filter2 := 'Excel format';
filter3 := 'Html documents';
filter4 := 'Comma-separated file';
filter5 := 'Semicolon-separated file';
filter6 := 'Pipe-separated file';
filter7 := 'Asterisk-separated file';
filter8 := 'Colon-separated file';
filter9 := 'Dollar sign-separated file';
filter10 := 'Tab-separated file';

inif := tinifile.Create(getuserdir()+conf);
saveexit.Checked := inif.ReadBool('Main','SettingsSave',saveexit.Checked);
top := inif.ReadInteger('Main','Top',top);
left := inif.ReadInteger('Main','Left',left);
width := inif.ReadInteger('Main','Width',width);
height := inif.ReadInteger('Main','Height',height);
lid := inif.ReadInteger('Main','LID',lid);
sid := inif.ReadInteger('Main','SID',sid);
eid := inif.ReadInteger('Main','EID',eid);

MenuItem17.Checked := inif.ReadBool('Settings','AutoResize',MenuItem17.Checked);
MenuItem18.Checked := inif.ReadBool('Settings','Label',MenuItem18.Checked);
if MenuItem17.Checked then
MenuItem17Click(Sender);
if MenuItem18.Checked then
MenuItem18Click(Sender);


MenuItem24.Checked := inif.ReadBool('View','ToolBar',MenuItem24.Checked);
toolbar1.Visible := MenuItem24.Checked;
MenuItem4.Checked := inif.ReadBool('View','StatusBar',MenuItem4.Checked);
statusbar1.Visible := MenuItem4.Checked;
{aotop.Checked := inif.ReadBool('View','AlwaysOnTop',aotop.Checked);
aotopClick(Sender);
aotopClick(Sender);}

stringgrid1.Color:= stringtocolor(inif.ReadString('Settings','Line1',colortostring(line1)));

stringgrid1.AlternateColor:= stringtocolor(inif.ReadString('Settings','Line2',colortostring(line2)));

stringgrid1.font.Color:= stringtocolor(inif.ReadString('Settings','Font',colortostring(sgfont)));
stringgrid1.FixedColor := stringtocolor(inif.ReadString('Settings','labelc',colortostring(labelc)));

stringgrid1.Repaint;
{stag := inif.ReadInteger('Search','Online',stag);
nyelv_id := inif.ReadInteger('Language','ID',0);
nyelv := inif.ReadString('Language','current',nyelv);}

inif.Free;
// INI READ END

end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  inif := tinifile.Create(getuserdir()+conf);
if saveexit.Checked then begin

inif.WriteBool('Settings','AutoResize',MenuItem17.Checked);
inif.WriteBool('Settings','Label',MenuItem18.Checked);
 showmessage(colortostring(sgfont));
inif.WriteString('Settings','Line1',colortostring(stringgrid1.Color));
inif.WriteString('Settings','Line2',colortostring(stringgrid1.AlternateColor));
inif.WriteString('Settings','Font',colortostring(stringgrid1.font.Color));
inif.WriteString('Settings','labelc',colortostring(stringgrid1.FixedColor));
inif.WriteBool('View','ToolBar',MenuItem24.Checked);
inif.WriteBool('View','StatusBar',MenuItem4.Checked);
//inif.WriteBool('View','AlwaysOnTop',aotop.Checked);

{
inif.WriteInteger('Search','Online',stag);
inif.WriteInteger('Language','ID',nyelv_id);
inif.WriteString('Language','Current',nyelv); }

inif.WriteInteger('Main','Left',left);
inif.WriteInteger('Main','Top',top);
inif.WriteInteger('Main','Width',width);
inif.WriteInteger('Main','Height',height);

inif.WriteInteger('Main','LID',lid);
inif.WriteInteger('Main','SID',sid);
inif.WriteInteger('Main','EID',eid);
end;
inif.WriteBool('Main','SettingsSave',saveexit.Checked);
inif.Free;
slista.Free;
tlista.Free;
end;

procedure TForm1.FormDropFiles(Sender: TObject; const FileNames: array of String
  );
var
   //fName : array [0..Max_Path] of char;
   //FileCount : integer;
   txt1:string;
begin

txt1 := copy(FileNames[0],strlen(pchar(FileNames[0]))-3,strlen(pchar(FileNames[0])));

if (AnsiLowerCase(txt1) <> '.csv') and (AnsiLowerCase(txt1) <> '.tab') and (AnsiLowerCase(txt1) <> '.tsv') and (AnsiLowerCase(txt1) <> '.txt') then begin
       application.MessageBox('This is file extension not supported!','Warning',0);
      end else begin

            Filename := FileNames[0];
            tlista.LoadFromFile(FileNames[0]);

            if elvalaszto(tlista.Strings[0]) <> '' then begin


              sep := elvalaszto(tlista.Strings[0]);


            end else begin

              sep := ',';

            end;

            feldolgozas(Filename);
            Statusbar1.Panels[2].Text := Filename;
            if sep = #9 then
              Statusbar1.Panels[1].Text := 'Tab'
            else
              Statusbar1.Panels[1].Text := sep;



      end;

end;




procedure TForm1.FormShow(Sender: TObject);
var
i:integer;
txt1:string;

begin
//langtolang(nyelv);
sep := ',';
Statusbar1.Panels[1].Text := sep;
if ParamCount > 0 then begin
   for i := 1 to ParamCount  do
     begin
        parameter:=parameter+ParamStr(i)+'';
     end;
   parameter:=''+parameter+'';


   if FileExists(parameter) then begin

        filename := parameter;
        txt1 := copy(parameter,strlen(pchar(parameter))-3,strlen(pchar(parameter)));

        if (AnsiLowerCase(txt1) = '.csv') or (AnsiLowerCase(txt1) = '.tab') or (AnsiLowerCase(txt1) = '.tsv') or (AnsiLowerCase(txt1) = '.txt') then begin
	     tlista.LoadFromFile(filename);
           if elvalaszto(tlista.Strings[0]) <> '' then begin

             sep := elvalaszto(tlista.Strings[0]);

          end else begin

            sep := ',';

          end;

            feldolgozas(filename);
            Statusbar1.Panels[2].Text := Filename;
            if sep = #9 then
              Statusbar1.Panels[1].Text := 'Tab'
            else
              Statusbar1.Panels[1].Text := sep;

        end else begin

           application.MessageBox('This is file extension not supported!','Warning',0);

        end;
   end;
end;
end;

procedure TForm1.FindDialog1Find(Sender: TObject);
var
  CurX, CurY, GridWidth, GridHeight: integer;
  X, Y: integer;
  TargetText: string;
  CellText: string;
  i: integer;
  GridRect: TGridRect;
label
  TheEnd;
begin
  CurX := StringGrid1.Selection.Left + 1;
  CurY := StringGrid1.Selection.Top;
  GridWidth := StringGrid1.ColCount;
  GridHeight := StringGrid1.RowCount;
  Y := CurY;
  X := CurX;
  if frMatchCase in FindDialog1.Options then
    TargetText := FindDialog1.FindText
  else
    TargetText := AnsiLowerCase(FindDialog1.FindText);
  while Y < GridHeight do
  begin
    while X < GridWidth do
    begin
      if frMatchCase in FindDialog1.Options then
        CellText := StringGrid1.Cells[X, Y]
      else
        CellText := AnsiLowerCase(StringGrid1.Cells[X, Y]);
      i := Pos(TargetText, CellText) ;

      if i > 0 then
      begin

        GridRect.Left := X;
        GridRect.Right := X;
        GridRect.Top := Y;
        GridRect.Bottom := Y;

          with StringGrid1 do
            Begin
            Col := GridRect.Left;
            Row := GridRect.Bottom;
            Selection := GridRect;
        {GetParentForm(StringGrid1).SetFocus;}
    SetFocus;
    s_true := true;
    //StringGrid1.EditorMode := true;
    {TCustomEdit(Components[0]).SelStart := i - 1;
    TCustomEdit(Components[0]).SelLength := length(TargetText);}

          end;

        goto TheEnd;
      end;
      inc(X);
    end;
    inc(Y);
    X := StringGrid1.FixedCols;

  end;
TheEnd:

end;

procedure TForm1.MenuItem18Click(Sender: TObject);
begin
  if stringgrid1.Fixedrows > 0 then begin
         stringgrid1.Fixedrows := 0;
  end else begin
        stringgrid1.Fixedrows := 1;
  end;

end;

procedure TForm1.MenuItem20Click(Sender: TObject);
begin
    stringgrid1.ColCount:= stringgrid1.ColCount + 1;

end;

procedure TForm1.MenuItem21Click(Sender: TObject);
begin
if stringgrid1.ColCount <> 1 then
stringgrid1.ColCount:= stringgrid1.ColCount - 1;
end;

procedure TForm1.MenuItem22Click(Sender: TObject);
begin
  stringgrid1.RowCount:=  stringgrid1.RowCount + 1;
end;

procedure TForm1.MenuItem23Click(Sender: TObject);
begin
if  stringgrid1.RowCount <> 1 then
  stringgrid1.RowCount:=  stringgrid1.RowCount - 1;
end;

procedure TForm1.MenuItem24Click(Sender: TObject);
begin
     if toolbar1.Visible then begin
  toolbar1.Visible:= false;

  end else begin
   toolbar1.Visible:= true;

  end;
end;

procedure TForm1.MenuItem25Click(Sender: TObject);
var
  v: THTMLBrowserHelpViewer;
  p: LongInt;
  URL: String;
  BrowserProcess: TProcessUTF8;
begin
  v:=THTMLBrowserHelpViewer.Create(nil);
  try
    v.FindDefaultBrowser(BrowserPath,BrowserParams);

    //debugln(['Path=',BrowserPath,' Params=',BrowserParams]);

    URL:='http://darhmedia.blogspot.hu';
    p:=System.Pos('%s', BrowserParams);
    System.Delete(BrowserParams,p,2);
    System.Insert(URL,BrowserParams,p);

    // start browser
    BrowserProcess:=TProcessUTF8.Create(nil);
    try
      BrowserProcess.CommandLine:=BrowserPath+' '+BrowserParams;
      BrowserProcess.Execute;
    finally
      BrowserProcess.Free;
    end;
  finally
    v.Free;
  end;

end;

procedure TForm1.MenuItem27Click(Sender: TObject);
var
  v: THTMLBrowserHelpViewer;
  d: string;
  p: LongInt;
  URL: String;
  BrowserProcess: TProcessUTF8;
begin
  v:=THTMLBrowserHelpViewer.Create(nil);
  try
    v.FindDefaultBrowser(BrowserPath,BrowserParams);

    //debugln(['Path=',BrowserPath,' Params=',BrowserParams]);
     d := 'mailto:darhmedia@gmail.com?subject="DMcsvEditor '+verz+' (linux)"';
     url := d;
    p:=System.Pos('%s', BrowserParams);
    System.Delete(BrowserParams,p,2);
    System.Insert(URL,BrowserParams,p);

    // start browser
    BrowserProcess:=TProcessUTF8.Create(nil);
    try
      BrowserProcess.CommandLine:=BrowserPath+' '+BrowserParams;
      BrowserProcess.Execute;
    finally
      BrowserProcess.Free;
    end;
  finally
    v.Free;
  end;

end;

procedure TForm1.MenuItem30Click(Sender: TObject);
begin


if printdialog1.Execute then begin

PrintGrid(StringGrid1) ;
end;


end;

procedure TForm1.MenuItem32Click(Sender: TObject);
var
  MyBitmap: TBitmap;
  MyDC: HDC;
  jpg :tjpegimage;
  save:tsavedialog;
begin

  MyDC := GetDC(form1.Handle);
  MyBitmap := TBitmap.Create;


    MyBitmap.LoadFromDevice(MyDC);

  mybitmap.Height:= height-50;


  jpg := tjpegimage.Create;
jpg.Assign(mybitmap);
save := tsavedialog.Create(self);
save.Title := 'Screenshot';
save.Filter := 'Jpeg format (*.jpg)|*.jpg|'+'Bitmap format (*.bmp)|*.bmp';


if save.Execute then begin

  if save.FilterIndex = 1 then begin

    if fileexists(save.FileName) or fileexists(save.FileName+'.jpg') then begin

      if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.filename])),pchar('Screenshot'),1) =1 then begin

        jpg.SaveToFile(save.FileName);

      end;

    end else begin

       jpg.SaveToFile(save.FileName+'.jpg');
    end;

  end else begin
      if fileexists(save.FileName) or fileexists(save.FileName+'.bmp') then begin

      if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.filename])),pchar('Screenshot'),1) =1 then begin

        mybitmap.SaveToFile(save.FileName);

      end;

    end else begin

       mybitmap.SaveToFile(save.FileName+'.bmp');
    end;

  end;

end;

ReleaseDC(form1.Handle, MyDC);
    FreeAndNil(MyBitmap);
save.Free;
jpg.Free;
//mybitmap.Free;

end;

procedure TForm1.MenuItem33Click(Sender: TObject);
begin
    clipboard.AsText := stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row];
end;

procedure TForm1.MenuItem35Click(Sender: TObject);
begin
if Clipboard.HasFormat(CF_TEXT) then
stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row] :=  clipboard.AsText;
end;

procedure TForm1.MenuItem37Click(Sender: TObject);
begin
  finddialog1.Execute;
end;

procedure TForm1.MenuItem39Click(Sender: TObject);
begin
if stringgrid1.RowCount <> 1 then begin
     if application.MessageBox(pchar('Are you sure delete #'+inttostr(stringgrid1.row+1)+' row?'),'Question',1) = 1 then begin

        stringgrid1.DeleteColRow(false,stringgrid1.row);

        end;
end;
end;

procedure TForm1.MenuItem41Click(Sender: TObject);
var
colorb :tcolordialog;
begin
colorb := tcolordialog.Create(self);
if colorb.Execute then begin
  stringgrid1.Color:= colorb.Color;
end;
  colorb.Free;
end;

procedure TForm1.MenuItem42Click(Sender: TObject);
var
colorb :tcolordialog;
begin
colorb := tcolordialog.Create(self);
if colorb.Execute then begin
  stringgrid1.AlternateColor:= colorb.Color;
end;
  colorb.Free;

end;

procedure TForm1.MenuItem43Click(Sender: TObject);
var
colorb :tcolordialog;
begin
colorb := tcolordialog.Create(self);
if colorb.Execute then begin
  stringgrid1.font.Color:= colorb.Color;
end;
  colorb.Free;

end;

procedure TForm1.MenuItem45Click(Sender: TObject);
begin
   stringgrid1.font.Color:= sgfont;
   stringgrid1.Color:=  line2;
   stringgrid1.AlternateColor:= line1;
   stringgrid1.FixedColor:= labelc;
end;

procedure TForm1.MenuItem46Click(Sender: TObject);

  var
colorb :tcolordialog;
begin
colorb := tcolordialog.Create(self);
if colorb.Execute then begin
  stringgrid1.FixedColor := colorb.Color;
end;
  colorb.Free;
end;

procedure TForm1.MenuItem49Click(Sender: TObject);
begin
  sortgrid(stringgrid1,stringgrid1.Col);
end;

procedure TForm1.MenuItem52Click(Sender: TObject);
var
  v: THTMLBrowserHelpViewer;
  p: LongInt;
  URL: String;
  BrowserProcess: TProcessUTF8;
begin
  v:=THTMLBrowserHelpViewer.Create(nil);
  try
    v.FindDefaultBrowser(BrowserPath,BrowserParams);


    URL:='https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=WMTEWGNAQ4Y9J';
    p:=System.Pos('%s', BrowserParams);
    System.Delete(BrowserParams,p,2);
    System.Insert(URL,BrowserParams,p);

    BrowserProcess:=TProcessUTF8.Create(nil);
    try
      BrowserProcess.CommandLine:=BrowserPath+' '+BrowserParams;
      BrowserProcess.Execute;
    finally
      BrowserProcess.Free;
    end;
  finally
    v.Free;
  end;

end;

procedure TForm1.saveexitClick(Sender: TObject);
begin
  if saveexit.Checked then begin
saveexit.Checked := false;
end else begin
saveexit.Checked := true;
end;
end;


procedure TForm1.MenuItem4Click(Sender: TObject);
begin
  if statusbar1.Visible then begin
  statusbar1.Visible:= false;

  end else begin
   statusbar1.Visible:= true;

  end;
end;

procedure TForm1.MenuItem8Click(Sender: TObject);
var
w:tform;
a,b,c,d,e:tlabel;
begin

w := tform.Create(self);

  with w do begin
    caption := 'DMcsvEditor '+stringreplace('About...','&','',[rfreplaceall]);
    position:= pomainformcenter;
    borderstyle:= bsdialog;
    width:= 300;
    height:= 150;
    color:= $00eeeeee;

  end;

a := tlabel.Create(self);
b := tlabel.Create(self);
c := tlabel.Create(self);
d := tlabel.Create(self);
e := tlabel.Create(self);
  with a do begin

    parent := w;
    top := 10;
    alignment := tacenter;
    font.Size := 12;
    font.Style := [fsbold];
    width:= 290;
    autosize:=false;
    font.color:=$00c76001;


  end;

  with b do begin

    parent := w;
    top := 40;
    alignment := tacenter;
    font.Size := 10;
    font.Style := [fsbold];
    width:= 290;
    autosize:=false;
    font.color:=$000174f0;


  end;

  with c do begin

    parent := w;
    top := 60;
    alignment := tacenter;
    font.Size := 10;
    font.Style := [fsunderline];
    width:= 290;
    autosize:=false;
    font.color:=$00000000;


  end;

  with d do begin

    parent := w;
    top := 80;
    alignment := tacenter;
    font.Size := 8;
    font.Style := [fsbold];
    width:= 290;
    autosize:=true;
    font.color:=$00666666;
    wordwrap := true;


  end;
   with e do begin

    parent := w;
    top := 100;
    alignment := tacenter;
    font.Size := 8;
    font.Style := [fsbold];
    width:= 290;
    autosize:=false;
    font.color:=$00666666;
    wordwrap := true;


  end;


a.Caption := 'DMcsvEditor v'+verz;
b.Caption := 'Darh Media - Tivadar';
c.Caption := 'Thanks to:';
d.Caption := 'Vladimir Zhirov, Christian Ebenegger, Pinvoke';
e.caption := 'Lazarus, UPX team';
w.ShowModal;

a.Free;
b.Free;
c.Free;
d.Free;
e.free;
w.Free;
end;

procedure TForm1.MenuItem9Click(Sender: TObject);
var
  i: integer;
begin
//sep := ',';
filename := '';
tlista.Clear;
Statusbar1.Panels[2].Text := '';
//Statusbar1.Panels[1].Text := ',';
for I := 0 to StringGrid1.RowCount - 1 do begin
 StringGrid1.Rows[I].Clear();
end;
StringGrid1.RowCount := 18;
StringGrid1.ColCount := 6;
StringGrid1.FixedRows := 0;


end;

procedure TForm1.StringGrid1Click(Sender: TObject);
begin
  s_true := false;
end;


procedure TForm1.StringGrid1DrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
begin
  Statusbar1.Panels[0].Text := 'X:' + IntToStr(stringgrid1.row+1) +
 '  Y:' + IntToStr(stringgrid1.Col+1);
with (Sender as TStringGrid) do begin


 if s_true then begin

    if (arow = stringgrid1.row)and (acol =stringgrid1.col) then begin
       Canvas.Brush.Color:= clyellow;
      Canvas.FillRect(aRect);

     end;
   end;
 end;
end;



initialization
  {$I unit1.lrs}

end.

