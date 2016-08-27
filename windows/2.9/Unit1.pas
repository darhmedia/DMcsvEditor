unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids,  ToolWin, Menus, ImgList,inifiles, HTTPGet,
  Clipbrd,shellapi, StdCtrls, jpeg,Printers,ShlObj, ComCtrls,
  ExtCtrls,unit2,csvdocument, XPMan;

type
  TForm1 = class(TForm)
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    ToolBar1: TToolBar;
    StatusBar1: TStatusBar;
    StringGrid1: TStringGrid;
    New1: TMenuItem;
    View1: TMenuItem;
    Edit1: TMenuItem;
    Rows1: TMenuItem;
    cols1: TMenuItem;
    Add1: TMenuItem;
    Remove1: TMenuItem;
    Add2: TMenuItem;
    Remove2: TMenuItem;
    oolbar1: TMenuItem;
    Statusbar2: TMenuItem;
    Colors1: TMenuItem;
    Line11: TMenuItem;
    Line21: TMenuItem;
    Settings1: TMenuItem;
    Label1: TMenuItem;
    Search1: TMenuItem;
    Find1: TMenuItem;
    Language1: TMenuItem;
    English1: TMenuItem;
    N2: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    Font1: TMenuItem;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ImageList1: TImageList;
    Open1: TMenuItem;
    Save1: TMenuItem;
    Saveas1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Export1: TMenuItem;
    Print1: TMenuItem;
    N5: TMenuItem;
    Exit1: TMenuItem;
    Help1: TMenuItem;
    About1: TMenuItem;
    Searchonline1: TMenuItem;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    PopupMenu1: TPopupMenu;
    Autoresize1: TMenuItem;
    N6: TMenuItem;
    Deafult1: TMenuItem;
    N7: TMenuItem;
    sort: TMenuItem;
    DeleteSelected1: TMenuItem;
    Screenshot1: TMenuItem;
    ToolButton15: TToolButton;
    Homepage1: TMenuItem;
    Email1: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    Checkversion1: TMenuItem;
    HTTPGet1: THTTPGet;
    HTTPGet2: THTTPGet;
    ToolButton16: TToolButton;
    saveexit: TMenuItem;
    aotop: TMenuItem;
    SearchEngines1: TMenuItem;
    FindDialog1: TFindDialog;
    Donate1: TMenuItem;
    XPManifest1: TXPManifest;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ranslators1: TMenuItem;
    PayPal1: TMenuItem;
    Flattr1: TMenuItem;
    CheckUpdatetoStart1: TMenuItem;

    procedure Resized(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure Add1Click(Sender: TObject);
    procedure Add2Click(Sender: TObject);
    procedure Remove1Click(Sender: TObject);
    procedure Remove2Click(Sender: TObject);
    procedure oolbar1Click(Sender: TObject);
    procedure Statusbar2Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
    Rect: TRect; State: TGridDrawState);
    procedure Line11Click(Sender: TObject);
    procedure Line21Click(Sender: TObject);
    procedure Font1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure langtolang(lang: string);
    procedure Find1Click(Sender: TObject);
    procedure menuitemsclick(Sender: TObject);
    procedure English1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Label1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Open1Click(Sender: TObject);
    procedure StringGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Save1Click(Sender: TObject);
    procedure Searchonline1Click(Sender: TObject);
    procedure netsearch(Sender: TObject);
    procedure SearchItemsclick(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure Export1Click(Sender: TObject);
    procedure Autoresize1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Deafult1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure sortClick(Sender: TObject);
    procedure DeleteSelected1Click(Sender: TObject);
    procedure Print1Click(Sender: TObject);
    procedure Screenshot1Click(Sender: TObject);
    procedure Saveas1Click(Sender: TObject);
    procedure Checkversion1Click(Sender: TObject);
    procedure HTTPGet1DoneString(Sender: TObject; Result: String);
    procedure HTTPGet1Error(Sender: TObject);
    procedure Email1Click(Sender: TObject);
    procedure Homepage1Click(Sender: TObject);
    procedure HTTPGet2Error(Sender: TObject);
    procedure HTTPGet2DoneFile(Sender: TObject; FileName: String;
      FileSize: Integer);
    procedure HTTPGet2Progress(Sender: TObject; TotalSize,
      Readed: Integer);
    procedure saveexitClick(Sender: TObject);
    procedure aotopClick(Sender: TObject);
    procedure SearchEngines1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure FindDialog1Find(Sender: TObject);
    procedure PayPal1Click(Sender: TObject);
    procedure Flattr1Click(Sender: TObject);
    procedure ranslators1Click(Sender: TObject);
    procedure CheckUpdatetoStart1Click(Sender: TObject);

  private
    { Private declarations }
    procedure ShowCellHint(X,Y:Integer);
    procedure WMDropFiles(var Msg: TWMDropFiles); message WM_DROPFILES;
    procedure feldolgozas(csv:string);
    procedure Sg2Csv(sep:char;ff:string);
    procedure Sg2Xml(F:string);
    procedure PrintGrid(var GenStrGrid : TStringGrid);

  public
    { Public declarations }
  end;

var
  Form1    : TForm1;
  line1    : Tcolor = $00eeeeee;
  line2    : Tcolor = clwhite;
  sgfont   : Tcolor = clblack;
  save     : Tsavedialog;
  lng,inif : tinifile;
  nyelv_id,stag,col,PosComma,LastRow, LastCol,sid,lid,eid,ATop, ABot: integer;
  nyelv,kereso,filename,sRecord,sField ,cv,cv2,alert_txt,notcsv,down,prt: string;
  onlyone,saved,overwrite,norecords,scr,updates,server:string;
  odir,location,downfail,notnew,newversion,exporting,screenshot,csvempty,translators:string;
  filter0,filter1,filter2,filter3,filter4,filter5,filter6,filter7,filter8,filter9,filter10,filter11,filter12,saveas,assoc,assoc_txt:string;
  slista,tlista   : tstringlist;
  sep : char = ',';
  s_true :boolean= false;
  csv1,csv2 : TCSVDocument;
  upd : boolean = false; // check update to start

const
  conf = 'settings.cfg';
  ver  = '2,9';
  verz  = '2.9';

implementation

{$R *.dfm}

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



function MyMessageDialog(const Msg: string; DlgType: TMsgDlgType;
  Buttons: TMsgDlgButtons; Captions: array of string): Integer;
var
  aMsgDlg: TForm;
  i,a: Integer;
  dlgButton: TButton;
  CaptionIndex: Integer;
begin
  aMsgDlg := CreateMessageDialog(Msg, DlgType, Buttons);
  captionIndex := 0;
  aMsgDlg.Caption := 'DMcsvEditor';
  amsgdlg.Position := pomainformcenter;
  a := amsgdlg.Width div 2 - 80;
  for i := 0 to aMsgDlg.ComponentCount - 1 do
  begin
    if (aMsgDlg.Components[i] is TButton) then
    begin
      dlgButton := TButton(aMsgDlg.Components[i]);
      dlgButton.Width := 80;
      if CaptionIndex > High(Captions) then Break;
      dlgButton.Caption := Captions[CaptionIndex];
      dlgButton.left := a;
      Inc(CaptionIndex);
      a := a + 83;
    end;

  end;

  Result := aMsgDlg.ShowModal;
end;

procedure TForm1.WMDropFiles ( var Msg : TWMDropFiles ) ;
 var
   fName : array [0..Max_Path] of char;
   FileCount : integer;
   txt1:string;

 begin
    FileCount := DragQueryFile( Msg.Drop,$FFFFFFFF,fName,MAX_PATH );

    if filecount > 1 then begin

     application.MessageBox(pchar(onlyone),pchar(alert_txt),MB_ICONEXCLAMATION);
    end else begin
       DragQueryFile( Msg.Drop,0,fName,MAX_PATH );



      txt1 := copy(fname,strlen(fname)-2,strlen(fname));

      if (AnsiLowerCase(txt1) <> 'csv') and (AnsiLowerCase(txt1) <> 'tab') and (AnsiLowerCase(txt1) <> 'tsv') and (AnsiLowerCase(txt1) <> 'txt') then begin
       application.MessageBox(pchar(notcsv),pchar(alert_txt),MB_ICONEXCLAMATION);
      end else begin

            Filename := fname;
            tlista.LoadFromFile(fname);

            if tlista.count < 1 then begin
                application.MessageBox(pchar(csvempty),pchar(alert_txt),MB_ICONEXCLAMATION);
                exit;
            end;

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
    DragFinish ( msg.Drop );
end;

// XML export

procedure tform1.Sg2Xml(F:string);
var
xml:tstringlist;
i,a:integer;
begin
xml := tstringlist.Create;

xml.Add('<?xml version="1.0" standalone="yes"?>');
xml.Add('<!-- Creator: DMcsvEditor v'+verz+' '+datetostr(now)+' -->');
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


procedure TForm1.netsearch(Sender: TObject);
var
i:integer;
mi:tmenuitem;
d:string;
begin

if fileexists(ExtractFilePath(Application.ExeName)+'search.egs') then begin
slista.Clear;
slista.LoadFromFile(ExtractFilePath(Application.ExeName)+'search.egs');
 for i:= 0 to slista.Count - 1 do begin

  if pos('|',slista.Strings[i]) > 0 then begin

        mi := TMenuItem.Create(popupmenu1);
        popupmenu1.Items.Add(mi);
        d := copy(slista.Strings[i],0,pos('|',slista.Strings[i])-1);
        mi.Caption := d;
        mi.Tag := i+1;
        mi.OnClick := SearchItemsclick;

  end else begin

    if slista.Strings[i] = '' then begin
       mi := TMenuItem.Create(popupmenu1);
        popupmenu1.Items.Add(mi);
        mi.Caption := '-';
        mi.Tag := i+1;
    end;

  end;
 end;
end;

end;

procedure TForm1.MenuItemsClick(Sender: TObject);
begin

with Sender as TMenuItem do
  begin
    langtolang(stringreplace(caption,'&','',[rfReplaceAll]));
    checked := true;
  end;

end;

procedure TForm1.langtolang(lang: string);
var
i:integer;

begin
// MENU ------>
// file
lng := tinifile.Create(ExtractFilePath(Application.ExeName)+'lang\'+lang+'.lng');

file1.Caption := lng.ReadString('Language','file','File');

new1.Caption := lng.ReadString('Language','new','New');
toolbutton1.Hint := lng.ReadString('Language','new','New');
open1.Caption := lng.ReadString('Language','open','Open...');
toolbutton2.Hint := lng.ReadString('Language','open','Open...');
save1.Caption := lng.ReadString('Language','save','Save');
toolbutton3.Hint := lng.ReadString('Language','save','Save');
saveas1.Caption := lng.ReadString('Language','saveas','Save as...');
saveas := lng.ReadString('Language','saveas','Save as...');
toolbutton16.Hint := saveas;
export1.Caption := lng.ReadString('Language','export','Export...');
exporting := lng.ReadString('Language','export','Export...');
toolbutton5.Hint := exporting;

screenshot1.Caption := lng.ReadString('Language','screenshot','Screenshot');
screenshot := lng.ReadString('Language','screenshot','Screenshot');
toolbutton17.Hint := screenshot;

print1.Caption := lng.ReadString('Language','print','Print');
prt := lng.ReadString('Language','print','Print');
toolbutton4.Hint := prt;
exit1.Caption := lng.ReadString('Language','exit','Exit');

edit1.Caption := lng.ReadString('Language','edit','Edit');

copy1.Caption := lng.ReadString('Language','copy','Copy');
toolbutton6.Hint := lng.ReadString('Language','copy','Copy');
paste1.Caption := lng.ReadString('Language','paste','Paste');
toolbutton7.Hint := lng.ReadString('Language','paste','Paste');
rows1.Caption := lng.ReadString('Language','rows','Vertical');
cols1.Caption := lng.ReadString('Language','cols','Horizontal');
add1.Caption := lng.ReadString('Language','add','Add line');
add2.Caption := add1.Caption;
remove1.Caption := lng.ReadString('Language','remove','Remove line');
remove2.Caption := remove1.Caption;
deleteselected1.Caption := lng.ReadString('Language','deleteselected','Delete Selected Line');
toolbutton15.Hint := lng.ReadString('Language','deleteselected','Delete Selected Line');
sort.Caption := lng.ReadString('Language','sort','Sort by selected col');
toolbutton18.Hint := sort.Caption;

search1.Caption := lng.ReadString('Language','search','Search');
find1.Caption := lng.ReadString('Language','find','Find');
Toolbutton9.Hint := lng.ReadString('Language','search','Search');
searchonline1.Caption := lng.ReadString('Language','searchonline','Search Online');
Toolbutton14.Hint := lng.ReadString('Language','searchonline','Search Online');

view1.Caption := lng.ReadString('Language','view','View');
oolbar1.Caption := lng.ReadString('Language','toolbar','Toolbar');
statusbar2.Caption := lng.ReadString('Language','statusbar','Statusbar');
aotop.Caption := lng.ReadString('Language','aotop','Always On Top');

language1.Caption := lng.ReadString('Language','Language','Language');

settings1.Caption := lng.ReadString('Language','settings','Settings');
autoresize1.Caption := lng.ReadString('Language','autoresize','Autoresize');
label1.Caption := lng.ReadString('Language','label','Label');


colors1.Caption := lng.ReadString('Language','colors','Colors');
line11.Caption := lng.ReadString('Language','line1','Line 1');
line21.Caption := lng.ReadString('Language','line2','Line 2');
font1.Caption := lng.ReadString('Language','font','Font');
deafult1.Caption := lng.ReadString('Language','default','Default');

saveexit.Caption := lng.ReadString('Language','saveexit','Save Settings when Exit');
//Associatedcsvfiles1.Caption := lng.ReadString('Language','associated','Associate with (*.csv) files');

help1.Caption := lng.ReadString('Language','help','Help');
about1.Caption := lng.ReadString('Language','about','About...');
homepage1.Caption := lng.ReadString('Language','homepage','Webpage');
email1.Caption := lng.ReadString('Language','email','Email');
checkversion1.Caption := lng.ReadString('Language','checkversion','Check for Updates...');
donate1.Caption := lng.ReadString('Language','donate','Donate');


alert_txt := lng.ReadString('Language','alert_txt','Warning');
notcsv := lng.ReadString('Language','notcsv','This is file extension not supported!');
onlyone := lng.ReadString('Language','onlyone','Only one a file into the editor!');

saved := lng.ReadString('Language','saved','File "%s" saved.');
filter0 := lng.ReadString('Language','filter0','Csv files [automatic recognition]');
filter1 := lng.ReadString('Language','filter1','Xml format');
filter2 := lng.ReadString('Language','filter2','Excel format');
filter3 := lng.ReadString('Language','filter3','Html documents');
filter4 := lng.ReadString('Language','filter4','Comma-separated file');
filter5 := lng.ReadString('Language','filter5','Semicolon-separated file');
filter6 := lng.ReadString('Language','filter6','Jpeg format');
filter7 := lng.ReadString('Language','filter7','Bmp format');

filter8 := lng.ReadString('Language','filter8','Pipe-separated file');
filter9 := lng.ReadString('Language','filter9','Asterisk-separated file');
filter10 := lng.ReadString('Language','filter10','Colon-separated file');

filter11 := lng.ReadString('Language','filter11','Dollar sign-separated file');
filter12 := lng.ReadString('Language','filter12','Tab-separated file');

overwrite := lng.ReadString('Language','overwrite','%s exists. Overwrite this file?');
scr := lng.ReadString('Language','scr','Search...');
norecords := lng.ReadString('Language','norecords','Not found records.');
updates := lng.ReadString('Language','updates','Updates');
server := lng.readstring('Language','server','Unable to connect Darh Media server.');
down := lng.ReadString('Language','down','download');
odir := lng.ReadString('Language','odir','Open directory?');
location := lng.ReadString('Language','location','New %s version download to:');
downfail := lng.ReadString('Language','downfail','Download failed!');
notnew := lng.ReadString('Language','notnew','There is no new version available this time!');
newversion := lng.ReadString('Language','newversion','New %s version available, download now?');

{assoc := lng.ReadString('Language','assoc','File Association');
assoc_txt := lng.ReadString('Language','assoc_txt','The (*.csv) files successfully DMcsvEditor they were associated.');}
SearchEngines1.Caption := lng.ReadString('Language','searchengines','Search Engines');
form2.Label1.Caption := lng.ReadString('Language','format','Format: Name|Url');
form2.Caption := 'DMcsvEditor - '+ SearchEngines1.Caption;
form2.Button1.Caption := lng.ReadString('Language','save','Save');
form2.Button2.Caption := lng.ReadString('Language','cancel','Cancel');
csvempty := lng.ReadString('Language','csvempty','This file is empty or corrupted!');
translators := lng.ReadString('Language','translators','Translators');
ranslators1.Caption := translators;
CheckUpdatetoStart1.Caption := lng.ReadString('Language','StartupUpdate','Startup Update Check');
lng.Free;


for i:=0 to mainmenu1.Items[4].Count - 1 do begin

  if mainmenu1.Items[4].Items[i].Caption = '&'+lang then begin
    mainmenu1.Items[4].Items[i].Checked := true;
    nyelv_id := i;
  end else begin

     mainmenu1.Items[4].Items[i].Checked := false;
  end;


end;
nyelv := lang;

end;

procedure TForm1.Resized(Sender: TObject);   // Cols auto resize BETA ;)
var
i: integer;
begin
  if autoresize1.Checked then begin
  for i := 0 to StringGrid1.ColCount -1 do begin

    stringgrid1.ColWidths[i] := (stringgrid1.Width div (StringGrid1.ColCount));

  end;



  if odd(stringgrid1.ColWidths[StringGrid1.ColCount-1]) then begin

    stringgrid1.ColWidths[StringGrid1.ColCount-1]  := stringgrid1.ColWidths[StringGrid1.ColCount-1] -25;

  end else begin

    stringgrid1.ColWidths[StringGrid1.ColCount-1]  := stringgrid1.ColWidths[StringGrid1.ColCount-1] -25;

  end;

  end;

end;

procedure TForm1.FormResize(Sender: TObject);
begin
  Resized(Sender);
end;

procedure TForm1.Add1Click(Sender: TObject);
begin
  StringGrid1.colCount := StringGrid1.colCount + 1;
end;

procedure TForm1.Add2Click(Sender: TObject);
begin
  StringGrid1.rowCount := StringGrid1.rowCount + 1;
 
  Resized(Sender);
end;

procedure TForm1.Remove1Click(Sender: TObject);
begin
  StringGrid1.colCount := StringGrid1.colCount - 1;
end;

procedure TForm1.Remove2Click(Sender: TObject);
begin
StringGrid1.rowCount := StringGrid1.rowCount - 1;

Resized(Sender);
end;

procedure TForm1.oolbar1Click(Sender: TObject);
begin
 if oolbar1.Checked then begin

   oolbar1.Checked := false;
   toolbar1.Visible := false;

 end else begin

   oolbar1.Checked := true;
   toolbar1.Visible := true;

 end;

end;

procedure TForm1.Statusbar2Click(Sender: TObject);
begin
 if Statusbar2.Checked then begin

   Statusbar2.Checked := false;
   Statusbar1.Visible := false;

 end else begin

   Statusbar2.Checked := true;
   Statusbar1.Visible := true;

 end;
end;


procedure TForm1.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin

  with (Sender as TStringGrid) do begin

    if Odd(ARow) then  begin

      Canvas.Brush.Color:= line2;
      canvas.Font.Color := sgfont;
      Canvas.TextRect(Rect,Rect.Left+2,Rect.Top+2,cells[acol,arow]);
      Canvas.FrameRect(Rect);

    end else begin
      Canvas.Brush.Color:= line1;
      canvas.Font.Color := sgfont;
      Canvas.TextRect(Rect,Rect.Left+2,Rect.Top+2,cells[acol,arow]);
      Canvas.FrameRect(Rect);
    end;

    if (stringgrid1.FixedRows = 1) and (arow = 0) then begin

     Canvas.Brush.Color:= form1.Color;
      canvas.Font.Color := clblack;
      Canvas.TextRect(Rect,Rect.Left+2,Rect.Top+2,cells[acol,arow]);
      Canvas.FrameRect(Rect);
    end;

   if s_true then begin

    if (arow = stringgrid1.row)and (acol =stringgrid1.col) then begin
       Canvas.Brush.Color:= clyellow;
      Canvas.FillRect(Rect);
      canvas.Font.Color := sgfont;
      canvas.Font.Style:=[fsBold];
      Canvas.TextOut(Rect.Left+1,Rect.Top+1,Cells[ACol,ARow]);
     end;
   end;

  end;

Statusbar1.Panels[0].Text := 'X:' + IntToStr(stringgrid1.row+1) +
 '  Y:' + IntToStr(stringgrid1.Col+1);
end;

procedure TForm1.Line11Click(Sender: TObject);
var
cb : Tcolordialog;
begin

cb := Tcolordialog.Create(self);

  if cb.Execute then begin

    line1 := cb.Color;
    stringgrid1.Repaint;

  end;

cb.Free;
end;

procedure TForm1.Line21Click(Sender: TObject);
var
cb : Tcolordialog;
begin

cb := Tcolordialog.Create(self);

  if cb.Execute then begin

    line2 := cb.Color;
    stringgrid1.Repaint;

  end;

cb.Free;

end;

procedure TForm1.Font1Click(Sender: TObject);
var
cb : Tcolordialog;
begin

cb := Tcolordialog.Create(self);

  if cb.Execute then begin

    sgfont := cb.Color;
    stringgrid1.Repaint;

  end;

cb.Free;

end;

procedure TForm1.FormActivate(Sender: TObject);
var
keres: TSearchRec;
it : tmenuitem;
i:integer;
begin



slista := tstringlist.Create;
netsearch(Sender);

if (stag > -1) and (popupmenu1.Items.Count > stag ) then begin
popupmenu1.Items[stag].Checked := true;
popupmenu1.Items[stag].Click;
end;

 if FindFirst(ExtractFilePath(Application.ExeName)+'lang\*.lng', faAnyFile, keres) = 0 then
begin
i:= 1;
      repeat
        it := TMenuItem.Create(mainmenu1);
        keres.Name := stringreplace(keres.Name,'.lng','',[rfReplaceAll]);
        mainmenu1.Items[4].Add(it);
        it.Caption := stringreplace(keres.Name,'_',' ',[rfReplaceAll]);
        it.Tag := i;
        it.Name := keres.Name+'1';
        it.OnClick := menuItemsClick;
        i:= i+1;
      until FindNext(keres) <> 0;
        FindClose(keres);
      end;

// lng itt volt
mainmenu1.Items[4].Items[nyelv_id].Checked := true;


//

end;

procedure TForm1.Find1Click(Sender: TObject);

begin
 FindDialog1.Execute;
end;

procedure TForm1.English1Click(Sender: TObject);
begin
langtolang('English');
end;

procedure TForm1.Exit1Click(Sender: TObject);
begin
application.Terminate;
end;

procedure TForm1.Label1Click(Sender: TObject);
begin
	if stringgrid1.rowCount <> 1 then begin
	
		if label1.Checked then begin
			label1.Checked := false;
			stringgrid1.FixedRows := 0;
		end else begin
			label1.Checked := true;
			stringgrid1.FixedRows := 1;
		end;
	end;
end;

procedure TForm1.Copy1Click(Sender: TObject);
begin
clipboard.AsText := stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row];
end;

procedure TForm1.Paste1Click(Sender: TObject);
begin
if Clipboard.HasFormat(CF_TEXT) then
stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row] :=  clipboard.AsText;
end;

procedure TForm1.Open1Click(Sender: TObject);
var
open : topendialog;
begin
open := topendialog.Create(self);
open.Filter := filter0+' (*.csv;*.tsv;*.tab;*.txt)|*.csv;*.tsv;*.tab;*.txt|'+filter4+' (*.csv)|*.csv|'+filter5+' (*.csv)|*.csv|'+filter8+' (*.csv)|*.csv|'+filter9+' (*.csv)|*.csv|'+filter10+' (*.csv)|*.csv|'+filter11+' (*.csv)|*.csv|'+filter12+' (*.tab;*.tsv)|*.tab;*.tsv';
open.FilterIndex := lid;
if open.Execute then begin


    Filename := open.FileName;
    tlista.LoadFromFile(open.FileName);
    if tlista.count < 1 then begin
      application.MessageBox(pchar(csvempty),pchar(alert_txt),MB_ICONEXCLAMATION);
      exit;
    end;
    if open.FilterIndex = 1 then begin

      sep := elvalaszto(tlista.strings[0]);

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

    feldolgozas(filename);
    Statusbar1.Panels[2].Text := Filename;
    if sep = #9 then
      Statusbar1.Panels[1].Text := 'Tab'
    else
      Statusbar1.Panels[1].Text := sep;

end;

open.Free;

end;

// SG2CSV

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


// CSV fájl feldogozása

procedure TForm1.feldolgozas(csv:string);
var
i,a : integer;
begin

csv1 := TCSVDocument.Create;
csv1.Delimiter:= sep;
csv1.LoadFromFile(csv);

    try
    stringgrid1.Font.CharSet:= 4;
    stringgrid1.RowCount := csv1.RowCount;
    stringgrid1.ColCount := csv1.ColCount[0];
   for i:=0 to csv1.RowCount -1 do begin

         for a:=0 to csv1.ColCount[i]-1 do begin

                 stringgrid1.Cells[a,i] := StringReplace(csv1.Cells[a,i],#13#10,' ',[rfReplaceAll]);

         end;

   end;

    finally
      //stringgrid1.EndUpdate;
      stringgrid1.FixedRows:= 0;
      stringgrid1.FixedCols:= 0;
      FreeAndNil(csv1);
    end;
end;


procedure TForm1.ShowCellHint(X,Y:Integer);
var
  ACol, ARow : Integer;
begin

  If StringGrid1.ShowHint = False Then
     StringGrid1.ShowHint := True;

  StringGrid1.MouseToCell(X, Y, ACol, ARow);

  If (ACol <> -1) And (ARow <> -1) Then
      StringGrid1.Hint:=StringGrid1.Cells[ACol,ARow];
  If (ACol<>LastCol) or (ARow<>LastRow) Then
  begin
    Application.CancelHint;
    LastCol:=ACol;
    LastRow:=ARow;
  end;
end;

procedure TForm1.StringGrid1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
 stringgrid1.ShowHint := true;
  ShowCellHint(X,Y);
end;

procedure TForm1.Save1Click(Sender: TObject);
begin
	if filename <> '' then begin
		Sg2Csv(sep,filename);
		//tlista.SaveToFile(filename);
		application.MessageBox(pchar(format(saved,[filename])),pchar(save1.Caption),MB_ICONINFORMATION);
	end else begin
		Saveas1Click(Sender);
	end;
end;

procedure TForm1.Searchonline1Click(Sender: TObject);
begin

	if stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row] <> '' then begin
		ShellExecute(Handle, nil, pchar(kereso+stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row]), nil, nil, SW_SHOW);
	end;
end;

procedure TForm1.SearchItemsclick(Sender: TObject);
var
i,b:integer;
a:string;
begin

with Sender as TMenuItem do
  begin
  a := stringreplace(caption,'&','',[rfReplaceAll]);
  for i:= 0 to slista.Count - 1 do begin
    if pos(a,slista.Strings[i]) > 0 then begin
    kereso := copy( slista.Strings[i] , pos('|',slista.Strings[i])+1 , strlen(pchar(slista.Strings[i])) );
      if pos('http://',kereso) > 0 then begin  // not found http:// ? added.
       kereso := kereso;
      end else begin
       kereso := 'http://'+ kereso;
      end;
    end;
  end;
   b := tag;
  end;
for i:= 0 to popupmenu1.Items.Count - 1 do begin

  if popupmenu1.Items[i].Tag = b then begin
   popupmenu1.Items[i].Checked := true;
   stag := i;
  end else begin
    popupmenu1.Items[i].Checked := false;
  end;


end;

end;

procedure TForm1.Close1Click(Sender: TObject);
var
  i: integer;
begin
filename := '';
Statusbar1.Panels[2].Text := '';
for I := 0 to StringGrid1.RowCount - 1 do begin
 StringGrid1.Rows[I].Clear();
end;
StringGrid1.RowCount := 24;
StringGrid1.ColCount := 6;
StringGrid1.FixedRows := 0;
label1.Checked := false;
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
  XlsStream.WriteBuffer(Pointer(AValue)^, L);
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
  Dest.Add('<html lang="en">');
  Dest.Add('<head>');
  Dest.Add('<meta charset="utf-8" />');
  Dest.Add('<title>DMcsvEditor v'+verz+' export: ' + ExtractFileName(filename) + '</title>');
  Dest.Add('<style type="text/css">');
  Dest.Add('<!--');
  Dest.Add('body{font-family: verdana, arial, helvetica, sans-serif;font-size: 12px;color: '+htmlcolor(sgfont)+';}');
  Dest.Add('a,a:hover {text-decoration: underline; color: '+htmlcolor(sgfont)+';}');
  Dest.Add('.center {margin: auto;text-align: left; background: '+htmlcolor(form1.Color)+'; border: 1px solid #eee;}');
  Dest.Add('.bg1 {background-color: '+htmlcolor(line1)+';}');
  Dest.Add('.bg1:hover {background-color: #eee;}');
  Dest.Add('.bg2 {background-color: '+htmlcolor(line2)+';}');
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
  Dest.Add('<div style="text-align: center;">Created <strong>DMcsvEditor v'+verz+'</strong> by: <a href="http://darhmedia.hu">Darh Media - Tivadar</a></div>');
  Dest.Add('</div>');
  Dest.Add('</body>');
  Dest.Add('</html>');
  dest.Text := UTF8encode(dest.Text);
   dest.SaveToFile(filenamex);
  dest.Free;
end;

procedure TForm1.Export1Click(Sender: TObject);
begin
save := tsavedialog.Create(self);
save.Title := exporting;
save.Filter := filter2+' (*.xls)|*.xls|'+filter3+' (*.html)|*.html|'+filter1+' (*.xml)|*.xml';
save.FilterIndex := eid;
if save.Execute then begin

  if save.FilterIndex = 1 then begin
    if fileexists(save.FileName) then begin

      if application.MessageBox(pchar(format(overwrite,[save.filename])),pchar(saveas),MB_YESNO+MB_ICONQUESTION) =IDYES then begin

        SaveAsExcelFile(StringGrid1, save.FileName);

      end;

    end else begin

      SaveAsExcelFile(StringGrid1, save.FileName+'.xls');

    end;

  end;

  if save.FilterIndex = 2 then begin
       if fileexists(save.FileName) then begin

          if application.MessageBox(pchar(format(overwrite,[save.filename])),pchar(saveas),MB_YESNO+MB_ICONQUESTION) =IDYES then begin

           SGridToHtml(Stringgrid1,save.FileName);

          end;

       end else begin

            SGridToHtml(Stringgrid1,save.FileName+'.html');
       end;

  end;

  if save.FilterIndex = 3 then begin
       if fileexists(save.FileName) then begin

          if application.MessageBox(pchar(format(overwrite,[save.filename])),pchar(saveas),MB_YESNO+MB_ICONQUESTION) =IDYES then begin

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



procedure TForm1.Autoresize1Click(Sender: TObject);
begin
if autoresize1.Checked then begin
autoresize1.Checked := false;

end else begin
autoresize1.Checked := true;

Resized(Sender);
end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin

inif := tinifile.Create(ExtractFilePath(Application.ExeName)+conf);

if saveexit.Checked then begin

inif.WriteBool('Settings','AutoResize',autoresize1.Checked);
inif.WriteBool('Settings','Label',label1.Checked);

inif.WriteString('Settings','Line1',colortostring(line1));
inif.WriteString('Settings','Line2',colortostring(line2));
inif.WriteString('Settings','Font',colortostring(sgfont));
inif.WriteBool('View','ToolBar',oolbar1.Checked);
inif.WriteBool('View','StatusBar',statusbar2.Checked);
inif.WriteBool('View','AlwaysOnTop',aotop.Checked);
inif.WriteBool('Settings','UpdateCheck',CheckUpdatetoStart1.Checked);


inif.WriteInteger('Search','Online',stag);
inif.WriteInteger('Language','ID',nyelv_id);
inif.WriteString('Language','Current',nyelv);

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

procedure TForm1.FormCreate(Sender: TObject);
var
upcheck : boolean;
begin

tlista := tstringlist.Create;
// INI READ START
inif := tinifile.Create(ExtractFilePath(Application.ExeName)+conf);
saveexit.Checked := inif.ReadBool('Main','SettingsSave',saveexit.Checked);
top := inif.ReadInteger('Main','Top',top);
left := inif.ReadInteger('Main','Left',left);
width := inif.ReadInteger('Main','Width',width);
height := inif.ReadInteger('Main','Height',height);
lid := inif.ReadInteger('Main','LID',lid);
sid := inif.ReadInteger('Main','SID',sid);
eid := inif.ReadInteger('Main','EID',eid);

autoresize1.Checked := inif.ReadBool('Settings','AutoResize',autoresize1.Checked);
label1.Checked := inif.ReadBool('Settings','Label',label1.Checked);
Label1Click(Sender);
Label1Click(Sender);


oolbar1.Checked := inif.ReadBool('View','ToolBar',oolbar1.Checked);
toolbar1.Visible := oolbar1.Checked;
statusbar2.Checked := inif.ReadBool('View','StatusBar',statusbar2.Checked);
statusbar1.Visible := statusbar2.Checked;
aotop.Checked := inif.ReadBool('View','AlwaysOnTop',aotop.Checked);
aotopClick(Sender);
aotopClick(Sender);

//if inif.ReadBool('Menu','Img',false) then
  //mainmenu1.Images := imagelist1;

line1 := stringtocolor(inif.ReadString('Settings','Line1',colortostring(line1)));
line2 := stringtocolor(inif.ReadString('Settings','Line2',colortostring(line2)));
sgfont := stringtocolor(inif.ReadString('Settings','Font',colortostring(sgfont)));
stag := inif.ReadInteger('Search','Online',stag);
nyelv_id := inif.ReadInteger('Language','ID',0);
nyelv := inif.ReadString('Language','current',nyelv);

 upcheck := inif.ReadBool('Settings','UpdateCheck',CheckUpdatetoStart1.Checked);
 CheckUpdatetoStart1.Checked := upcheck;
inif.Free;
// INI READ END

DragAcceptFiles(Handle,true);
if upcheck then
Checkversion1Click(Sender)

end;

procedure TForm1.Deafult1Click(Sender: TObject);
begin
line1    := $00eeeeee;
line2    := clwhite;
sgfont   := clblack;
stringgrid1.Repaint;
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

procedure TForm1.FormShow(Sender: TObject);
var
i:integer;
parameter:string;
begin
langtolang(nyelv);
	if ParamCount > 0 then begin
   for i := 1 to ParamCount  do
     begin
        parameter:=parameter+ParamStr(i)+' ';
     end;
      parameter:=''+parameter+'';

		if FileExists(parameter) then begin

        filename := parameter;
				tlista.LoadFromFile(filename);
        if tlista.count < 1 then begin
                application.MessageBox(pchar(csvempty),pchar(alert_txt),MB_ICONEXCLAMATION);
                exit;
        end;
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

		end;
	end;
end;

procedure TForm1.sortClick(Sender: TObject);
begin
sortgrid(stringgrid1,stringgrid1.Col);
end;

procedure TForm1.DeleteSelected1Click(Sender: TObject);
var
  CurrentRow, R, C: integer;
begin

  if ActiveControl <> StringGrid1 then begin
		  exit;
		end;


 CurrentRow := StringGrid1.Row;

  if StringGrid1.RowCount > 2 then begin

    if CurrentRow< StringGrid1.RowCount - 1 then begin

      for R := CurrentRow to StringGrid1.RowCount - 1 do
        for C := 0 to StringGrid1.ColCount - 1 do
          StringGrid1.Cells[C, R] := StringGrid1.Cells[C, R + 1];
    end;

    StringGrid1.RowCount := StringGrid1.RowCount - 1;
  end
  else
    
    for C := 0 to StringGrid1.ColCount - 1 do
      StringGrid1.Cells[C, CurrentRow] := '';

end;

procedure TForm1.Print1Click(Sender: TObject);
 var
p : TPrinterSetupDialog;
begin
p := TPrinterSetupDialog.Create(self);

if p.Execute then begin

PrintGrid(StringGrid1) ;
end;

p.Free;
end;

procedure TForm1.Screenshot1Click(Sender: TObject);
var
kep :tbitmap;
jpg :tjpegimage;
DCDesk: HDC;

begin
kep := tbitmap.Create;

kep.Height := Height;
kep.width := width;
DCDesk := GetWindowDC(GetDesktopWindow);
BitBlt(kep.Canvas.Handle, 0, 0, Width, Height,DCDesk, left, top, SRCCOPY);

jpg := tjpegimage.Create;
jpg.Assign(kep);
save := tsavedialog.Create(self);
save.Title := screenshot;
save.Filter := filter6+' (*.jpg)|*.jpg|'+filter7+' (*.bmp)|*.bmp';


if save.Execute then begin

  if save.FilterIndex = 1 then begin

    if fileexists(save.FileName) then begin

      if application.MessageBox(pchar(format(overwrite,[save.filename])),pchar(screenshot),MB_YESNO+MB_ICONQUESTION) =IDYES then begin

        jpg.SaveToFile(save.FileName);

      end;

    end else begin

       jpg.SaveToFile(save.FileName+'.jpg');
    end;

  end else begin
      if fileexists(save.FileName) then begin

      if application.MessageBox(pchar(format(overwrite,[save.filename])),pchar(screenshot),MB_YESNO+MB_ICONQUESTION) =IDYES then begin

        kep.SaveToFile(save.FileName);

      end;

    end else begin

       kep.SaveToFile(save.FileName+'.bmp');
    end;

  end;

end;
ReleaseDC(GetDesktopWindow, DCDesk);

save.Free;
jpg.Free;
kep.Free;

end;

procedure TForm1.Saveas1Click(Sender: TObject);
begin
save := tsavedialog.Create(self);
save.Filter := filter4+' (*.csv)|*.csv|'+filter5+' (*.csv)|*.csv|'+filter8+' (*.csv)|*.csv|'+filter9+' (*.csv)|*.csv|'+filter10+' (*.csv)|*.csv|'+filter11+' (*.csv)|*.csv|'+filter12+' (*.tab)|*.tab|'+filter12+' (*.tsv)|*.tsv';
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
     if application.MessageBox(pchar(format(overwrite,[save.filename])),pchar(saveas),MB_YESNO+MB_ICONQUESTION) =IDYES then begin
        sg2csv(sep,save.FileName);
        //tlista.SaveToFile(save.FileName);
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

procedure TForm1.Checkversion1Click(Sender: TObject);
begin

httpget1.PostQuery := 'app=dmcsveditor_donate';
httpget1.URL := 'http://darhmedia.hu/version.php';
httpget1.GetString;


end;

procedure TForm1.HTTPGet1DoneString(Sender: TObject; Result: String);
var

tmp : Tstringlist;
begin
caption := 'DMcsvEditor';
tmp := Tstringlist.Create;
tmp.Text := result;
if tmp.Count > 4 then
tmp.Text := '1'+#10+' ';

cv := tmp.Strings[0];
cv2 := tmp.Strings[1];

if (strtofloat(ver) < strtofloat(cv)) then begin
     if application.MessageBox(pchar(format(newversion,[cv2])),pchar(updates),MB_YESNO+MB_ICONQUESTION) = IDYES then begin

      httpget2.URL := tmp.Strings[3];
      httpget2.FileName := ExtractFilePath(Application.ExeName)+tmp.Strings[2];
      httpget2.GetFile;
     end;
end else begin
  if upd then
    application.MessageBox(pchar(notnew),pchar(updates),MB_ICONINFORMATION)

end;

tmp.Free;
upd := true;
end;

procedure TForm1.HTTPGet1Error(Sender: TObject);
begin

  if upd then
  application.MessageBox(pchar(server),pchar(updates),MB_ICONEXCLAMATION);

upd := true;
end;

procedure TForm1.Email1Click(Sender: TObject);
begin
ShellExecute(Handle, nil, pchar('mailto:info@darhmedia.hu?subject=DMcsvEditor%20v'+verz), nil, nil, SW_SHOW);
end;

procedure TForm1.Homepage1Click(Sender: TObject);
begin
ShellExecute(Handle, nil, 'http://darhmedia.hu', nil, nil, SW_SHOW);
end;


procedure TForm1.HTTPGet2Error(Sender: TObject);
begin
application.MessageBox(pchar(downfail),pchar(updates),MB_ICONEXCLAMATION);
caption := 'DMcsvEditor';

end;

procedure TForm1.HTTPGet2DoneFile(Sender: TObject; FileName: String;
  FileSize: Integer);
begin
if application.MessageBox(pchar(format(location,[cv2])+#10+ExtractFilePath(Application.ExeName)+#10+odir),pchar(updates),MB_YESNO+MB_ICONINFORMATION) = IDYES then begin
ShellExecute(Handle, nil, pchar(ExtractFilePath(Application.ExeName)), nil, nil, SW_SHOW);
end;
caption := 'DMcsvEditor';

end;

procedure TForm1.HTTPGet2Progress(Sender: TObject; TotalSize,
  Readed: Integer);
  var
  t,c:integer;
begin
t := totalsize div 1024;
c := Readed div 1024;
caption := 'DMcsvEditor - v'+cv2+' '+down+': '+inttostr(t)+' KB / '+inttostr(c)+' KB';

end;

procedure TForm1.saveexitClick(Sender: TObject);
begin
if saveexit.Checked then begin
saveexit.Checked := false;
end else begin
saveexit.Checked := true;
end;
end;

procedure TForm1.PrintGrid(var GenStrGrid : TStringGrid);
const // Size of the Page for European Paper Size
FinalBot = 1100; // Portrait and  783 for LandScape;
var
OldCursor : TCursor;
PrnRect : TRect;
I, J, K,  NC, NR, XScale, YScale, LBot : Integer;
GString : String;
LRMatrix : Variant;
begin
OldCursor := Screen.Cursor;
Screen.Cursor := crHourGlass;
NC := (GenStrGrid.ColCount - 1);
NR := (GenStrGrid.RowCount - 1);
LRMatrix := VarArrayCreate([0, NC, 0,1], VarInteger);
I := 0;
try
  begin
   repeat
    if I = 0 then
     begin
      LRMatrix[I,0] := 0;
      LRMatrix[I,1] := GenStrGrid.ColWidths[0];
      Inc(I,1);
     end
    else
     begin
      LRMatrix[I,0] := LRMatrix[(I - 1),0] + GenStrGrid.ColWidths[(I - 1)] + 1;
      LRMatrix[I,1] := LRMatrix[(I - 1),1] + GenStrGrid.ColWidths[I] + 1;
      Inc(I,1);
     end;
   until
    I > NC;

   with Printer do
    begin
     XScale := GetDeviceCaps(Printer.Handle,LogPixelsX) div Self.PixelsPerInch;
     YScale := GetDeviceCaps(Printer.Handle,LogPixelsY) div Self.PixelsPerInch;

     Title := extractfilename(filename);
     Canvas.Font := Self.Font;
     Orientation := printer.Orientation;

     Begindoc;
     J := 0;
     repeat
       if J = 0 then
        begin
         ATop := 0;
         ABot := GenStrGrid.RowHeights[J];
         LBot := ABot;
        end
       else
        begin
         ATop := ATop + GenStrGrid.RowHeights[(J - 1)] + 1;
         ABot := ABot + GenStrGrid.RowHeights[J] + 1;
         LBot := ABot;
        end;
       if LBot > printer.PageWidth  then
        begin
         ATop := 0;
         ABot := GenStrGrid.RowHeights[0];
         NewPage;
         for K := 0 to NC do
          begin
           PrnRect := Rect(LRMatrix[K,0], ATop, LRMatrix[K,1], ABot);
            with PrnRect do
             begin
              Left := Left * XScale;
              Right := (Right - 3) * XScale;
              Top := (20 + Top) * YScale;
              Bottom := (20 + Bottom) * YScale;
             end;
            GString := GenStrGrid.Cells[K,0];
            Canvas.Font.Style := [fsBold] + [fsUnderline];
            Canvas.Pen.Style := psSolid;
            Canvas.Pen.Color := clBlack;
            Canvas.FrameRect(PrnRect);
            DrawText(Canvas.Handle, PChar(GString), StrLen(PChar(GString)), PrnRect, DT_WORDBREAK);  
         end;
          Dec(J,1);
          ABot := GenStrGrid.RowHeights[J];
        end
       else
        begin
         for K := 0 to NC do
          begin
           PrnRect := Rect(LRMatrix[K,0], ATop, LRMatrix[K,1], ABot);
           with PrnRect do
            begin
             Left := Left * XScale;
             Right := (Right - 3) * XScale;
             Top := (20 + Top) * YScale;
             Bottom := (20 + Bottom) * YScale;
            end;
            GString := GenStrGrid.Cells[K,J];
            if J = 0 then
             Canvas.Font.Style := [fsBold] + [fsUnderline]
            else
             Canvas.Font.Style := [fsBold];

            Canvas.Pen.Style := psSolid;
            Canvas.Pen.Color := clBlack;

            //Canvas.Brush.Color := line1;
            Canvas.FrameRect(PrnRect);

            DrawText(Canvas.Handle, PChar(GString), StrLen(PChar(GString)), PrnRect, DT_WORDBREAK);  //DT_WORDBREAK
          end;
        end;
        Inc(J,1);
     until
      J > NR;
     EndDoc;
    end;
  end;
finally
  begin
   Screen.Cursor := OldCursor;
   LRMatrix := UnAssigned;
  end;
end;
end;

procedure TForm1.aotopClick(Sender: TObject);
begin
if aotop.Checked then begin
 aotop.Checked := false;
 SetWindowPos(Handle, // handle to window
               HWND_NOTOPMOST, // placement-order handle {*}
               Left,  // horizontal position
               Top,   // vertical position
               Width,
               Height,
               // window-positioning options
               SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);
end else begin
aotop.Checked := true;
 SetWindowPos(Handle, // handle to window
               HWND_TOPMOST, // placement-order handle {*}
               Left,  // horizontal position
               Top,   // vertical position
               Width,
               Height,
               // window-positioning options
               SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);
end;
end;

procedure TForm1.SearchEngines1Click(Sender: TObject);
begin
 form2.ShowModal;
popupmenu1.Items.Clear;
 netsearch(Sender);

if (stag > -1) and (popupmenu1.Items.Count > stag ) then begin
popupmenu1.Items[stag].Checked := true;
popupmenu1.Items[stag].Click;
end;
end;

procedure TForm1.About1Click(Sender: TObject);
var
w:tform;
a,b,c,d,e,f:tlabel;
begin

w := tform.Create(self);

  with w do begin
    caption := 'DMcsvEditor '+stringreplace(about1.Caption,'&','',[rfreplaceall]);
    position:= pomainformcenter;
    borderstyle:= bsdialog;
    width:= 300;
    height:= 170;
    color:= $00eeeeee;

  end;

a := tlabel.Create(self);
b := tlabel.Create(self);
c := tlabel.Create(self);
d := tlabel.Create(self);
e := tlabel.Create(self);
f := tlabel.Create(self);
  with a do begin

    parent := w;
    top := 10;
    alignment := tacenter;
    font.Size := 14;
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
    font.Style := [fsbold,fsunderline];
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
    autosize:=false;
    font.color:=$00666666;
    wordwrap := true;


  end;
  with e do begin

    parent := w;
    top := 95;
    alignment := tacenter;
    font.Size := 8;
    font.Style := [fsbold];
    width:= 290;
    autosize:=false;
    font.color:=$00666666;
    wordwrap := true;


  end;
  with f do begin

    parent := w;
    top := 110;
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
d.Caption := 'Vladimir Zhirov, UtilMind, Prog.hu,';
e.Caption := 'Christian Ebenegger, Pinvoke, UPX team';
f.caption := 'and my muse Monesz!';
w.ShowModal;

a.Free;
b.Free;
c.Free;
d.Free;
e.Free;
f.free;
w.Free;

end;

procedure TForm1.StringGrid1Click(Sender: TObject);
begin
s_true := false;
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
            StringGrid1.Selection := GridRect;
        GetParentForm(StringGrid1).SetFocus;
    StringGrid1.SetFocus;
    StringGrid1.EditorMode := true;
    TCustomEdit(StringGrid1.Components[0]).SelStart := i - 1;
    TCustomEdit(StringGrid1.Components[0]).SelLength := length(TargetText);

          end;
         //showmessage(inttostr(GridRect.Bottom));

        //s_true := true;
        goto TheEnd;
      end;
      inc(X);
    end;
    inc(Y);
    X := StringGrid1.FixedCols;

  end;
TheEnd:

end;

procedure TForm1.PayPal1Click(Sender: TObject);
begin
ShellExecute(Handle, nil, 'https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=WMTEWGNAQ4Y9J', nil, nil, SW_SHOW);
end;

procedure TForm1.Flattr1Click(Sender: TObject);
begin
ShellExecute(Handle, nil, 'https://flattr.com/thing/418687/DMcsvEditor', nil, nil, SW_SHOW);
end;

procedure TForm1.ranslators1Click(Sender: TObject);
var
f:tform;
m:tmemo;
begin

f:= tform.Create(self);

with f do begin
    caption := 'DMcsvEditor '+translators;
    position:= pomainformcenter;
    borderstyle:= bsdialog;
    width:= 300;
    height:= 250;


end;

m := tmemo.Create(self);

with m do begin
parent := f;
align := alclient;
scrollbars := ssvertical;
readonly := true;
color:= $00eeeeee;
lines.Add('Thanks to:');
lines.Add('');
lines.Add('Valerij Romanovskij - Russian');
lines.Add('Tilt. - Japanese');
lines.Add('Zdenìk Chalupský - Czech');
lines.Add('Almir Bispo - Portuguese');
lines.Add('Rafael Grau & Eduardo Ponce Garcia - Spanish');
lines.Add('Christoph Stephan - German');
lines.Add('Ronald Reerds - Dutch');
lines.Add('MuiCG - Thai');
lines.Add('Ake Engelbrektson - Swedish');
lines.Add('Milan Legény - Slovak');
lines.Add('Antoine Parra - French');
lines.Add('Grazia Messineo - Italiano');
lines.Add('jyokei - Simplified Chinese');
cursor := crarrow;

end;
HideCaret(m.Handle);
f.ShowModal;


m.Free;
f.Free;
end;
procedure TForm1.CheckUpdatetoStart1Click(Sender: TObject);
begin
if CheckUpdatetoStart1.Checked then begin
CheckUpdatetoStart1.Checked := false;
end else begin
CheckUpdatetoStart1.Checked := true;
end;
end;

initialization
decimalSeparator:=',';
end.
