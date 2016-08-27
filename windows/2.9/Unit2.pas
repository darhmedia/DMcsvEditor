{*********************************
  Name    : DMcsvEditor
  Version : 2.2
  Author  : Darh Media - Tivadar
  Date    : 2009
  Licence : GPL v2.0
  Contact : info@darhmedia.hu
**********************************}
unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm2 = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
implementation

{$R *.dfm}

procedure TForm2.Button2Click(Sender: TObject);
begin
close;
end;

procedure TForm2.Button1Click(Sender: TObject);
begin
memo1.Lines.SaveToFile(extractfilepath(application.ExeName)+'search.egs');
close;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
if fileexists(extractfilepath(application.ExeName)+'search.egs') then
  form2.Memo1.Lines.LoadFromFile(extractfilepath(application.ExeName)+'search.egs');
end;

end.
