object Form2: TForm2
  Left = 524
  Top = 223
  BorderStyle = bsDialog
  Caption = 'DMcsvEditor - Search Engines'
  ClientHeight = 227
  ClientWidth = 304
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 102
    Height = 13
    Caption = 'Format: Name|Url'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Memo1: TMemo
    Left = 0
    Top = 32
    Width = 304
    Height = 153
    Align = alCustom
    Lines.Strings = (
      'Google|http://www.google.hu/search?q='
      ''
      'Bing|http://www.bing.com/search?q='
      'Yahoo|search.yahoo.com/search?p=')
    TabOrder = 0
  end
  object Button1: TButton
    Left = 48
    Top = 192
    Width = 89
    Height = 25
    Caption = 'Save'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 160
    Top = 192
    Width = 97
    Height = 25
    Caption = 'Cancel'
    TabOrder = 2
    OnClick = Button2Click
  end
end
