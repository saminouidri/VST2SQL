object VST2SQL: TVST2SQL
  Left = 0
  Top = 0
  Caption = 'VST2SQL [Disconnected]'
  ClientHeight = 242
  ClientWidth = 259
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnClose = OnClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 38
    Width = 31
    Height = 13
    BiDiMode = bdLeftToRight
    Caption = 'Page :'
    ParentBiDiMode = False
  end
  object Label2: TLabel
    Left = 8
    Top = 57
    Width = 246
    Height = 13
    Caption = '_________________________________________'
  end
  object Label3: TLabel
    Left = 8
    Top = 224
    Width = 12
    Height = 13
    Caption = '-/-'
  end
  object VisioBox: TComboBox
    Left = 8
    Top = 8
    Width = 243
    Height = 21
    TabOrder = 0
    Text = 'Select Visio Doc...'
  end
  object OnlyBox: TCheckBox
    Left = 123
    Top = 35
    Width = 128
    Height = 17
    Caption = 'Page ShapeData Only'
    TabOrder = 1
  end
  object PageNo: TEdit
    Left = 45
    Top = 35
    Width = 72
    Height = 21
    NumbersOnly = True
    TabOrder = 2
    Text = '1'
  end
  object FillFormat: TCheckBox
    Left = 123
    Top = 78
    Width = 97
    Height = 17
    Caption = 'FillFormat'
    TabOrder = 3
  end
  object LineFormat: TCheckBox
    Left = 8
    Top = 101
    Width = 97
    Height = 17
    Caption = 'LineFormat'
    TabOrder = 4
  end
  object Character: TCheckBox
    Left = 8
    Top = 78
    Width = 97
    Height = 17
    Caption = 'Character'
    TabOrder = 5
  end
  object Scratch: TCheckBox
    Left = 123
    Top = 101
    Width = 97
    Height = 17
    Caption = 'Scratch'
    TabOrder = 6
  end
  object ShapeTransform: TCheckBox
    Left = 8
    Top = 124
    Width = 97
    Height = 17
    Caption = 'ShapeTransform'
    TabOrder = 7
  end
  object Textfield: TCheckBox
    Left = 123
    Top = 124
    Width = 97
    Height = 17
    Caption = 'Textfield'
    TabOrder = 8
  end
  object UserDefined: TCheckBox
    Left = 8
    Top = 147
    Width = 97
    Height = 17
    Caption = 'UserDefined'
    TabOrder = 9
  end
  object btnCommit: TButton
    Left = 8
    Top = 170
    Width = 243
    Height = 25
    Caption = 'Commit'
    TabOrder = 10
    OnClick = btnCommitClick
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 201
    Width = 243
    Height = 17
    Step = 1
    TabOrder = 11
  end
  object FDConnection1: TFDConnection
    Params.Strings = (
      'Database=C:\Users\Public\Documents\mydb.db'
      'DriverID=SQLite')
    Connected = True
    Left = 136
    Top = 224
  end
  object FDQuery1: TFDQuery
    Connection = FDConnection1
    Left = 56
    Top = 224
  end
  object DataSource1: TDataSource
    DataSet = FDQuery1
    Left = 96
    Top = 224
  end
  object MainMenu1: TMainMenu
    Left = 176
    Top = 224
    object File1: TMenuItem
      Caption = 'File'
      object OpenDatabase: TMenuItem
        Caption = 'Open Database'
        OnClick = OpenDatabaseClick
      end
      object CloseDatabase1: TMenuItem
        Caption = 'Close Database'
        OnClick = CloseDatabase1Click
      end
      object Exit: TMenuItem
        Caption = 'Exit'
        OnClick = ExitClick
      end
    end
    object About: TMenuItem
      Caption = 'About'
      OnClick = AboutClick
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 216
    Top = 224
  end
end
