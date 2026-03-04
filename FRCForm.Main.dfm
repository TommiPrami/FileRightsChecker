object FRCMainForm: TFRCMainForm
  Left = 0
  Top = 0
  Caption = 'FRCMainForm'
  ClientHeight = 685
  ClientWidth = 1031
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  TextHeight = 15
  object PanelButtons: TPanel
    Left = 819
    Top = 0
    Width = 212
    Height = 685
    Align = alRight
    BevelOuter = bvNone
    ShowCaption = False
    TabOrder = 0
    object ButtonRun: TButton
      AlignWithMargins = True
      Left = 3
      Top = 8
      Width = 206
      Height = 25
      Margins.Top = 8
      Action = ActionRun
      Align = alTop
      TabOrder = 0
    end
    object CheckBoxOpenFilesLongFileAndPathNameSupport: TCheckBox
      AlignWithMargins = True
      Left = 3
      Top = 46
      Width = 206
      Height = 19
      Margins.Top = 10
      Align = alTop
      Caption = 'Long filename support'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object CheckBoxCheckProcessBackupPrivileges: TCheckBox
      AlignWithMargins = True
      Left = 3
      Top = 71
      Width = 206
      Height = 19
      Align = alTop
      Caption = 'Check process backup privileges'
      TabOrder = 2
    end
  end
  object PanelLeft: TPanel
    Left = 0
    Top = 0
    Width = 819
    Height = 685
    Align = alClient
    BevelOuter = bvNone
    ShowCaption = False
    TabOrder = 1
    object PanelTop: TPanel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 813
      Height = 107
      Align = alTop
      BevelOuter = bvNone
      ShowCaption = False
      TabOrder = 0
      object LabelMustHaveWriteRights: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 63
        Width = 807
        Height = 15
        Margins.Top = 8
        Align = alTop
        Caption = 'Must have write rights'
      end
      object LabelMustHaveReadRights: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 8
        Width = 807
        Height = 15
        Margins.Top = 8
        Align = alTop
        Caption = 'Must have read rights'
      end
      object EditReadOnlyCheck: TEdit
        AlignWithMargins = True
        Left = 3
        Top = 29
        Width = 807
        Height = 23
        Align = alTop
        TabOrder = 0
      end
      object EditReadWrtiteChecks: TEdit
        AlignWithMargins = True
        Left = 3
        Top = 84
        Width = 807
        Height = 23
        Align = alTop
        TabOrder = 1
      end
    end
    object PanelLog: TPanel
      Left = 0
      Top = 113
      Width = 819
      Height = 572
      Align = alClient
      BevelOuter = bvNone
      ShowCaption = False
      TabOrder = 1
      object MemoLog: TMemo
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 813
        Height = 566
        Align = alClient
        TabOrder = 0
      end
    end
  end
  object ActionList: TActionList
    Left = 830
    Top = 235
    object ActionRun: TAction
      Category = 'File'
      Caption = 'Run'
      ShortCut = 16397
      OnExecute = ActionRunExecute
    end
  end
end
