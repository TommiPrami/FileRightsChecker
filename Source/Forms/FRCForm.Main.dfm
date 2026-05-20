object FRCMainForm: TFRCMainForm
  Left = 0
  Top = 0
  Caption = 'Fire Rights Checker'
  ClientHeight = 685
  ClientWidth = 1210
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  TextHeight = 15
  object PanelButtons: TPanel
    Left = 864
    Top = 0
    Width = 346
    Height = 685
    Align = alRight
    BevelOuter = bvNone
    ShowCaption = False
    TabOrder = 0
    object ButtonRun: TButton
      AlignWithMargins = True
      Left = 3
      Top = 8
      Width = 340
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
      Width = 340
      Height = 19
      Margins.Top = 10
      Align = alTop
      Caption = 'Long filename support'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object CheckBoxProcessBackupPrivileges: TCheckBox
      AlignWithMargins = True
      Left = 3
      Top = 71
      Width = 340
      Height = 19
      Align = alTop
      Caption = 'Check process backup privileges'
      TabOrder = 2
    end
    object CheckBoxRunFileGetEffectiveRightsShortfallTests: TCheckBox
      AlignWithMargins = True
      Left = 3
      Top = 121
      Width = 340
      Height = 19
      Align = alTop
      Caption = 'Run file GetEffectiveRightsShortfall tests (Slow)'
      TabOrder = 3
    end
    object CheckBoxRunDirectoryGetEffectiveRightsShortfallTests: TCheckBox
      AlignWithMargins = True
      Left = 3
      Top = 96
      Width = 340
      Height = 19
      Align = alTop
      Caption = 'Run directory GetEffectiveRightsShortfall tests (Slow)'
      TabOrder = 4
    end
    object CheckBoxRunCurrentUserIsOwnerTests: TCheckBox
      AlignWithMargins = True
      Left = 3
      Top = 146
      Width = 340
      Height = 19
      Align = alTop
      Caption = 'Run CurrentUserIsOwner tests (are basically false positives)'
      TabOrder = 5
    end
  end
  object PanelLeft: TPanel
    Left = 0
    Top = 0
    Width = 864
    Height = 685
    Align = alClient
    BevelOuter = bvNone
    ShowCaption = False
    TabOrder = 1
    object PanelTop: TPanel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 858
      Height = 107
      Align = alTop
      BevelOuter = bvNone
      ShowCaption = False
      TabOrder = 0
      object LabelMustHaveWriteRights: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 63
        Width = 852
        Height = 15
        Margins.Top = 8
        Align = alTop
        Caption = 'Must have write rights'
      end
      object LabelMustHaveReadRights: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 8
        Width = 852
        Height = 15
        Margins.Top = 8
        Align = alTop
        Caption = 'Must have read rights'
      end
      object EditReadOnlyCheck: TEdit
        AlignWithMargins = True
        Left = 3
        Top = 29
        Width = 852
        Height = 23
        Align = alTop
        TabOrder = 0
      end
      object EditReadWriteChecks: TEdit
        AlignWithMargins = True
        Left = 3
        Top = 84
        Width = 852
        Height = 23
        Align = alTop
        TabOrder = 1
      end
    end
    object PanelLog: TPanel
      Left = 0
      Top = 113
      Width = 864
      Height = 572
      Align = alClient
      BevelOuter = bvNone
      ShowCaption = False
      TabOrder = 1
      object MemoLog: TMemo
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 858
        Height = 566
        Align = alClient
        ScrollBars = ssBoth
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
