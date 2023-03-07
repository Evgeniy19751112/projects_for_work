object frmServerDB: TfrmServerDB
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1044#1086#1089#1090#1091#1087' '#1082' '#1089#1077#1088#1074#1077#1088#1091' '#1040#1057#1055
  ClientHeight = 172
  ClientWidth = 594
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object grbParams: TGroupBox
    Left = 8
    Top = 8
    Width = 185
    Height = 129
    CustomHint = BalloonHint1
    Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1087#1086#1076#1082#1083#1102#1095#1077#1085#1080#1103':'
    TabOrder = 0
    object imgIndicator: TImage
      Left = 20
      Top = 100
      Width = 16
      Height = 16
      CustomHint = BalloonHint1
      Proportional = True
      Transparent = True
    end
    object txtSrv: TLabeledEdit
      Left = 54
      Top = 16
      Width = 121
      Height = 21
      CustomHint = BalloonHint1
      EditLabel.Width = 44
      EditLabel.Height = 13
      EditLabel.CustomHint = BalloonHint1
      EditLabel.Caption = #1057#1077#1088#1074#1077#1088': '
      LabelPosition = lpLeft
      TabOrder = 0
      OnChange = MozhnoZayti
    end
    object txtLogin: TLabeledEdit
      Left = 54
      Top = 43
      Width = 121
      Height = 21
      CustomHint = BalloonHint1
      EditLabel.Width = 37
      EditLabel.Height = 13
      EditLabel.CustomHint = BalloonHint1
      EditLabel.Caption = #1051#1086#1075#1080#1085': '
      LabelPosition = lpLeft
      TabOrder = 1
      OnChange = MozhnoZayti
    end
    object txtPass: TLabeledEdit
      Left = 54
      Top = 70
      Width = 121
      Height = 21
      CustomHint = BalloonHint1
      EditLabel.Width = 44
      EditLabel.Height = 13
      EditLabel.CustomHint = BalloonHint1
      EditLabel.Caption = #1055#1072#1088#1086#1083#1100': '
      LabelPosition = lpLeft
      PasswordChar = '*'
      TabOrder = 2
      OnChange = MozhnoZayti
    end
    object cmdConnect: TButton
      Left = 54
      Top = 96
      Width = 121
      Height = 25
      CustomHint = BalloonHint1
      Caption = #1055#1086#1076#1082#1083#1102#1095#1080#1090#1100#1089#1103
      TabOrder = 3
      OnClick = cmdConnectClick
    end
  end
  object grbDB: TGroupBox
    Left = 200
    Top = 8
    Width = 140
    Height = 160
    CustomHint = BalloonHint1
    Caption = #1041#1072#1079#1099' '#1083#1086#1103' '#1088#1072#1073#1086#1090#1099':'
    TabOrder = 1
    object lst_chk_Bases: TCheckListBox
      Left = 2
      Top = 15
      Width = 136
      Height = 143
      CustomHint = BalloonHint1
      Align = alClient
      ItemHeight = 13
      TabOrder = 0
      OnClick = lst_chk_BasesClick
    end
  end
  object grbReports: TGroupBox
    Left = 346
    Top = 8
    Width = 242
    Height = 129
    CustomHint = BalloonHint1
    Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1092#1086#1088#1084#1080#1088#1086#1074#1072#1085#1080#1103' '#1086#1090#1095#1105#1090#1072':'
    TabOrder = 2
    DesignSize = (
      242
      129)
    object lblReportBaseInfo: TLabel
      Left = 4
      Top = 102
      Width = 234
      Height = 20
      CustomHint = BalloonHint1
      Alignment = taCenter
      Anchors = [akLeft, akTop, akRight]
      AutoSize = False
      Caption = 'lblReportBaseInfo'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ExplicitWidth = 196
    end
    object txtPattern: TLabeledEdit
      Left = 4
      Top = 35
      Width = 234
      Height = 21
      CustomHint = BalloonHint1
      Anchors = [akLeft, akTop, akRight]
      EditLabel.Width = 115
      EditLabel.Height = 13
      EditLabel.CustomHint = BalloonHint1
      EditLabel.Caption = #1064#1072#1073#1083#1086#1085' '#1080#1084#1077#1085#1080' '#1092#1072#1081#1083#1072': '
      TabOrder = 0
      OnChange = txtPatternChange
    end
    object txtExample: TLabeledEdit
      Left = 4
      Top = 76
      Width = 234
      Height = 21
      CustomHint = BalloonHint1
      Anchors = [akLeft, akTop, akRight]
      EditLabel.Width = 112
      EditLabel.Height = 13
      EditLabel.CustomHint = BalloonHint1
      EditLabel.Caption = #1055#1088#1080#1084#1077#1088' '#1080#1084#1077#1085#1080' '#1092#1072#1081#1083#1072': '
      ReadOnly = True
      TabOrder = 1
    end
  end
  object cmdSelectSQL: TBitBtn
    Left = 8
    Top = 143
    Width = 186
    Height = 25
    Hint = 
      #1042#1099#1073#1088#1072#1090#1100' '#1080' '#1087#1088#1086#1090#1077#1089#1090#1080#1088#1086#1074#1072#1090#1100' '#1088#1072#1073#1086#1090#1091' '#1089#1082#1088#1080#1087#1090#1072' SQL. '#1055#1072#1087#1082#1072' '#1089' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1086#1084 +
      ' '#1073#1091#1076#1077#1090' '#1086#1090#1082#1088#1099#1090#1072' '#1072#1074#1090#1086#1084#1072#1090#1080#1095#1077#1089#1082#1080
    CustomHint = BalloonHint1
    Caption = #1058#1077#1089#1090#1086#1074#1099#1081' '#1079#1072#1087#1088#1086#1089
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00500000000000
      055557777777777775F508888888888880557F5FFFFFFFFFF75F080000000000
      88057577777777775F755080FFFFFF05088057F7FFFFFF7575F70000000000F0
      F08077777777775757F70FFFFFFFFF0F008075F5FF5FF57577F750F00F00FFF0
      F08057F775775557F7F750FFFFFFFFF0F08057FF5555555757F7000FFFFFFFFF
      0000777FF5FFFFF577770900F00000F000907F775777775777F7090FFFFFFFFF
      00907F7F555555557757000FFFFFFFFF0F00777F5FFF5FF57F77550F000F00FF
      0F05557F777577557F7F550FFFFFFFFF0005557F555FFFFF7775550FFF000000
      05555575FF777777755555500055555555555557775555555555}
    NumGlyphs = 2
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    OnClick = cmdSelectSQLClick
  end
  object BalloonHint1: TBalloonHint
    Left = 240
    Top = 48
  end
  object ImagesIndicator: TImageList
    Left = 8
    Top = 112
    Bitmap = {
      494C010108001800480010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000003000000001002000000000000030
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000FFFFFF00808080008080800080808000808080008080800000000000FFFF
      FF00FFFFFF000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000FFFFFF00808080008080800080808000808080008080800000000000FFFF
      FF00FFFFFF000000000000000000000000000000000000000000000000000000
      000000000000C0C0C000C0C0C000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      8000808080000000000000000000FFFFFF00FFFFFF00FFFFFF00808080008080
      800000000000FFFFFF0000000000000000000000000000000000000000000000
      000000000000C0C0C000C0C0C000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      8000808080000000000000000000FFFFFF00FFFFFF00FFFFFF00808080008080
      800000000000FFFFFF000000000000000000000000000000000000000000C0C0
      C000C0C0C0008080800000000000000000000000000080808000808080008080
      8000000000000000000000000000000000000000000000000000808080000000
      000000000000FFFFFF00808080008080800080808000FFFFFF00FFFFFF00FFFF
      FF008080800000000000FFFFFF0000000000000000000000000000000000C0C0
      C000C0C0C0008080800000000000000000000000000080808000808080008080
      8000000000000000000000000000000000000000000000000000808080000000
      000000000000FFFFFF00808080008080800080808000FFFFFF00FFFFFF00FFFF
      FF008080800000000000FFFFFF00000000000000000000000000C0C0C000C0C0
      C000000000000000000000000000000000000000000000000000000000008080
      8000808080000000000000000000000000000000000080808000FFFFFF000000
      000080808000808080008080800080808000808080008080800080808000FFFF
      FF00FFFFFF0080808000FFFFFF00000000000000000000000000C0C0C000C0C0
      C00000000000FFFFFF0000FF0000FFFFFF0000FF0000FFFFFF00000000008080
      8000808080000000000000000000000000000000000080808000FFFFFF000000
      000080808000808080000000000000000000000000008080800080808000FFFF
      FF00FFFFFF0080808000FFFFFF00000000000000000000000000C0C0C0000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000808080000000000000000000000000000000000080808000000000008080
      8000808080008080800080808000808080008080800080808000808080008080
      8000FFFFFF008080800000000000FFFFFF000000000000000000C0C0C0000000
      0000FFFFFF000000000000000000000000000000000000000000FFFFFF000000
      0000808080000000000000000000000000000000000080808000000000008080
      8000000000000000000080808000808080008080800000000000000000008080
      8000FFFFFF008080800000000000FFFFFF0000000000C0C0C000808080000000
      0000000000000000000000800000008000000080000000000000000000000000
      00008080800080808000000000000000000080808000FFFFFF00000000008080
      8000808080008080800000000000000000000000000080808000808080008080
      8000FFFFFF00FFFFFF0080808000FFFFFF0000000000C0C0C00080808000FFFF
      FF00000000000000000000FF000000800000008000000000000000000000FFFF
      FF008080800080808000000000000000000080808000FFFFFF00000000008080
      8000000000008080800000000000000000000000000080808000000000008080
      8000FFFFFF00FFFFFF0080808000FFFFFF0000000000C0C0C000000000000000
      0000000000000080000000800000008000000080000000800000000000000000
      00000000000080808000000000000000000080808000FFFFFF00808080008080
      800080808000FFFFFF0000000000000000000000000000000000808080008080
      800080808000FFFFFF0080808000FFFFFF0000000000C0C0C0000000000000FF
      00000000000000FF00000080000000FF000000800000008000000000000000FF
      00000000000080808000000000000000000080808000FFFFFF00808080000000
      000080808000FFFFFF0000000000000000000000000000000000808080000000
      000080808000FFFFFF0080808000FFFFFF0000000000FFFFFF00000000000000
      0000000000000080000000800000008000000080000000800000000000000000
      00000000000080808000000000000000000080808000FFFFFF00808080008080
      800080808000FFFFFF0000000000000000000000000000000000808080008080
      800080808000FFFFFF0080808000FFFFFF0000000000FFFFFF0000000000FFFF
      FF000000000000FF000000FF000000FF000000FF00000080000000000000FFFF
      FF000000000080808000000000000000000080808000FFFFFF00808080000000
      000080808000FFFFFF0000000000000000000000000000000000808080000000
      000080808000FFFFFF0080808000FFFFFF0000000000FFFFFF00000000000000
      000000000000C0C0C00000800000008000000080000000800000000000000000
      000000000000C0C0C000000000000000000080808000FFFFFF00808080008080
      800080808000FFFFFF00FFFFFF00000000000000000000000000808080008080
      8000808080000000000080808000FFFFFF0000000000FFFFFF000000000000FF
      000000000000FFFFFF0000FF000000FF00000080000000FF00000000000000FF
      000000000000C0C0C000000000000000000080808000FFFFFF00808080000000
      000080808000FFFFFF00FFFFFF00000000000000000000000000808080000000
      0000808080000000000080808000FFFFFF0000000000FFFFFF00808080000000
      00000000000000000000C0C0C000C0C0C0000080000000000000000000000000
      000080808000C0C0C00000000000000000008080800000000000FFFFFF008080
      80008080800080808000FFFFFF00FFFFFF00FFFFFF0080808000808080008080
      8000FFFFFF0000000000808080000000000000000000FFFFFF0080808000FFFF
      FF000000000000000000FFFFFF00FFFFFF0000FF00000000000000000000FFFF
      FF0080808000C0C0C00000000000000000008080800000000000FFFFFF008080
      80000000000080808000FFFFFF00FFFFFF00FFFFFF0080808000000000008080
      8000FFFFFF000000000080808000000000000000000000000000C0C0C0000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000C0C0C0000000000000000000000000000000000080808000FFFFFF008080
      8000808080008080800080808000808080008080800080808000808080008080
      80000000000080808000FFFFFF00000000000000000000000000C0C0C0000000
      0000FFFFFF000000000000000000000000000000000000000000FFFFFF000000
      0000C0C0C0000000000000000000000000000000000080808000FFFFFF008080
      8000000000000000000080808000808080008080800000000000000000008080
      80000000000080808000FFFFFF00000000000000000000000000FFFFFF00C0C0
      C00000000000000000000000000000000000000000000000000000000000C0C0
      C000C0C0C000000000000000000000000000000000008080800000000000FFFF
      FF00808080008080800080808000808080008080800080808000808080000000
      0000000000008080800000000000000000000000000000000000FFFFFF00C0C0
      C00000000000FFFFFF0000FF0000FFFFFF0000FF0000FFFFFF0000000000C0C0
      C000C0C0C000000000000000000000000000000000008080800000000000FFFF
      FF00808080008080800000000000000000000000000080808000808080000000
      000000000000808080000000000000000000000000000000000000000000FFFF
      FF00C0C0C0008080800000000000000000000000000080808000C0C0C000C0C0
      C000000000000000000000000000000000000000000000000000808080000000
      0000FFFFFF00FFFFFF008080800080808000808080000000000000000000FFFF
      FF0080808000000000000000000000000000000000000000000000000000FFFF
      FF00C0C0C0008080800000000000000000000000000080808000C0C0C000C0C0
      C000000000000000000000000000000000000000000000000000808080000000
      0000FFFFFF00FFFFFF008080800080808000808080000000000000000000FFFF
      FF00808080000000000000000000000000000000000000000000000000000000
      000000000000FFFFFF00FFFFFF00FFFFFF00C0C0C000C0C0C000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      80008080800000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00808080008080
      8000000000000000000000000000000000000000000000000000000000000000
      000000000000FFFFFF00FFFFFF00FFFFFF00C0C0C000C0C0C000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      80008080800000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00808080008080
      8000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000008080800080808000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000008080800080808000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000FFFFFF00808080008080800080808000808080008080800000000000FFFF
      FF00FFFFFF000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000FFFFFF00808080008080800080808000808080008080800000000000FFFF
      FF00FFFFFF000000000000000000000000000000000000000000000000000000
      000000000000C0C0C000C0C0C000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      8000808080000000000000000000FFFFFF00FFFFFF00FFFFFF00808080008080
      800000000000FFFFFF0000000000000000000000000000000000000000000000
      000000000000C0C0C000C0C0C000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      8000808080000000000000000000FFFFFF00FFFFFF00FFFFFF00808080008080
      800000000000FFFFFF000000000000000000000000000000000000000000C0C0
      C000C0C0C0008080800000000000000000000000000080808000808080008080
      8000000000000000000000000000000000000000000000000000808080000000
      000000000000FFFFFF00808080008080800080808000FFFFFF00FFFFFF00FFFF
      FF008080800000000000FFFFFF0000000000000000000000000000000000C0C0
      C000C0C0C0008080800000000000000000000000000080808000808080008080
      8000000000000000000000000000000000000000000000000000808080000000
      000000000000FFFFFF00808080008080800080808000FFFFFF00FFFFFF00FFFF
      FF008080800000000000FFFFFF00000000000000000000000000C0C0C000C0C0
      C000000000000000000000000000000000000000000000000000000000008080
      8000808080000000000000000000000000000000000080808000FFFFFF000000
      000080808000808080008080800080808000808080008080800080808000FFFF
      FF00FFFFFF0080808000FFFFFF00000000000000000000000000C0C0C000C0C0
      C00000000000FFFFFF000000FF00FFFFFF000000FF00FFFFFF00000000008080
      8000808080000000000000000000000000000000000080808000FFFFFF000000
      000080808000808080000000000000000000000000008080800080808000FFFF
      FF00FFFFFF0080808000FFFFFF00000000000000000000000000C0C0C0000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000808080000000000000000000000000000000000080808000000000008080
      8000808080008080800080808000808080008080800080808000808080008080
      8000FFFFFF008080800000000000FFFFFF000000000000000000C0C0C0000000
      0000FFFFFF000000000000000000000000000000000000000000FFFFFF000000
      0000808080000000000000000000000000000000000080808000000000008080
      8000000000000000000080808000808080008080800000000000000000008080
      8000FFFFFF008080800000000000FFFFFF0000000000C0C0C000808080000000
      0000000000000000000000008000000080000000800000000000000000000000
      00008080800080808000000000000000000080808000FFFFFF00000000008080
      8000808080008080800000000000000000000000000080808000808080008080
      8000FFFFFF00FFFFFF0080808000FFFFFF0000000000C0C0C00080808000FFFF
      FF0000000000000000000000FF0000008000000080000000000000000000FFFF
      FF008080800080808000000000000000000080808000FFFFFF00000000008080
      8000000000008080800000000000000000000000000080808000000000008080
      8000FFFFFF00FFFFFF0080808000FFFFFF0000000000C0C0C000000000000000
      0000000000000000800000008000000080000000800000008000000000000000
      00000000000080808000000000000000000080808000FFFFFF00808080008080
      800080808000FFFFFF0000000000000000000000000000000000808080008080
      800080808000FFFFFF0080808000FFFFFF0000000000C0C0C000000000000000
      FF00000000000000FF00000080000000FF000000800000008000000000000000
      FF000000000080808000000000000000000080808000FFFFFF00808080000000
      000080808000FFFFFF0000000000000000000000000000000000808080000000
      000080808000FFFFFF0080808000FFFFFF0000000000FFFFFF00000000000000
      0000000000000000800000008000000080000000800000008000000000000000
      00000000000080808000000000000000000080808000FFFFFF00808080008080
      800080808000FFFFFF0000000000000000000000000000000000808080008080
      800080808000FFFFFF0080808000FFFFFF0000000000FFFFFF0000000000FFFF
      FF00000000000000FF000000FF000000FF000000FF000000800000000000FFFF
      FF000000000080808000000000000000000080808000FFFFFF00808080000000
      000080808000FFFFFF0000000000000000000000000000000000808080000000
      000080808000FFFFFF0080808000FFFFFF0000000000FFFFFF00000000000000
      000000000000C0C0C00000008000000080000000800000008000000000000000
      000000000000C0C0C000000000000000000080808000FFFFFF00808080008080
      800080808000FFFFFF00FFFFFF00000000000000000000000000808080008080
      8000808080000000000080808000FFFFFF0000000000FFFFFF00000000000000
      FF0000000000FFFFFF000000FF000000FF00000080000000FF00000000000000
      FF0000000000C0C0C000000000000000000080808000FFFFFF00808080000000
      000080808000FFFFFF00FFFFFF00000000000000000000000000808080000000
      0000808080000000000080808000FFFFFF0000000000FFFFFF00808080000000
      00000000000000000000C0C0C000C0C0C0000000800000000000000000000000
      000080808000C0C0C00000000000000000008080800000000000FFFFFF008080
      80008080800080808000FFFFFF00FFFFFF00FFFFFF0080808000808080008080
      8000FFFFFF0000000000808080000000000000000000FFFFFF0080808000FFFF
      FF000000000000000000FFFFFF00FFFFFF000000FF000000000000000000FFFF
      FF0080808000C0C0C00000000000000000008080800000000000FFFFFF008080
      80000000000080808000FFFFFF00FFFFFF00FFFFFF0080808000000000008080
      8000FFFFFF000000000080808000000000000000000000000000C0C0C0000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000C0C0C0000000000000000000000000000000000080808000FFFFFF008080
      8000808080008080800080808000808080008080800080808000808080008080
      80000000000080808000FFFFFF00000000000000000000000000C0C0C0000000
      0000FFFFFF000000000000000000000000000000000000000000FFFFFF000000
      0000C0C0C0000000000000000000000000000000000080808000FFFFFF008080
      8000000000000000000080808000808080008080800000000000000000008080
      80000000000080808000FFFFFF00000000000000000000000000FFFFFF00C0C0
      C00000000000000000000000000000000000000000000000000000000000C0C0
      C000C0C0C000000000000000000000000000000000008080800000000000FFFF
      FF00808080008080800080808000808080008080800080808000808080000000
      0000000000008080800000000000000000000000000000000000FFFFFF00C0C0
      C00000000000FFFFFF000000FF00FFFFFF000000FF00FFFFFF0000000000C0C0
      C000C0C0C000000000000000000000000000000000008080800000000000FFFF
      FF00808080008080800000000000000000000000000080808000808080000000
      000000000000808080000000000000000000000000000000000000000000FFFF
      FF00C0C0C0008080800000000000000000000000000080808000C0C0C000C0C0
      C000000000000000000000000000000000000000000000000000808080000000
      0000FFFFFF00FFFFFF008080800080808000808080000000000000000000FFFF
      FF0080808000000000000000000000000000000000000000000000000000FFFF
      FF00C0C0C0008080800000000000000000000000000080808000C0C0C000C0C0
      C000000000000000000000000000000000000000000000000000808080000000
      0000FFFFFF00FFFFFF008080800080808000808080000000000000000000FFFF
      FF00808080000000000000000000000000000000000000000000000000000000
      000000000000FFFFFF00FFFFFF00FFFFFF00C0C0C000C0C0C000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      80008080800000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00808080008080
      8000000000000000000000000000000000000000000000000000000000000000
      000000000000FFFFFF00FFFFFF00FFFFFF00C0C0C000C0C0C000000000000000
      0000000000000000000000000000000000000000000000000000000000008080
      80008080800000000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00808080008080
      8000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000008080800080808000808080008080800080808000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000008080800080808000808080008080800080808000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000300000000100010000000000800100000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000FFFFFC1FFFFFFC1FF83FF027F83FF027
      E00FE60BE00FE60BC007D805C007D80580039001800393818003A0028003AC62
      0001238000012BA0000103C0000113D0000103C0000113D0000101C4000111D4
      00014005000148258003800980038C698003A01B8003A39BC007D067C007D067
      E00FE40FE00FE40FF83FF83FF83FF83FFFFFFC1FFFFFFC1FF83FF027F83FF027
      E00FE60BE00FE60BC007D805C007D80580039001800393818003A0028003AC62
      0001238000012BA0000103C0000113D0000103C0000113D0000101C4000111D4
      00014005000148258003800980038C698003A01B8003A39BC007D067C007D067
      E00FE40FE00FE40FF83FF83FF83FF83F00000000000000000000000000000000
      000000000000}
  end
end
