object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'EcoDTO Generator'
  ClientHeight = 233
  ClientWidth = 310
  Color = clWindow
  Ctl3D = False
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GridPanel1: TGridPanel
    Left = 0
    Top = 0
    Width = 310
    Height = 233
    Align = alClient
    ColumnCollection = <
      item
        Value = 100.000000000000000000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = pnlSuperior
        Row = 0
      end
      item
        Column = 0
        Control = pnlInferior
        Row = 1
      end>
    Ctl3D = False
    ParentCtl3D = False
    RowCollection = <
      item
        Value = 85.397096498719040000
      end
      item
        Value = 14.602903501280960000
      end>
    TabOrder = 0
    object pnlSuperior: TPanel
      Left = 1
      Top = 1
      Width = 308
      Height = 197
      Align = alClient
      Ctl3D = False
      ParentCtl3D = False
      TabOrder = 0
      DesignSize = (
        308
        197)
      object lblDestino: TLabel
        Left = 8
        Top = 138
        Width = 36
        Height = 13
        Caption = 'Destino'
      end
      object lblPassword: TLabel
        Left = 8
        Top = 88
        Width = 30
        Height = 13
        Caption = 'Senha'
      end
      object lblUser: TLabel
        Left = 8
        Top = 63
        Width = 36
        Height = 13
        Caption = 'Usuario'
      end
      object lbl2: TLabel
        Left = 8
        Top = 38
        Width = 20
        Height = 13
        Caption = 'Tipo'
      end
      object lbl1: TLabel
        Left = 8
        Top = 13
        Width = 34
        Height = 13
        Caption = 'Origem'
      end
      object Label1: TLabel
        Left = 8
        Top = 113
        Width = 32
        Height = 13
        Caption = 'Tabela'
      end
      object Label2: TLabel
        Left = 8
        Top = 163
        Width = 3
        Height = 13
      end
      object SbDestino: TSearchBox
        Left = 55
        Top = 135
        Width = 238
        Height = 19
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 5
        OnInvokeSearch = SbDestinoInvokeSearch
      end
      object EdtPassword: TEdit
        Left = 55
        Top = 85
        Width = 157
        Height = 19
        PasswordChar = '*'
        TabOrder = 3
        OnExit = EdtPasswordExit
      end
      object EdtUser: TEdit
        Left = 55
        Top = 60
        Width = 157
        Height = 19
        TabOrder = 2
      end
      object Tipo: TComboBox
        Left = 55
        Top = 35
        Width = 238
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 1
      end
      object SbOrigem: TSearchBox
        Left = 55
        Top = 10
        Width = 238
        Height = 19
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 0
        OnInvokeSearch = SbOrigemInvokeSearch
      end
      object CBTabelas: TComboBox
        Left = 55
        Top = 110
        Width = 238
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        CharCase = ecUpperCase
        TabOrder = 4
      end
      object DBGerarDAO: TCheckBox
        Left = 55
        Top = 165
        Width = 130
        Height = 17
        Caption = 'Gerar objetos DAO?'
        Checked = True
        State = cbChecked
        TabOrder = 6
      end
    end
    object pnlInferior: TPanel
      Left = 1
      Top = 198
      Width = 308
      Height = 34
      Align = alClient
      Ctl3D = False
      ParentCtl3D = False
      TabOrder = 1
      object GridPanelBotoes: TGridPanel
        Left = 1
        Top = 1
        Width = 306
        Height = 32
        Margins.Left = 0
        Margins.Top = 0
        Margins.Right = 0
        Margins.Bottom = 0
        Align = alClient
        ColumnCollection = <
          item
            Value = 50.000000000000000000
          end
          item
            Value = 50.000000000000000000
          end>
        ControlCollection = <
          item
            Column = 0
            Control = BtnCancelar
            Row = 0
          end
          item
            Column = 1
            Control = BtnGerar
            Row = 0
          end>
        Ctl3D = False
        ParentCtl3D = False
        RowCollection = <
          item
            Value = 100.000000000000000000
          end>
        TabOrder = 0
        object BtnCancelar: TBitBtn
          Left = 1
          Top = 1
          Width = 152
          Height = 30
          Margins.Left = 0
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'Cancelar'
          TabOrder = 0
          OnClick = BtnCancelarClick
        end
        object BtnGerar: TBitBtn
          Left = 153
          Top = 1
          Width = 152
          Height = 30
          Margins.Left = 0
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'Gerar'
          TabOrder = 1
          OnClick = BtnGerarClick
        end
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 32
    Top = 256
  end
  object FDConnection: TFDConnection
    Left = 88
    Top = 256
  end
end
