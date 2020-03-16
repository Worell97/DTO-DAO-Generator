unit DTOGenerator;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.FB, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.StdCtrls,
  Vcl.Buttons,
  {$WARN UNIT_PLATFORM OFF}
  VCL.FileCtrl,
  {$WARN UNIT_PLATFORM ON}
  Vcl.WinXCtrls, Vcl.ExtCtrls, FireDAC.Phys.Intf, Vcl.DBCtrls;

type
  EGerarFields = class(System.SysUtils.Exception);
  EGerarDelarationsProcedures = class(System.SysUtils.Exception);
  EGerarDelarationsFunctions = class(System.SysUtils.Exception);
  EGerarProperty = class(System.SysUtils.Exception);
  EGerarCabecalho = class(System.SysUtils.Exception);
  EGerarPublicDeclarations = class(System.SysUtils.Exception);
  EGerarClassItem = class(System.SysUtils.Exception);
  EGerarClassLista = class(System.SysUtils.Exception);
  EGerarImplementation = class(System.SysUtils.Exception);
  EGerarFunctions = class(System.SysUtils.Exception);   
  EGerarProcedures = class(System.SysUtils.Exception);

  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    FDConnection: TFDConnection;
    GridPanel1: TGridPanel;
    pnlSuperior: TPanel;
    pnlInferior: TPanel;
    lblDestino: TLabel;
    SbDestino: TSearchBox;
    EdtPassword: TEdit;
    lblPassword: TLabel;
    lblUser: TLabel;
    EdtUser: TEdit;
    Tipo: TComboBox;
    lbl2: TLabel;
    lbl1: TLabel;
    SbOrigem: TSearchBox;
    GridPanelBotoes: TGridPanel;
    BtnCancelar: TBitBtn;
    BtnGerar: TBitBtn;
    CBTabelas: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    DBGerarDAO: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
    procedure SbOrigemInvokeSearch(Sender: TObject);
    procedure SbDestinoInvokeSearch(Sender: TObject);
    procedure BtnCancelarClick(Sender: TObject);
    procedure EdtPasswordExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  strict private
    const
      SelectTabelas =
        'select rdb$relation_name as Tabela from rdb$relation_fields ';

      wherePadrao =
        'where rdb$relation_name not like ''%$%'' group by rdb$relation_name';

      whereCustom =
        'where rdb$relation_name = :Tabela group by rdb$relation_name';

      SelectColunas =
        'select R.rdb$field_name as Coluna,    					  	'+
        '	case F.RDB$FIELD_TYPE             						    '+
        '     when 7 then ''System.Integer''      					'+
        '     when 8 then ''System.Integer''       					'+
        '     when 10 then ''System.Double''        				'+
        '     when 12 then ''System.TDate''         				'+
        '     when 13 then ''System.TTime''         				'+
        '     when 14 then ''System.UnicodeString''       	'+
        '     when 16 then ''System.Int64''        					'+
        '     when 27 then ''System.Double''       					'+
        '     when 35 then ''System.TDateTime''    					'+
        '     when 37 then ''System.UnicodeString''					'+
        '     when 261 then ''TBlobField''       					  '+
        '     else ''UNKNOWN''             							    '+
        '   end as Tipo,	            							        '+
        '	case F.RDB$FIELD_TYPE             						    '+
        '     when 7 then ''Integer''      							    '+
        '     when 8 then ''Integer''       						    '+
        '     when 10 then ''Currency''        						  '+
        '     when 12 then ''Date''         						    '+
        '     when 13 then ''Time''         						    '+
        '     when 14 then ''String''         						  '+
        '     when 16 then ''Int64''        						    '+
        '     when 27 then ''Currency''       						  '+
        '     when 35 then ''DateTime''    						    '+
        '     when 37 then ''String''								        '+
        '     when 261 then ''TBlobField''       					  '+
        '     else ''UNKNOWN''             							    '+
        '   end as TipoDAO		            						      '+
        'from rdb$relation_fields R									        '+
        '	left join RDB$FIELDS F   								          '+
        '		on R.RDB$FIELD_SOURCE = F.RDB$FIELD_NAME 			  '+
        '	left join RDB$COLLATIONS COLL 							      '+
        '	  on F.RDB$COLLATION_ID = COLL.RDB$COLLATION_ID 	'+
        '	left join RDB$CHARACTER_SETS CSET 						    '+
        '	  on F.RDB$CHARACTER_SET_ID = CSET.RDB$CHARACTER_SET_ID	'+
        'where rdb$relation_name= :Tabela 							';
  private
    { Private declarations }
    NomeTabela : System.UnicodeString;
    function Gerar() : Boolean;

    procedure GerarCabecalho(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarClassItem(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarFields(AArquivo : TStringList; AQueryColunas : TFDQuery);
    procedure GerarDeclaracoesProcedures(AArquivo : TStringList; AQueryColunas : TFDQuery);
    procedure GerarDeclaracoesFunctions(AArquivo : TStringList; AQueryColunas : TFDQuery);
    procedure GerarPublicDeclarations(AArquivo : TStringList; AQueryTabelas : TFDQuery);
    procedure GerarPropertys(AArquivo : TStringList; AQueryColunas : TFDQuery);
    procedure GerarClassLista(AArquivo : TStringList; AQueryTabelas : TFDQuery);
    procedure GerarImplementation(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarConstructor(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarClear(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarFunctions(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarInternalClear(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure GerarProcedures(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
    procedure Clear();

    function ValidaCampos(): Boolean;
  public
    { Public declarations }
    function SetParametrosConexao(Con : TFDConnection): Boolean;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}
uses
  DAOGenerator;

procedure TForm1.FormCreate(Sender: TObject);
begin
{$IFDEF DEBUG}
  SbOrigem.Text := 'C:\Ecosis\dados\ECODADOS.ECO';
  Tipo.Text := 'FireBird';
  EdtUser.Text := 'sysdba';
  EdtPassword.Text := 'masterkey';
  SbDestino.Text := 'C:\Projetos\Templates';
{$ENDIF}
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  Tipo.Items.Add('FireBird');
  Tipo.Items.Add('PostgreSQL');
  SbOrigem.SetFocus;
end;

function TForm1.Gerar: Boolean;
var
  LQueryTabelas: TFDQuery;
  LQueryColunas: TFDQuery;
  LArquivoDTO,
  LArquivoDAO : TStringList;
  LDaoGenerator : TDAOGenerator;
begin
  try
    LDaoGenerator := TDAOGenerator.Create;
    LQueryTabelas := TFDQuery.Create(nil);

    if Trim(CBTabelas.Text) <> '' then
    begin
      LQueryTabelas.SQL.Add(SelectTabelas + whereCustom);
      LQueryTabelas.Params.ParamByName('Tabela').AsString := CBTabelas.Text;
    end
    else
       LQueryTabelas.SQL.Add(SelectTabelas+wherePadrao);

    LQueryTabelas.Connection := FDConnection;

    if LQueryTabelas.OpenOrExecute then
    begin
      LQueryTabelas.First;
      while not(LQueryTabelas.EOF) do
      begin
        LArquivoDTO := TStringList.Create;
        NomeTabela := Copy(LQueryTabelas.FieldByName('Tabela').AsString, 2, Length(LQueryTabelas.FieldByName('Tabela').AsString));
        NomeTabela := Copy(NomeTabela,0,4)+LowerCase(Copy(NomeTabela,5,Length(NomeTabela)));
        if Trim(NomeTabela) <>'' then
        begin
          LQueryColunas := TFDQuery.Create(nil);
          LQueryColunas.SQL.Add(SelectColunas);
          LQueryColunas.ParamByName('Tabela').Value := LQueryTabelas.FieldByName('Tabela').AsString;
          LQueryColunas.Connection := FDConnection;
          if LQueryColunas.OpenOrExecute then
          begin
            Self.GerarCabecalho(LArquivoDTO, LQueryTabelas, LQueryColunas);
            Self.GerarImplementation(LArquivoDTO, LQueryTabelas, LQueryColunas);
            if DBGerarDAO.Checked then
            begin
              LArquivoDAO := TStringList.Create;
              LDaoGenerator.GerarDAO(LArquivoDAO, LQueryTabelas, LQueryColunas, FDConnection);
              LArquivoDAO.SaveToFile(SbDestino.Text+'\EcoDAO.'+NomeTabela+'.pas');
            end;
          end;
          LArquivoDTO.SaveToFile(SbDestino.Text+'\EcoDTO.'+NomeTabela+'.pas');
        end;
        LQueryTabelas.Next;
        FreeAndNil(LArquivoDAO);
        FreeAndNil(LArquivoDTO);
      end;
    end;
    Showmessage('Concluído');
    Result := True;
  except
    raise Exception.Create('Erro ao gerar DTO');
  end;
end;

procedure TForm1.GerarCabecalho(AArquivo: TStringList; AQueryTabelas, AQueryColunas: TFDQuery);
begin 
  try
    AArquivo.Add('unit EcoDTO.'+NomeTabela+';');
    AArquivo.Add('');
    AArquivo.Add('interface');
    AArquivo.Add('');
    AArquivo.Add('uses');
    AArquivo.Add('  EcoRT.Objects,');
    AArquivo.Add('  EcoRT.Data;');
    AArquivo.Add('');
    AArquivo.Add('type');
    AArquivo.Add('  TDto'+NomeTabela+' = class;');
    AArquivo.Add('');
    AArquivo.Add('  TDto'+NomeTabela+'s = class;');
    AArquivo.Add('');
    Self.GerarClassItem(AArquivo, AQueryTabelas, AQueryColunas);
    Self.GerarClassLista(AArquivo, AQueryTabelas);
  except
    Exception.ThrowOuterException(EGerarCabecalho.Create('Erro ao gerar o cabeçalho!'));
  end;
end;

procedure TForm1.GerarClassItem(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
begin
  try
    AArquivo.Add('  TDto'+NomeTabela+
                 ' = class(EcoRT.Objects.TErtCollectionItem<TDto'+NomeTabela+'s>)');
    AArquivo.Add('  private');
    AArquivo.Add('    { private declarations }');
    AArquivo.Add('    //----------------------------------------------------------------------//');    
    Self.GerarFields(AArquivo, AQueryColunas);
    Self.GerarDeclaracoesProcedures(AArquivo, AQueryColunas);
    Self.GerarDeclaracoesFunctions(AArquivo, AQueryColunas);
    Self.GerarPublicDeclarations(AArquivo, AQueryTabelas);
    Self.GerarPropertys(AArquivo, AQueryColunas);
    AArquivo.Add('  end;');
  except
    Exception.ThrowOuterException(EGerarClassItem.Create('Erro ao gerar o cabeçalho!'));
  end;
end;

procedure TForm1.GerarClassLista(AArquivo : TStringList; AQueryTabelas : TFDQuery);
begin
  try
    AArquivo.Add('  TDto'+NomeTabela+'s = class(EcoRT.Objects.TErtCollection<'+
                 'TDto'+NomeTabela+'>)');
    AArquivo.Add('  end;');
    AArquivo.Add('');
  except
    Exception.ThrowOuterException(EGerarClassLista.Create('Erro ao gerar a lista do objeto!'));
  end;
end;

procedure TForm1.GerarClear(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  try
    AArquivo.Add('procedure TDto'+NomeTabela+'.Clear;');
    AArquivo.Add('begin');
    AArquivo.Add('  Self.InternalClear();');
    AArquivo.Add('end;'); 
    AArquivo.Add('');             
  except
    Exception.ThrowOuterException(EGerarFunctions.Create('Erro ao gerar as procedures!'));
  end;
end;

procedure TForm1.GerarConstructor(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin    
  try
    AArquivo.Add('constructor TDto'+NomeTabela+'.Create(const AOwner: EcoRT.Objects.TErtPersistent);');
    AArquivo.Add('begin');
    AArquivo.Add('  Self.InternalClear();');
    AArquivo.Add('  inherited Create(AOwner);');
    AArquivo.Add('end;'); 
    AArquivo.Add('');             
  except
    Exception.ThrowOuterException(EGerarFunctions.Create('Erro ao gerar as procedures!'));
  end;
end;

procedure TForm1.GerarFields(AArquivo: TStringList; AQueryColunas: TFDQuery);
begin
  try
    AQueryColunas.First;
    while not (AQueryColunas.Eof) do
    begin
      AArquivo.Add('    F'+AQueryColunas.FieldByName('Coluna').AsString+': '+AQueryColunas.FieldByName('Tipo').AsString+';');
      AQueryColunas.Next;
    end;
    AArquivo.Add('');
    AArquivo.Add('    //----------------------------------------------------------------------//');
  except
    Exception.ThrowOuterException(EGerarFields.Create('Erro ao gerar os campos!'));
  end;
end;

procedure TForm1.GerarFunctions(AArquivo: TStringList; AQueryTabelas, AQueryColunas: TFDQuery);
begin
  try
    AQueryColunas.First;
    while not(AQueryColunas.Eof) do
    begin
      AArquivo.Add('function TDto'+NomeTabela+
                    '.Get'+AQueryColunas.FieldByName('Coluna').AsString+': '+AQueryColunas.FieldByName('Tipo').AsString+';');
      AArquivo.Add('begin');
      AArquivo.Add('  Result := F'+AQueryColunas.FieldByName('Coluna').AsString+';');
      AArquivo.Add('end;'); 
      AArquivo.Add('');
      AQueryColunas.Next;
    end;
  except
    Exception.ThrowOuterException(EGerarFunctions.Create('Erro ao gerar as funções!'));
  end;
end;

procedure TForm1.GerarImplementation(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
begin
  try
    AArquivo.Add('implementation');
    AArquivo.Add('');
    AArquivo.Add('{ TDto'+NomeTabela+' }');
    AArquivo.Add('');
    Self.GerarClear(AArquivo, AQueryTabelas, AQueryColunas);
    Self.GerarConstructor(AArquivo, AQueryTabelas, AQueryColunas);
    Self.GerarFunctions(AArquivo, AQueryTabelas, AQueryColunas);
    Self.GerarInternalClear(AArquivo, AQueryTabelas, AQueryColunas);
    Self.GerarProcedures(AArquivo, AQueryTabelas, AQueryColunas);
    AArquivo.Add('end.');
  except   
    Exception.ThrowOuterException(EGerarImplementation.Create('Erro ao gerar os campos!'));
  end;
end;

procedure TForm1.GerarInternalClear(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  try
    AQueryColunas.First; 
    AArquivo.Add('procedure TDto'+NomeTabela+'.InternalClear;');
    AArquivo.Add('begin');
    while not(AQueryColunas.Eof) do
    begin
      AArquivo.Add('  ErtData.Clear(F'+AQueryColunas.FieldByName('Coluna').AsString+');');
      AQueryColunas.Next;
    end;
    AArquivo.Add('end;'); 
    AArquivo.Add('');             
  except
    Exception.ThrowOuterException(EGerarFunctions.Create('Erro ao gerar as funções!'));
  end;
end;

procedure TForm1.GerarProcedures(AArquivo: TStringList; AQueryTabelas, AQueryColunas: TFDQuery);
begin
  try
    AQueryColunas.First;
    while not(AQueryColunas.Eof) do
    begin
      AArquivo.Add('procedure TDto'+NomeTabela+
                    '.Set'+AQueryColunas.FieldByName('Coluna').AsString+'(const A'+AQueryColunas.FieldByName('Coluna').AsString+': '+AQueryColunas.FieldByName('Tipo').AsString+');');
      AArquivo.Add('begin');
      AArquivo.Add('  FieldSet(F'+AQueryColunas.FieldByName('Coluna').AsString+',A'+AQueryColunas.FieldByName('Coluna').AsString+');');
      AArquivo.Add('end;');
      AArquivo.Add('');
      AQueryColunas.Next;
    end;
  except
    Exception.ThrowOuterException(EGerarFunctions.Create('Erro ao gerar as funções!'));
  end;
end;

procedure TForm1.GerarPropertys(AArquivo: TStringList; AQueryColunas: TFDQuery);
begin
  try
    AQueryColunas.First;
    while not (AQueryColunas.Eof) do
    begin
      AArquivo.Add('    property '+AQueryColunas.FieldByName('Coluna').AsString+
                   ': '+AQueryColunas.FieldByName('tipo').AsString+
                   ' read Get'+AQueryColunas.FieldByName('Coluna').AsString+
                   ' write Set'+AQueryColunas.FieldByName('Coluna').AsString+';');
      AQueryColunas.Next;
    end;
    AArquivo.Add('');
    AArquivo.Add('    //----------------------------------------------------------------------//');
  except
    Exception.ThrowOuterException(EGerarProperty.Create('Erro ao gerar as propertys!'));
  end;
end;

procedure TForm1.GerarPublicDeclarations(AArquivo: TStringList;  AQueryTabelas: TFDQuery);
begin
  try
    AArquivo.Add('  public																					');
    AArquivo.Add('    /// <summary>                                                                         ');
    AArquivo.Add('    ///   Classe responsável pelas informações da tabela '+NomeTabela+' do sistema ECO.   ');
    AArquivo.Add('    /// </summary>                                                                        ');
    AArquivo.Add('    constructor Create(const AOwner: EcoRT.Objects.TErtPersistent); override;             ');
    AArquivo.Add('                                                                                          ');
    AArquivo.Add('    /// <summary>                                                                         ');
    AArquivo.Add('    ///   Método responsável por limpar as propriedades do objeto.                        ');
    AArquivo.Add('    /// </summary>                                                                        ');
    AArquivo.Add('    procedure Clear; inline;                                                              ');
    AArquivo.Add('                                                                                          ');
    AArquivo.Add('    //----------------------------------------------------------------------//            ');
  except
    Exception.ThrowOuterException(EGerarPublicDeclarations.Create('Erro ao gerar as declarações publicas!'));
  end;
end;

procedure TForm1.GerarDeclaracoesFunctions(AArquivo: TStringList; AQueryColunas: TFDQuery);
begin
  try
    AQueryColunas.First;
    while not (AQueryColunas.Eof) do
    begin;
      AArquivo.Add('    function Get'+AQueryColunas.FieldByName('Coluna').AsString+'(): '+AQueryColunas.FieldByName('Tipo').AsString+';');
      AQueryColunas.Next;
    end;
    AArquivo.Add('');
    AArquivo.Add('    //----------------------------------------------------------------------//');
  except
    Exception.ThrowOuterException(EGerarDelarationsFunctions.Create('Erro ao gerar as funções!'));
  end;
end;

procedure TForm1.GerarDeclaracoesProcedures(AArquivo: TStringList; AQueryColunas: TFDQuery);
begin
  try
    AArquivo.Add('    procedure InternalClear();');
    AQueryColunas.First;
    while not (AQueryColunas.Eof) do
    begin
      AArquivo.Add('    procedure Set'+AQueryColunas.FieldByName('Coluna').AsString+'(const A'+
                          AQueryColunas.FieldByName('Coluna').AsString+': '+AQueryColunas.FieldByName('Tipo').AsString+');');
      AQueryColunas.Next;
    end;
    AArquivo.Add('');
    AArquivo.Add('    //----------------------------------------------------------------------//');
  except
    Exception.ThrowOuterException(EGerarDelarationsProcedures.Create('Erro ao gerar as procedures!'));
  end;
end;

procedure TForm1.SbOrigemInvokeSearch(Sender: TObject);
begin
  if OpenDialog1.Execute then
    SbOrigem.Text := OpenDialog1.FileName;
end;

procedure TForm1.SbDestinoInvokeSearch(Sender: TObject);
var
  selDir : string;
begin
  SelectDirectory('Selecione uma pasta', 'C:\', selDir);
  SbDestino.Text := selDir;
end;

function TForm1.SetParametrosConexao(Con : TFDConnection): Boolean;
begin
  try
    if Tipo.Text = 'FireBird' then
    begin
      Con.DriverName      := 'FB';
      Con.Params.DriverID := 'FB';
    end else
    begin
      Con.DriverName      := 'PG';
      Con.Params.DriverID := 'PG';
    end;
    Con.Params.Database := SbOrigem.Text;
    Con.Params.UserName := EdtUser.Text;
    Con.Params.Password := EdtPassword.Text;
    Con.Connected := True;
    Result := True;
  except
    raise Exception.Create('Usuario ou senha incorretos!');
  end;
end;

function TForm1.ValidaCampos: Boolean;
begin
  Result := ((Trim(SbOrigem.Text) <> '') and
            (Trim(Tipo.Text)<> '' ) and
            (Trim(EdtUser.Text)<> '' ) and
            (Trim(EdtPassword.Text)<> '' ) and
            (Trim(SbDestino.Text) <> ''));
end;

procedure TForm1.BtnCancelarClick(Sender: TObject);
begin
	Clear;
end;

procedure TForm1.BtnGerarClick(Sender: TObject);
begin
  if ValidaCampos then
  begin
  	if Gerar then
    	Clear();
  end else
  begin
    ShowMessage('Preencha os campos corretamente');
    Exit;
  end;
end;

procedure TForm1.Clear;
begin
  SbOrigem.Text    := '';
  Tipo.Text        := '';
  EdtUser.Text     := '';
  EdtPassword.Text := '';
  SbDestino.Text   := '';
  SbOrigem.SetFocus;
end;

procedure TForm1.EdtPasswordExit(Sender: TObject);
var
  AQueryTabelas: TFDQuery;
begin
  if SetParametrosConexao(FDConnection) then
  begin
    AQueryTabelas := TFDQuery.Create(nil);
    AQueryTabelas.SQL.Add(SelectTabelas + wherePadrao);
    AQueryTabelas.Connection := FDConnection;
    try
      if AQueryTabelas.OpenOrExecute then
      begin
        AQueryTabelas.First;
        while not(AQueryTabelas.Eof) do
        begin
          CBTabelas.Items.Add(AQueryTabelas.FieldByName('Tabela').AsString);
          AQueryTabelas.Next;
        end;
      end;
    finally
      FreeAndNil(AQueryTabelas);
    end;
  end else
    Exit;
end;

end.

