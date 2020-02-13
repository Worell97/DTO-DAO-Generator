unit DAOGenerator;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Buttons;
type
  EGerarCabecalho = class(System.SysUtils.Exception);
  EGerarSelect = class(System.SysUtils.Exception);
  EGerarInsert = class(System.SysUtils.Exception);
  TDAOGenerator = class
    private
    { private declarations }
      NomeTabela : System.UnicodeString;
      FQueryKeys : TFDQuery;
      procedure GerarCabecalho(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarClass(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarImplementation(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarUses(AArquivo : TStringList);
      procedure GerarConstructor(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarInternalLoad(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarSave(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarUnsafeSave(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarSelect(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarInsert(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarUpdate(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarStrictPrivateDeclarations(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      procedure GerarPublicDeclarations(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
      function  GetKey(AQueryKeys : TFDQuery): Boolean;

    public
    { public declarations }
      procedure GerarDAO(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery; FDConnection: TFDConnection);

  end;
implementation

{ TDAOGenerator }

procedure TDAOGenerator.GerarCabecalho(AArquivo : TStringList; AQueryTabelas, AQueryColunas : TFDQuery);
begin
  try
    AArquivo.Add('unit EcoDAO.'+Copy(NomeTabela, 2, Length(NomeTabela))+';');
    AArquivo.Add('');
    AArquivo.Add('interface');
    GerarUses(AArquivo);
    AArquivo.Add('');
    AArquivo.Add('type');
  except
    Exception.ThrowOuterException(EGerarCabecalho.Create('Erro ao gerar o cabeçalho!'));
  end;
end;

procedure TDAOGenerator.GerarClass(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  AArquivo.Add('');
  AArquivo.Add('  TDao'+Copy(NomeTabela, 2, Length(NomeTabela))+' = class(TDaoObjectFb)');
  GerarStrictPrivateDeclarations(AArquivo, AQueryTabelas, AQueryColunas);
  GerarPublicDeclarations(AArquivo, AQueryTabelas, AQueryColunas);
  AArquivo.Add('  end;');
end;

procedure TDAOGenerator.GerarConstructor(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  AArquivo.Add('constructor TDao'+Copy(NomeTabela, 2, Length(NomeTabela))+'.Create;');
  AArquivo.Add('begin');
  AArquivo.Add('  inherited;');
  AArquivo.Add('');
  AArquivo.Add('end;');
end;

procedure TDAOGenerator.GerarDAO(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery; FDConnection: TFDConnection);
begin
  FQueryKeys := TFDQuery.Create(nil);
  FQueryKeys.Connection := FDConnection;
  NomeTabela := AQueryTabelas.FieldByName('tabela').AsString;
  if (GetKey(FQueryKeys)) then
  begin
    GerarCabecalho(AArquivo, AQueryTabelas, AQueryColunas);
    GerarClass(AArquivo, AQueryTabelas, AQueryColunas);
    GerarImplementation(AArquivo, AQueryTabelas, AQueryColunas);
  end else
    Raise Exception.Create('Não foi possível encontrar as chaves da tabela');

end;

procedure TDAOGenerator.GerarImplementation(AArquivo: TStringList;
  AQueryTabelas, AQueryColunas: TFDQuery);
begin
  AArquivo.Add('implementation');
  AArquivo.Add('');
  AArquivo.Add('  uses');
  AArquivo.Add('    EcoDAO.Utils,');
  AArquivo.Add('    EcoRT.Types,');
  AArquivo.Add('    System.TypInfo;');
  AArquivo.Add('');
  AArquivo.Add('{ TDao'+Copy(NomeTabela, 2, Length(NomeTabela))+' }');
  GerarConstructor(AArquivo, AQueryTabelas, AQueryColunas);
  GerarInternalLoad(AArquivo, AQueryTabelas, AQueryColunas);
  GerarSave(AArquivo, AQueryTabelas, AQueryColunas);
  GerarUnsafeSave(AArquivo, AQueryTabelas, AQueryColunas);
  AArquivo.Add('');
  AArquivo.Add('end.');
end;

procedure TDAOGenerator.GerarInsert(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
var
  Coluna : System.UnicodeString;
begin
  try
    AArquivo.Add('');
    AArquivo.Add('      Insert: System.UnicodeString =');
    AArquivo.Add('          '+QuotedStr('INSERT INTO '+ AQueryTabelas.FieldByName('tabela').AsString+'(')+' +');
    AQueryColunas.First;
    while not(AQueryColunas.Eof) do
    begin
      Coluna := '  '+AQueryColunas.FieldByName('coluna').AsString;
      AQueryColunas.Next;
      if AQueryColunas.Eof then
        AArquivo.Add('          '+QuotedStr(Coluna+')')+'+')
      else
        AArquivo.Add('          '+QuotedStr(Coluna+',')+'+');
    end;
    AQueryColunas.First;
    AArquivo.Add('          '+QuotedStr('VALUES(')+'+');
    while not(AQueryColunas.Eof) do
    begin
      Coluna := '  :'+AQueryColunas.FieldByName('coluna').AsString;
      AQueryColunas.Next;
      if AQueryColunas.Eof then
        AArquivo.Add('          '+QuotedStr(Coluna+')')+';')
      else
        AArquivo.Add('          '+QuotedStr(Coluna+',')+'+');
    end;
  except
    Exception.ThrowOuterException(EGerarInsert.Create('Erro ao gerar a constante Insert!'));
  end;
end;

procedure TDAOGenerator.GerarInternalLoad(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  AArquivo.Add('');
  AArquivo.Add('class procedure TDao'+Copy(NomeTabela, 2, Length(NomeTabela))+'.InternalLoad(var AQuery: TEcoDataQueryMarshall;const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+');');
  AArquivo.Add('begin');
  AQueryColunas.First;
  while not(AQueryColunas.Eof) do
  begin
    if (AQueryColunas.FieldByName('tipoDAO').AsString = 'Date') or
       (AQueryColunas.FieldByName('tipoDAO').AsString = 'Time') or
       (AQueryColunas.FieldByName('tipoDAO').AsString = 'DateTime') then
    begin
      AArquivo.Add('  AObject.'+AQueryColunas.FieldByName('coluna').AsString+
                   ' := DBField.DBToT'+AQueryColunas.FieldByName('tipoDAO').AsString+
                   '(AQuery.FieldByName('+QuotedStr(AQueryColunas.FieldByName('coluna').AsString)+'));');
    end else
      AArquivo.Add('  AObject.'+AQueryColunas.FieldByName('coluna').AsString+
                   ' := DBField.DBTo'+AQueryColunas.FieldByName('tipoDAO').AsString+
                   '(AQuery.FieldByName('+QuotedStr(AQueryColunas.FieldByName('coluna').AsString)+'));');

    AQueryColunas.Next;
  end;
  AArquivo.Add('end;');
end;

procedure TDAOGenerator.GerarPublicDeclarations(AArquivo: TStringList;
  AQueryTabelas, AQueryColunas: TFDQuery);
begin
  AArquivo.Add('');
  AArquivo.Add('  public');
  AArquivo.Add('      constructor Create; override;');
  AArquivo.Add('      procedure UnsafeSave(const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+');');
  AArquivo.Add('      procedure Save(const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+');');
  AArquivo.Add('      procedure Load   (const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+
               '; const AEmpresa: System.UnicodeString; const ACodigoCC : System.UnicodeString);');
  AArquivo.Add('      procedure UnSafeDelete (const AEmpresa: System.UnicodeString; const AGID : System.UnicodeString);');
  AArquivo.Add('      procedure SafeDelete (const AEmpresa: System.UnicodeString; const AGID : System.UnicodeString);');
  AArquivo.Add('');
  AArquivo.Add('      function  TryLoad(const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+
               '; const AEmpresa: System.UnicodeString; const ACodigoCC : System.UnicodeString): Boolean;');
end;

procedure TDAOGenerator.GerarStrictPrivateDeclarations(AArquivo: TStringList;
  AQueryTabelas, AQueryColunas: TFDQuery);
begin
  AArquivo.Add('  strict private');
  AArquivo.Add('    const');
  GerarSelect(AArquivo, AQueryTabelas, AQueryColunas);
  GerarInsert(AArquivo, AQueryTabelas, AQueryColunas);
  GerarUpdate(AArquivo, AQueryTabelas, AQueryColunas);
  AArquivo.Add('    class procedure InternalLoad(var AQuery: TEcoDataQueryMarshall; const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+'); static;');
end;

procedure TDAOGenerator.GerarSave(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  AArquivo.Add('');
  AArquivo.Add('procedure TDao'+Copy(NomeTabela, 2, Length(NomeTabela))+'.Save(const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+');');
  AArquivo.Add('begin');
  AArquivo.Add('   FConnection.StartTransaction;');
  AArquivo.Add('   try');
  AArquivo.Add('      UnsafeSave(AObject);');
  AArquivo.Add('      FConnection.Commit;');
  AArquivo.Add('   except');
  AArquivo.Add('      FConnection.Rollback;');
  AArquivo.Add('      raise;');
  AArquivo.Add('   end;');
  AArquivo.Add('end;');
end;

procedure TDAOGenerator.GerarSelect(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
var
  LColuna : System.UnicodeString;
begin
  try
    AArquivo.Add('      Select: System.UnicodeString =');
    AArquivo.Add('          '+QuotedStr('SELECT ')+' +');
    AQueryColunas.First;
    while not(AQueryColunas.Eof) do
    begin
      LColuna := AQueryColunas.FieldByName('coluna').AsString;
      AQueryColunas.Next;
      if AQueryColunas.Eof then
        AArquivo.Add('          '+QuotedStr('  '+LColuna)+'+')
      else
        AArquivo.Add('          '+QuotedStr('  '+LColuna+',')+'+');
    end;
    AArquivo.Add('          '+QuotedStr('FROM ')+' +');
    AArquivo.Add('          '+QuotedStr('   '+AQueryTabelas.FieldByName('tabela').AsString+' ')+' ;');
  except
    Exception.ThrowOuterException(EGerarSelect.Create('Erro ao gerar a constante Select!'));
  end;
end;

procedure TDAOGenerator.GerarUnsafeSave(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
begin
  AArquivo.Add('');
  AArquivo.Add('procedure TDao'+Copy(NomeTabela, 2, Length(NomeTabela))+'.UnsafeSave(const AObject: TDto'+Copy(NomeTabela, 2, Length(NomeTabela))+');');
  AArquivo.Add('var');
  AArquivo.Add('  LQuery: TEcoDataQueryMarshall;');
  AArquivo.Add('  LGId: System.UnicodeString;');
  AArquivo.Add('begin');
  AArquivo.Add('  if (AObject.'+FQueryKeys.FieldByName('RDB$FIELD_NAME').AsString+' > 0) then ');
  AArquivo.Add('  begin');
  AArquivo.Add('    LQuery.Create(FConnection, Insert)');
  AQueryColunas.First;
  while not(AQueryColunas.Eof) do
  begin
    AArquivo.Add('      .ParamSet'+AQueryColunas.FieldByName('tipoDAO').AsString+'('+QuotedStr(AQueryColunas.FieldByName('coluna').AsString)+', AObject.'+AQueryColunas.FieldByName('coluna').AsString+')');
    AQueryColunas.Next;
  end;
  AArquivo.Add('      .ExecSQL;');
  AArquivo.Add('  end else');
  AArquivo.Add('  begin');
  AArquivo.Add('    LQuery.Create(FConnection, Update)');
  AQueryColunas.First;
  while not(AQueryColunas.Eof) do
  begin
    AArquivo.Add('      .ParamSet'+AQueryColunas.FieldByName('tipoDAO').AsString+'('+QuotedStr(AQueryColunas.FieldByName('coluna').AsString)+', AObject.'+AQueryColunas.FieldByName('coluna').AsString+')');
    AQueryColunas.Next;
  end;
  AArquivo.Add('      .ExecSQL;');
  AArquivo.Add('  end;');
  AArquivo.Add('end;');
end;

procedure TDAOGenerator.GerarUpdate(AArquivo: TStringList; AQueryTabelas,
  AQueryColunas: TFDQuery);
var
  LColuna : System.UnicodeString;
begin
  try
    AArquivo.Add('');
    AArquivo.Add('      Update: System.UnicodeString =');
    AArquivo.Add('          '+QuotedStr('UPDATE ')+' +');
    AArquivo.Add('          '+QuotedStr(' '+AQueryTabelas.FieldByName('tabela').AsString+' ')+' +');
    AArquivo.Add('          '+QuotedStr('SET ')+' +');
    AQueryColunas.First;
    while not(AQueryColunas.Eof) do
    begin
      LColuna := AQueryColunas.FieldByName('coluna').AsString;
      AQueryColunas.Next;
      if AQueryColunas.Eof and FQueryKeys.IsEmpty then
        AArquivo.Add('          '+QuotedStr('  '+LColuna+' = :'+LColuna)+';')
      else
        AArquivo.Add('          '+QuotedStr('  '+LColuna+' = :'+LColuna+', ')+'+');
    end;
    AArquivo.Add('          '+QuotedStr('WHERE ')+' +');
    FQueryKeys.First;
    while not(FQueryKeys.Eof) do
    begin

      LColuna := FQueryKeys.FieldByName('RDB$FIELD_NAME').AsString;
      FQueryKeys.Next;
      if FQueryKeys.Eof then
        AArquivo.Add('          '+QuotedStr('  '+LColuna+' = :'+LColuna)+';')
      else
        AArquivo.Add('          '+QuotedStr('  '+LColuna+' = :'+LColuna+', ')+'+');
    end;

    AArquivo.Add('');
    AArquivo.Add('');
  except
    Exception.ThrowOuterException(EGerarSelect.Create('Erro ao gerar o cabeçalho!'));
  end;
end;

procedure TDAOGenerator.GerarUses(AArquivo: TStringList);
begin
    AArquivo.Add('');
    AArquivo.Add('uses');
    AArquivo.Add('  EcoRT.Objects,');
    AArquivo.Add('  EcoRT.Data,');
    AArquivo.Add('  EcoDB.Objects,');
    AArquivo.Add('  EcoDAO.Objects,');
    AArquivo.Add('  System.SysUtils,');
    AArquivo.Add('  Data.DB,');
    AArquivo.Add('  EcoDTO.'+Copy(NomeTabela, 2, Length(NomeTabela))+';');
end;

function TDAOGenerator.GetKey(AQueryKeys: TFDQuery): Boolean;
const
  SelectKeys = 'SELECT RDB$FIELD_NAME,C.RDB$CONSTRAINT_TYPE '+
               'FROM                                  '+
               '  RDB$RELATION_CONSTRAINTS C,         '+
               '  RDB$INDEX_SEGMENTS S                '+
               'WHERE                                 '+
               '  C.RDB$RELATION_NAME = :Tabela AND   '+
               '  S.RDB$INDEX_NAME = C.RDB$INDEX_NAME ';
begin
  AQueryKeys.SQL.Add(SelectKeys);
  AQueryKeys.Params.ParamByName('Tabela').AsString := NomeTabela;
  Result := AQueryKeys.OpenOrExecute;
end;

end.
