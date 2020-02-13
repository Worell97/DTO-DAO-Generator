// - 12/11/2003 - Kleber - Alteração do tipo de pessoa na função EditaCGCCPF, pessoa do tipo P (Produtor rural)


(* --------------------------------------------------------
   Rotinas gerais para o sistema ECO
   --------------------------------------------------------
   Empresa....: Eco Sistemas Informática
   Sistema....: Eco
   Autor......: Equipe ECO
   Linguagem..: Delphi6 / Kylix2

   historicos:
   -----------
   13-Mar-02 - Inclusao das rotinas de extenso
   14-Mar-02 - ajustes na biblioteca de extenso quanto ao tipo de retorno etc...
   14-Mar-02 - funcao justifica texto
   14-Mar-02 - acerto nas rotinas de pascoa
   14-Mar-02 - ajustes para conversao no CLX - linux - incompleto
   02-Ago-02 - teio - adicao de novas funcoes
   14-Ago-02 - teio - adicao de novas funcoes
   04-Apr-05 - Funcao que retorna o nome do computador.
   11-May-05 - Nova variavel NumeroOrdemGlobal.
   16/05/2005 - Nova variavel string ObservacaoDoItemNoPedido, serve pra Ven601RA, RB.
   28/05/2005 - Daniel-Nova funcao: ChamaTelaDeExportar: Agora, em qualquer pesquisa que tenha
                uma query, pode ser chamado esta funcao passando apenas o nome da Lista (string)
                e o nome da query (TSQLQuery). Sistema abre tela padrao pra usuario escolher em
                qual tipo de arquivo quer salvar.
*)

unit Geral;

interface
          
uses
     {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }

     DnPrn, cxButtons, DnPanel, cxPC, DnEdit, cxGroupBox,
     EditBox, DnBox, NumEdit, QuickRpt, extctrls, DBCtrls,
     FMTBcd, DBGrids,dnDBGrid, CheckLst, Buttons, DBTables, ACBrCHQ, DBGridHack,
     cxCheckComboBox, advlued, AdvDateTimePicker, DnComboBox,
     cxGrid, cxGridDBTableView,
     {$ENDIF}

     IniFiles, Variants, stdctrls, classes, Forms, Dialogs,
     Controls, Windows, Sysutils, Messages, Graphics,
     StrUtils, DateUtils, SimpleDS, Math, Grids, SqlExpr, DBClient,
     DB, Mask, Registry, ComCtrls, Printers, DBXCommon, Contnrs, WinSpool,
     FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
     FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client,
     System.UITypes, System.Generics.Collections, System.RegularExpressions;

type

   TModoCadastro = (tInclusao, tAlteracao, tConsulta, tExclusao);
   TTipoResgate      = (trFrente, trSaldo);
   TTipoBonus        = (tbVenda, tbDevolucao);
   TTipoMsg          = (tmConfirma, tmConfirmaN, tmInforma, tmErro, tmESC, tmNada, tmContinuarCancelar, tmConfirmaAlertaN);
   TTipoAcesso       = (TmIncluir, TmEditar, TmDeletar, TmConsultar, TmGeral);
   TAlinhamento      = (taEsquerdo, taCentro, taDireito);
   TDestinoRelatorio = (Video, Impressora);
   TModoRelatorio    = (Texto, Grafico);
   TSistema          = (tsCaixa, tsBanco, tsReceber, tsPagar, tsEstoque, tsVenda, tsGeral, tsAcessorios, tsContabilidade, tsExportacao, tsPCP, tsDRH);
   TMarcaDesmarca    = (Marca, Desmarca);
   TImpressora       = (tiPedidoNormal,
                        tiPedidoPrazo,
                        tiPedidoReduzido,
                        tiOrcamentoNormal,
                        tiOrcamentoReduzido,
                        tiCondicionalNormal,
                        tiCondicionalReduzida,
                        tiOrdemServico,
                        tiNotaPromissoria,
                        tiDuplicata,
                        tiEntregaFutura,
                        tiRecibo,
                        tiListaSeparacao,
                        tiBoleto,
                        tiCarne,
                        tiNotaFiscal,
                        tiNotaServico,
                        tiContrato,
                        tiCheque,
                        tiEtiqueta);

   TPrintCmd         = (tpCmd20cpi,
                        tpCmd17cpi,
                        tpCmd12cpi,
                        tpCmd10cpi,
                        tpCmd06cpi,
                        tpCmd05cpi,
                        tpCmdLgNeg,
                        tpCmdDgNeg,
                        tpCmdComum,
                        tpCmdNeg10,
                        tpCmdNeg12);


    TTipoParamEco = (tpeBoolean, tpeString, tpeInteger, tpeNumeric);

    TModulosEco = (meCaixa,
                   meBanco,
                   mePagar,
                   meReceber,
                   meVenda,
                   meExportacao,
                   meMadeireira,
                   meEstoque,
                   mePCP,
                   meGeral);

   { variaveis definidas como constantes para leitura de paramtros do DB }
    TParmsEco = (
         // Parametro do contas Receber
         peAutenticaDocumento,

         // Parametros da Venda
         peUsaRegistradora,
         peUsaDeptoCrediario,

         peEmiteBoleto,
         peEmitePedido,
         peEmitePedidoRecibo,
         peEmiteContrato,
         peEmiteNotaPromissoria,
         peEmiteEntrega,
         peEmiteListaSeparacao,
         peEmiteRecibo,
         peEmiteDuplicata,
         peEmiteCarne,
         peEmiteNotaFiscal,
         peEditaObservacaoNota,
         peEmiteCupomFiscal,
         peEmiteCupomVinculado,

         peConfirmaBoleto,
         peConfirmaCarne,
         peConfirmaDuplicata,
         peConfirmaPromissoria,
         peConfirmaContrato,
         peConfirmaPedido,
         peConfirmaPedidoRecibo,
         peConfirmaEntrega,
         peConfirmaSeparacao,
         peConfirmaRecibo,

         peBoletoNaVenda,
         peCarneNaVenda,
         peDuplicataNaVenda,
         pePromissoriaNaVenda,
         peContratoNaVenda,
         pePedidoNaVenda,
         peReciboNaVenda,
         peComprovanteNaVenda,
         peListaNaVenda,
         peApresentaTelaDetalhes,

         // usados na venda
         peImpNomeCliente,
         peImpEnderecoCliente,
         peImpCidadeCliente,
         peImpCnpjCpfCliente,
         peImpDocumentoOrigem,
         peImpComprador,
         peImpParcelas,
         peImpMensagem,
         peMensagem1,
         peMensagem2,
         peImpMediaConsumo,

         // estoques
         peEstoqueNegativo,
         peUsaOrdemProducao,

         // madeireira
         peUsaRomaneio,
         peEditaRomaneio,
         peImprimeRomaneio,
         peEditaPauta,
         peIncentivoFiscal,
         peFundei,
         peFethab,
         peAlteraPauta,

         peNull // nao usado
      );

   TRecAtributo = record
       Nome     : String;
       Attrib   : String;
   end;

   TPapel = record
     Size: SmallInt;
     Width: SmallInt;
     Height: SmallInt;
   end;

   TDataEco = record
                 Dia: Word;
                 Mes: Word;
                 Ano: Word;
              end;

   TBoleto  = record
                 Cliente      :String;
                 AbrevTipo    :String;
                 Documento    :String;
                 Parcela      :String;
                 NomeCliente  :String;
                 Portador     :String;
                 Endereco     :String;
                 Bairro       :String;
                 Cidade       :String;
                 Estado       :String;
                 CepCobranca  :String;
                 CgcCpf       :String;
                 RgIe         :String;
                 CaixaPostal  :String;
                 Valor        :Currency;
                 Juros        :Currency;
                 Multa        :Currency;
                 Desconto     :Currency;
                 Emissao      :TDate;
                 Vencimento   :TDate;
                 QtdParcela   :String;
              end;

   TPrecoProduto = record
       Produto                  : String;
       AliqCompraIcms           : Currency;
       AliqCompraReducao        : Currency;
       AliqCompraMargem         : Currency;
       AliqVendaIcms1           : Currency;
       AliqVendaIcms2           : Currency;
       AliqVendaReducao         : Currency;
       AliqVendaProtege1        : Currency;
       AliqVendaProtege2        : Currency;
       PautaCompra              : Currency;
       PautaVenda               : Currency;
       MargemLucroGarantido     : Currency;
       VlrIcmsCompra            : Currency;
       VlrIcmsVenda             : Currency;
       VlrIcmsGarantidoIntegral : Currency;
       CargaPS                  : Currency;
       MargemLucro              : Currency;
       PrecoVenda               : Currency;
       PrecoAtacado             : Currency;
       CustoFabrica             : Currency;
       CustoOrigem              : Currency;
       CustoReposicao           : Currency;
       CustoMedioReposicao      : Currency;
       CustoFinal               : Currency;
       CustoUtilizado           : Integer;
       ClassificacaoEstadual    : Integer;
       Decimais                 : Integer;
       EstadoOrigem             : String;
       Tributacao               : String;
       CompraCSF                : String;
       VendaCSF1                : String;
   end;

  TIcms = record
             Classificacao       :Integer;
             ClassificacaoCompra :Byte;
             MargemGarantido     :Currency;

             CompraCsf           :Integer;
             CompraIcms          :Currency;
             CompraReducao       :Currency;
             CompraMargem        :Currency;
             CompraIcmsSubst     :Currency;
             CompraReducaoSubst  :Currency;
             CompraClassificacao :Integer;
             CompraPDif          :Currency;

             VendaCsf            :Integer;
             VendaIcms           :Currency;
             VendaReducao        :Currency;
             VendaDiferimento    :Currency;
             VendaProtege        :Currency;
             VendaMargem         :Currency;
             VendaIcmsSubst      :Currency;
             VendaReducaoSubst   :Currency;

             CSOSNConsumidor     :Integer;
             CSOSNCompra         :Integer;
             CSOSNRevenda        :Integer;
             CSOSNSuperSimples   :Integer;

             AliqInterna: Currency;
             AliqFcp: Currency;
          end;

   TCaixa = record
               Dinheiro    :Currency;
               Cheque      :Currency;
               ChequeVista :Currency;
               ChequePrazo :Currency;
               Credito     :Currency;
               ValeTroco   :Currency;
               CartaFrete  :Currency;
               ClientePermuta: String;
               Permuta     :Currency;
               Total       :Currency;
            end;

   TChequeMovimento = record
               Sequencia      :Integer;
               IdCheque       :Integer;
               IdMovimento    :Integer;
               IdEmpresa      :Integer;
               Identificador  :Integer;
               IdRegistradora :Integer;
               IdPeriodo      :Integer;
               IdCaixa        :String;
               DataCaixa      :TDate;
               Status         :String;
               Historico      :String;
               Usuario        :String;
               IdBaixaReplica :Integer;
            end;

   TFabricanteTintometrico = (ftSuvinil, ftColorQuick, ftSW);

var
   TipoMsg             :TTipoMsg; { Tipo de mensagens usadas na função Mensagem }
   TipoAcesso          :TTipoAcesso;
   SistemaAtivo        :TSistema; { Tipo de sistema ativo }

   CorTitulo           :TColor;
   CorFonteTitulo      :TColor;
   CorRodape           :TColor;
   CorFonteRodape      :TColor;
   CorAtivada          :TColor;
   CorDesativada       :TColor;
   CorLabel            :TColor;

   UltimoCodigoUsado   :string;

   { Variaveis de uso global }

   RefreshMenuPopup    :boolean;
   Usuario             :String;
   IdUsuario           :Integer;
   Desenvolvimento     :boolean;
   DiretorioDoExeNovo  :string;      // local onde fica o exectuavel de distribuicao.
   DiretorioDeRelatorios:string;    // local no servidor onde vão ficar os arquivos rpt
   PathLogoCrystal     :string;
   Empresa             :string;
   qstrEmpresa         :string;      // quoted string; ''01'' ou ''05''
   NomeEmpresa         :string;
   RazaoEmpresa        :string;
   EnderecoEmpresa     :string;
   BairroEmpresa       :string;
   CidadeEmpresa       :String;
   CodCidadeEmpresa    :String;
   FoneEmpresa         :string;
   EmpresaUsaNfe       :boolean;
   Industria           :boolean;
   CEPEmpresa          :string;
   UFEmpresa           :string;
   EmailEmpresa        :string;
   CpfCnpjEmpresa      :string;
   InscrEstEmpresa     :string;
   CodigoCRTEmpresa    :Shortint;
   TipoAtividadeEmpresa:Shortint;
   CargaMediaEstimativaEmpresa:Currency;
   EmpresaRegimeISSQN  :Shortint;
   sNumeroLicenca      :string;
   sNomeRepresentante  :string;
   sProdutoLMC         :String;
   CodigoEmpresa       :integer;
   CodigoECO           :integer;
   ModeloDoPedido      :integer;    // usado como modelo de pedido ven406la
   ModeloPedidoReduzido :integer;
   TipoAtividade       :String;      { Tipo de atividade do cliente, CONVENIENCIA, OFICINA etc.. Usado na OS nova }
   ObservacaoDoItemNoPedido:String;
   PelaRegistradora    :Boolean=False;
   NaoAddPisCofinsCusto :Boolean;

   RotinaSelecionada   :integer;    // Programa corrente / rontina selecionada
   RotinaAuxSel        :integer;
   S_RotinaSelecionada :integer;
   ProgramaSelecionado :string;     // nome do programa
   ProgramaPadrao      :integer;
   ItemSelecionado     :string;

   UsuarioLog          :string;     // usuario da estacao - default
   EmpresaLog          :string;     // empresa da estacao - default
   ProgramaLog         :string;     // programa padrao de execucao-desenvolvimento

   PrimeiraLinha       :string;
   SegundaLinha        :string;
   NumeroOrdemGlobal   :Integer;
   GlobalIDCupom       :String;

   UsuarioSistema      :String;
   UsuarioCliente      :String;
   SenhaSistema        :String;
   RetornoPesquisa     :String;
   SqlText             :String;
   TipoDocBoleto       :String;
   TipoDocVenda        :String;
   TipoDocConvenio     :String;
   TipoDocJuro         :String;
   AtrasoMaxParametro  :Integer;
   BUsuario            :String;   // Usuário de bloqueio do registro

   BaixaConvenio       :Boolean;
   BaixaPorLote        :Boolean;

   BoletoFoiImpresso   :Boolean; //Flag pra ver se confirmou ou nao a impressao de boleto bancario no Ven601RA;
   IDBloqueioRemoto    :Integer;
   SemRestricoesDeBloqueio:Boolean=True;
   //Impressao de duplicatas
   FazendoVenda        :Boolean;  //Usado ao imprimir duplicata na hora da venda.
   DupEmpresa          :String;
   DupCliente          :String;
   DupTipo             :String;
   DupDocumento        :String;

   EntregaFutura       :Boolean;  //Na impressao do relatorio de entrega, define se vai ser futura ou de pedido.
   PedidoPraImprimir   :String;

   PesquisaIncrementalDosProdutos :Boolean;  { usado no f2-Produtos em clientes que possuem +10000 itens ou a maquina lenta..}
   PesquisaIncrementalDosClientes :Boolean;  { usado no f2-Clientes em clientes que possuem muitos clientes cadastrados (Dorane) }
   SelecionaQtdeProdutoAutomatico :Boolean;  { Quando não for incremental dos produtos, se já seleciona a quantidade do item selecionado }
   UnificaRecebimento             :Boolean;  { Define se o contas a receber será unificado }
   UnificaPagamento               :Boolean;  { Define se o contas a pagar é unificado }

   MotivoBloqChDevProprio  :String;
   MotivoBloqChDevTerceiro :String;

   //Posto
   MostrandoAbastecimentos :Boolean; //Controla se tela que mostra abastecimentos pendentes esta aberta pra nao abrir novamente.
   ProdutoDaBomba          :String;  //Produto selecionado na bomba de combustivel
   QtdeProdutoBomba        :Currency;  //Qtde do produto selecionado na bomba de combustivel
   EscolheuBicoComF1       :Boolean;
   cNumeroDigitosBomba     :byte;      // o numero de digitos das bombas de abastecimento nos postos;

   EscolheuAbastecimento   :Boolean;
   TelaAbastFechada        :Boolean;
   NumeroBicoGlobal        :String;
   IDAbastecimentoGlobal   :Integer; //Usado na baixa do abastecimento.
   PdtCustoUtilizado       :Integer; // Tipo de custo utilizado como base para formar preco de venda dos produtos
                                     // 0 = Custo reposicao (Padrao)
                                     // 1 = Custo merio reposicao

   //Campanha
   AceitouCartao           : Boolean;
   UsouFidelidade          : Boolean;
   CampanhaEmAberto        : Integer;
   cValorEntrada           : Currency;
   cValorLiquido           : Currency;
   TipoDeResgate           : TTipoResgate;
   TipoBonus               : TTipoBonus;

   cTotalBonusVista        : Currency;
   cTotalBonusPrazo        : Currency;

   Devolucao               : Boolean;
   PelaFolha               : Boolean=False;
   ExisteTransacaoComCheque: Boolean;
   { Variaveis de Impressão de NF }
   slDEF, slNF         :TStringList;  // StringList de Definições e StringList de NF
   Linha               :String;
   cLin,k,
   iSize, iWidth,
   iHeight,
   iAlturaLinha,
   LarguraCaracter,
   LarguraPagina,
   AlturaLinha,
   AlturaPagina,
   QtdeProdutos,
   QtdeServicos,
   PCol, PRow          :Integer;
   QuebrouPagina       :Boolean;
   Titulo,
   TipoLinha,
   Tamanho,
   NormalBold          :String;
   Orientacao          :TPrinterOrientation;
   i, j,                                                        // declaração de variáveis auxiliares
   Lin, Col, Tam,
   LIni, LFim,
   LdtCont,
   ContLinha,
   Contador,
   QtdeLinhas    : integer;
   Multiplicador,
   Multiplica    : Real;
//   Sx, X, Y,     Campo,
   sMascara,
   Atributo      : String;
   iLinhasPolegada :Integer;
   iLargura        :Integer;

   CodImpressoraDanfe  :String;
   CodImpressoraDanfeNFC  :String;
   CodImpressoraDanfeNFSE: string;
   CodImpressoraEcoProntaEntrega: string;

   BancoDeDados        :String;

   ServidorScripts     :String;  // Endereço do servidor de scripts
   ServidorLog         :String;  // Endereço do Servidor de log

   Transporte          :String; // Variavel de uso geral
   Transporte1         :String; // Variavel de uso geral
   Transporte2         :String; // Variavel de uso geral
   Transporte3         :String; // Variavel de uso geral
   Transporte4         :String; // Variavel de uso geral
   Transporte5         :String; // Variavel de uso geral

   ConsiderarReservados: Boolean = true; //utilizado para indicar nas telas de pesquisa
                                         //se deverá ser retornado os registros reservados para o sistema
                                         //cada tela de pesquisa deverá implementar a sua regra de validação
   ReservadosExcecao: string;            //lista de códigos dos registros da tela de pesquisa que serão retornados
                                         //independentemente de serem reservados ao sistema

   VendoDetalhe        :Boolean;
   DocDetalhe          :String;

   MascConta           :String;  // Usada pelo Fluxo Financeiro
   EmpresaUsaCC        :Boolean; // Usada pelo Fluxo Financeiro
   UsaContaRevisao     :Boolean; // Usada pelo Fluxo Financeiro
   ContaRevisao        :String;  // Usada pelo Fluxo Financeiro
   
   { Variaveis de relatório }

   ImpressoraRelatorio :Integer;
   DestinoRelatorio    :TDestinoRelatorio;
   ModoRelatorio       :TModoRelatorio;
   ImprimeSemSpoolWin  :boolean;          // usado para controle na impressao direta quando
                                          // a impressora for uma lq-570+
   { Variaveis de integração }

   IntegradoCompraPagar,
   IntegradoPagarCaixa,
   IntegradoPagarBanco,
   IntegradoVendaCaixa,
   IntegradoVendaReceber,
   IntegradoReceberCaixa,
   IntegradoReceberBanco,
   IntegradoBancoCaixa   :Boolean;
   FSetObrigatorio: Boolean;

   // Lincenças para a empresa
   Licenca    :integer;

   { Variaveis Cxa }

   DataCaixa      :TDateTime;
   CaixaTrabalho  :String;
   DescricaoCaixa :String;

   { Variaveis Ban }

   DataBanco       :TDateTime;
   CodigoConta     :String;
   DefinicaoCheque :String;

   { Variaveis Rec }

   ControlaCredito,
   PermiteCPFDuplicado,
   TrabalhaComIndice,
   DesbloqueioAutomaticoCh :Boolean;
   PortadorPadrao          :String;
   PortadorVenda           :String;
   TipoJuro                :String;
   Juros                   :Currency;
   Multa                   :Currency;
   Desconto                :Currency;
   DiaMaxDesconto          :Byte;
   DescontoCliente         :Currency;
   DiaMaxDescontoCliente   :Byte;
   JurosCliente            :Currency;

   { Variaveis Pag }

    TipoDocEntrada      :String;

   { Variaveis Ven }


   PortaSerial          :Integer;     // Porta com p/ ECF(Zanthus)
   rMaxDescBaixa        :Real;        // Máximo de desconto na baixa de documentos...
   sIDConsumidor        :string;      // Cliente consumidor
   sUsaRegistradora,
   sUsaDeptoCrediario,
   sEmiteBoleto,
   sEmiteCarne,
   sEmiteDuplicata,
   sEmiteNotaPromissoria,
   sEmiteNotaPromissoriaDoc,
   sUsaDiaUtil,
   sEmiteContrato,
   sEmitePedido,
   sEmitePedidoRecibo,
   sEmitePedidoOS,
   sEmiteOSDetalhe,
   sOrdemDetalheNaOS,
   sEmiteEntregaFutura,
   sEmiteRecibo,
   sEmiteComprovante,
   sEmiteListaSeparacao,
   sEmiteNotaFiscal,
   sEmiteCupomFiscal,
   sImprimeCheque,
   sEmiteCupomVinculado,
   sConfirmaBoleto,
   sConfirmaCarne,
   sConfirmaDuplicata,
   sConfirmaPromissoria,
   sConfirmaContrato,
   sConfirmaPedido,
   sConfirmaPedidoRecibo,
   sConfirmaPedidoOS,
   sConfirmaOSDetalhe,
   sConfirmaOrcamento,
   sConfirmaCondicional,
   sConfirmaEntrega,
   sConfirmaSeparacao,
   sConfirmaRecibo,
   sConfirmaComprovante,

   sBoletoNaVenda,
   sCarneNaVenda,
   sDuplicataNaVenda,
   sPromissoriaNaVenda,
   sContratoNaVenda,
   sPedidoNaVenda,
   sPedidoNaOS,
   sReciboNaVenda,
   sComprovanteNaVenda,
   sUsaAgenteVendas,
   sUsaPesquisaDiretaProdutos,

   sListaNaVenda,

   sGeraContrapartidaCtaRec,
   sGeraContrapartidaCheque,
   sGeraContrapartidaCtaPag :Boolean;

   sQuantidadeMaximaItem: Integer;
   iTempoEspera: Integer;
   bFecharEnvioSefazAutomatico: Boolean;

   sCaptorTCPIP: string;
   sCaptorPorta: Integer;

   lReciboReduzido                  :Boolean;
   lSelecionaImpPedidoNormal        :Boolean;
   lSelecionaImpOS: Boolean;
   lSelecionaImpPedidoPrazo         :Boolean;
   lSelecionaImpOrcamentoNormal     :Boolean;
   lSelecionaImpCondicionalNormal   :Boolean;
   lSelecionaImpEntregaMercadoria   :Boolean;
   lSelecionaImpRecibo              :Boolean;
   lSelecionaImpBoleto              :Boolean;
   sModeloImpRecibo                 :String;
   PaginaPedidoOrcamento            :string;

   { Variáveis Credsat}
//   DiretorioCredSat       :String;
//   AnalisaCredito         :Boolean;

   { Variaveis Vendas }
   sAlmoxPadrao                :String;
   sUsaAlmox                   :Boolean;
   AlmoxVenda                  :String;
   RegistradoraTrabalho        :String;
   PeriodoTrabalho             :Integer;
   RegistradoraAgrupamentoSemFinanceiro: string;
   DescricaoRegistradoraAgrupamentoSemFinanceiro: string;
   PeriodoAgrupamentoSemFinanceiro: Integer;
   UsuarioPadraoRegistradoraAgrupamentoSemFinanceiro: string;
   IdRegistro                  :Integer;
   RegistradoraPossuiECF       :Boolean;
   AlmoxTrabalho               :String;
   PeriodoNome                 :String;
   PeriodoData                 :TDate;
   ExisteExcessaoDeadLock      :Boolean;
   MonitorarTempoSelects       :boolean;
   MenuRibbonAtivo             :boolean; 
   VendaPosto                  :Boolean; { Verdadeiro se tiver usando a venda de posto }
   IDEstacao                   :Integer; { Numero identificador da estacao }
   IdSessao                    :Integer;
   PesquisaIncremental         :Boolean; { Se digitacao da pesquisa na venda vai procurando o item a cada tecla digitada }
   NotaServicoSeparada         :Boolean; { Se nota fiscal de servico é emitida em separado }
   MudaVendedorOrcamento       :Boolean; { Quando for orcamento se pode ou não mudar o vendedor }
   LiberaRateioComissao        :Boolean; { Se o total de rateio de comissao na ordem de servico pode ou não utrapassar 100% }
   SugereCemPorcento           :Boolean; { Se o sistema sugere 100% de rateio para cada mecanico }
   EditaPrecoVendaPosto        :Boolean; { Se edita ou nao o preco do produto na venda posto }
   ControleTempo               :Integer; { Utilizado no timer para controlar o tempo do usuario na nova ordem de servico    }
   SolicitaSenhaVendedor       :Boolean; { Se solicita ou nao a senha do vendedor na ordem de servico                       }
   SolicitaSenhaAlterarVendedor:Boolean;
   PercentualIssQn             :Extended;
   RequisicaoObrigatoria       :Boolean;
   GeraLogRegistroVenda        :Boolean;

   EstaFechandoTodasJanelas    :Boolean;

   IsImportadora: Boolean;
   PercDescImportadora: Currency;
   ProdutosInImportadora: string;
   NFSeEco: TNFSeEco;

   sFavoritosDoUsuario : array[0..9] of integer;

   //-> ATENÇÃO: Especifico para o controle da lista de objetos na unitPisCofins
   FListaGrupoTribPisConfis: TObjectList;
   FListaGrupoTribPisCofinsNatureza: TObjectList;
   FListaGrupoTribPisCofinsDadosProduto: TObjectList;

   FStreamConfiguracaoAdicional: TStrings;
   FIniMemoFile: TMemInifile;
   FIdUsuarioConfiguracaoPost, FIdUsuarioConfiguracaoReplicacao: Integer;
   fVersaoBanco: string;
   fVersaoBancoInt: Integer;

   const
        sTipoFreteContaEmitente: Integer = 0;
        sTipoFreteContaDestinatario: Integer = 1;
        sTipoFreteContaTerceiro: Integer = 2;
        sTipoFreteProprioRemetente: Integer = 3;
        sTipoFreteProprioDestinatario: Integer = 4;
        sTipoFreteSemCobranca: Integer = 9;
        Aspa = Chr($27);
        {Tabela de Origens de mercadorias}
        TributacaoOrigem :Array[0..8] of String = ('0-Nacional',
                                                   '1-Estrangeira - importação direta',
                                                   '2-Estrangeira - adquirido no mercado interno',
                                                   '3-Nacional, mercadoria ou bem com conteúdo de importação superior a 40% e inferior ou igual a 70%',
                                                   '4-Nacional, cuja produção tenha sido feita em conformidade com os processos produtivos básicos de que tratam as legislações citadas nos Ajustes',
                                                   '5-Nacional, mercadoria ou bem com conteúdo de importação inferior ou igual a 40%',
                                                   '6-Estrangeira - Importação direta, sem similar nacional, constante em lista da CAMEX',
                                                   '7-Estrangeira - Adquirida no mercado interno, sem similar nacional, constante em lista da CAMEX',
                                                   '8-Nacional, mercadoria ou bem com conteúdo de importação superior a 70%');

        TributacaoTipo   :Array[0..10] of String = ('00-Tributada integralmente',
                                                    '10-Tributada e com cobrança do ICMS por substituição tributária',
                                                    '20-Com redução na base de cálculo',
                                                    '30-Isenta ou não tributada e c/ cobrança do ICMS p/ subst. trib',
                                                    '40-Isenta',
                                                    '41-Não tributada',
                                                    '50-Suspensão',
                                                    '51-Diferimento',
                                                    '60-ICMS cobrado anteriormente por subs. tributária',
                                                    '70-Com redução de base de cálculo e cobrança sob ICMS p/ subst. trib',
                                                    '90-OUTROS');

        Limpa        = #27#64;                                   // inicializa a impressora
        lgExpand     = #27#15;                                   // liga expandido
        dgExpand     = #27#14;                                   // desliga expandido
        lgNegrito    = #27#71;                                   // liga negrito
        dgNegrito    = #27#72;                                   // desliga negrito
        lgEnfatizado = #27#69;                                   // liga enfatizado
        dgEnfatizado = #27#70;                                   // desliga enfatizado
        lgCondensado = chr($1B)+chr($0F);                        // seleciona modo condensado
        dgCondensado = chr($12);                                 // desliga  modo condensado
        lgOitavos    = #27#48;                                   // liga oitavos
        lgSextos     = #27#50;                                   // liga sextos
        lg10CPP      = chr($1B)+chr($50)+chr($12);               //  80 colunas
        lg12CPP      = chr($1B)+chr($4D)+chr($12);               //  96 colunas
        lg17CPP      = chr($1B)+chr($50)+chr($0F);               // 136 colunas
        lg20CPP      = chr($1B)+chr($4D)+chr($0F);               // 160 colunas
        lgDRAFT      = chr($1B)+chr($78)+chr($00);               // liga modo qualidade rascunho
        lgNLQ        = chr($1B)+chr($78)+chr($01);               // liga modo qualidade Carta

        AttribMap : array[1..15] of TRecAtributo =
            ((Nome: '#A,10CPP'; Attrib: lg10CPP     ),
             (Nome: '#A,12CPP'; Attrib: lg12CPP     ),
             (Nome: '#A,17CPP'; Attrib: lg17CPP     ),
             (Nome: '#A,20CPP'; Attrib: lg20CPP     ),
             (Nome: '#A,LGENF'; Attrib: lgEnfatizado),
             (Nome: '#A,DGENF'; Attrib: dgEnfatizado),
             (Nome: '#A,LGEXP'; Attrib: lgExpand    ),
             (Nome: '#A,DGEXP'; Attrib: lgExpand    ),
             (Nome: '#A,LGNEG'; Attrib: lgNegrito   ),
             (Nome: '#A,DGNEG'; Attrib: dgNegrito   ),
             (Nome: '#A,LGDRF'; Attrib: lgDRAFT     ),
             (Nome: '#A,LGNLQ'; Attrib: lgNLQ       ),
             (Nome: '#A,LGOIT'; Attrib: lgOitavos   ),
             (Nome: '#A,LGSEX'; Attrib: lgSextos    ),
             (Nome: '#A,LIMPA'; Attrib: Limpa       ));

        TipoMovimentoCheque :Array[0..34] of String =
                                  {00} ('                                                            ',
                                  {01}  'RECEBIDO NO CONTAS A RECEBER                                ',
                                  {02}  'DADO COMO TROCO NO CONTAS A RECEBER                         ',
                                  {03}  'RECEBIDO NA REGIST. REGISTRO DE PEDIDO                      ',
                                  {04}  'DADO COMO TROCO NA REGIST. REGISTRO DE PEDIDO               ',
                                  {05}  'UTILIZACAO DE VALOR DISPONIVEL DO CHEQUE NA REGISTRADORA    ',
                                  {06}  'DEVOLUCAO 1ª APRESENTACAO                                   ',
                                  {07}  'DEVOLUCAO 2ª APRESENTACAO                                   ',
                                  {08}  'UTILIZADO EM SAIDAS DA REGISTRADORA                         ',
                                  {09}  'TRANSF. P/ CAIXA (FECHAMENTO REGISTRADORA)                  ',
                                  {10}  'TRANSF. P/ CONTAS RECEBER (FECHAMENTO REGIST.)              ',
                                  {11}  'TRANSF. P/ CAIXA (SANGRIA)                                  ',
                                  {12}  'TRANSF. P/ CONTAS RECEBER (SANGRIA)                         ',
                                  {13}  'UTILIZACAO VALOR DISPONIVEL BAIXA DE DCTOS CONT. RECEBER    ',
                                  {14}  'ENTRADA ATRAVES DA DIGITACAO DE CHEQUE                      ',
                                  {15}  'TRANSFERENCIA PARA CAIXA TRABALHO                           ',
                                  {16}  'NEGOCIACAO C/ FACTORING                                     ',
                                  {17}  'ENTRADA ATRAVES DE LANCAMENTO DO CAIXA                      ',
                                  {18}  'SAIDA ATRAVES DE LANCAMENTO DO CAIXA                        ',
                                  {19}  'DEPOSITO                                                    ',
                                  {20}  'RECEBIDO COMO TROCO NO CONTAS A PAGAR                       ',
                                  {21}  'TRANSF. AUTOMATICA P/ CAIXA DE TRAB. ATRAVES CONTAS A PAGAR ',
                                  {22}  'UTILIZACAO NA BAIXA DE CONTAS A PAGAR                       ',
                                  {23}  'NEGOCIACAO C/ BANCO                                         ',
                                  {24}  'RECEBIDO NA QUITACAO DE CHEQUE                              ',
                                  {25}  'UTILIZACAO VALOR DISPONIVEL QUITACAO DE CHEQUE              ',
                                  {26}  'DADO COMO TROCO NA QUITACAO DE CHEQUE                       ',
                                  {27}  'RECEBIDO NA REGIST. BAIXA CONTAS A RECEBER                  ',
                                  {28}  'UTILIZACAO DE CHEQUE NA ANTECIPACAO DE PAGTO A FORNECEDOR   ',
                                  {29}  'QUITAÇÃO                                                    ',
                                  {30}  'TRANSFERENCIA ENTRE CAIXAS                                  ',
                                  {31}  'TRANSFERENCIA ENTRE EMPRESAS                                ',
                                  {32}  'QUITAÇÃO PARCIAL                                            ',
                                  {33}  'CHEQUE CUSTODIADO                                           ',
                                  {34}  'TRANSF EMPRESAS P/ CONTROLE DE CHEQUES                      ');

    cSecaoConectPost = 'CONECT POST';
    cSecaoEco360 = 'ECO 360';

   function CreateFormMensagem(const AMensagem: System.UnicodeString; TipoMsg: TTipoMsg; QtdeCaracterLinha: Integer = 0): TForm;
   function Mensagem(msg : string; TipoMsg : TTipoMsg; qtdeCaracterLinha : integer = 0):boolean; { Mostra mensagens em cx. de dialogo }
   function MensagemSize(msg : string; TipoMsg : TTipoMsg; width, height: Integer):boolean;
   function MensErro(Msg :String):Boolean; overload;
   function MensErro(Msg :String; UserCourier :boolean):Boolean; overload;

   function MensOK(Msg :String):Boolean;
   function ConfirmaSN(Msg :String):Boolean;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function MensagemList(msg : String; TipoMsg : TTipoMsg; StringList : TStringList; qtdeCaracteresLinha : Integer = 0):Boolean;
   {$ENDIF}

   function Replique(Caracter :String; Tamanho :Byte) :String; { Replica um caracter 'n' vezes }
   function StrZero(Conteudo :Variant; Tamanho :Byte) :String; { Preenche c/ 0 a esquerda }

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function CampoEmBranco(ForcaFoco :TWinControl) :Boolean; { Verifica se existe campos em branco }
   {$ENDIF}

   function UnicaCopia(Classe :TComponentClass) :Boolean; { Verifica se já existe uma copia do form a ser chamado }
   function FormCriado(fForm :TFormClass) :Boolean;
   function EditaCGCCPF(const CGCCPF: string; TipoPessoa: string): string;
   function ChecaCgcCpf(Aux:String):Boolean; { Verifica se CPF e CGC são válidos }

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   procedure LimpaCampoGrupo(Agrupador :TWinControl); { Limpa os campos de ediçao de um grupo }
   procedure Abilita(Agrupador :TWinControl; Tipo :Boolean; AbilitaLabel :Boolean); { Abilita/Desabilita grupo componente }
   procedure Habilita(Agrupador :TWinControl; Acao :Boolean); { Abilita/Desabilita grupo componente }
   procedure AbilitaComp(Componente :TComponent; Tipo :Boolean); { Abilita/Desabilita um componente }
   {$ENDIF}

   function TeclaPress(Tecla :Char) :Char; { Permite somente letra de A a Z}
   function TeclaNumPress(Tecla :Char) :Char; { Permite somente numero de 0 a 9}
   procedure ChamaPrograma(Classe :TComponentClass; Form :TForm); {Cria e mostra forms}
   function MesExtenso(Mes :Byte) :String; { Retorna o mes por extendo }
   function AnoExtenso(Ano: Word): string;
   function StrToExtended2Dec(Text: String):extended;
   function ValZerosLeft(Valor : Extended; TamanhoTotal, Decimais :Byte) :string; forward;
   function StrZerosDec(Valor: Extended; Decimais :integer) :string;
   function SQLValorDec(Valor: Extended; Decimais :integer) :string; forward;
   function QDataSQL(const Data: TDateTime):string; forward;
   function ValorNaFaixa(const Valor, ValorMinimo, ValorMaximo: Variant): Variant; forward;

   function TestaUsuario(TesUsuario :String): Boolean;                  // 16-01-03 Dirceu testa aexistencia de um usuario

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function SolicitaSenha(Tipo:Word):Boolean;
   {$ENDIF}

   function StrRight(ValueString :string; Tamanho :Integer) :string;

   function DiaExtenso(DataX :TDate) :String; overload; { Retorna o dia da semana por extendo }
   function DiaExtenso(Dia: Word): string; overload;
   function MascaraFloat(Valor :String) :Extended; { Mascara uma string de números no formato flutuante }
   function DiasNoMes(Mes, Ano :Word) :Word;  { Retorna o número de dias em um mes específico }
   function DiaMaximoNoMes(const Data: TDateTime): TDateTime; {volta a maior data no mes enviado}
   function UltimoDiaNoMes(const Data: TDateTime): TDateTime; {volta a maior data no mes enviado}
   function RetornaHoraMinuto(Tempo:Currency):String;
   function DesmascaraNumero(Value :String) :String; { Desmascara uma strinf númerica mantendo apenas o ponto decimal}
   function DesmascaraNumero1(Value :String) :String; { Desmascara uma strinf númerica mantendo apenas o ponto decimal}
   function MascaraVLR(Valor:Variant; Decimais:Byte=2) : Extended;
   function RemoveCaracter(ATarget, AString : String) : String; {Remove o caracter selecionado da string}

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   procedure MudaCor; { Muda a cor do label do componente EditBox}
   procedure MudaCorFrame(Sender :TObject);  { Muda a cor do frame }
   procedure SalvaPadraoRelatorio(Classe:TForm;Query:TSQLQuery;Usuario,Programa,sNome,sObs:String); overload;
   procedure SalvaPadraoRelatorio(Classe:TForm;Query:TFDQuery;Usuario,Programa,sNome,sObs:String); overload;

   procedure LePadraoDeRelatorio(Classe:TForm;Query:TSQLQuery;IDRelatorio:Integer);  overload;
   procedure LePadraoDeRelatorio(Classe:TForm;Query:TFDQuery;IDRelatorio:Integer);  overload;
   {$ENDIF}

   procedure LimpaLinhaStringGrid(StringGrid :TStringGrid; Linha :Integer);
   function DataExtenso(Data:TDate):String;
   procedure LimpaStringGrid(StringGrid :TStringGrid); { Limpa todas as células de um StringGrid}
   procedure LimpaStringGridPTL(StringGrid :TStringGrid; RowRemanescente:Integer); { Limpa todas as células de um StringGrid}
   procedure ExcluiLinhaStringGrid(StringGrid :TStringGrid); { Exclui a linha atual de uma StringGrid até ficar 2 linhas; cabec + 1 lin}
   procedure ExcluiLinhaStringGridPTL(StringGrid :TStringGrid; Linha:Integer);

   function BuildNotifyEvent(AData, ACode: Pointer): TNotifyEvent;
   function FecharTodasJanelas:Boolean; { Fecha todas as janelas MDI abertas } overload;
   function FecharTodasJanelas(const OnCompleted: TNotifyEvent; QtdForms: Integer = 1):Boolean; { Fecha todas as janelas MDI abertas } overload;
   function MascaraFone(const Fone : string) : string; { Mascara um telfone }
   function Alinha(Conteudo :String; Alinhamento :TAlinhamento; Tamanho :Integer) :String; { Alinha uma string }
   function Espaco(Tamanho :Byte) :String; { retornas espacos }
   function MascaraValor(Valor :Extended; Decimais :Byte) :String; { Mascara um valor retornando uma string }
   function DiaPascoa(Ano :Integer) :TDate; { Retorna a data da pascoa }
   function DiaSextaFeiraSanta(Ano :Integer) :TDate; { Retorna a data da sexta-feira santa }
   function DiaCarnaval(Ano :Integer) :TDate; { Retorna a data do caranaval }
   function ProximoDiaUtil(Data :TDate) :TDate; { Retorna o próximo dia util }
   function DiaFeriadoMunicipal(DataTeste :TDateTime; out msgRetorno: string) :boolean; {retorna se um data cadastrada é um feriado}

   function DataValida(Data :String) :TDateTime; { Retorna uma data valida }
   function CalculaJuros(Valor :Currency; Vencimento, Recebimento :TDate) :Currency; { Calcula juros }
   function CalculaMulta(Valor, Multa :Currency; Vencimento, Recebimento :TDate) :Currency; {Calcula multa}
   function CalculaJurosPag(Valor :Currency; Vencimento :TDate; Juros :Currency) :Currency; {Calcula juros para pagamento}

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function CalculaDesconto(Valor :Currency; Vencimento, Recebimento: TDate) :Currency; overload; { Calcula desconto }
   function CalculaDesconto(Valor, Percentual: Currency; Vencimento, Recebimento: TDate): Currency; overload;
   {$ENDIF}

   function ValorIndice(Indice :String; Data :TDateTime) :Extended; { Retorna valor do indice }
   function Arred(Valor :Extended) :Extended; { Arredonda valor }
   function Arredonda(Valor :Extended; Decimais:Integer;Truncar:Boolean=False) :Extended; { Arredonda valor dependendo no nº de decimais}
   function Trunca(Valor :Extended; Decimais :Integer) :Extended; { Trunca valores, ignorando decimais acima do especificado }
   function TruncTo(const AValue: Extended; const ADigit: TRoundToRange = -2): Extended;
   function ReadPrintIni(const Ident, Default: string): string;
   procedure WritePrintIni(const Ident, Default: string);
   function PortaImpressora(TipoImpressora :TImpressora): String;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function EAN13DV(EAN : string) : Char; { Retorna o digito verificado do EAN }
   function EAN14DV(EAN : string) : Char; { Retorna o digito verificado do EAN14 }
   function GeraDigitoEAN13(EAN : string) : String; { Insere o DV no EAN 13 }
   function GeraDigitoEAN13Grade(Produto : string; Sequencia: integer) : String;
   function TestaEAN13(EAN : String) : Boolean;
   function CalcularDVGtin(ACodigoGTIN: String): String;
   function ValidarGTIN(const ACodigoEAN: String; var AMensagemErro: String; const AGTINTributavel: Boolean = False): Boolean;
   function ValidarDigitoVerificadorGTIN(const ACodigoEAN: String; var ADigCalculado: String): Boolean;
   function ValidarPrefixoGTIN(const ACodigoEAN: String) :Boolean;
   function BetweenEco(const AValorVerificar, ValorInicial, ValorFinal: Integer): Boolean;
   function IsNumero(S : String) : Boolean;
   function ExisteEAN(EAN: String; QueryPesquisa: TSQLQuery): Boolean; overload;
   function ExisteEAN(EAN: String; QueryPesquisa: TFDQuery): Boolean; overload;
   function ExtraiVendaDoEAN(sEAN13: string): Extended;
   function IsProdutoPesado(sEAN13: string): boolean;
   {$ENDIF}

   function RetornaTipo(Codigo : String): String;
   function RetornaOrigem(Codigo : String): String;
   function GeraSequencia(Opcao: string; QueryPesquisa: TSQLQuery; AtualizaSequencia: Boolean = False): Integer; overload;
   function GeraSequencia(Opcao: string; QueryPesquisa: TFDQuery; AtualizaSequencia: Boolean = False): Integer; overload;

   procedure GravaSequencia(Opcao, Valor: string; QueryPesquisa: TSQLQuery);   overload;
   procedure GravaSequencia(Opcao, Valor: string; QueryPesquisa: TFDQuery);   overload;

   function ValidaHora(Texto : String): TTime;
   procedure PegaParamCC(QueryPesquisa : TSQLQuery); overload;
   procedure PegaParamCC(QueryPesquisa : TFDQuery); overload;
   function BancoFazLancamentoNoFluxoFinanceiro(QueryAux: TSQLQuery): boolean; overload;
   function BancoFazLancamentoNoFluxoFinanceiro(QueryAux: TFDQuery): boolean; overload;

   function MontaContaCC(sMascaraDaConta, sContaCC: string): string;

   function PegaAProximaSequenciaDoCaixa(var QueryPesquisa: TSQLQuery; caixa: string; DataCaixa:TDateTime): Integer;  overload;
   function PegaAProximaSequenciaDoCaixa(var QueryPesquisa: TFDQuery; caixa: string; DataCaixa:TDateTime): Integer;  overload;


   function PegaAProximaSequenciaDoBanco(var QueryPesquisa: TSQLQuery; Banco: string; DataBanco:TDateTime): Integer; overload;
   function PegaAProximaSequenciaDoBanco(var QueryPesquisa: TFDQuery; Banco: string; DataBanco:TDateTime): Integer; overload;

   procedure GeraImpressao2(Arq: String; TipoImp: TImpressora; Query1, Query2: TSimpleDataSet); overload;
   procedure GeraImpressao2(Arq: String; TipoImp: TImpressora; Query1, Query2: TFDMemTable); overload;
   procedure PrintCmd(Value: TPrintCmd; Const ImprimeDireto: Boolean=false);
   function  extenso(valor:extended; Tipo:Byte):string;
   procedure MontaExtensoCheques(const Valor: extended; var vExtenso1, vExtenso2: string);

   function  PreencheExtenso(Tam:Byte;Valor:Extended;Tipo:Byte):String;
   function  Justifica(Texto:String; LengthLines:Integer; LengthPriLine:Integer; LineReturn:Integer; JustUltLine:Boolean):String;
   function  JustificaTexto(Texto:String; TamLines:Integer; TabPrimLine:Integer):TStringList;
   function  Codificar(vsStr, vsCod :String) :String;
   function  IIF(Value:Boolean; Verdadeiro, Falso : Variant): Variant;
   function  Lower(Text:String): String;
   function  Upper(Text:String): String;
   function  BetWeen(Const Compare, Valor1, Valor2:Variant):Boolean;
   function  FindPrinter(const Impressora : String): Integer;
   function  MudaTamPapel(const Paper: TPapel): TPapel;
   function  TipoCmd(Comando: String): TPrintCmd;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   procedure ImprimeNF(oImp: TImprimeNota);
   procedure Verifica_TabSheet(Componente : TComponent);
   {$ENDIF}

   function  LeFavoritosDoUsuario : boolean;
   function  GravaFavoritosDoUsuario : boolean;
   function  TotalDeDiasPeriodo(DataInicio, DataFim :TDateTime;DiaInicial01:Boolean = true;DiaFinal31:Boolean = false) :Integer;
   function  VerificaVcto(Dia, Mes, Ano:Integer):TDateTime;

   function  Aspas(StringValue :string) : string;                         // 02-08-02 teio
   function  ZerosLeft(ValueString :string; Tamanho :Integer) :string;  overload;  // 24-08-02 teio
   function  ZerosLeft(const Valor :Variant; Tamanho :Integer = 5) :string; overload;
   function  QZerosLeft(const Valor :Variant; Tamanho :Integer = 5) :string;

   function  ParaOTamanho(ValueString :string; Tamanho :Integer) :string; // 24-08-02 teio
   function  EcoOpen(Query : TSQLQuery): boolean; overload                // 26-08-02 Triburtini - teio 20/06/2003
   function  EcoOpenOS(Query : TSQLQuery): boolean;
   function  EcoOpen(Query: TFDQuery): Boolean; overload;
   procedure EcoExecSql(Query : TSQLQuery); overload;                     // 26-08-02 Triburtini
   procedure EcoExecSql(AQuery: TFDQuery); overload;
   function  EcoClientOpen(ClientDataSet : TClientDataSet):Boolean;       // 26-08-02 Triburtini
   function  StrToInteiro(Value :String):Integer;                         // 15-10-02 Triburtini
   function  TestaAcesso(TesPrograma :String; TipoAcesso :TTipoAcesso = TmGeral ): Boolean;                   // 03-10-02 Dirceu
   function  DataMDA(Data: TDateTime) :string;
   function  DataDMA(Data: TDateTime) :string;
   function  DataTimeSQL(Data: TDateTime) :string;
   function  Substitui(TextoOld :String ; TextoNew: String; TextoPai: String) :String; // Substitui strings dentro de um texto
   function  CalculaPercentual(Valor1,Valor2:Extended):Extended;
   function  MontaGrupoEmpresas(sGrupo, sTabela, sComeco:String):String;

   function PesoPrincipal(Produto: string; TipoPeso: string): Real; overload;
   procedure PesoPrincipal(const Produto: string; var pesoBruto, pesoLiquido: Real); overload;

   { Cadastro de produto  :Triburtini }
   procedure CalculaMargemLucro(Var PrecoProduto:TPrecoProduto);
   procedure CalculaPrecoVenda(Var PrecoProduto:TPrecoProduto);
   procedure CalculaCustoFinal(Var PrecoProduto:TPrecoProduto);
   procedure CalculaIcmsCompra(Var PrecoProduto:TPrecoProduto);
   procedure CalculaIcmsGarantidoIntegral(Var PrecoProduto:TPrecoProduto);
   procedure CalculaIcmsVenda(Var PrecoProduto:TPrecoProduto;Opcao:Integer=1);
   function  CustoBaseCalculo(var PrecoProduto:TPrecoProduto):Currency;
   function  RetornaBaseCalculo(var PrecoProduto:TPrecoProduto):Currency;

   function ParamAtivo(const ecPar: TParmsEco): Variant; forward;
   function GetParmEco(Const ModuloEco: TModulosEco; Parametro:string; TipoCampo: TFieldType=ftBoolean): Variant; forward;
   function ParametroAtivo(Tabela, Parametro, Empresa :String; TipoCampo: TFieldType = ftBoolean): Variant; forward;
   function TrimDup(const Str :string): string;
   procedure SetPrinterPage1(Width, Height : LongInt);
   function MontaCodigoChegada(Cod, Dig: string): string;

   function NomeDeComputador : string;

   procedure ChangeDefaultPrinter(const Name: string);
   function GetDefaultPrinterName:string;

   function GetNotaFiscal(pEmpresa, pCliente, pDocumento, pTipo: String; pIdDocumento: Integer; pMascarar: Boolean = False): String;
   function GetNotaFiscalServico(pEmpresa, pCliente, pDocumento, pTipo: String; pIdDocumento: Integer): String;
   function DocumentoPossuiAgrupamento(pEmpresa, pDocumento: String): Boolean;

   function GetProximoCodigoLivre(pCampo, pTabela: String; pTamanho: Integer): String;
   function GetQuantidadeCartoesReservados : Integer;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function ChamaTelaDeExportar(NomeDaLista:String;NomeDaQuery:TSQLQuery):boolean;
   {$ENDIF}

   { Valores p/ TAG = 0-Confere campo em branco
                      1-Não confere campo em branco
                      2-Campo desabilitado
                      3-Campo desabilitado não muda cor do label}

   { Motivos de bloqueios padrao
                      01-Cadastro desatualizado
                      02-Documento em atrazo
                      03-Limite de credito Excedido}

   { MovimentoId - tabela de tipo de movimentacao do extrato do produto
                      01-Pedido de venda  (TVenPedido quando Status for EFE)
                      02-Entrada de mercadoria (TEstNfe)
                      03-Transferencia entre almoxarifados (TEstTransfAlmox)
                      04-Transferencia do produto principal p/ agregado (TEstTransfProd) }

   { Tipo venda - Campo tipo de venda do tvenproduto
                      'N' := 'Venda (Normal)';
                      'A' := 'Venda (Merc a retirar)';
                      'C' := 'Venda (Condicional)';
                      'O' := 'Orçamento';   }

   function StrToByte(sStr: string): byte;
   function  GeralTimeSQL(const Data: TDateTime):string; forward;
   function GeralQTimeSQL(const Data: TDateTime):string; forward;
   function GeralDataSQL(const Data: TDateTime):string; forward;
   function GeralQDataSQL(const Data: TDateTime):string; forward;
   function GeralHexChar(const c: Char): Byte; forward;
   function DataEstaNoPeriodo(DataINI, DataFim, Data: TDateTime): boolean; forward;
   procedure GravaArquivoDeLog(const Arquivo, StrLog : string); forward;
   function ValorEntre(ValorInicio,ValorFim,ValorPraTeste:Currency):Boolean; forward;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   procedure ImprimirRelatorioCrystal(const CaminhoRelatorio: string; const Parametros: array of variant; const EmVideo: Boolean = True;
                                      const EscolheImpressora: Boolean = True; const NomeImpressora: string = ''; const QuantidadeVias: Integer = 1;
                                      const OrientacaoPapel: string = 'R');

   procedure ChamaFormCrystal(Relatorio: string; Parametros: array of variant; EmVideo: Boolean=True;
                              EscolheImpressora: Boolean=True; NomeImpressora:String=''; QuantidadeVias: Integer=1;
                              OrientacaoPapel: String = 'R');

   {$ENDIF}

   procedure CreateODBCDriver;
   function  DescricaoDaGrade(Query:TFDQuery;sProduto:String):String;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   procedure ImprimeCheque(Numero: string; Banco: string; Valor: Double; Data: TDate; Cidade: string; Favorecido: string);
   {$ENDIF}

   function  DivideValor(Valor1, Valor2: Extended): Extended;

   function ValidaNCMDigitado(const NcmInformado : String):Boolean;

   function TipoTributacaoPeloCSF(CSF : String):Byte;

   {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
   function BuscaComponentTela(Agrupador : TWinControl; Nome : String) : TControl;
   {$ENDIF}

   //-> Agrupamento de pedidos registrados gerando um pedido sem financeiro
   procedure AbrePeriodoAgrupamentoSemFinanceiro(var Registradora: string; var Periodo: Integer);

   function GetSQLCancelamentoConcluido: string;
   function GetSQLItemPendenteExcluir: string;
   function GetSQLItemPendenteDenegado: string;
   function ExistePedidoCancelamentoNotaFiscalPendente(const AOrigem, AIdOrigem: Integer): Boolean; overload;
   function ExistePedidoCancelamentoNotaFiscalPendente(const ARegistradora: string; const AIdPeriodo: Integer): Boolean; overload;
   function ExistePedidoNotaFiscalPendente(const AOrigem, AIdOrigem: Integer): Boolean; overload;
   function ExistePedidoNotaFiscalPendente(const ARegistradora: string; const AIdPeriodo: Integer): Boolean; overload;
   function ExisteCancelamentoRegistroPendente(const AIdOrigem: Integer): Boolean; overload;
   function ExisteCancelamentoRegistroPendente(const ARegistradora: string; const AIdPeriodo: Integer): Boolean; overload;
   function ExisteCancelamentoConcluido(const AOrigem, AIdOrigem: Integer): Boolean;
   function ExisteItemPendenteExcluir(const AOrigem, AIdOrigem: Integer): Boolean;
   function ExisteEmailPendente(const AOrigem, AIdOrigem: Integer): Boolean;
   function ExisteReimpressaoPendente(const AOrigem, AIdOrigem: Integer): Boolean;
   function ExisteItemPendenteDenegado(const AOrigem, AIdOrigem: Integer): Boolean; overload;
   function ExisteItemPendenteDenegado(const ARegistradora: string; const AIdPeriodo: Integer): Boolean; overload;
   function CodigoCombustivelExigeTransportador(const ACodigoCombustivel: string): Boolean;

   function GetQuery: TSQLQuery;
   function GetFDQuery: TFDQuery;
   procedure GravaLogEco(Tela, Log: String; GravaLog: Boolean);
   function FileInUse(FileName: TFileName): Boolean;
   function impressoraMatricial(nomeImpressora : String) : Boolean;
   function DataAtual: TDateTime;
   function DataHoraAtual: TDateTime;
   function IsUrl(S: string): Boolean;
   function GetoFStreamConfiguracaoAdicional: TStrings;
   function GetoFMemoIniFileGerConfiguracao: TMemInifile;
   procedure GravarConfiguracaonoMemoFileConfigPostEco360(const aIpServidor, aPortaConexao, aUsuario, aSenha, aNome, aCaminhoPost, aLinkServido360: String);
   procedure CarregarConfiguracadoMemoFile(var aIpServidor, aPortaConexao, aUsuario, aSenha, aNomeBanco, aCaminhoPasta, aLinkServidor360: String);
   procedure GravarConfiguracoesAdicionais(const aIpServidor, aPortaConexao, aUsuario, aSenha, aNome, aCaminhoPost, aLinkServido360: String; const aIdTipoServidor: TTipoServidor; const aIdConfiguracao: Integer = 0);
   procedure CarregarConfiguracoesAdicionais(var aIpServidor, aPortaConexao, aUsuario, aSenha, aNomeBanco, aCaminhoPasta, aLinkServidor360: String; var aIdConfiguracao: Integer);
   function ValidaCompatibilidadeVersoes(const aValidarSomenteVersaoRelease: Boolean = True): Boolean;
   function ValidarLoginEmpresaPrincipal(const AValidar: Boolean): Boolean;
   function PermitirUtilizarFuncao(const MostrarMensagem: Boolean = True): Boolean;
   function TemEmpresaPrincipal: Boolean;
   procedure BuscarAlmoxPadrao;
   function VerificacoesPlaca(Placa: string): Boolean;
   function ValidarCNPJ(numCNPJ: string): boolean;
   function SomenteNumeros(const S: String): String;
   procedure AtivarDesativarLogChangesDataSet(ACDS: TClientDataSet; const AAtivar: Boolean = True);
var
  MDICloseForm: procedure(Form: TForm);

implementation

uses
  Utilitarios, DataModule, dtMensagem,uEcoFDTransaction
  {$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
  , CampoBranco, ven601rf, ExportarLista,dtMensagem_Nova,
  PreviewCrystal, PesquisaRelatorioCrystal, dtMensagemList, dtMensagemList_Nova,
  UFAlerta,EcoMenu, cxDropDownEdit
  {$ENDIF};

function FileInUse(FileName: TFileName): Boolean;
var
   HFileRes: HFILE;
begin
   Result := False;

   if not FileExists(FileName) then
      Exit;

   HFileRes := CreateFile(PChar(FileName), GENERIC_READ or GENERIC_WRITE, 0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
   Result := (HFileRes = INVALID_HANDLE_VALUE);

   if not Result then
      CloseHandle(HFileRes);
end;

procedure GravaLogEco(Tela, Log: String; GravaLog: Boolean);
var
   FNomeArqLog:String;
   Arq: TextFile;
   PathEco: string;
begin
   if not GravaLog then
      Exit;

   PathEco := ExtractFilePath(Application.ExeName) + 'EcoLog\';

   try
      if not DirectoryExists(PathEco) then
         ForceDirectories(PathEco);

      FNomeArqLog := PathEco + 'LogVerificacaoECO_' + FormatDateTime('DDMMYYYY',Date)+'.txt';

      AssignFile(Arq, FNomeArqLog);
      if not FileExists(FNomeArqLog) then
      begin
         Rewrite(arq);
      end else
      begin
         if not FileInUse(FNomeArqLog) then
         begin
            Reset(Arq);
            Append(Arq);

            try
               Log := FormatDateTime('HH:MM:SS = ', Now) + Tela + ':' + Log;
               Writeln(Arq, Log);
            finally
               CloseFile(Arq);
            end;
         end;
      end;
   except
   end;
end;

function impressoraMatricial(nomeImpressora : String) : Boolean;
var
  Device,
  Driver,
  Port : array[0..255] of char;
  HdeviceMode: Thandle;

  resolucao: Integer;
  impressora : TPrinter;

  possuiLetraDraft : Boolean;
  maiorQue16Cores : Boolean;
begin
  {ESTE MÉTODO IDENTIFICA SE A IMPRESSORA RECEBIDA NO PARÂMETRO É MATRICIAL OU NÃO}

  //CRIA IMPRESSORA TEMPORARIA
  impressora := TPrinter.Create;

  //CARREGA A IMPRESSORA PASSADA NO PARAMETRO
  impressora.printerindex := FindPrinter(nomeImpressora);

  //CARREGA AS PROPRIEDADES DA IMPRESSORA
  impressora.getprinter(Device, Driver, Port, HdeviceMode);

  //CARREGA RESOLUCAO DA IMPRESSORA
  resolucao := GetDeviceCaps(impressora.Handle, ASPECTX);

  //VERIFICA SE A IMPRESSORA POSSUI LETRA TIPO "DRAFT"
  if impressora.Fonts.IndexOf('Draft 10cpi') >= 0 then
    possuiLetraDraft := True
  else
    possuiLetraDraft := False;

  //VERIFICA SE A IMPRESSORA TEM RESOLUCAO MAIOR QUE 16 CORES
  if (GetDeviceCaps(impressora.Handle,NUMCOLORS)<=16) and (DeviceCapabilities(device, port, DC_TRUETYPE,nil,nil) = DCTT_BITMAP) then
    maiorQue16Cores := False
  else
    maiorQue16Cores := True;

  //APÓS VERIFICAÇÕES ANTERIORES DETERMINA SE É UMA IMPRESSORA MATRICIAL OU NÃO
  if (resolucao <= 144) and (possuiLetraDraft) and (not maiorQue16Cores) then
    Result := True
  else
    Result := False;

  //DESALOCA DA MEMÓRIA A IMPRESSORA TEMPORARIA
  FreeAndNil(impressora);

end;

function GetDefaultPrinterName : string;
begin
   if(Printer.PrinterIndex >= 0)then begin
      Result := Printer.Printers[Printer.PrinterIndex];
   end else begin
      {Retorna uma string vazia quando não encontra nenhuma impressora}
      Result := '';
   end;
end;

function GetProximoCodigoLivre(pCampo, pTabela: String; pTamanho: Integer): String;
var q: TSQLQuery;
   anterior, atual: Integer;
   maxCodigo: Integer;

   function GetMaxCodigo(pTamnho: Integer): Integer;
   var temp: string;
       I: integer;
   begin
      temp := '';
      for I := 0 to pTamanho - 1 do begin
         temp := temp + '9';
      end;
      Result := StrToInt(temp);
   end;

begin
   Result := '';
   anterior := 0;
   atual := 0;
   q := TSQLQuery.Create(nil);
   q.SQLConnection := FDataModule.sConnection;
   try
      q.Close;
      q.SQL.Clear;
      maxCodigo := GetMaxCodigo(pTamanho);
      q.SQL.Add('SELECT '+pCampo+' FROM '+pTabela+' ORDER BY '+pCampo+' ;');
      if EcoOpen(q) then begin
         while not q.Eof do begin
            atual := q.FieldByName(pCampo).AsInteger;
            if atual - anterior >= 2 then begin
               Result := ZerosLeft(IntToStr(anterior + 1), pTamanho);
               Exit;
            end;
            anterior := atual;
            q.Next;
         end;
      end else begin
         Result := ZerosLeft('1', pTamanho);
         Exit;
      end;
      if atual < maxCodigo then begin
         Result := ZerosLeft(atual + 1, pTamanho);
      end;
   finally
      q.Free;
   end;   
end;

function GetQuantidadeCartoesReservados : Integer;
var
   query: TSQLQuery;
begin
   Result := 0;

   query := TSQLQuery.Create(nil);
   try
      query.SQLConnection := FDataModule.sConnection;

      query.SQL.Clear;
      query.SQL.Add('select coalesce(count(codigo),0) as qtd from TRECTIPODOCUMENTO where cartao = '+QuotedStr('S'));

      if EcoOpen(query) then
         Result := query.Fields[0].AsInteger;
   finally
     Query.Close;
     FreeAndNil(query);
   end;
end;

function GetNotaFiscal(pEmpresa, pCliente, pDocumento, pTipo: string; pIdDocumento: Integer; pMascarar: Boolean = False): String;
var
   q: TSQLQuery;
   nf: string;
   nfTemp: string;
   nfCandidata: string;
   fEmpresa, fCliente, fDocumento, fTipo: string;
   fIdDocumento: Integer;
   PedidosOrigemList, NFUnicasLocalizadas, DocumentosAgrupados: TStringList;
   I: Integer;

   function MascaraNFCe(pNFCe, pSerieNFCe: string): string;
   begin
      Result := ZerosLeft(pNFCe, 9);
      Result := Result + '-' + ZerosLeft(pSerieNFCe, 3);
      Result := Result + '-' + 'NFCe';
   end;

   function MascaraNFe(pNotaNFe, pSerieNFe: string): string;
   begin
      Result := ZerosLeft(pNotaNFe, 9);
      Result := Result + '-' + ZerosLeft(pSerieNFe, 3);
      Result := Result + '-' + 'NFe';
   end;

   function NFJaExiste(pNF: string): Boolean;
   var
      index: Integer;
   begin
      Result := False;

      for index := 0 to NFUnicasLocalizadas.Count - 1 do
      begin
         if NFUnicasLocalizadas[index] = pNF then
         begin
            Result := True;
            Break;
         end;
      end;
   end;

   function GetPedidosAPartirDaRenegociacao(pEmpresa: string; pIdDocumento:
      Integer): TStringList;
   begin
      Result := TStringList.Create;
      Result.Clear;

      q.SQL.Clear;
      q.SQL.Add('SELECT DISTINCT PAR.DOCUMENTO                              ');
      q.SQL.Add('FROM TRECPARCELA PAR                                       ');
      q.SQL.Add('INNER JOIN TRecDocumento DOC ON (Doc.Empresa = Par.Empresa ');
      q.SQL.Add('  AND  Doc.Cliente = Par.Cliente                           ');
      q.SQL.Add('  AND  Doc.Tipo = Par.Tipo                                 ');
      q.SQL.Add('  AND  Doc.Documento = Par.Documento)                      ');
      q.SQL.Add('WHERE PAR.EMPRESA = :Empresa                               ');
      q.SQL.Add('   AND PAR.IDRENEGOCIACAO = :IDRENEGOCIACAO                ');
      q.SQL.Add('   AND Doc.Origem = ''VEN''                                ');
      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('IdRenegociacao').AsInteger := pIdDocumento;
      if EcoOpen(q) then
      begin
         while not q.Eof do
         begin
            Result.Add(q.FieldByName('Documento').AsString);
            q.Next;
         end;
         q.Close;
      end;
   end;

   function GetPedidosRenegociacao(pEmpresa, pCliente, pDocumento, pTipo: string;
      pIdDocumento: Integer): TStringList;
   begin
      Result := TStringList.Create;
      Result.Clear;
      q.SQL.Clear;
      q.SQL.Add('SELECT PAR.DOCUMENTO                      ');
      q.SQL.Add('FROM TRECPARCELA PAR                      ');
      q.SQL.Add('WHERE PAR.EMPRESA = :Empresa              ');
      q.SQL.Add('  AND PAR.CLIENTE = :Cliente              ');
      q.SQL.Add('  AND PAR.DOCUMENTO = :Documento          ');
      q.SQL.Add('  AND PAR.TIPO = :Tipo                    ');
      q.SQL.Add('  AND PAR.IDRENEGOCIACAO = :IdRenegociacao');
      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Cliente').AsString := pCliente;
      q.ParamByName('Documento').AsString := pDocumento;
      q.ParamByName('Tipo').AsString := pTipo;
      q.ParamByName('IdRenegociacao').AsInteger := pIdDocumento;
      if EcoOpen(q) then
      begin
         while not q.Eof do
         begin
            Result.Add(q.FieldByName('Documento').AsString);
            q.Next;
         end;

         q.Close;
      end;
   end;

   function GetPedidoAgrupado(pEmpresa, pDocumento: string): string;
   begin
      Result := '';
      q.SQL.Clear;
      q.SQL.Add('SELECT PED.PEDIDOAGRUPAMENTO ');
      q.SQL.Add('FROM TVENPEDIDO PED          ');
      q.SQL.Add('WHERE PED.EMPRESA = :Empresa ');
      q.SQL.Add('  AND PED.CODIGO = :Pedido   ');
      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Pedido').AsString := pDocumento;
      if EcoOpen(q) then
      begin
         Result := q.FieldByName('PedidoAgrupamento').AsString;
         q.Close;
      end;
   end;

   function IsTipoDocBoleto(pTipo: string): Boolean;
   begin
      Result := (pTipo = TipoDocBoleto);
   end;

   function GetDocumentosAgrupados(pEmpresa, pCliente, pDocumento, pTipo: string):
      TStringList;
   begin
      Result := TStringList.Create;
      Result.Clear;
      if (not IsTipoDocBoleto(pTipo)) then
         Exit;

      q.SQL.Clear;
      q.SQL.Add('SELECT PAR.DOCUMENTO           ');
      q.SQL.Add('FROM TRECPARCELA PAR           ');
      q.SQL.Add('WHERE PAR.EMPRESA = :Empresa   ');
      q.SQL.Add(' AND PAR.CLIENTE = :Cliente    ');
      q.SQL.Add(' AND PAR.IDBOLETO = :Documento ');
      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Cliente').AsString := pCliente;
      q.ParamByName('Documento').AsString := pDocumento;
      if EcoOpen(q) then
      begin
         while not q.Eof do
         begin
            Result.Add(q.FieldByName('Documento').AsString);
            q.Next;
         end;
         q.Close;
      end;
   end;

   function GetNFFromPedidoAgrupado(pEmpresa, pDocumento: string): string;
   var
      PedAgrupado: string;
   begin
      Result := '';

      PedAgrupado := GetPedidoAgrupado(pEmpresa, pDocumento);
      if (Trim(PedAgrupado) = '') then
         Exit;

      q.SQL.Clear;
      q.SQL.Add('SELECT PED.NUMERONFCE,         ');
      q.SQL.Add('       PED.SERIENFCE,          ');
      q.SQL.Add('       PED.NOTANFE,            ');
      q.SQL.Add('       PED.SERIENFE,           ');
      q.SQL.Add('       PED.NOTAFISCAL          ');
      q.SQL.Add('FROM TVENPEDIDO PED            ');
      q.SQL.Add('WHERE PED.EMPRESA = :Empresa   ');
      q.SQL.Add('  AND PED.CODIGO = :PedAgrupado');
      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('PedAgrupado').AsString := PedAgrupado;

      if EcoOpen(q) then
      begin
         if (not q.FieldByName('NUMERONFCE').IsNull) and (q.FieldByName('NUMERONFCE').AsInteger
            > 0) then
         begin
            if pMascarar then
            begin
               Result := MascaraNFCe(IntToStr(q.FieldByName('NUMERONFCE').AsInteger),
                  IntToStr(q.FieldByName('SERIENFCE').AsInteger));
            end
            else
            begin
               Result := IntToStr(q.FieldByName('NUMERONFCE').AsInteger);
            end;
         end
         else if (not q.FieldByName('NOTANFE').IsNull) and (q.FieldByName('NOTANFE').AsInteger
            > 0) then
         begin
            if pMascarar then
            begin
               Result := MascaraNFe(IntToStr(q.FieldByName('NOTANFE').AsInteger),
                  IntToStr(q.FieldByName('SERIENFE').AsInteger));
            end
            else
            begin
               Result := IntToStr(q.FieldByName('NOTANFE').AsInteger);
            end
         end
         else if (not q.FieldByName('NOTAFISCAL').IsNull) and (q.FieldByName('NOTAFISCAL').AsString
            <> '') then
         begin
            Result := q.FieldByName('NOTAFISCAL').AsString;
         end;

         q.Close;
      end;
   end;

   function GetNFFromPedido(pEmpresa, pDocumento: string): string;
   begin
      Result := '';
      q.SQL.Clear;
      q.SQL.Add('SELECT                       ');
      q.SQL.Add('PED.NUMERONFCE,              ');
      q.SQL.Add('PED.SERIENFCE,               ');
      q.SQL.Add('PED.NOTANFE,                 ');
      q.SQL.Add('PED.SERIENFE,                ');
      q.SQL.Add('PED.NOTAFISCAL               ');
      q.SQL.Add('FROM TVENPEDIDO PED          ');
      q.SQL.Add('WHERE PED.EMPRESA = :Empresa ');
      q.SQL.Add('  AND PED.CODIGO = :Codigo   ');
      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Codigo').AsString := pDocumento;

      if EcoOpen(q) then
      begin
         if (not q.FieldByName('NUMERONFCE').IsNull) and
            (q.FieldByName('NUMERONFCE').AsInteger > 0) then
         begin
            if pMascarar then
            begin
               Result := MascaraNFCe(IntToStr(q.FieldByName('NUMERONFCE').AsInteger),
                  IntToStr(q.FieldByName('SERIENFCE').AsInteger));
            end
            else
            begin
               Result := IntToStr(q.FieldByName('NUMERONFCE').AsInteger);
            end;
         end
         else if (not q.FieldByName('NOTANFE').IsNull) and (q.FieldByName('NOTANFE').AsInteger
            > 0) then
         begin
            if pMascarar then
            begin
               Result := MascaraNFe(IntToStr(q.FieldByName('NOTANFE').AsInteger),
                  IntToStr(q.FieldByName('SERIENFE').AsInteger));
            end
            else
            begin
               Result := IntToStr(q.FieldByName('NOTANFE').AsInteger);
            end;
         end
         else if (not q.FieldByName('NOTAFISCAL').IsNull) and (q.FieldByName('NOTAFISCAL').AsString
            <> '') then
         begin
            Result := q.FieldByName('NOTAFISCAL').AsString;
         end;

         q.Close;
      end;
   end;


   //busca dados do documento e atribui para as variáveis:
   //fEmpresa, fCliente, fTipo, fDocumento, fIdDocumento
   procedure CarregaDadosDoDocumento(pDocumento: string);
   begin
      q.SQL.Clear;
      q.SQL.Add('SELECT PAR.DOCUMENTO,                                               ');
      q.SQL.Add('       PAR.EMPRESA,                                                 ');
      q.SQL.Add('       PAR.CLIENTE,                                                 ');
      q.SQL.Add('       PAR.TIPO,                                                    ');
      q.SQL.Add('       DOC.IDDOCUMENTO                                              ');
      q.SQL.Add('FROM TRECPARCELA PAR                                                ');
      q.SQL.Add('     INNER JOIN TRECDOCUMENTO DOC ON (DOC.EMPRESA = PAR.EMPRESA     ');
      q.SQL.Add('                                  AND DOC.DOCUMENTO = PAR.DOCUMENTO ');
      q.SQL.Add('                                  AND DOC.CLIENTE = PAR.CLIENTE     ');
      q.SQL.Add('                                  AND DOC.TIPO = PAR.TIPO)          ');
      q.SQL.Add('WHERE PAR.DOCUMENTO = :Documento                                    ');
      q.SQL.Add('  AND PAR.EMPRESA = :Empresa                                        ');
      q.ParamByName('Documento').AsString := pDocumento;
      q.ParamByName('Empresa').AsString := Empresa;

      if EcoOpen(q) then
      begin
         fDocumento := q.FieldByName('Documento').AsString;
         fEmpresa := q.FieldByName('Empresa').AsString;
         fCliente := q.FieldByName('Cliente').AsString;
         fTipo := q.FieldByName('Tipo').AsString;
         fIdDocumento := q.FieldByName('IdDocumento').AsInteger;

         q.Close;
      end;
   end;
begin
   Result := '';

   if (pEmpresa = '') or (pDocumento = '') or (pCliente = '') or (pTipo = '') then
      Exit;

   q := TSQLQuery.Create(nil);
   q.SQLConnection := FDataModule.sConnection;
   NFUnicasLocalizadas := TStringList.Create();
   try
      {Busca o número da nota fiscal ou da nota fiscal agrupada pela ordem de precedência abaixo}

      {Verifica se o documento foi originado de um agrupamento}
      DocumentosAgrupados := GetDocumentosAgrupados(pEmpresa, pCliente,
         pDocumento, pTipo);
      if DocumentosAgrupados.Count > 0 then
      begin
         for I := 0 to DocumentosAgrupados.Count - 1 do
         begin
            CarregaDadosDoDocumento(DocumentosAgrupados[I]);
            nfCandidata := GetNotaFiscal(fEmpresa, fCliente, fDocumento, fTipo,
               fIdDocumento, pMascarar);
            if Pos(',', Result) > 0 then
            begin
               if ((Pos(nfCandidata + ',', Result) = 0) and (Pos(', ' +
                  nfCandidata, Result) = 0)) then
               begin
                  if (Trim(nfCandidata) <> '') and (Trim(Result) <> '') then
                  begin
                     Result := Result + ', ';
                  end;
                  Result := Result + nfCandidata;
               end;
            end
            else
            begin
               if (Pos(nfCandidata, Result) = 0) then
               begin
                  if (Trim(nfCandidata) <> '') and (Trim(Result) <> '') then
                  begin
                     Result := Result + ', ';
                  end;
                  Result := Result + nfCandidata;
               end;
            end;
         end;
      end;
      {Verifica se o documento foi originado de uma renegociacao}
      if (Copy(pDocumento, 1, 1) = 'R') then
      begin
         PedidosOrigemList := GetPedidosAPartirDaRenegociacao(pEmpresa, pIdDocumento);
      end
      else
      begin
         PedidosOrigemList := GetPedidosRenegociacao(pEmpresa, pCliente,
            pDocumento, pTipo, pIdDocumento);
      end;

      if PedidosOrigemList.Count > 0 then
      begin
         {Busca todas as NF dos documentos origem da renegociação}
         {Os resultados são concatenados na variável nf}
         for I := 0 to PedidosOrigemList.Count - 1 do
         begin
            nfTemp := GetNFFromPedido(pEmpresa, PedidosOrigemList[I]);
            if nfTemp = '' then
            begin
               nfTemp := GetNFFromPedidoAgrupado(pEmpresa, PedidosOrigemList[I]);
            end;
            if not NFJaExiste(nfTemp) then
            begin
               NFUnicasLocalizadas.Add(nfTemp);
               if nf <> '' then
               begin
                  nf := nf + ', ';
               end;
               nf := nf + nfTemp;
            end;
         end;
      end
      else
      begin
         {Busca pelos campos NotaNfe ou NotaFiscal ou NfAgrupada a partir do documento}
         nf := GetNFFromPedido(pEmpresa, pDocumento);
         if nf = '' then
         begin
            nf := GetNFFromPedidoAgrupado(pEmpresa, pDocumento);
         end;
      end;

      if ((Result <> '') and (nf <> '')) then
      begin
         Result := Result + ', ';
      end;

      Result := Result + nf;
   finally
      FreeAndNil(PedidosOrigemList);
      FreeAndNil(DocumentosAgrupados);
      FreeAndNil(NFUnicasLocalizadas);
      FreeAndNil(q);
   end;
end;

function DocumentoPossuiAgrupamento(pEmpresa, pDocumento: String): Boolean;
var q: TSQLQuery;
begin
   Result := False;
   q := TSQLQuery.Create(nil);
   q.SQLConnection := FDataModule.sConnection;
   try
      q.Close;
      q.SQL.Clear;
      q.SQL.Add(  'SELECT COUNT(*) AS QTD_AGRUPAMENTO                                  ' +
                  'FROM TRECPARCELA PAR                                                ' +
                  '     INNER JOIN TRECDOCUMENTO DOC ON (DOC.EMPRESA = PAR.EMPRESA     ' +
                  '                                  AND DOC.DOCUMENTO = PAR.DOCUMENTO ' +
                  '                                  AND DOC.CLIENTE = PAR.CLIENTE     ' +
                  '                                  AND DOC.TIPO = PAR.TIPO)          ' +
                  'WHERE PAR.IDBOLETO = :Documento                                     ' +
                  '  AND PAR.EMPRESA = :Empresa                                        ');
      q.ParamByName('Documento').AsString := pDocumento;
      q.ParamByName('Empresa').AsString := pEmpresa;
      if EcoOpen(q) then begin
         Result := iif(q.FieldByName('Qtd_Agrupamento').AsInteger > 0, True, false);
      end;
   finally
      q.Close;
   end;
end;

function GetNotaFiscalServico(pEmpresa, pCliente, pDocumento, pTipo: String; pIdDocumento: Integer): String;
var
   q: TSQLQuery;
   nf: string;
   nfTemp: string;
   nfCandidata: string;
   fEmpresa,fCliente, fDocumento, fTipo: string;
   fIdDocumento: Integer;
   PedidosOrigemList, NFUnicasLocalizadas, DocumentosAgrupados: TStringList;
   I: Integer;

   function NFJaExiste(pNF:String): Boolean;
   var index: Integer;
   begin
      Result := False;
      for index := 0 to NFUnicasLocalizadas.Count - 1 do begin
         if NFUnicasLocalizadas[index] = pNF then begin
            Result := True;
            Break;
         end;
      end;
   end;

   function GetPedidosRenegociacao(pEmpresa, pCliente, pDocumento, pTipo: String; pIdDocumento: Integer): TStringList;
   begin
      Result := TStringList.Create;
      Result.Clear;

      q.SQL.Clear;
      q.SQL.Add('SELECT PAR.DOCUMENTO                      ' +
                'FROM TRECPARCELA PAR                      ' +
                'WHERE PAR.EMPRESA = :Empresa              ' +
                '  AND PAR.CLIENTE = :Cliente              ' +
                '  AND PAR.DOCUMENTO = :Documento          ' +
                '  AND PAR.TIPO = :Tipo                    ' +
                '  AND PAR.IDRENEGOCIACAO = :IdRenegociacao');

      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Cliente').AsString := pCliente;
      q.ParamByName('Documento').AsString := pDocumento;
      q.ParamByName('Tipo').AsString := pTipo;
      q.ParamByName('IdRenegociacao').AsInteger := pIdDocumento;

      if EcoOpen(q) then begin
         while not q.Eof do begin
            Result.Add(q.FieldByName('Documento').AsString);
            q.Next;
         end;

         q.Close;
      end;
   end;

   function GetPedidosAPartirDaRenegociacao(pEmpresa: string; pIdDocumento: Integer): TStringList;
   begin
      Result := TStringList.Create;
      Result.Clear;

      q.SQL.Clear;
      q.SQL.Add('SELECT DISTINCT PAR.DOCUMENTO ' +
                '  FROM TRECPARCELA PAR ' +
                '       INNER JOIN TRecDocumento DOC ON (Doc.Empresa = Par.Empresa ' +
                '                                   AND  Doc.Cliente = Par.Cliente ' +
                '                                   AND  Doc.Tipo = Par.Tipo ' +
                '                                   AND  Doc.Documento = Par.Documento) ' +
                ' WHERE PAR.EMPRESA = :Empresa ' +
                '   AND PAR.IDRENEGOCIACAO = :IDRENEGOCIACAO ' +
                '   AND Doc.Origem = ''VEN''');

      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('IdRenegociacao').AsInteger := pIdDocumento;

      if EcoOpen(q) then begin
         while not q.Eof do begin
            Result.Add(q.FieldByName('Documento').AsString);
            q.Next;
         end;

         q.Close;
      end;
   end;

   function GetPedidoAgrupado(pEmpresa, pDocumento: String): String;
   begin
      Result := '';
      
      q.SQL.Clear;
      q.SQL.Add('SELECT PED.PEDIDOAGRUPAMENTO   ' +
                'FROM TVENPEDIDO PED            ' +
                'WHERE PED.EMPRESA = :Empresa   ' +
                '  AND PED.CODIGO = :Pedido     ');

      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Pedido').AsString := pDocumento;

      if EcoOpen(q) then begin
         Result := q.FieldByName('PedidoAgrupamento').AsString;
         q.Close;
      end;
   end;

   function GetNFServicoFromPedidoAgrupado(pEmpresa, pDocumento: String): string;
   var
      PedAgrupado: string;
   begin
      Result := '';

      PedAgrupado := GetPedidoAgrupado(pEmpresa, pDocumento);
      if (Trim(PedAgrupado) = '') then
         Exit;

      q.SQL.Clear;
      q.SQL.Add('SELECT PED.NOTAFISCALSERVICO, PED.NFAGRUPADASERVICO ' +
                'FROM TVENPEDIDO PED                                 ' +
                'WHERE PED.EMPRESA = :Empresa                        ' +
                '  AND PED.CODIGO = :PedAgrupado                     ');

      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('PedAgrupado').AsString := PedAgrupado;

      if EcoOpen(q) then begin
         if (not q.FieldByName('NOTAFISCALSERVICO').IsNull) and (q.FieldByName('NOTAFISCALSERVICO').AsString <> '') then begin
            Result := q.FieldByName('NOTAFISCALSERVICO').AsString;
         end else if (not q.FieldByName('NFAGRUPADASERVICO').IsNull) and (q.FieldByName('NFAGRUPADASERVICO').AsString <> '') then begin
            Result := q.FieldByName('NFAGRUPADASERVICO').AsString;
         end;

         q.Close;
      end;
   end;

   function GetNFServicoFromPedido(pEmpresa, pDocumento: String): string;
   begin
      Result := '';

      q.SQL.Clear;
      q.SQL.Add('SELECT                       ' +
                'PED.NOTAFISCALSERVICO,       ' +
                'PED.NFAGRUPADASERVICO        ' +
                'FROM TVENPEDIDO PED          ' +
                'WHERE PED.EMPRESA = :Empresa ' +
                '  AND PED.CODIGO = :Codigo   ');

      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Codigo').AsString := pDocumento;

      if EcoOpen(q) then begin
         if (not q.FieldByName('NOTAFISCALSERVICO').IsNull) and (q.FieldByName('NOTAFISCALSERVICO').AsString <> '') then begin
            Result := q.FieldByName('NOTAFISCALSERVICO').AsString;
         end else if (not q.FieldByName('NFAGRUPADASERVICO').IsNull) and (q.FieldByName('NFAGRUPADASERVICO').AsString <> '') then begin
            Result := q.FieldByName('NFAGRUPADASERVICO').AsString;
         end;
         
         q.Close;
      end;
   end;

   function IsTipoDocBoleto(pTipo: string): Boolean;
   begin
      Result := (pTipo = TipoDocBoleto);
   end;

   function GetDocumentosAgrupados(pEmpresa, pCliente, pDocumento, pTipo: string): TStringList;
   begin
      Result := TStringList.Create;
      Result.Clear;

      if (not IsTipoDocBoleto(pTipo)) then
         Exit;

      q.SQL.Clear;
      q.SQL.Add('SELECT PAR.DOCUMENTO           ' +
                'FROM TRECPARCELA PAR           ' +
                'WHERE PAR.EMPRESA = :Empresa   ' +
                '  AND PAR.CLIENTE = :Cliente   ' +
                '  AND PAR.IDBOLETO = :Documento');

      q.ParamByName('Empresa').AsString := pEmpresa;
      q.ParamByName('Cliente').AsString := pCliente;
      q.ParamByName('Documento').AsString := pDocumento;

      if EcoOpen(q) then begin
         while not q.Eof do begin
            Result.Add(q.FieldByName('Documento').AsString);
            q.Next;
         end;

         q.Close;
      end;
   end;

   //busca dados do documento e atribui para as variáveis:
   //fEmpresa, fCliente, fTipo, fDocumento, fIdDocumento
   procedure CarregaDadosDoDocumento(pDocumento: string);
   begin
      q.SQL.Clear;
      q.SQL.Add('SELECT PAR.DOCUMENTO,                                               ' +
                '       PAR.EMPRESA,                                                 ' +
                '       PAR.CLIENTE,                                                 ' +
                '       PAR.TIPO,                                                    ' +
                '       DOC.IDDOCUMENTO                                              ' +
                'FROM TRECPARCELA PAR                                                ' +
                '     INNER JOIN TRECDOCUMENTO DOC ON (DOC.EMPRESA = PAR.EMPRESA     ' +
                '                                  AND DOC.DOCUMENTO = PAR.DOCUMENTO ' +
                '                                  AND DOC.CLIENTE = PAR.CLIENTE     ' +
                '                                  AND DOC.TIPO = PAR.TIPO)          ' +
                'WHERE PAR.DOCUMENTO = :Documento                                    ' +
                '  AND PAR.EMPRESA = :Empresa                                        ');

      q.ParamByName('Documento').AsString := pDocumento;
      q.ParamByName('Empresa').AsString := Empresa;

      if EcoOpen(q) then begin
         fDocumento := q.FieldByName('Documento').AsString;
         fEmpresa := q.FieldByName('Empresa').AsString;
         fCliente := q.FieldByName('Cliente').AsString;
         fTipo := q.FieldByName('Tipo').AsString;
         fIdDocumento := q.FieldByName('IdDocumento').AsInteger;

         q.Close;
      end;
   end;
begin
   Result := '';

   if (pEmpresa = '') or (pDocumento = '') or (pCliente = '') or (pTipo = '') then
      Exit;

   q := TSQLQuery.Create(nil);
   q.SQLConnection := FDataModule.sConnection;
   NFUnicasLocalizadas := TStringList.Create();
   try
      {Verifica se o documento foi originado de um agrupamento}
      DocumentosAgrupados := GetDocumentosAgrupados(pEmpresa, pCliente, pDocumento, pTipo);
      if DocumentosAgrupados.Count > 0 then begin
         for I := 0 to DocumentosAgrupados.Count - 1 do begin
            CarregaDadosDoDocumento(DocumentosAgrupados[I]);
            nfCandidata := GetNotaFiscalServico(fEmpresa, fCliente, fDocumento, fTipo, fIdDocumento);
            if Pos(',', Result) > 0 then begin
               if ((Pos(nfCandidata+',', Result) = 0) and (Pos(', '+nfCandidata, Result) = 0)) then begin
                  if Result <> '' then begin
                     Result := Result +', ';
                  end;
                  Result := Result + nfCandidata;
               end;
            end else begin
               if (Pos(nfCandidata, Result) = 0) then begin
                  if Result <> '' then begin
                     Result := Result +', ';
                  end;
                  Result := Result + nfCandidata;
               end;
            end;
         end;
      end;

      {Verifica se o documento foi originado de uma renegociacao}
      if (Copy(pDocumento, 1, 1) = 'R') then begin
         PedidosOrigemList := GetPedidosAPartirDaRenegociacao(pEmpresa, pIdDocumento);
      end else begin
         PedidosOrigemList := GetPedidosRenegociacao(pEmpresa, pCliente, pDocumento, pTipo, pIdDocumento);
      end;

      if PedidosOrigemList.Count > 0 then begin
         {Busca todas as NF dos documentos origem da renegociação}
         {Os resultados são concatenados na variável nf}
         for I := 0 to PedidosOrigemList.Count - 1 do begin
            nfTemp := GetNFServicoFromPedido(pEmpresa, PedidosOrigemList[I]);
            if nfTemp = '' then begin
               nfTemp := GetNFServicoFromPedidoAgrupado(pEmpresa, PedidosOrigemList[I]);
            end;
            if not NFJaExiste(nfTemp) then begin
               NFUnicasLocalizadas.Add(nfTemp);
               if nf <> '' then begin
                  nf := nf + ', ';
               end;
               nf := nf + nfTemp;
            end;
         end;
      end else begin
         {Busca pelos campos NotaNfe ou NotaFiscal ou NfAgrupada a partir do documento}
         nf := GetNFServicoFromPedido(pEmpresa, pDocumento);
         if nf = '' then begin
            nf := GetNFServicoFromPedidoAgrupado(pEmpresa, pDocumento);
         end;
      end;
      if ((Result <> '') and (nf <> '')) then begin
         Result := Result + ', ';
      end;
      Result := Result + nf;
   finally
      FreeAndNil(PedidosOrigemList);
      FreeAndNil(DocumentosAgrupados);
      FreeAndNil(NFUnicasLocalizadas);
      FreeAndNil(q);
   end;
end;

procedure ChangeDefaultPrinter(const Name: string);
var
    W2KSDP: function(pszPrinter: PChar): Boolean; stdcall;
    H: THandle;
    Size, Dummy: Cardinal;
    PI: PPrinterInfo2;
begin
   if (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 5) then begin
      @W2KSDP := GetProcAddress(GetModuleHandle(winspl), 'SetDefaultPrinterA');
      if @W2KSDP = nil then RaiseLastOSError;
      if not W2KSDP(PChar(Name)) then RaiseLastOSError;
   end else begin
      if not OpenPrinter(PChar(Name), H, nil) then RaiseLastOSError;
      try
         GetPrinter(H, 2, nil, 0, @Size) ;
         if GetLastError <> ERROR_INSUFFICIENT_BUFFER then RaiseLastOSError;
         GetMem(PI, Size) ;
         try
            if not GetPrinter(H, 2, PI, Size, @Dummy) then RaiseLastOSError;
            PI^.Attributes := PI^.Attributes or PRINTER_ATTRIBUTE_DEFAULT;
            if not SetPrinter(H, 2, PI, PRINTER_CONTROL_SET_STATUS) then
               RaiseLastOSError;
         finally
            FreeMem(PI) ;
         end;
      finally
        ClosePrinter(H) ;
      end;
    end;
end;

function MensagemSize(msg : string; TipoMsg : TTipoMsg; width, height: Integer):boolean;
begin
   application.createform(TFDtMensagem, FDtMensagem);
  try
    fdtMensagem.Mensagem := Msg;
    fdtMensagem.TipoMsg  := TipoMsg;
    fdtMensagem.Width := width;
    fdtMensagem.Height := height;
    fdtMensagem.ShowModal;
    Result := fdtMensagem.Tag = 1;
  finally
    fdtMensagem.Free;
  end;

end;

function CreateFormMensagem(const AMensagem: System.UnicodeString; TipoMsg: TTipoMsg; QtdeCaracterLinha: Integer = 0): TForm;
begin
//{$IFDEF DEPURACAO}
//   Application.CreateForm(TFDtMensagem_Nova, FDtMensagem_Nova);
//
//   FDtMensagem_Nova.Mensagem := AMensagem;
//   FDtMensagem_Nova.TipoMsg := TipoMsg;
//
//   Result := FDtMensagem_Nova;
//{$ELSE}
   Application.CreateForm(TFDtMensagem, FDtMensagem);

   FDtMensagem.Mensagem := AMensagem;
   FDtMensagem.TipoMsg := TipoMsg;

   if (QtdeCaracterLinha > 0) then
    FDtMensagem.NumeroCaracteresPorLinha := QtdeCaracterLinha;

   Result := FDtMensagem;
//{$ENDIF}
end;

function Mensagem(Msg :String; TipoMsg: TTipoMsg; qtdeCaracterLinha : integer = 0):Boolean;
var
  Form: TForm;
begin
  Form := CreateFormMensagem(Msg, TipoMsg, qtdeCaracterLinha);
  try
    Form.ShowModal;
    Result := (Form.Tag = 1);
  finally
    Form.Free;
  end;
end;

function MensErro(Msg :String):Boolean;
begin
  Result := Mensagem(msg, tmErro);
end;

function MensErro(Msg :String; UserCourier :boolean):Boolean;
begin
   Result := False;
  Application.createform(TFDtMensagem, FDtMensagem);
  try
    fdtMensagem.Mensagem := Msg;
    fdtMensagem.TipoMsg  := tmErro;
    {$IFNDEF PAFECOBPL}
    fdtMensagem.RichEdit.Style.Font.Name := 'Courier';
    {$ENDIF}
    fdtMensagem.ShowModal;
  finally
    fdtMensagem.Free;
  end;
end;

function MensOK(Msg :String): boolean;
begin
   Result := False;
  Application.createform(TFDtMensagem, FDtMensagem);
  try
    fdtMensagem.Mensagem := Msg;
    fdtMensagem.TipoMsg  := tmInforma;
    {$IFNDEF PAFECOBPL}
    fdtMensagem.RichEdit.Style.Font.Name := 'Courier';
    {$ENDIF}
    fdtMensagem.ShowModal;
  finally
    fdtMensagem.Free;
  end;
end;

function ConfirmaSN(Msg :String):Boolean;
begin
  Result := Mensagem(msg, tmConfirma);
end;

function isInteiro(Value: Variant):boolean;
begin
  Result := VarType(Value) in [varSmallint, varInteger, varShortInt, varWord, varLongWord, varInt64, varByte];
end;

function isNumeric(Value: Variant):boolean;
begin
  Result := VarType(Value) in [varSingle, varDouble, varCurrency]
end;


function RetornaTipo(Codigo : String): String;
begin
   Result := '';
   if Copy(Codigo,2,2) = '00' then Result := TributacaoTipo[00];
   if Copy(Codigo,2,2) = '10' then Result := TributacaoTipo[01];
   if Copy(Codigo,2,2) = '20' then Result := TributacaoTipo[02];
   if Copy(Codigo,2,2) = '30' then Result := TributacaoTipo[03];
   if Copy(Codigo,2,2) = '40' then Result := TributacaoTipo[04];
   if Copy(Codigo,2,2) = '41' then Result := TributacaoTipo[05];
   if Copy(Codigo,2,2) = '50' then Result := TributacaoTipo[06];
   if Copy(Codigo,2,2) = '51' then Result := TributacaoTipo[07];
   if Copy(Codigo,2,2) = '60' then Result := TributacaoTipo[08];
   if Copy(Codigo,2,2) = '70' then Result := TributacaoTipo[09];
   //if Copy(Codigo,2,2) = '80' then Result := TributacaoTipo[10];
   if Copy(Codigo,2,2) = '90' then Result := TributacaoTipo[10];
end;

function ValidaHora(Texto : String): TTime;
Var Hora, Min, Seg : Word;
begin
   Hora := StrToIntDef(Copy(Texto, 0,2),0);
   Min  := StrToIntDef(Copy(Texto, 3,2),0);
   Seg  := StrToIntDef(Copy(Texto, 5,2),0);
   ValidaHora := StrToTime(StrZero(Hora,2)+':'+StrZero(Min,2)+':'+StrZero(Seg,2));
end;

// Retorna um valor entre apostrofos;
function Aspas(StringValue :String) :String;
begin
   result := Aspa+StringValue+Aspa;
end;

function StrRight(ValueString :string; Tamanho :Integer) :string;
var X :integer;
begin
   Result := Trim(ValueString);
   X      := Length(Result);
   if X > 0 then
     if X > Tamanho then
       Result := Copy(Result, 1, Tamanho)
     else
       if X < Tamanho then
         Result := StringOfChar(' ', (Tamanho - X)) + Result;
end;


// retorna um string com zeros a esquerda
function ZerosLeft(ValueString :string; Tamanho :Integer) :string;
var Aux :string;
    K   :integer;
begin
  Result := Trim(ValueString);
  Aux    := '';
  for K  := 1 to (Tamanho - Length(Result)) do
      Aux := Aux + '0';
  Result := Aux + Result;
end;

function ZerosLeft(const Valor :Variant; Tamanho :Integer = 5) :string; overload;
var K:Integer;
begin
  if VarType(Valor) in[varNull, varEmpty] then
     Result := '0';

  if (VarType(Valor) = varString) or (VarType(Valor) = varUString) then
     Result := ErtStrUtils.OnlyNumbers(Valor);

  if isInteiro(Valor) then
     Result := IntToStr(Valor);

  K := Length(Result);
  if K < Tamanho then
     Result := StringOfChar('0', Tamanho - K) + Result;
end;



function QZerosLeft(const Valor :Variant; Tamanho :Integer = 5) :string;
var K:Integer;
begin
  if VarType(Valor) in[varNull, varEmpty] then
     Result := '0';

  if (VarType(Valor) = varString) or (VarType(Valor) = varUString) then
     Result := ErtStrUtils.OnlyNumbers(Valor);

  if VarType(Valor) in [varInteger, varWord, varSmallInt] then
     Result := IntToStr(Valor);
  K := Length(Result);
  if K < Tamanho then
     Result := StringOfChar('0', Tamanho - K) + Result;
  Result := QuotedStr(Result);
end;


{ justifica um valor com zeros a esquerda sem o FormatSettings.DecimalSeparator }
function StrZerosDec(Valor: Extended; Decimais :integer) :string;
begin
  try
    Case Decimais of
      0: Result:= FormatFloat('##0',  Valor);
  -1, 1: Result:= FormatFloat('##0.0', Valor);
  -2, 2: Result:= FormatFloat('##0.00', Valor);
      3: Result:= FormatFloat('##0.000', Valor);
      4: Result:= FormatFloat('##0.0000', Valor);
      5: Result:= FormatFloat('##0.00000', Valor);
      6: Result:= FormatFloat('##0.000000', Valor);
      7: Result:= FormatFloat('##0.0000000', Valor);
      8: Result:= FormatFloat('##0.00000000', Valor);
      9: Result:= FormatFloat('##0.000000000', Valor);
    end;
  except
    Result := '';
  end;
end;


// retorna um valor para usar em clausulas sql validas
//
// ex : 12.129,99 -> 12129.99
//         9,9978 ->     9.9978
//         0,99   ->     0.99
//         12,00  ->    12.00

function SQLValorDec(Valor: Extended; Decimais :integer) :string;
begin
  Result := StrZerosDec(Valor, Decimais);
  if Result = '' then
     Result := '0';
  Result := AnsiReplaceText(Result, ',', '.');
end;

function QDataSQL(const Data: TDateTime): string;
begin
  try
    Result := QuotedStr(FormatDateTime('dd.mm.yyyy',Data));
  except
    Result := 'Null';
  end;
end;

function ValorNaFaixa(const Valor, ValorMinimo, ValorMaximo: Variant): Variant;
begin
  Result := Valor;
  if Valor < ValorMinimo then
    Result := ValorMinimo
  else
    if Valor > ValorMaximo then
      Result := ValorMaximo;
end;


function  StrToInteiro(Value :String):Integer;
begin
   if Value='' then begin
      Result := 0;
      Exit;
   end;
   try
      Result := StrToInt(Value);
   except
      Result := 0;
   end;
end;

{ retorna uma string com o espacos a direita limitada ao tamanho }
function ParaOTamanho(ValueString :string; Tamanho :Integer) :string;
var Aux  :string;
    K    :integer;
begin
  Result := Trim(ValueString);
  Aux    := '';
  for K  := 1 to (Tamanho - Length(Result)) do
      Aux := Aux + ' ';
  Result := Copy((Result+Aux), 1, Tamanho);
end;


function RetornaOrigem(Codigo : String) :String;
begin
   Result := '';
   if Length(Codigo)>0 then
   begin
      if CharInSet(Codigo[1], ['0'..'8']) then
         Result := TributacaoOrigem[StrToInt(Codigo[1])];
   end;

end;

function Replique(Caracter :String; Tamanho :Byte) :String;
var
   Contador :Byte;
   Retorno  :String;
begin
   for Contador := 1 to Tamanho do begin
      Retorno := Retorno + Caracter;
   end;
   Replique := Retorno;
end;

function StrZero(Conteudo :Variant; Tamanho :Byte) :String;
var ConteudoAux :String;
Begin
   ConteudoAux := String(Conteudo);
   if Tamanho < Length(ConteudoAux) then begin
      Tamanho := Length(ConteudoAux);
   end;
   StrZero := Replique('0', (Tamanho - Length(ConteudoAux))) + ConteudoAux;
end;

function StrToExtended2Dec(Text: String):extended;
var x : integer;
begin
  result := 0;
  if Text = '' then exit;
  Text := ErtStrUtils.OnlyNumbers(Text);
  X := Length(Text);
  try
    Result := StrToFloat(Copy(Text, 1, X-2) + FormatSettings.DecimalSeparator + Copy(Text,  X-1, 2));
  except
    result := 0;
  end;
end;

{ justifica um valor com zeros a esquerda sem o FormatSettings.DecimalSeparator }
function ValZerosLeft(Valor : Extended; TamanhoTotal, Decimais :Byte) :string;
begin
  Result:=ZerosLeft(ErtStrUtils.OnlyNumbers(StrZerosDec(Valor, Decimais)), TamanhoTotal);
end;

function DesmascaraNumero(Value : string):string;
var K : byte;
begin
  result := '';
  for K := 1 to Length(value) do
     if CharInSet(value[k], ['0'..'9']) or (Value[k] = ',')then
        result := result + value[k];
end;

// copiado de 'Utilitarios.DesmascaraNumero' - nao permite repeticoes dos sinais ".,-"
function DesmascaraNumero1(Value : string):string;
var K   : byte;
    ch  : char;
    Aux : string;

    function TestaSinal(oChar: char): string;
    begin
      TestaSinal := '';
      if Pos(oChar, Aux)=0 then
        TestaSinal := oChar;
    end;

begin
  Aux := '';
  for K := 1 to Length(Value) do begin
    ch := Value[k];
    case ch of
      '0'..'9' : Aux := Aux + Ch;
      '.'      : Aux := Aux + TestaSinal(FormatSettings.DecimalSeparator);
      ','      : Aux := Aux + TestaSinal(FormatSettings.DecimalSeparator);
      '-'      : Aux := Aux + TestaSinal('-');
    end
  end;
  if Aux = '' then Aux := '0';
  Result := Aux;
end;

function MascaraVLR(Valor:Variant; Decimais:Byte=2) : Extended;
begin
   Result := 0.00;
   if (VarType(Valor) = varString) or (VarType(Valor) = varUString) then begin
      if Valor = '' then Exit;
      Result := StrToFloat(DesmascaraNumero(Valor))
   end else
      if VarType(Valor) in [varSingle, varDouble, varCurrency] then
         Result := StrToFloat(DesmascaraNumero(FloatToStr(Valor)))
      else
         if VarType(Valor) in [varSmallint, varInteger, varShortInt,varByte, varWord, varLongWord, varInt64] then
            Result := StrToFloat(DesmascaraNumero(FloatToStr(Valor)))
end;

function RemoveCaracter(ATarget, AString :String): String;{Remove o caracter selecionado da string}
var
  LTmp, LCodigoTmp: string;
  LPosEspace : Integer;
begin
  LPosEspace    := Pos(ATarget, AString);
  LCodigoTmp := AString;
  if Length(AString) >= 1 then
  begin
    LTmp := Copy(AString, 0, 1);
    if LPosEspace > 0 then
    begin
      while (LPosEspace > 0) do
      begin
         LCodigoTmp := Copy(LCodigoTmp, 0, LPosEspace - 1) + Copy(LCodigoTmp, LPosEspace + 1, Length(LCodigoTmp));
         LPosEspace := Pos(ATarget, LCodigoTmp);
      end;
    end;
  end;
  Result := LCodigoTmp;
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
function CampoEmBranco(ForcaFoco :TWinControl) :Boolean;
var
   Contador   :Integer;
   Componente :Tcomponent;
   HaFoco     :boolean;
   QtdComp    :Integer;

   function LocalizaFocusControl(NomedoEdit :TComponent):String;
   var bx       : integer;
       comp     : TComponent;
       TemLabel : Boolean;
   begin
     TemLabel := False;
     with Screen.ActiveForm do begin
        for bx := QtdComp downto 0 do begin
           Comp := Components[bx];
           if (Comp is TLabel) then begin
              TemLabel := True;
              if TLabel(Comp).FocusControl = TEdit(NomedoEdit) then begin
                 result := TLabel(comp).Caption;
                 Exit;
              end
              else begin
                 result := 'Falta focus control ' + TEdit(NomedoEdit).Name + ' - ' + (Comp).Name + '  ' + IntToStr(bx);
              end;
           end
           else if (Comp is TDnBox) then begin
              TemLabel := True;
              if TDnBox(Comp).FocusControl = TDnBox(NomedoEdit) then begin
                 result := TDnBox(comp).LabelText;
                 Exit;
              end
              else begin
                 result := 'Falta focus control ' + TDnBox(NomedoEdit).Name + ' - ' + (Comp).Name + '  ' + IntToStr(bx);
              end;
           end;
        end;
        if Not TemLabel then begin
           result := 'Não existe Label no form' + TEdit(NomedoEdit).Name;
        end;
     end;
  end;

  function TiraE(Texto :String):String;
  var
     TextoAux  :String;
     Contador :Integer;
     Caracter :String;
  begin
     Contador := 1;
     TextoAux := '';
     while Contador <= Length(Texto) do begin
        Caracter := Copy(Texto, Contador, 1);
        Inc(Contador);
        if Caracter = '&' then Continue;
        TextoAux := TextoAux + Caracter;
     end;
     Result := TextoAux;
  end;

begin
  CampoEmBranco := False;
  HaFoco := False;
  Application.CreateForm(TFCampoBranco, FCampoBranco);

  with Screen.ActiveForm do
  begin
    QtdComp := ComponentCount - 1;
    if ForcaFoco is TButton then
    begin
      if (ForcaFoco.CanFocus) then
        ForcaFoco.SetFocus;
    end;
    for Contador := 0 to QtdComp do
    begin
      Componente := Components[Contador];
      if (Componente.Tag <> 0) then
        Continue;
      if Componente is TEdit then
      begin
        if Trim(TEdit(Componente).Text) = '' then
        begin
          FCampoBranco.Memo.Lines.Add(TiraE(LocalizaFocusControl(Componente)));
          FCampoBranco.Memo.Lines.Add('');
          if not HaFoco then
          begin
            if TEdit(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TEdit(Componente).CanFocus then
                TEdit(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if Componente is TNumEdit then
      begin
        if (TNumEdit(Componente).Value = 0) and (TNumEdit(Componente).Enabled) then
        begin
          FCampoBranco.Memo.Lines.Add(TiraE(LocalizaFocusControl(Componente)));
          FCampoBranco.Memo.Lines.Add('');
          if not HaFoco then
          begin
            if TNumEdit(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if (TNumEdit(Componente).Parent.CanFocus) then
                TNumEdit(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if Componente is TDnEdit then
      begin
        if (Trim(TDnEdit(Componente).Text) = '') and (TDnEdit(Componente).Enabled) then
        begin
          FCampoBranco.Memo.Lines.Add(TiraE(LocalizaFocusControl(Componente)));
          FCampoBranco.Memo.Lines.Add('');
          if not HaFoco then
          begin
            if TEdit(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TEdit(Componente).CanFocus then
                TEdit(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if (Componente is TEditBox) and (TEditBox(Componente).Enabled) then
      begin
        if (Trim(TEditBox(Componente).Text) = '') then
        begin
          if TEditBox(Componente).Hint <> '' then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TEditBox(Componente).Hint));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end
          else if TiraE(TEditBox(Componente).LabelText) <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TEditBox(Componente).LabelText));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end;
          if not HaFoco then
          begin
            if TDnEdit(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TDnEdit(Componente).CanFocus then
                TDnEdit(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if (Componente is TSynEdit) and (TSynEdit(Componente).Enabled) then
      begin
        if Trim(TSynEdit(Componente).Lines.Text) = EmptyStr then
        begin
          if TSynEdit(Componente).Hint <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TSynEdit(Componente).Hint));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end
          else if TiraE(TSynEdit(Componente).Name) <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TSynEdit(Componente).Name));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end;
          if not HaFoco then
          begin
            if TSynEdit(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TSynEdit(Componente).CanFocus then
                TSynEdit(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if (Componente is TAdvLUEdit) and (TAdvLUEdit(Componente).Enabled) then
      begin
        if Trim(TAdvLUEdit(Componente).Text) = EmptyStr then
        begin
          if TAdvLUEdit(Componente).Hint <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TAdvLUEdit(Componente).Hint));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end
          else if TiraE(TAdvLUEdit(Componente).Name) <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TAdvLUEdit(Componente).Name));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end;
          if not HaFoco then
          begin
            if TAdvLUEdit(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TAdvLUEdit(Componente).CanFocus then
                TAdvLUEdit(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if (Componente is TAdvDateTimePicker) and (TAdvDateTimePicker(Componente).Enabled) then
      begin
        if ((TAdvDateTimePicker(Componente).Date) <= ZeroValue) or ((TAdvDateTimePicker(Componente).Date) = Null) then
        begin
          if TAdvDateTimePicker(Componente).Hint <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TAdvDateTimePicker(Componente).Hint));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end
          else if TiraE(TAdvDateTimePicker(Componente).Name) <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TAdvDateTimePicker(Componente).Name));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end;
          if not HaFoco then
          begin
            if TAdvDateTimePicker(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TAdvDateTimePicker(Componente).CanFocus then
                TAdvDateTimePicker(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
      else if (Componente is TcxCheckComboBox) and (TcxCheckComboBox(Componente).Enabled) then
      begin
        if Trim(TcxCheckComboBox(Componente).Text) = EmptyStr then
        begin
          if TcxCheckComboBox(Componente).Hint <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TcxCheckComboBox(Componente).Hint));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end
          else if TiraE(TcxCheckComboBox(Componente).Name) <> EmptyStr then
          begin
            FCampoBranco.Memo.Lines.Add(TiraE(TcxCheckComboBox(Componente).Name));
            FCampoBranco.Memo.Lines.Add(EmptyStr);
          end;
          if not HaFoco then
          begin
            if TcxCheckComboBox(Componente).Enabled then
            begin
              Verifica_TabSheet(Componente);

              if TcxCheckComboBox(Componente).CanFocus then
                TcxCheckComboBox(Componente).SetFocus;
            end;
            HaFoco := True;
          end;
          CampoEmBranco := True;
        end;
      end
    end;
  end;
  if FCampoBranco.Memo.Lines.Count > 1 then
  begin
    FCampoBranco.ShowModal;
  end;
  FCampoBranco.Free;
end;
{$ENDIF}

procedure SalvaLogDeTempo(Local: string; TempoTotal: Integer; sql: string);
var s, filename : string;
    Str : TStrings;
    Linhas : integer;
begin
  if not Assigned(Screen.ActiveForm) then exit;
  if Screen.FormCount <= 0 then exit;

  S := Screen.ActiveForm.ClassName;
  FileName := 'c:\Ecosis\' + S + '_times.txt';

  Str := TStringList.Create;
  if FileExists(Filename) then
    Str.LoadFromFile(FileName);

  Str.Insert(0, TrimDup(SQL));
  Str.Insert(0, StringOfChar('=' ,60));
  Str.Insert(0, S + ' Local: ' + Local + FormatDateTime(' dd/mm/yy ', Now) + ' Tempo..: ' + IntToStr(TempoTotal));
  Str.Insert(0, StringOfChar('=' ,60));

  Linhas := str.Count - 1;
  while Linhas > 1000 do begin
     Str.Delete(1000);
     linhas := str.Count - 1;
  end;
  Str.SaveToFile(FileName);
  Str.Free;
end;


// retorna false se nao ha registros e true se ha registros lidos
function EcoOpen(Query : TSQLQuery): Boolean;
var
   Ini: TDateTime;
   changeCursorBack: Boolean;
begin
   Ini := Now;

   changeCursorBack := SetCursorBusy();

   try
      Result := ecoOpenOs(Query)
   finally
      SetCursorDefault(changeCursorBack);
   end;

   if MonitorarTempoSelects then
//      if Abs(MilliSecondsBetween(Ini, Now)) > 05 then // 100 - 0,1 segundo   {mudado para 05 - 0.005}
         SalvaLogDeTempo('Geral_EcoOpen', MilliSecondsBetween(Ini, Now), Query.Text);
end;

function EcoOpenOS(Query : TSQLQuery): Boolean;
begin
  Query.Close;
  Query.Open;
  Result := not (Query.Eof and Query.Bof);
end;

function  EcoOpen(Query: TFDQuery): Boolean;
var
   Ini: TDateTime;
   changeCursorBack: Boolean;
begin
   Ini := Now;

   changeCursorBack := SetCursorBusy();

   try
     Query.Close;
     Query.Open;
   finally
     SetCursorDefault(changeCursorBack);
   end;

   Result := not (Query.Bof and Query.Eof);

   if MonitorarTempoSelects then
      SalvaLogDeTempo('Geral_EcoOpen', MilliSecondsBetween(Ini, Now), Query.Sql.Text);
end;

procedure EcoExecSql(Query : TSQLQuery);
var
  Ini: TDateTime;
//  changeCursorBack: Boolean;
begin
  Ini := Now;

//  changeCursorBack := SetCursorBusy();

//  try
    Query.ExecSql;
//  finally
//    SetCursorDefault(changeCursorBack);
//  end;

  if MonitorarTempoSelects then
//     if Abs(MilliSecondsBetween(Ini, Now)) > 05 then  // 100 - 0,1 segundo   {mudado para 05 - 0.005}
        SalvaLogDeTempo('Geral_EcoExecSql', MilliSecondsBetween(Ini, Now), Query.Text);
end;

procedure EcoExecSql(AQuery: TFDQuery);
var
  Ini: TDateTime;
begin
  Ini := Now;
  AQuery.ExecSQL;

  if MonitorarTempoSelects then
     SalvaLogDeTempo('Geral_EcoExecSql', MilliSecondsBetween(Ini, Now), AQuery.Sql.Text);
end;

function EcoClientOpen(ClientDataSet : TClientDataSet):Boolean;
var
  Ini: TDateTime;
  changeCursorBack: Boolean;
begin
  Ini := Now;

  changeCursorBack := SetCursorBusy();

  try
    ClientDataSet.Close;
    ClientDataSet.Open;
  finally
    SetCursorDefault(changeCursorBack);
  end;

  Result := not(ClientDataSet.Bof and ClientDataSet.Eof);

  if MonitorarTempoSelects then
//     if Abs(MilliSecondsBetween(Ini, Now)) > 05 then // 100 - 0,1 segundo   {mudado para 05 - 0.005}
        SalvaLogDeTempo('Geral_EcoClientOpen', MilliSecondsBetween(Ini, Now), ClientDataSet.CommandText);
end;

{

 processo transferido para o PaiDeTodos, para evitar LimpaCampo(Self)

procedure LimpaCampo(Form :TForm=nil);
var
   Contador   :Integer;
   Componente :TComponent;
begin
   if not Assigned(Form) then Form := Screen.ActiveForm;
   with Form do begin
      for Contador := 0 to ComponentCount - 1 do begin
         Componente := Components[Contador];
         if Componente is TEdit then begin
            TEdit(Componente).Text := '';
         end
         else if Componente is TDnEdit then begin
            TDnEdit(Componente).Text    := '';
            TDnEdit(Componente).SqlText := '';
         end
         else if (Componente is TCheckBox) and (TCheckBox(Componente).tag = 0) then begin
            TCheckBox(Componente).Checked := False;
         end
         else if Componente is TEditBox then begin
            TEditBox(Componente).Text    := '';
            TEditBox(Componente).SqlText := '';
         end
         else if Componente is TMemo  then begin
            TMemo(Componente).Text    := '';
         end
// linux
         else if Componente is TNumEdit then begin
            TNumEdit(Componente).Value := 0;
         end
         else if Componente is TComboBox then begin
            TComboBox(Componente).Style     := csDropDown;
            TComboBox(Componente).ItemIndex := 0;
            TComboBox(Componente).Style     := csDropDownList;
         end;
      end;
   end;
end;
}

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure LimpaCampoGrupo(Agrupador :TWinControl);
var
   Contador   :Integer;
   Componente :TControl;
begin
   with Agrupador do begin
      for Contador := 0 to ControlCount - 1 do begin
         Componente := Controls[Contador];
         if Componente.Tag = 3 then Continue;

         if Componente is TDnBox then LimpaCampoGrupo(TDnBox(Componente));
         if Componente is TDnPanel then LimpaCampoGrupo(TDnPanel(Componente));
         

         if Componente is TEdit then begin
            TEdit(Componente).Text :=  '';
         end
         else if (Componente is TCheckBox) and (TCheckBox(Componente).tag = 0) then begin
            TCheckBox(Componente).Checked := False;
         end
         else if Componente is TDnEdit then begin
            TDnEdit(Componente).Text    := '';
            TDnEdit(Componente).SqlText := '';
         end
         else if Componente is TEditBox then begin
            TEditBox(Componente).Text    := '';
            TEditBox(Componente).SqlText := '';
         end
         else if Componente is TNumEdit then begin
            TNumEdit(Componente).Value := 0;
         end
         else if Componente is TComboBox then begin
            TComboBox(Componente).Style     := csDropDown;
            TComboBox(Componente).ItemIndex := 0;
            TComboBox(Componente).Style     := csDropDownList;
         end;
      end;
   end;
end;

procedure Abilita(Agrupador :TWinControl; Tipo :Boolean; AbilitaLabel :Boolean);
var
   Contador   :Integer;
   Componente :TControl;
begin
  with Agrupador do begin
      for Contador := 0 to (ControlCount - 1) do begin
         Componente := Controls[Contador];
         if (Componente.Tag = 2) or (Componente.Tag = 3) or (Componente.Tag >= 10) then Continue;
         if (Componente is TEdit) or (Componente is TDnEdit)  then begin
            TEdit(Componente).ReadOnly := not Tipo;
            TEdit(Componente).TabStop  := Tipo;
            if Tipo then begin
               TEdit(Componente).Color := clWhite;
            end
            else begin
               TEdit(Componente).Color := clBtnFace;
            end;
         end
         else if Componente is TDBMemo then begin
             TDBMemo(Componente).Enabled := Tipo;
             if Tipo then begin
                TDBMemo(Componente).Color := clWhite;
             end
             else begin
                TDBMemo(Componente).Color := $00EEEEEE;
             end;

         end
         else if Componente is TCheckListBox then begin
             TCheckListBox(Componente).Enabled := Tipo;
             if Tipo then begin
                TCheckListBox(Componente).Color := clWhite;
             end
             else begin
                TCheckListBox(Componente).Color := $00EEEEEE;
             end;
         end
         else if Componente is TSpeedButton then begin
             TSpeedButton(Componente).Enabled := Tipo;
         end
         else if Componente is TDBGrid then begin
             TDBGrid(Componente).Enabled := Tipo;
             if Tipo then begin
                TDBGrid(Componente).Color := clWhite;
             end
             else begin
                TDBGrid(Componente).Color := $00EEEEEE;
             end;

         end
         else if Componente is TDnBox then begin
             TDnBox(Componente).Enabled := Tipo;
         end
         else if Componente is TEditBox then begin
             TEditBox(Componente).Enabled := Tipo;
         end
         else if Componente is TRadioButton then begin
            TRadioButton(Componente).Enabled := Tipo;
         end
         else
         if Componente is TcxButton then
         begin
            TcxButton(Componente).Enabled := Tipo;
         end
         else
         if Componente is TcxPageControl then
         begin
            TcxPageControl(Componente).Enabled := Tipo;
         end
         else
         if Componente is TcxTabSheet then
         begin
            TcxTabSheet(Componente).Enabled := Tipo;
         end
         else
         if Componente is TMemo then
         begin
            TMemo(Componente).Enabled := Tipo;
            if Tipo then begin
               TMemo(Componente).Color := clWhite;
            end
            else begin
               TMemo(Componente).Color := $00EEEEEE;
            end;
         end
         else if Componente is TLabel then begin
            TLabel(Componente).Enabled := AbilitaLabel;
         end
         else if Componente is TCheckBox then begin
            TCheckBox(Componente).Enabled := Tipo;
         end
         else if Componente is TComboBox then begin
            TComboBox(Componente).Enabled := Tipo;
            if Tipo then begin
               TComboBox(Componente).Color := clWhite;
            end
            else begin
               TComboBox(Componente).Color := clBtnFace;
            end;
         end;
      end;
   end;
end;


procedure Habilita(Agrupador :TWinControl; Acao :Boolean);
var Contador   :Integer;
    Componente :TControl;
begin
  with Agrupador do begin
      for Contador := 0 to (ControlCount - 1) do begin
         Componente := Controls[Contador];
         if (Componente.Tag = 2) or (Componente.Tag = 3) or (Componente.Tag >= 10) then Continue;

         if (Componente is TEdit) or (Componente is TDnEdit)  then begin
            TEdit(Componente).ReadOnly := not Acao;
            TEdit(Componente).TabStop  := Acao;
            TEdit(Componente).Color    := iif(Acao, clWhite, clBtnFace);
         end
         else if Componente is TDBMemo then begin
             TDBMemo(Componente).Enabled := Acao;
             TDBMemo(Componente).Color := iif(Acao, clWhite, $00EEEEEE);
         end
         else if Componente is TCheckListBox then begin
             TCheckListBox(Componente).Enabled := Acao;
             TCheckListBox(Componente).Color   := iif(Acao, clWhite, $00EEEEEE);
         end
         else if Componente is TSpeedButton then begin
             TSpeedButton(Componente).Enabled := Acao;
         end
         else if Componente is TDBGrid then begin
             TDBGrid(Componente).Enabled := Acao;
             TDBGrid(Componente).Color   := iif(Acao, clWhite, $00EEEEEE);
         end
         else if Componente is TDnBox then begin
             TDnBox(Componente).Enabled := Acao;
         end
         else if Componente is TEditBox then begin
             TEditBox(Componente).Enabled := Acao;
         end
         else if Componente is TComboBox then begin
            TComboBox(Componente).Enabled := Acao;
            TComboBox(Componente).Color   := iif(Acao, clWhite, clBtnFace);
         end
         else if Componente is TBitBtn then begin
            TBitBtn(Componente).Enabled := Acao;
         end
         else if Componente is TcxPageControl then begin
            TcxPageControl(Componente).Enabled := Acao;
         end
         else if Componente is TcxTabSheet then begin
            TcxTabSheet(Componente).Enabled := Acao;
         end
         else if Componente is TcxButton then begin
            TcxButton(Componente).Enabled := Acao;
         end;

         { Agrupadores }
         if (Componente.ClassType = TPageControl)   then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TcxPageControl) then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TcxTabSheet)    then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TDnBox)         then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TPanel)         then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TDnPanel)       then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TTabSheet)      then Habilita(TWinControl(Componente), Acao);
         if (Componente.ClassType = TcxGroupBox)    then 
            Habilita(TWinControl(Componente), Acao);

      end;
   end;
end;

procedure AbilitaComp(Componente :TComponent; Tipo :Boolean);
begin
   if (Componente is TEdit) or (Componente is TDnEdit) then begin
      TEdit(Componente).ReadOnly := not Tipo;
      TEdit(Componente).TabStop  := Tipo;
      if Tipo then begin
         TEdit(Componente).Color := clWhite;
      end
      else begin
         TEdit(Componente).Color := clBtnFace;
      end;
   end
   else if Componente is TDnBox then begin
     TDnBox(Componente).Enabled := Tipo;
   end
   else if Componente is TEditBox then begin
       TEditBox(Componente).Enabled := Tipo;
   end
   else if Componente is TLabel then begin
      TEdit(Componente).Enabled := Tipo;
   end
   else if Componente is TComboBox then begin
      TComboBox(Componente).Enabled   := Tipo;
      if Tipo then begin
         TComboBox(Componente).Color := clWhite;
      end
      else begin
         TComboBox(Componente).Color := clBtnFace;
      end;
   end;
end;
{$ENDIF}

function UnicaCopia(Classe :TComponentClass) :Boolean;
var
   Contador :Integer;
begin
   UnicaCopia := True;
   Contador   := 0;
   while Contador < (Screen.FormCount)  do begin
      if Screen.Forms[Contador].ClassType = Classe then begin
         UnicaCopia := False;
         Screen.Forms[Contador].BringToFront;
         Exit;
      end;
      Inc(Contador);
   end;
end;

function FormCriado(fForm :TFormClass) :Boolean;
var I :Integer;
begin
   Result := False;
   for I := 0 to Screen.FormCount - 1 do begin
      if Screen.Forms[I].ClassType = fForm then begin
         Result := True;
         Exit;
      end;
   end;
end;

function EditaCGCCPF(const CGCCPF: string; TipoPessoa: string): string;
var Aux :string;
begin
  Result := '';
  Aux := ErtStrUtils.OnlyNumbers(CGCCPF);
  if (Uppercase(TipoPessoa) = 'F') or (Uppercase(TipoPessoa) = 'P') then
    Result := ErtStrUtils.MaskFormatText('999.999.999-99;0; ', Aux)
  else
    if Uppercase(TipoPessoa) = 'J' then
      Result := ErtStrUtils.MaskFormatText('99.999.999/9999-99;0; ', Aux);
end;

// retorna se um CNPJ/cpf digitado e válido
Function ChecaCgcCpf(Aux:String):Boolean;
Var Bx,I: Byte;
    Tot : real;
    dg1 : Char;
    dg2 : Char;

    function Mod11(xx:Real):Byte;
    begin
       result := Trunc(Tot - (int(Tot / 11) * 11));
       if result = 10 then result := 0;
    end;

    function Ordem(Letra:char):Byte;
    begin
       result := (ord(Letra)-48);
    end;

    function MontaCPF:Boolean;
    var k : byte;
    begin
      Tot := 0;
      for k := 1 to 9 do Tot := Tot + (ordem(aux[k]) * k);

      dg1 := chr(Mod11(Tot)+48);

      Tot := 0;
      for k := 2 to 9 do Tot := Tot + (ordem(aux[k]) * (k - 1));

      Tot := Tot + (ordem(dg1) * 9);       // mais o digito Ver. anterior

      dg2 := chr(Mod11(Tot)+48);

      result := (dg1  = aux[10]) and (dg2 = Aux[11]);

      if not(result) then
         Mensagem('O C.P.F digitado é inválido, os digitos ' +
                     'verificadores não coincidem', tmErro);
    end;
    //
    function montaCGC : boolean;
    Var k,x : byte;
    begin
      Tot := 0;
      X   := 6;
      for k := 1 to 12 do begin
         if X > 9 then X := 2;
         Tot := Tot + (ordem(aux[k]) * X);
         Inc(X);
         end;
      dg1 := chr(Mod11(Tot)+48);

      Tot := 0;
      X   := 5;
      for k := 1 to 12 do begin
         if X > 9 then X := 2;
         Tot := Tot + (ordem(aux[k]) * X);
         Inc(X);
         end;

      Tot := Tot + (ordem(dg1) * 9);       // mais o digito Ver. anterior
      dg2 := chr(Mod11(Tot)+48);

      result := (dg1 = aux[13]) and (dg2 = Aux[14]);
      if not(result) then
         mensagem('O C.N.P.J. digitado é inválido, os dígitos '+
                     'verificadores não coincidem!', tmErro);
    end;

begin
  Result := false;
  Bx  := Length(Aux);
  if bx = 0 then begin
     result := true;     // se o campo for em branco é valido
     exit;
  end;
  for I := 1 to bx do
     if not CharInSet(aux[I], ['0'..'9']) then begin
        Mensagem('Existem caracteres ilegais no CNPJ/CPF.',tmErro);
        exit;
        end;
  case bx of
    11 : result := MontaCPF;
    14 : result := MontaCGC;
  else
     Mensagem('Número incorreto de dígitos no CNPJ/CPF.',tmErro);
  end;
end;

function TeclaPress(Tecla :Char) :Char;
begin
   TeclaPress := Tecla;
   if CharInSet(Tecla, ['1'..'9']) then
      TeclaPress := #0;
end;

function TeclaNumPress(Tecla :Char) :Char;
begin
   TeclaNumPress := Tecla;
   if not CharInSet(Tecla, ['0'..'9',',','-', #8, #127]) and not (Tecla = ',') then begin
      if Tecla = '.' then begin
         TeclaNumPress := ',';
      end
      else begin
         TeclaNumPress := #0;
      end;
   end;
end;

procedure ChamaPrograma(Classe :TComponentClass; Form :TForm);
begin
   if UnicaCopia(Classe) then begin
      Application.CreateForm(Classe, Form);
      if Form.Width  > 775 then Form.Left := 1;
      if Form.Height > 510 then Form.Top  := 1;
      if Form.BorderStyle = {$IFDEF LINUX} fbsDialog {$ELSE} bsDialog {$ENDIF} then begin
         Form.ShowModal;
         Form.Free;
      end else begin
         Form.Show;
//         Form.Top := ((Application.MainForm.ClientHeight - Form.Height) div 2) - (22 + 13);
      end;
   end;
end;

function MesExtenso(Mes :Byte) :String;
begin
   case Mes of
      1:  MesExtenso := 'Janeiro';
      2:  MesExtenso := 'Fevereiro';
      3:  MesExtenso := 'Março';
      4:  MesExtenso := 'Abril';
      5:  MesExtenso := 'Maio';
      6:  MesExtenso := 'Junho';
      7:  MesExtenso := 'Julho';
      8:  MesExtenso := 'Agosto';
      9:  MesExtenso := 'Setembro';
      10: MesExtenso := 'Outubro';
      11: MesExtenso := 'Novembro';
      12: MesExtenso := 'Dezembro';
   end;
end;

function AnoExtenso(Ano: Word): string;
begin
   case Ano of
      2012: Result := 'Dois mil e doze';
      2013: Result := 'Dois mil e treze';
      2014: Result := 'Dois mil e quatorze';
      2015: Result := 'Dois mil e quize';
      2016: Result := 'Dois mil e dezesseis';
      2017: Result := 'Dois mil e dezessete';
      2018: Result := 'Dois mil e dezoito';
      2019: Result := 'Dois mil e dezenove';
      2020: Result := 'Dois mil e vinte';
   end;
end;

function DiaExtenso(DataX :TDate) :String;
var
  Dias: array[1..7] of string;
begin
  Dias[1] := 'Domingo';
  Dias[2] := 'Segunda-Feira';
  Dias[3] := 'Terça-Feira';
  Dias[4] := 'Quarta-Feira';
  Dias[5] := 'Quinta-Feira';
  Dias[6] := 'Sexta-Feira';
  Dias[7] := 'Sabado';
  DiaExtenso :=  Dias[DayOfWeek(DataX)];
end;

function DiaExtenso(Dia: Word): string;
begin
   case Dia of
      1: Result := 'Um';
      2: Result := 'Dois';
      3: Result := 'Três';
      4: Result := 'Quatro';
      5: Result := 'Cinco';
      6: Result := 'Seis';
      7: Result := 'Sete';
      8: Result := 'Oito';
      9: Result := 'Nove';
      10: Result := 'Dez';
      11: Result := 'Onze';
      12: Result := 'Doze';
      13: Result := 'Treze';
      14: Result := 'Quatorze';
      15: Result := 'Quize';
      16: Result := 'Dezesseis';
      17: Result := 'Dezessete';
      18: Result := 'Dezoito';
      19: Result := 'Dezenove';
      20: Result := 'Vinte';
      21: Result := 'Vinte e um';
      22: Result := 'Vinte e dois';
      23: Result := 'Vinte e três';
      24: Result := 'Vinte e quatro';
      25: Result := 'Vinte e cinco';
      26: Result := 'Vinte e seis';
      27: Result := 'Vinte e sete';
      28: Result := 'Vinte e oito';
      29: Result := 'Vinte e nove';
      30: Result := 'Trinta';
      31: Result := 'Trinta e um';
   end;
end;

function MascaraFloat(Valor :String) :Extended;
var
   ValorAux :String;
   Contador :Byte;
   Posicao  :Byte;
begin
   Try  
      Result := 0;
      if Valor = '' then Exit;
      ValorAux := Valor;
      for Contador := 1 to Length(ValorAux) do begin
         Posicao := Pos(FormatSettings.ThousandSeparator, Valor);
         if Posicao > 0 then begin
            Delete(Valor, Posicao, 1)
         end;
         Posicao := Pos(FormatSettings.DecimalSeparator+FormatSettings.DecimalSeparator, Valor);
         if Posicao > 0 then begin
            Delete(Valor, Posicao, 1)
         end;
      end;
      Result := StrToFloat(Valor);
   except
      Result := 0;
   end;
end;

function DiasNoMes(Mes,Ano: Word) : Word;
const
  NDias:array[1..12] of Byte=(31,28,31,30,31,30,31,31,30,31,30,31);

  function Modulo4(AAno: Word): Boolean;
  begin
    Result:=(AAno mod 4=0) and ((AAno mod 100<>0) or (AAno mod 400=0));
  end;

begin
  Result := NDias[Mes];
  if (Mes = 2) and Modulo4(Ano) then Inc(Result);
end;

function DataExtenso(Data:TDate):String;
begin
  Result := FormatDateTime('dd "de" mmmm "de" yyyy', Data)
end;

// retorna o ultimo dia do mes que foi enviado.
function DiaMaximoNoMes(const Data: TDateTime): TDateTime;
var dia, mes, ano : word;
begin
   DecodeDate(Data, Ano, Mes, Dia);
   Dia := DiasNoMes(Mes, Ano);
   Result := EncodeDate(Ano, Mes, Dia);
end;

function UltimoDiaNoMes(const Data: TDateTime): TDateTime;
begin
  Result := DiaMaximoNoMes(Data);
end;


{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure MudaCor;
var
   Contador, i :Integer;
   Componente  :TComponent;
begin
   if not Assigned(Screen.ActiveForm) then
      Exit;
   with Screen.ActiveForm do
   begin
      for Contador := 0 to ComponentCount - 1 do begin
         Componente := Components[Contador];
         if Componente.Tag = 3 then Continue;
         if (Componente is TDnBox) then begin
            if TDnBox(Componente).Tag = 3 then Continue;
            TDnBox(Componente).LabelFont.Color := clNavy;
            TDnBox(Componente).LabelFont.Style := [];
            TDnBox(Componente).ColorEnabled    := CorAtivada;
            TDnBox(Componente).ColorDisabled   := CorDesativada;
            if (TDnBox(Componente).Tag = 10) or (TDnBox(Componente).Tag = 11) then begin
               if (TDnBox(Componente).Tag = 10) then begin
                  TDnBox(Componente).ColorDisabled   := CorTitulo;
                  TDnBox(Componente).LabelFont.Color := CorFonteTitulo;//iif(CorTitulo=7034685, clwhite, CorFonteTitulo);
               end;
               if (TDnBox(Componente).Tag = 11) then begin
                  TDnBox(Componente).ColorDisabled   := CorRodape;
                  TDnBox(Componente).LabelFont.Color := CorFonteRodape;//iif(CorTitulo=7034685, clwhite, CorFonteRodape);
               end;
               TDnBox(Componente).LabelFont.Style    := [fsBold];
            end;
         end
         else if Componente is TEditBox then begin
            if TEditBox(Componente).Tag = 3 then Continue;
            TEditBox(Componente).LabelFont.Color := CorLabel;
         end
         else if (Componente is TDBGrid) then begin
            if not TDBGrid(Componente).ParentColor then begin
               TDBGrid(Componente).Color      := CorDesativada;
            end;
            TDBGrid(Componente).FixedColor := CorDesativada;
            for i := 0 to TDBGrid(Componente).Columns.Count - 1 do begin

               if TDBGrid(Componente).DataSource <> nil then begin

                  TDBGrid(Componente).Columns[i].Color            := CorDesativada;
                  TDBGrid(Componente).Columns[i].Title.Color      := CorTitulo;

                  if CorFonteTitulo <> 0 then begin
                     TDBGrid(Componente).Columns[i].Title.Font.Color := CorFonteTitulo;
                     TDBGrid(Componente).Columns[i].Title.Font.Style := [fsBold];
                  end else begin
                     if TDBGrid(Componente).Columns[i].Title.Font.Color >= 0 then begin // este teste evita
                        TDBGrid(Componente).Columns[i].Title.Font.Color := clBlack;     // um erro de read em situacoes especiais
                        TDBGrid(Componente).Columns[i].Title.Font.Style := [fsBold];    // nao mexer...
                     end;
                  end;
               end;
            end;
         end
         else if (Componente is TDBZGrid) then begin
            if not TDBZGrid(Componente).ParentColor then begin
               TDBZGrid(Componente).Color      := CorDesativada;
            end;
            TDBZGrid(Componente).FixedColor := CorDesativada;
            for i := 0 to TDBZGrid(Componente).Columns.Count - 1 do begin

               if TDBZGrid(Componente).DataSource <> nil then begin

                  TDBZGrid(Componente).Columns[i].Color            := CorDesativada;
                  TDBZGrid(Componente).Columns[i].Title.Color      := CorTitulo;

                  if CorFonteTitulo <> 0 then begin
                     TDBZGrid(Componente).Columns[i].Title.Font.Color := CorFonteTitulo;
                     TDBZGrid(Componente).Columns[i].Title.Font.Style := [fsBold];
                  end else begin
                     if TDBZGrid(Componente).Columns[i].Title.Font.Color >= 0 then begin // este teste evita
                        TDBZGrid(Componente).Columns[i].Title.Font.Color := clBlack;     // um erro de read em situacoes especiais
                        TDBZGrid(Componente).Columns[i].Title.Font.Style := [fsBold];    // nao mexer...
                     end;
                  end;
               end;
            end;
         end

//         else if (Componente is TcxGridDBTableView) then
//         begin
////            if not TcxGridDBTableView(Componente).ParentColor then begin
////               TcxGridDBTableView(Componente).Color      := CorDesativada;
////            end;
////            TcxGrid(Componente).FixedColor := CorDesativada;
//            for i := 0 to TcxGridDBTableView(Componente).Columns - 1 do
//            begin
//
//               if TcxGridDBTableView(Componente).DataSource <> nil then begin
//
//                  TcxGridDBTableView(Componente).Columns[i].Color            := CorDesativada;
//                  TcxGridDBTableView(Componente).Columns[i].Title.Color      := CorTitulo;
//
//                  if CorFonteTitulo <> 0 then begin
//                     TcxGridDBTableView(Componente).Columns[i].Title.Font.Color := CorFonteTitulo;
//                     TcxGridDBTableView(Componente).Columns[i].Title.Font.Style := [fsBold];
//                  end else begin
//                     if TcxGridDBTableView(Componente).Columns[i].Title.Font.Color >= 0 then begin // este teste evita
//                        TcxGridDBTableView(Componente).Columns[i].Title.Font.Color := clBlack;     // um erro de read em situacoes especiais
//                        TcxGridDBTableView(Componente).Columns[i].Title.Font.Style := [fsBold];    // nao mexer...
//                     end;
//                  end;
//               end;
//            end;
//         end

         else if Componente is TStringGrid then begin
              TStringGrid(Componente).FixedColor := CorTitulo;
              TStringGrid(Componente).Color      := CorDesativada;
         end
         else if Componente is TLabel then begin
            if (TLabel(Componente).Tag = 10) or (TLabel(Componente).Tag = 11) then begin
               if (TLabel(Componente).Tag = 10) then TLabel(Componente).Font.Color := CorFonteTitulo;
               if (TLabel(Componente).Tag = 11) then TLabel(Componente).Font.Color := CorFonteRodape;
               TLabel(Componente).Font.Style := [fsBold];
            end;
         end
         else if Componente is TPanel then begin
            if (TPanel(Componente).Tag = 10) or (TPanel(Componente).Tag = 11) then begin
               if (TPanel(Componente).Tag = 10) then TPanel(Componente).Color      := CorTitulo;
               if (TPanel(Componente).Tag = 10) then TPanel(Componente).Font.Color := CorFonteTitulo;
               if (TPanel(Componente).Tag = 11) then TPanel(Componente).Color      := CorRodape;
               if (TPanel(Componente).Tag = 11) then TPanel(Componente).Font.Color := CorFonteRodape;
               TLabel(Componente).Font.Style := [fsBold];
            end;
         end
         else if Componente is TFrame then
            MudaCorFrame(Componente);
      end;
   end;
end;

procedure MudaCorFrame(Sender :TObject);
var
   Contador, i :Integer;
   Componente  :TComponent;
begin
   with TFrame(Sender) do begin
      for Contador := 0 to ComponentCount - 1 do begin
         Componente := Components[Contador];
         if Componente.Tag = 3 then Continue;
         if Componente is TDnBox then begin
            if TDnBox(Componente).Tag = 3 then Continue;
            TDnBox(Componente).LabelFont.Color := CorLabel;
            TDnBox(Componente).LabelFont.Style := [];
            TDnBox(Componente).ColorEnabled    := CorAtivada;
            TDnBox(Componente).ColorDisabled   := CorDesativada;
            if (TDnBox(Componente).Tag = 10) or (TDnBox(Componente).Tag = 11) then begin
               if (TDnBox(Componente).Tag = 10) then begin
                  TDnBox(Componente).ColorDisabled   := CorTitulo;
                  TDnBox(Componente).LabelFont.Color := CorFonteTitulo;
                  TDnBox(Componente).ColorEnabled    := CorTitulo;
                  TDnBox(Componente).LabelFont.Color := CorFonteTitulo;//iif(CorTitulo=7034685, clwhite, CorFonteTitulo);
               end;
               if (TDnBox(Componente).Tag = 11) then begin
                  TDnBox(Componente).ColorDisabled   := CorRodape;
                  TDnBox(Componente).ColorEnabled    := CorRodape;
                  TDnBox(Componente).LabelFont.Color := CorFonteRodape;//iif(CorTitulo=7034685, clwhite, CorFonteRodape);
               end;
               TDnBox(Componente).LabelFont.Style    := [fsBold];
            end;
         end
         else if Componente is TEditBox then begin
            if TEditBox(Componente).Tag = 3 then Continue;
            TEditBox(Componente).LabelFont.Color := CorLabel;
         end
         else if Componente is TDBGrid then begin
            TDBGrid(Componente).Color      := CorDesativada;
            TDBGrid(Componente).FixedColor := CorDesativada;
            for i := 0 to TDBGrid(Componente).Columns.Count - 1 do begin
               TDBGrid(Componente).Columns[i].Color            := CorDesativada;
               TDBGrid(Componente).Columns[i].Title.Color      := CorTitulo;
               TDBGrid(Componente).Columns[i].Title.Font.Color := CorFonteTitulo;
               TDBGrid(Componente).Columns[i].Title.Font.Style := [fsBold];
            end;
         end
         else if Componente is TLabel then begin
            if (TLabel(Componente).Tag = 10) or (TLabel(Componente).Tag = 11) then begin
               if (TLabel(Componente).Tag = 10) then TLabel(Componente).Font.Color := CorFonteTitulo;
               if (TLabel(Componente).Tag = 11) then TLabel(Componente).Font.Color := CorFonteRodape;
               TLabel(Componente).Font.Style := [fsBold];
            end
            else begin
               TLabel(Componente).Font.Color := CorLabel;
            end;
         end
         else if Componente is TPanel then begin
            if (TPanel(Componente).Tag = 10) then begin
               TPanel(Componente).Color      := CorTitulo;
               TPanel(Componente).Font.Color := CorFonteTitulo;
            end;
            if (TPanel(Componente).Tag = 11) then begin
               TPanel(Componente).Color      := CorRodape;
               TPanel(Componente).Font.Color := CorFonteRodape;
            end;
         end
      end;
   end;
end;
{$ENDIF}

procedure LimpaStringGrid(StringGrid :TStringGrid);
var
   ContLin  :Integer;
   ContCol  :Integer;
begin
   with StringGrid do begin
      for ContLin := 1 to (RowCount - 1) do begin
         for ContCol := 0 to 215 do begin
            Cells[ContCol, ContLin] := '';
         end;
      end;
      RowCount := 2;
   end;
end;

procedure LimpaStringGridPTL(StringGrid :TStringGrid; RowRemanescente:Integer);
var
   ContLin  :Integer;
   ContCol  :Integer;
begin
   if RowRemanescente <= 0 then RowRemanescente := 1;
   with StringGrid do begin
      for ContLin :=  RowRemanescente-1 to (RowCount - 1) do begin
         for ContCol := 0 to ColCount + 204 do begin
            Cells[ContCol, ContLin] := '';
         end;
      end;
      RowCount := RowRemanescente;
   end;
end;

procedure ExcluiLinhaStringGrid(StringGrid :TStringGrid); { Exclui a linha atual de uma StringGrid}
var
   ContLin  :integer;
   ContCol  :integer;
   Linha    :integer;
begin
   with StringGrid do begin
      Linha := Row;
      if RowCount = 2 then begin
         for ContCol := 0 to ColCount + 204 do begin
             Cells[ContCol, 1] := '';
         end;
         Exit;
      end;
      for ContLin := Linha to RowCount - 1 do begin
         for ContCol := 0 to ColCount + 204 do begin
            Cells[ContCol, ContLin] := Cells[ContCol, ContLin + 1];
         end;
      end;
      RowCount := RowCount - 1;
   end;
end;

procedure ExcluiLinhaStringGridPTL(StringGrid :TStringGrid; Linha:Integer); { Exclui a linha atual de uma StringGrid até ficar uma linha}
var
   ContLin  :Integer;
   ContCol  :Integer;
begin
   with StringGrid do begin
      if RowCount = 1 then begin
         for ContCol := 0 to ColCount + 204 do begin
            Cells[ContCol, 0] := '';
         end;
         Exit;
      end;
      for ContLin := Linha  to (RowCount - 1) do begin
         for ContCol := 0 to ColCount + 204 do begin
            Cells[ContCol, ContLin] := Cells[ContCol, ContLin + 1];
         end;
      end;
      if (RowCount-1)>StringGrid.FixedRows then begin
         RowCount := RowCount - 1;
      end;
   end;
end;

procedure FecharTodasJanelasMDI;
var
   i: Integer;
   MainForm, Form: TForm;
begin
   MainForm := Application.MainForm;

   for i := (MainForm.MDIChildCount - 1) downto 0 do
   begin
      Form := MainForm.MDIChildren[i];
      if (Form.Name <> 'F_Alerta') then
      begin
        if Assigned(MDICloseForm) then
          MDICloseForm(Form)
        else
          Form.Close;
      end;
   end;
end;

type
  PMethod = ^TMethod;
  TDelayInvMethod = class(TObject)
  strict private
    fMethod: PMethod;
    fWndHandle: HWND;
    procedure MainWndProc(var AMessage: TMessage);
    procedure DoTimer;
  public
    constructor Create(AMethod: PMethod);
    destructor Destroy; override;
  end;

{ TDelayInvMethod }

constructor TDelayInvMethod.Create(AMethod: PMethod);
begin
  fMethod := AMethod;
  fWndHandle := AllocateHWnd(MainWndProc);
  SetTimer(fWndHandle, 1, 250, nil);
end;

destructor TDelayInvMethod.Destroy;
begin
  if (fWndHandle <> 0) then
  begin
    DeallocateHWnd(fWndHandle);
    fWndHandle := 0;
  end;

  if (fMethod <> nil) then
    Dispose(fMethod);
end;

procedure TDelayInvMethod.DoTimer;
begin
  KillTimer(fWndHandle, 1);

  try
     FecharTodasJanelasMDI;
     if Assigned(fMethod) then
       TNotifyEvent(fMethod^)(Self);
     Destroy;
  finally
     EstaFechandoTodasJanelas := False;
  end;
end;

procedure TDelayInvMethod.MainWndProc(var AMessage: TMessage);
begin
  try
    if (AMessage.Msg = WM_TIMER) then
      DoTimer
    else
      AMessage.Result := DefWindowProc(fWndHandle, AMessage.Msg, AMessage.WParam, AMessage.LParam);
  except
  end;
end;

procedure ApplicationModalEnd(Self, Sender: TObject);
begin
  Application.OnModalEnd := nil;
  TDelayInvMethod.Create(Pointer(Self));
end;

function GetApplication_OnModalEnd(const OnCompleted: TNotifyEvent): TNotifyEvent;
var
   P: ^TMethod;
begin
   if Assigned(OnCompleted) then
   begin
      New(P);
      P^ := TMethod(OnCompleted);
   end else
      P := nil;

   System.TMethod(Result).Code := @ApplicationModalEnd;
   System.TMethod(Result).Data := Pointer(P);
end;

function BuildNotifyEvent(AData, ACode: Pointer): TNotifyEvent;
begin
  TMethod(Result).Code := ACode;
  TMethod(Result).Data := AData;
end;

function FecharTodasJanelas:Boolean;
begin
   Result := FecharTodasJanelas(nil);
end;

function FecharTodasJanelas(const OnCompleted: TNotifyEvent; QtdForms: Integer = 1): Boolean;
var
   i, ModalCount: Integer;
   Form: TForm;
begin
    {$IFNDEF PAFECOBPL}
    if FEcoMenu.NumeroFormsVisiveisValidos = QtdForms then
    begin
       if Assigned(OnCompleted) then
          OnCompleted(Application);
       //Exit(False);
       Exit(True);
    end;
   {$ENDIF}

    if not Mensagem('Todas janelas serão fechadas!'#13#10'Deseja fechá-las?', tmConfirma) then
    begin
       if Assigned(OnCompleted) then
          OnCompleted(Application);
       Exit(False);
    end;

   Result := True;
   ModalCount := 0;

   for i := (Screen.FormCount - 1) downto 0 do
   begin
      Form := Screen.Forms[i];
      if (fsModal in Form.FormState) then
      begin
         Form.ModalResult := mrCancel;
         Inc(ModalCount);
      end;
   end;

   {
      Foi Criado esta variável, para corrigir uma falha com o seguinte cenário:
      é aberto uma tela MDI e no showform dessa tela é criado uma tela em modal
      então é aberto a tela de fechamento de todas as telas do sistema

      #Solicitação 145722
   }
   EstaFechandoTodasJanelas := True;
   if (ModalCount = 0) then
   begin
      { Se não há janelas em ShowModal.. }
      try
         FecharTodasJanelasMDI;

         if Assigned(OnCompleted) then
           OnCompleted(Application);
      finally
         EstaFechandoTodasJanelas := False;
      end;
   end else
   begin
      { Se há janelas em modal.
        Neste caso tem que esperar que todas as janelas em modal sejam fechadas.. }
      Application.OnModalEnd := GetApplication_OnModalEnd(OnCompleted);
   end;
end;

function MascaraFone(const Fone :string) :string;
var Tam       : Byte;
    Telefone  : TMaskEdit;
    FFone: string;
begin
  result := '';
  FFone := ErtStrUtils.OnlyNumbers(Fone);
  if Trim(FFone) = '' then Exit;
  Telefone := TMaskEdit.Create(nil);
  try
    With Telefone do begin
       Tam := length(FFone);
       if Tam < 8 then begin
          EditMask := '999-9999;0; ';
          Text := StrZero(StrToInt64(Copy(FFone,1,7)),7)
       end
       else if Tam = 8 then begin
             EditMask := '9999-9999;0; ';
             Text := StrZero(StrToInt64(Copy(FFone,1,8)),8);
       end
       else if Tam = 9 then begin
                EditMask := '(99) 999-9999;0; ';
                Text := StrZero(StrToInt64(Copy(FFone,1,9)),9)
       end
       else if Tam = 10 then begin

          if StrZero(StrToInt64(Copy(FFone,1,4)),4) = '0800' then
          begin
             EditMask := '(9999) 999999;0; ';
             Text := StrZero(StrToInt64(Copy(FFone,1,10)),10);
          end
          else if Copy(FFone,1,1) = '0' then
          begin
             EditMask := '(999) 999-9999;0; ';
             Text := StrZero(StrToInt64(Copy(FFone,1,10)),10);
          end
          else
          begin
             EditMask := '(99) 9999-9999;0; ';
             Text := StrZero(StrToInt64(Copy(FFone,1,10)),10)
          end
       end else if Tam = 11 then
       begin
          EditMask := '(999) 9999-9999;0; ';
          Text := StrZero(StrToInt64(Copy(Trim(FFone),1,11)),11)
       end
       else if Tam = 12  then
       begin

         EditMask := '(999) 99999-9999;0; ';
         Text := StrZero(StrToInt64(Copy(FFone,1,12)),12)

       end else
       begin
          EditMask := '(999) 99999-9999;0; ';
          //Text := StrZero(StrToInt64(Copy(Trim(FFone),1,11)),11)
          Text := StrZero(StrToInt64(Copy(Trim(FFone),1,12)),12)
       end;
       result := Telefone.EditText;
    end
  finally
    Telefone.Free;
  end;
                                     
end;

function Alinha(Conteudo :String; Alinhamento :TAlinhamento; Tamanho :Integer) :String;
begin
   if Tamanho < Length(Conteudo) then begin
      Conteudo := Copy(Conteudo, 1, Tamanho);
   end;
   if Alinhamento = taEsquerdo then begin
      Alinha := Conteudo + Espaco(Tamanho - Length(Conteudo));
   end
   else if Alinhamento = taCentro then begin
      Alinha := Espaco((Tamanho - Length(Conteudo)) div 2) + Conteudo + Espaco((Tamanho - Length(Conteudo)) div 2);
   end
   else if Alinhamento = taDireito then begin
      Alinha := Espaco(Tamanho - Length(Conteudo)) + Conteudo;
   end;
end;

function Espaco(Tamanho :Byte) :String;
begin
   Espaco := StringOfChar(' ', Tamanho);
end;

function MascaraValor(Valor :Extended; Decimais :Byte) :String;
begin
  MascaraValor := FormatFloat('#,##' + '0.' + Replique('0', Decimais), Valor);
end;

function DiaPascoa(Ano :Integer) :TDate;
var
   Numero    :Byte;
   Dia       :Byte;
   Mes       :Byte;
   DiaSemana :Byte;
const
  ArrayDia : array[1..19] of byte = (14,03,23,11,31,18,08,28,16,05,25,13,02,22,10,30,17,07,27);
  ArrayMes : array[1..19] of byte = (04,04,03,04,03,04,04,03,04,04,03,04,04,03,04,03,04,04,03);

begin
   Numero := (Ano mod 19) + 1;
   Dia    := ArrayDia[Numero];
   Mes    := ArrayMes[Numero];

   DiaSemana := DayOfWeek(StrToDate(IntToStr(Dia) + '/' + IntToStr(Mes) + '/' + IntToStr(Ano)));
   Dia := Dia + (7 - DiaSemana) + 1;
   if Dia > DiasNoMes(Mes, Ano) then begin
      Dia := Dia - DiasNoMes(Mes, Ano);
      Inc(Mes);
   end;
   DiaPascoa := StrToDate(IntToStr(Dia) + '/' + IntToStr(Mes) + '/' + IntToStr(Ano))
end;

function DiaSextaFeiraSanta(Ano :Integer) :TDate;
begin
  DiaSextaFeiraSanta := DiaPascoa(Ano) - 2;
end;

function DiaCarnaval(Ano :Integer) :TDate;
begin
  DiaCarnaval := DiaPascoa(Ano) - 47;
end;

function LeFavoritosDoUsuario : boolean;
var QueryPesquisa :TSQLQuery;
begin
   Result := false;
   if Usuario = '' then exit;
   QueryPesquisa               := TSQLQuery.Create(nil);
   QueryPesquisa.SQLConnection := FDataModule.sConnection;

   try
      with QueryPesquisa do begin
         Sql.Clear;
         Sql.Add('SELECT Usuario,       ' +
                 '       Fav0,          ' +
                 '       Fav1,          ' +
                 '       Fav2,          ' +
                 '       Fav3,          ' +
                 '       Fav4,          ' +
                 '       Fav5,          ' +
                 '       Fav6,          ' +
                 '       Fav7,          ' +
                 '       Fav8,          ' +
                 '       Fav9           ' +
                 'FROM   TGerUsuario    ' +
                 'WHERE                 ' +
                 'Usuario = :Usuario    ' );
         ParamByName('Usuario').asString := Usuario;
         Open;
         if (Bof and Eof) then begin
            Mensagem('Usuário não encontrado lendo os favoritos!', tmInforma);
         end else begin
            result := true;
            sFavoritosDoUsuario[0] := FieldByName('Fav0').asInteger;
            sFavoritosDoUsuario[1] := FieldByName('Fav1').asInteger;
            sFavoritosDoUsuario[2] := FieldByName('Fav2').asInteger;
            sFavoritosDoUsuario[3] := FieldByName('Fav3').asInteger;
            sFavoritosDoUsuario[4] := FieldByName('Fav4').asInteger;
            sFavoritosDoUsuario[5] := FieldByName('Fav5').asInteger;
            sFavoritosDoUsuario[6] := FieldByName('Fav6').asInteger;
            sFavoritosDoUsuario[7] := FieldByName('Fav7').asInteger;
            sFavoritosDoUsuario[8] := FieldByName('Fav8').asInteger;
            sFavoritosDoUsuario[9] := FieldByName('Fav9').asInteger;
         end;
      end;
   finally
      QueryPesquisa.Free;
      RefreshMenuPopup := true;
   end;
end;

function GravaFavoritosDoUsuario : boolean;
var QueryPesquisa :TSQLQuery;
begin
   Result := false;
   if Usuario = '' then exit;

   QueryPesquisa               := TSQLQuery.Create(nil);
   QueryPesquisa.SQLConnection := FDataModule.sConnection;

   try
     with QueryPesquisa do begin
        Sql.Clear;
        Sql.Add('UPDATE TGerUsuario SET   ' +
                 '  Fav0 = :Fav0,          ' +
                 '  Fav1 = :Fav1,          ' +
                 '  Fav2 = :Fav2,          ' +
                 '  Fav3 = :Fav3,          ' +
                 '  Fav4 = :Fav4,          ' +
                 '  Fav5 = :Fav5,          ' +
                 '  Fav6 = :Fav6,          ' +
                 '  Fav7 = :Fav7,          ' +
                 '  Fav8 = :Fav8,          ' +
                 '  Fav9 = :Fav9           ' +
                 'WHERE                    ' +
                 'Usuario = :Usuario       ' );
        ParamByName('Fav0').asInteger   := sFavoritosDoUsuario[0];
        ParamByName('Fav1').asInteger   := sFavoritosDoUsuario[1];
        ParamByName('Fav2').asInteger   := sFavoritosDoUsuario[2];
        ParamByName('Fav3').asInteger   := sFavoritosDoUsuario[3];
        ParamByName('Fav4').asInteger   := sFavoritosDoUsuario[4];
        ParamByName('Fav5').asInteger   := sFavoritosDoUsuario[5];
        ParamByName('Fav6').asInteger   := sFavoritosDoUsuario[6];
        ParamByName('Fav7').asInteger   := sFavoritosDoUsuario[7];
        ParamByName('Fav8').asInteger   := sFavoritosDoUsuario[8];
        ParamByName('Fav9').asInteger   := sFavoritosDoUsuario[9];
        ParamByName('Usuario').asString := Usuario;
        ExecSql;
        result := true;
     end;
   finally
     QueryPesquisa.Free;
   end;
end;

function DiaFeriadoMunicipal(DataTeste :TDateTime; out msgRetorno: string) :boolean;
const cSQL='SELECT Descricao FROM TGerFeriado WHERE DiaMes= ''%S'' ';
var Qry     :TSQLQuery;
    DiaMes  :String;
begin
   Result := False;
   Qry := TSQLQuery.Create(nil);
   Qry.SQLConnection := FDataModule.sConnection;
   try
      DiaMes := FormatDateTime('DD/MM',DataTeste);
      Qry.Sql.Clear;
      Qry.Sql.Add(Format(cSQL, [DiaMes]));
      if EcoOpen(Qry) then begin
         Result := True;
         msgRetorno := Qry.FieldByName('Descricao').asString;
      end;
   finally
      Qry.Free;
   end;
end;


function ProximoDiaUtil(Data :TDate) :TDate;
var
   QueryPesquisa :TSQLQuery;
   DataUtil      :TDate;
begin
   QueryPesquisa               := TSQLQuery.Create(nil);
   QueryPesquisa.SQLConnection := FDataModule.sConnection;
   try
      with QueryPesquisa do begin
         DataUtil := Utilitarios.ProximoDiaUtil(Data);
         while True do begin
            Sql.Clear;
            Sql.Add('SELECT DiaMes,   ' +
                    '       Ano       ' +
                    'FROM TGerFeriado ' +
                    'WHERE            ' +
                    'DiaMes = :DiaMes ' );
            ParamByName('DiaMes').asString := Copy(DateToStr(DataUtil), 1, 5);
            if not EcoOpen(QueryPesquisa) then begin
               Result := DataUtil;
               Exit;
            end else begin
               if StrToIntDef(FieldByName('Ano').AsString,0)<>0 then begin
                  if StrZero(YearOf(DataUtil),4)<>FieldByName('Ano').asString then begin
                     Result := DataUtil;
                     Exit;
                  end;
               end;
            end;
            DataUtil := DataUtil + 1;
         end;
      end;
   finally
      QueryPesquisa.Free;
   end;
end;

function DataValida(Data :String) :TDateTime;
var DiaX,MesX,AnoX : word;  // auxiliares-para conversao e teste
    Dia1,Mes1,Ano1 : word;  // auxiliares-data do sistema
begin
   DecodeDate(Date, Ano1, Mes1, Dia1);  // data atual

   Data := ErtStrUtils.OnlyNumbers(Data);

   DiaX := 0;
   MesX := 0;
   AnoX := 0;

   if Copy(Data, 1, 2) <> '' then begin
      DiaX := StrToInt(Copy(Data, 1, 2));
   end;
   if Copy(Data, 3, 2) <> '' then begin
      MesX := StrToInt(Copy(Data, 3, 2));
   end;
   if Copy(Data, 5, 4) <> '' then begin
      AnoX := StrToInt(Copy(Data, 5, 4));
   end;

   { faixa valida 01/01/1900 ... 31/12/2199 }
   if DiaX = 00 then begin     {digitado 00 no dia será a data atual}
      DiaX := Dia1;
      MesX := Mes1;
      AnoX := Ano1;
   end;

   if  MesX > 12 then MesX := 12;

{   if (MesX = 00) and (Copy(Data, 3, 2) = '') then MesX := Mes1;
   if (AnoX = 00) and (Copy(Data, 5, 4) = '') then AnoX := Ano1; }

   if (MesX = 00) then MesX := Mes1;
   if (AnoX = 00) then AnoX := Ano1;

   if  DiaX > DiasNoMes(MesX,AnoX) then DiaX := DiasNoMes(MesX,AnoX);

   if  AnoX <= 0079 then
       AnoX := 2000 + AnoX
   else
      if (AnoX >= 0080) and (AnoX < 0100) then
          AnoX := 1900 + AnoX
      else
          if (AnoX >= 0100) and (AnoX < 1900) then
              AnoX := 1900
          else
              if  AnoX > 2199 then AnoX := 2199;


   Data := StrZero(DiaX, 2) + '/' + StrZero(MesX, 2) + '/' + StrZero(AnoX, 4);

   DataValida := StrToDate(Data);
end;

function CalculaMulta(Valor, Multa :Currency; Vencimento, Recebimento :TDate) :Currency;
begin
   Result := Trunca(TFinanceiro_Calculo.CalculaMulta(Valor,Multa,Vencimento, Recebimento),2);
//   if Vencimento >= Recebimento then
//   begin
//      Result := 0;
//      Exit;
//   end;
//   Result := Trunca((Valor * Multa) / 100, 2);

end;

function CalculaJuros(Valor :Currency; Vencimento, Recebimento :TDate) :Currency;
var
//   Contador       :Integer;
//   ValorAcumulado :Currency;
//   JurosAtual     :Currency;
     LJurosMens     :Currency;
begin
   if JurosCliente > 0 then
   begin
      LJurosMens := JurosCliente;
   end
   else begin
     LJurosMens := Juros;
   end;

   Result := TFinanceiro_Calculo.CalculaJuros(Valor,LJurosMens, Vencimento, Recebimento,TipoJuro);

//   Result := 0;
//   if Vencimento >= Recebimento then
//      Exit;
//
//   ValorAcumulado := 0;
//   JurosMens := 0;
//
//   if JurosCliente > 0 then
//   begin
//      JurosAtual := JurosCliente;
//   end
//   else begin
//      JurosAtual := Juros;
//   end;
//
//   Try
//     if TipoJuro = 'S' then
//     begin
//        Result := (Valor * (JurosAtual * (Recebimento - Vencimento))) / 100;
//     end
//     else begin
//        ValorAcumulado := Valor;
//        if TipoJuro = 'D' then begin
//           for Contador := 1 to Trunc(Recebimento-Vencimento) do begin
//              ValorAcumulado := ((ValorAcumulado * JurosAtual) / 100 ) + ValorAcumulado;
//           end;
//        end
//        else begin
//           JurosMens := JurosAtual * 30;
//           Contador  := Trunc(Recebimento-Vencimento);
//           while Contador > 30 do begin
//              ValorAcumulado := ((ValorAcumulado * JurosMens) / 100 ) + ValorAcumulado;
//              Contador := Contador - 30;
//           end;
//           if Contador > 0 then begin
//              ValorAcumulado := ((ValorAcumulado * (JurosAtual * Contador)) / 100 ) + ValorAcumulado;
//           end;
//        end;
//        Result := ValorAcumulado - Valor;
//     end;
//   Except
//   on e:Exception do
//   begin
//     ShowMessage('Erro ao calcular: Acum: '+FloatToStr(ValorAcumulado)+' Jurosmens '+FloatToStr(JurosMens)+' JurosAtu '+FloatToStr(JurosAtual)+#13+#10+
//                 e.Message);
//   end;
// end;
end;

function CalculaJurosPag(Valor :Currency; Vencimento :TDate; Juros :Currency) :Currency;
begin
   Result := TFinanceiro_Calculo.CalculaJurosPag(Valor,Juros,Vencimento);
//   if Vencimento >= Date then begin
//      Result := 0;
//      Exit;
//   end;
//   Result := ((Valor * Juros) / 100 ) *  Trunc(Date-Vencimento);
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
function CalculaDesconto(Valor :Currency; Vencimento, Recebimento: TDate) :Currency;
var
   Percentual :Currency;
begin
   if DescontoCliente > 0 then
      Percentual := DescontoCliente
   else
      Percentual := Desconto;

   Result := CalculaDesconto(Valor, Percentual, Vencimento, Recebimento);
end;

function CalculaDesconto(Valor, Percentual: Currency; Vencimento, Recebimento: TDate): Currency;
var
   diasCarencia, diasAVencer: Integer;
begin

   Result := 0;

   { Verifica se o parâmetro que indica desconto proporcional aos dias
     adiantados está ativado }
   if (ecPrm.ParamAtivo(peDescontoProporcionalAntecipado)) and (Vencimento > Recebimento) then
   begin
      diasAVencer := DaysBetween(Recebimento, Vencimento);
      diasCarencia := ecPrm.ParamAtivo(peDiasCarenciaDescontoAntecipado);

      if (diasAVencer <= diasCarencia) then
         exit;

      Percentual := Percentual * diasAVencer;
      Result := Arredonda((Valor * Percentual) / 100, 2);

      if (Result >= Valor) then
         Result := 0;
   end;
end;
{$ENDIF}

function ValorIndice(Indice :String; Data :TDateTime) :Extended;
var
   QueryPesquisa :TSQLQuery;
begin
   QueryPesquisa               := TSQLQuery.Create(nil);
   QueryPesquisa.SQLConnection := FDataModule.sConnection;

   try 
      with QueryPesquisa do begin
         Sql.Clear;
         Sql.Add('SELECT Data,         ' +
                 '       Valor         ' +
                 'FROM TGerValorIndice ' +
                 'WHERE                ' +
                 'Indice = :Indice AND ' +
                 'Data   <= :Data AND  ' +
                 'Valor  > 0           ' +
                 'ORDER BY             ' +
                 'Data DESC            ' );
         ParamByName('Indice').asString := Indice;
         ParamByName('Data').asDate     := Data;
         Open;
         if (Bof and Eof) then begin
            Result := 0;
         end
         else begin
           Transporte := FieldByName('Data').asString;
           Result     := FieldByName('Valor').asCurrency;
         end;
         Close;
      end;
   finally
      QueryPesquisa.Free;
   end;
end;

function Arred(Valor :Extended) :Extended;
begin
   Result := StrToFloat(FloatToStr(Valor));
end;

function Arredonda(Valor :Extended; Decimais :Integer;Truncar:Boolean=False) :Extended;
begin
   if Truncar then
      Arredonda := Trunca(Valor,Decimais)
   else
      Arredonda := MascaraFloat(FormatFloat('#,##' + '0.' + Replique('0', Decimais), Valor));
end;

function Trunca(Valor :Extended; Decimais :Integer) :Extended;
var
   Var1: String;
begin
   Var1 := FloatToStr(Valor);
   if (Pos(FormatSettings.DecimalSeparator,Var1) > 0) then
      Trunca := StrToFloat(Copy(Var1, 1, Pos(FormatSettings.DecimalSeparator,Var1)) + Copy(Var1,Pos(FormatSettings.DecimalSeparator,Var1) + 1, Decimais))
   else
      Trunca := Valor;
end;

function TruncTo(const AValue: Extended; const ADigit: TRoundToRange = -2): Extended;
var LFactor: Extended;
begin
   LFactor := IntPower(10, ADigit);
   Result := Trunc((AValue / LFactor)) * LFactor;
end;

function PortaImpressora(TipoImpressora :TImpressora): String;
var
   Impressora: string;
begin
   if TipoImpressora = tiBoleto              then Impressora := ReadPrintIni('ImpBoleto',              'ImpBoleto'             );
   if TipoImpressora = tiNotaFiscal          then Impressora := ReadPrintIni('ImpNotaFiscal',          'ImpNotaFiscal'         );
   if TipoImpressora = tiNotaServico         then Impressora := ReadPrintIni('ImpNotaServico',         'ImpNotaFiscal'         );
   if TipoImpressora = tiCarne               then Impressora := ReadPrintIni('ImpCarne',               'ImpCarne'              );
   if TipoImpressora = tiCheque              then Impressora := ReadPrintIni('ImpCheque',              'ImpCheque'             );
   if TipoImpressora = tiPedidoNormal        then Impressora := ReadPrintIni('ImpPedidoNormal',        'ImpPedidoNormal'       );
   if TipoImpressora = tiPedidoPrazo         then Impressora := ReadPrintIni('ImpPedidoPrazo',         'ImpPedidoPrazo'        );
   if TipoImpressora = tiPedidoReduzido      then Impressora := ReadPrintIni('ImpPedidoReduzido',      'ImpPedidoReduzido'     );
   if TipoImpressora = tiOrcamentoNormal     then Impressora := ReadPrintIni('ImpOrcamentoNormal',     'ImpOrcamentoNormal'    );
   if TipoImpressora = tiOrcamentoReduzido   then Impressora := ReadPrintIni('ImpOrcamentoReduzido',   'ImpOrcamentoReduzido'  );
   if TipoImpressora = tiCondicionalNormal   then Impressora := ReadPrintIni('ImpCondicionalNormal',   'ImpCondicionalNormal'  );
   if TipoImpressora = tiCondicionalReduzida then Impressora := ReadPrintIni('ImpCondicionalReduzida', 'ImpCondicionalReduzida');
   if TipoImpressora = tiDuplicata           then Impressora := ReadPrintIni('ImpDuplicata',           'ImpDuplicata'          );
   if TipoImpressora = tiEntregaFutura       then Impressora := ReadPrintIni('ImpEntrega',             'ImpEntrega'            );
   if TipoImpressora = tiRecibo              then Impressora := ReadPrintIni('ImpRecibo',              'ImpRecibo'             );
   if TipoImpressora = tiNotaPromissoria     then Impressora := ReadPrintIni('ImpNotaPromissoria',     'ImpNotaPromissoria'    );
   if TipoImpressora = tiListaSeparacao      then Impressora := ReadPrintIni('ImpListaSeparacao',      'ImpListaSeparacao'     );
   if TipoImpressora = tiContrato            then Impressora := ReadPrintIni('ImpContrato',            'ImpContrato'           );
   if TipoImpressora = tiEtiqueta            then Impressora := ReadPrintIni('ImpEtiqueta',            'ImpEtiqueta'           );
   if TipoImpressora = tiOrdemServico        then Impressora := ReadPrintIni('ImpOrdemServico',        'ImpOrdemServico'       );

   Result := Impressora;
end;

function ReadPrintIni(const Ident, Default: string): string;
var
   Aux: string;
   IniFile: TIniFile;
   Section: string;
begin
    Aux := ChangeFileExt(Application.ExeName, '.ini');

    IniFile := TIniFile.Create(UpperCase(Aux));
    try
       Section := 'Impressoras';
       if IniFile.SectionExists(Section + EmpresaLog) then
          Section := Section + EmpresaLog;

       Result := IniFile.ReadString(Section, Ident, Default);
    finally
       FreeAndNil(IniFile);
    end;
end;

procedure WritePrintIni(const Ident, Default: string);
var
   Aux: string;
   IniFile: TIniFile;
   Section: string;
begin
   Aux := ChangeFileExt(Application.ExeName, '.ini');

   IniFile := TIniFile.Create(UpperCase(Aux));
   try
      Section := 'Impressoras';
      if IniFile.SectionExists(Section + EmpresaLog) then
         Section := Section + EmpresaLog;

      IniFile.WriteString(Section, Ident, Default);
   finally
      FreeAndNil(IniFile);
   end;
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
function EAN14DV(EAN : string) : Char;
var Par, Impar, digito : integer;
    sEan : string;

    function N(I: byte): byte;
    begin
      Result := strToInt(sEAN[i]);
    end;

begin
   sEan   := ZerosLeft(ErtStrUtils.OnlyNumbers(EAN), 13);
   Par    := N(02)+N(04)+N(06)+N(08)+N(10)+N(12);
   Impar  := N(01)+N(03)+N(05)+N(07)+N(09)+N(11)+N(13);
   Impar  := Impar * 3;
   Par    := Par + Impar;
   Digito := Par + 10;
   Digito := digito div 10; // retirar a 3 casa
   Digito := digito *   10;
   Digito := digito - Par;
   if Digito = 10 then
      Result := '0'
   else
      Result := Char(Digito + 48);
end;

function EAN13DV(EAN : string) : Char;
var C : array[2..13] of Char;
    N : array[2..13] of Byte;
    F : Byte;
    D : string;
begin
   if Length(EAN) = 7 then begin
      EAN := StrZero(StrToFloat(EAN), 12);
   end else if Length(EAN) = 13 then begin
      Result := EAN14DV(EAN);
      Exit;
   end else if Length(EAN) <> 12 then begin
      EAN := StrZero(StrToFloat(EAN), 12);
   end;

   for F := 2 to 13 do
     C[F] := EAN[14-F];

   for F := 2 to 13 do
     N[F] := StrToInt(C[F]);

   D := IntToStr(10 - (((N[13]+N[11]+N[9]+N[7]+N[5]+N[3]) +
                       ((N[12]+N[10]+N[8]+N[6]+N[4]+N[2])*3)) mod 10));

   if Length(D) = 2 then
      Result := '0'
   else
      Result := D[1];
end;

function GeraDigitoEAN13(EAN : string) : String;
begin
   Result := EAN+EAN13DV(EAN);
end;

function GeraDigitoEAN13Grade(Produto : string; Sequencia: integer) : String;
var i : integer;
    s : string;
begin
   I := (99 * 10000) + Sequencia;
   S := ZerosLeft(I, 6) + ZerosLeft(Produto, 6);
   Result := GeraDigitoEAN13(s);
end;

function TestaEAN13(EAN:String) : Boolean;
var EANX :String;
begin
   if ecPrm.ParamAtivo(peLiberaEANProdutos) or (Copy(EAN,1,3) = '999') then begin
      Result := True;
      Exit;
   end;

   if Length(EAN) = 12 then begin
      Result := True;
   end else begin
      Try
         StrToFloat(EAN);
         EANX   := Copy(EAN, 1, Length(EAN) - 1);
         Result := GeraDigitoEAN13(EANX) = EAN;
      Except
         Result := False;
      End;
   end;
end;

function IsNumero(S : String) : Boolean;
var i: Integer;
begin
  Result := True;
  for i:= 1 to Length(s) do
  begin
    if (NOT (CharInSet(s[i],['0'..'9']))) and (s[i] <> '.') then
    begin
       Result := False;
       exit;
    end;
  end;
end;

function BetweenEco(const AValorVerificar, ValorInicial, ValorFinal: Integer): Boolean;
begin
   result := ((AValorVerificar >= ValorInicial) and (AValorVerificar <= ValorFinal));
end;

function ValidarPrefixoGTIN(const ACodigoEAN: String) :Boolean;
Const
    {Solitiação: 151802 Data: 05/06/2018 14:41:10 Usuário: Carlos.Schroder
     Detalhes: Lista de Prefixos Válidos por País disponivel em: http://www.nfe.fazenda.gov.br/portal/exibirArquivo.aspx?conteudo=mpYVEbsVRuE=}

  cPrefixo = Concat('000, 001, 002, 003, 004, 019, 020, 029, 030, 039, 040, 049, 050, 059, 060, 139, 200, 299, 300, 379, 380, 380, 383, 383, ',
                    '385, 385, 387, 387, 389, 389, 400, 440, 450, 459, 490, 499, 460, 469, 470, 470, 471, 471, 474, 474, ',
                    '475, 475, 476, 476, 477, 477, 478, 478, 479, 479, 480, 480, 481, 481, 482, 482, 483, 483, 484, 484, ',
                    '485, 485, 486, 486, 487, 487, 488, 488, 489, 489, 500, 509, 520, 521, 528, 528, 529, 529, 530, 530, ',
                    '531, 531, 535, 535, 539, 539, 540, 549, 560, 560, 569, 569, 570, 579, 590, 590, 594, 594, 599, 599, ',
                    '600, 601, 603, 603, 604, 604, 608, 608, 609, 609, 611, 611, 613, 613, 615, 615, 616, 616, 618, 618, ',
                    '619, 619, 620, 620, 621, 621, 622, 622, 623, 623, 624, 624, 625, 625, 626, 626, 627, 627, 628, 628, ',
                    '629, 629, 640, 649, 690, 699, 700, 709, 729, 729, 730, 739, 740, 740, 741, 741, 742, 742, 743, 743, ',
                    '744, 744, 745, 745, 746, 746, 750, 750, 754, 755, 759, 759, 760, 769, 770, 771, 773, 773, 775, 775, ',
                    '777, 777, 778, 779, 780, 780, 784, 784, 786, 786, 789, 790, 800, 839, 840, 849, 850, 850, 858, 858, ',
                    '859, 859, 860, 860, 865, 865, 867, 867, 868, 869, 870, 879, 880, 880, 884, 884, 885, 885, 888, 888, ',
                    '890, 890, 893, 893, 896, 896, 899, 899, 900, 919, 930, 939, 940, 949, 950, 950, 951, 951, 955, 955, ',
                    '958, 958, 960, 969, 977, 977, 978, 979, 980, 980, 981, 984, 990, 999' );
var
  vPrefixo: String;
  vPrefixoInt: Integer;
begin
  if Length(ACodigoEAN) = 14 then
  begin
    vPrefixo := Copy(ACodigoEAN, 2,3);
  end
  else
  begin
    vPrefixo := Copy(ACodigoEAN, 1, 3);
  end;
  Result := (Pos(vPrefixo, cPrefixo) > ZeroValue);

  if not Result then
  begin
    vPrefixoInt := StrToInteiro(vPrefixo);
    Result := ((BetweenEco(vPrefixoInt,0,19)) or  //GS1 US
               (BetweenEco(vPrefixoInt,20,29)) or //Números de circulação restrita dentro da região
               (BetweenEco(vPrefixoInt,30,39)) or //GS1 US
               (BetweenEco(vPrefixoInt,40,49)) or //GS1 Números de circulação restrita dentro da empresa
               (BetweenEco(vPrefixoInt,50,59)) or //GS1 US reserved for future use
               (BetweenEco(vPrefixoInt,60,139)) or //GS1 US
               (BetweenEco(vPrefixoInt,200,299)) or //GS1 Números de circulação restrita dentro da região
               (BetweenEco(vPrefixoInt,300,379)) or //GS1 France
               (BetweenEco(vPrefixoInt,400,440)) or //GS1 Germany
               (BetweenEco(vPrefixoInt,450,459)) or  //GS1 Japan
               (BetweenEco(vPrefixoInt,490,499)) or  //GS1 Japan
               (BetweenEco(vPrefixoInt,460,469)) or //GS1 Russia
               (BetweenEco(vPrefixoInt,500,509)) or //GS1 UK
               (BetweenEco(vPrefixoInt,540,549)) or //GS1 Belgium & Luxembourg
               (BetweenEco(vPrefixoInt,570,579)) or //GS1 Denmark
               (BetweenEco(vPrefixoInt,640,649)) or //GS1 Finland
               (BetweenEco(vPrefixoInt,690,699)) or //GS1 China
               (BetweenEco(vPrefixoInt,700,709)) or //GS1 Norway
               (BetweenEco(vPrefixoInt,730,739)) or //GS1 Sweden
               (BetweenEco(vPrefixoInt,760,769)) or //GS1 Schweiz, Suisse, Svizzera
               (BetweenEco(vPrefixoInt,800,839)) or //GS1 Italy
               (BetweenEco(vPrefixoInt,840,849)) or //GS1 Spain
               (BetweenEco(vPrefixoInt,870,879)) or //GS1 Netherlands
               (BetweenEco(vPrefixoInt,900,919)) or //GS1 Austria
               (BetweenEco(vPrefixoInt,930,939)) or //GS1 Australia
               (BetweenEco(vPrefixoInt,940,949)) or //GS1 New Zealand
               (BetweenEco(vPrefixoInt,960,969)) or //Global Office (GTIN-8s)
               (BetweenEco(vPrefixoInt,981,984)) or //GS1 Coupon identification for common currency areas
               (BetweenEco(vPrefixoInt,990,999))); //GS1 Coupon identification
  end;

end;

function CalcularDVGtin(ACodigoGTIN: String): String;
var
 Dig, I, DV: Integer;
begin
   DV := ZeroValue;
   Result := EmptyStr;
   // adicionar os zeros a esquerda, se não fizer isso o cálculo não bate
   // limite = tamanho maior codigo (gtin14) - 1 (digito)
   ACodigoGTIN := StrZero(ACodigoGTIN, 13);

   for I := Length(ACodigoGTIN) downto 1 do
   begin
     Dig := StrToInt(ACodigoGTIN[I]);
     DV  := DV + (Dig * IfThen(odd(I), 3, 1));
   end;

   DV := (Ceil(DV / 10) * 10) - DV ;
   Result := IntToStr(DV);
end;

function ValidarDigitoVerificadorGTIN(const ACodigoEAN: String; var ADigCalculado: String): Boolean;
var
  DigOriginal, Codigo: String;
begin
  if ((not(Length(ACodigoEAN) in [8, 12, 13, 14]))) then
  begin
    Result := False;
  end
  else
  begin
    Codigo       := Copy(ACodigoEAN, 1, Length(ACodigoEAN) - 1);
    DigOriginal  := ACodigoEAN[Length(ACodigoEAN)];
    ADigCalculado := CalcularDVGtin(Codigo);
    Result := (DigOriginal = ADigCalculado);
  end;
end;

function ValidarGTIN(const ACodigoEAN: String; var AMensagemErro: String; const AGTINTributavel: Boolean = False): Boolean;
var
  DigCalculado: String;
begin
  Result := True;
  AMensagemErro := EmptyStr;

  if ((not ecPrm.ParamAtivo(peLiberaEANProdutos)) AND
      (ACodigoEAN <> EmptyStr)) then
  begin
     if not IsNumero(ACodigoEAN) then
     begin
       AMensagemErro := Concat('Código GTIN (EAN)', IIF(AGTINTributavel, ' Tributável', EmptyStr), ' é inválido, o código GTIN (EAN) deve conter somente números.') ;
       Result := False;
     end
     else
     if (not AGTINTributavel) and ((not(Length(ACodigoEAN) in [8, 12, 13, 14]))) then
     begin
       AMensagemErro := 'Código GTIN (EAN) é inválido, o código GTIN deve ter 8, 12, 13 ou 14 caracteres.';
       Result := False;
     end
     else
     if (AGTINTributavel) and (not(Length(ACodigoEAN) in [8, 12, 13, 14])) then
     begin
       AMensagemErro := Concat('Código GTIN (EAN) Tributável é inválido, o código GTIN (EAN) Tributável deve ter 8, 12, 13 ou 14 caracteres.') ;
       Result := False;
     end
     else
     begin
        if not ValidarDigitoVerificadorGTIN(ACodigoEAN, DigCalculado) then
        begin
          AMensagemErro := COncat('O Dígito Verificador do Código GTIN (EAN)', IIF(AGTINTributavel, ' Tributável', EmptyStr), ' é inválido.');
          AMensagemErro := AMensagemErro + ' Dígito calculado: ' + DigCalculado;
          Result := False;
        end;
     end;

     if Result then
     begin
        if (Copy(ACodigoEAN, 1, 3) <> '999') and (not ValidarPrefixoGTIN(ACodigoEAN)) then
        begin
          AMensagemErro := Concat('Prefixo do Código GTIN (EAN)', IIF(AGTINTributavel, ' Tributável', EmptyStr), ' é inválido.');
          Result := False;
        end
     end;
  end;
end;

{$ENDIF}

function ExisteEAN(EAN: String; QueryPesquisa: TSQLQuery): Boolean;
const
  ASQL: String = 'SELECT GER.CODIGO, GER.CODIGOBARRA FROM TESTPRODUTOGERAL GER' + sLineBreak +
                 'INNER JOIN TESTPRODUTO PRD ON GER.CODIGO = PRD.PRODUTO' + sLineBreak +
                 'WHERE GER.CODIGOBARRA = :EAN AND PRD.EMPRESA = :EMPRESA' + sLineBreak +
                 'ORDER BY GER.CODIGOBARRA';
begin
   { SO 131275 - Alterado a SQL para que seja a feita a consulta de existência de um determinado
    EAN pela Empresa também...atualmente a consulta é feita no banco de dados inteiro, e agora
     é permitido ter o mesmo código de barra cadastrado em empresas diferentes }

   QueryPesquisa.Sql.Clear;
   QueryPesquisa.Sql.Add(ASQL);
   QueryPesquisa.ParamByName('EAN').AsString := EAN;
   QueryPesquisa.ParamByName('EMPRESA').AsString := Empresa;

   Result := EcoOpen(QueryPesquisa);

   if Result then
      Transporte := QueryPesquisa.FieldByName('CodigoBarra').asString;

end;

function ExisteEAN(EAN: String; QueryPesquisa: TFDQuery): Boolean;
const
  ASQL: String = 'SELECT GER.CODIGO, GER.CODIGOBARRA FROM TESTPRODUTOGERAL GER' + sLineBreak +
                 'INNER JOIN TESTPRODUTO PRD ON GER.CODIGO = PRD.PRODUTO' + sLineBreak +
                 'WHERE GER.CODIGOBARRA = :EAN AND PRD.EMPRESA = :EMPRESA' + sLineBreak +
                 'ORDER BY GER.CODIGOBARRA';
begin
   { SO 131275 - Alterado a SQL para que seja a feita a consulta de existência de um determinado
    EAN pela Empresa também...atualmente a consulta é feita no banco de dados inteiro, e agora
     é permitido ter o mesmo código de barra cadastrado em empresas diferentes }

   QueryPesquisa.Sql.Clear;
   QueryPesquisa.Sql.Add(ASQL);
   QueryPesquisa.ParamByName('EAN').AsString := EAN;
   QueryPesquisa.ParamByName('EMPRESA').AsString := Empresa;

   Result := EcoOpen(QueryPesquisa);

   if Result then
      Transporte := QueryPesquisa.FieldByName('CodigoBarra').asString;

end;

function GeraSequencia(Opcao: string; QueryPesquisa: TSQLQuery; AtualizaSequencia: Boolean = False): Integer;  overload;
var
   SequenciaTransaction: TDBXTransaction;
   Valor: string;
begin
   Result := 1;
   try
      SequenciaTransaction := FDataModule.sConnection.BeginTransaction;

      with QueryPesquisa do
      begin
         SQL.Clear;
         SQL.Add('SELECT ' + Opcao + ' FROM TGerSequencia WITH LOCK');

         if EcoOpen(QueryPesquisa) then
            Result := (FieldByName(Opcao).AsInteger + 1)
         else
         begin
            SQL.Clear;
            SQL.Add('INSERT INTO TGerSequencia (PRODUTO,CLIENTE,CONSUMIDOR,TRANSPORTADOR) VALUES (0, 0, 0, 0)');
            EcoExecSql(QueryPesquisa);
         end;
      end;

      if AtualizaSequencia then
      begin
         Valor := QZerosLeft(Result, 6);

         if (LowerCase(Opcao) = 'cliente') or (LowerCase(Opcao) = 'transportador') or (LowerCase(Opcao) = 'consumidor') then
            Valor := QZerosLeft(Result, 5);

         GravaSequencia(Opcao, Valor, QueryPesquisa);
      end;

      QueryPesquisa.Close;
      FDataModule.sConnection.CommitFreeAndNil(SequenciaTransaction);
   except
      QueryPesquisa.Close;
      FDataModule.sConnection.RollbackFreeAndNil(SequenciaTransaction);
      Result := 0;
   end;
end;

function GeraSequencia(Opcao: string; QueryPesquisa: TFDQuery; AtualizaSequencia: Boolean = False): Integer;  overload;
var
   LTransaction : IEcoTransaction;
   Valor: string;
begin
   Result := 1;
   try
      LTransaction := FDataModule.FDConnection.BeginTransaction;

      with QueryPesquisa do
      begin
         SQL.Clear;
         SQL.Add('SELECT ' + Opcao + ' FROM TGerSequencia WITH LOCK');

         if EcoOpen(QueryPesquisa) then
            Result := (FieldByName(Opcao).AsInteger + 1)
         else
         begin
            SQL.Clear;
            SQL.Add('INSERT INTO TGerSequencia (PRODUTO,CLIENTE,CONSUMIDOR,TRANSPORTADOR) VALUES (0, 0, 0, 0)');
            EcoExecSql(QueryPesquisa);
         end;
      end;

      if AtualizaSequencia then
      begin
         Valor := QZerosLeft(Result, 6);

         if (LowerCase(Opcao) = 'cliente') or (LowerCase(Opcao) = 'transportador') or (LowerCase(Opcao) = 'consumidor') then
            Valor := QZerosLeft(Result, 5);

         GravaSequencia(Opcao, Valor, QueryPesquisa);
      end;

      QueryPesquisa.Close;
      FDataModule.FDConnection.CommitFreeAndNil(LTransaction);
   except
      QueryPesquisa.Close;
      FDataModule.FDConnection.RollbackFreeAndNil(LTransaction);
      Result := 0;
   end;
end;


procedure GravaSequencia(Opcao, Valor: string; QueryPesquisa: TSQLQuery); overload;
begin
   with QueryPesquisa do
   begin
      Sql.Clear;
      Sql.Add(Format('UPDATE TGerSequencia SET %S = %S', [Opcao, Valor]));
      EcoExecSql(QueryPesquisa);
   end;
end;

procedure GravaSequencia(Opcao, Valor: string; QueryPesquisa: TFDQuery); overload;
begin
   with QueryPesquisa do
   begin
      Sql.Clear;
      Sql.Add(Format('UPDATE TGerSequencia SET %S = %S', [Opcao, Valor]));
      EcoExecSql(QueryPesquisa);
   end;
end;

Procedure PegaParamCC(QueryPesquisa : TSQLQuery);
begin
   with QueryPesquisa do begin
      Sql.Clear;
      Sql.Add('SELECT * FROM TGERCCPARAMETRO ' +
              'WHERE                         ' +
              'Empresa = :Empresa            ' );
      ParamByName('Empresa').asString := Empresa;
      if EcoOpen(QueryPesquisa) then begin
         MascConta       := FieldByName('Mascara').asString;
         EmpresaUsaCC    := iif(FieldByName('UsaCC').asString = 'S', True, False);
         UsaContaRevisao := iif(FieldByName('UsaContaRevisao').asString = 'S', True, False);
         ContaRevisao    := FieldByName('ContaRevisao').asString;
         Close;
      end;
   end;
end;

Procedure PegaParamCC(QueryPesquisa : TFDQuery);
begin
   with QueryPesquisa do begin
      Sql.Clear;
      Sql.Add('SELECT * FROM TGERCCPARAMETRO ' +
              'WHERE                         ' +
              'Empresa = :Empresa            ' );
      ParamByName('Empresa').asString := Empresa;
      if EcoOpen(QueryPesquisa) then begin
         MascConta       := FieldByName('Mascara').asString;
         EmpresaUsaCC    := iif(FieldByName('UsaCC').asString = 'S', True, False);
         UsaContaRevisao := iif(FieldByName('UsaContaRevisao').asString = 'S', True, False);
         ContaRevisao    := FieldByName('ContaRevisao').asString;
         Close;
      end;
   end;
end;

function BancoFazLancamentoNoFluxoFinanceiro(QueryAux: TSQLQuery): boolean;
var Aux : string;
begin
  Result := False;
  Aux := Format('SELECT p.OPDEPEMPRESA, o.OPERUSACC FROM TBANPARAMETRO p       '+
                'left outer join TBANOPERACAO o on (o.CODIGO = p.OPDEPEMPRESA) '+
                'where p.Empresa = %S', [QuotedStr(Empresa)]);
  with QueryAux do begin
    sql.clear;
    sql.add(Aux);
    if EcoOpen(QueryAux) then
      Result  := FieldByName('OPERUSACC').AsString = 'S';
  end;
end;

function BancoFazLancamentoNoFluxoFinanceiro(QueryAux: TFDQuery): boolean;
var Aux : string;
begin
  Result := False;
  Aux := Format('SELECT p.OPDEPEMPRESA, o.OPERUSACC FROM TBANPARAMETRO p       '+
                'left outer join TBANOPERACAO o on (o.CODIGO = p.OPDEPEMPRESA) '+
                'where p.Empresa = %S', [QuotedStr(Empresa)]);
  with QueryAux do begin
    sql.clear;
    sql.add(Aux);
    if EcoOpen(QueryAux) then
      Result  := FieldByName('OPERUSACC').AsString = 'S';
  end;
end;

function MontaContaCC(sMascaraDaConta, sContaCC: string): string;
var aux : string;
    K   : word;
begin
   Aux := ErtStrUtils.MaskFormatText(sMascaraDaConta + ';0; ', sContaCC);
   Aux := Trim(Aux);
   for k := Length(Aux) downto 1 do begin
      if (not CharInSet(Aux[k], ['0'..'9'])) then
         Aux[K] := ' '
      else
         Break;
   end;
   Result := aux;
end; 
// procedure para resolver o problema de foco da rotina CampoEmBranco
// para os casos de telas com TabSheet
// Moncks - Janeiro/2002

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure Verifica_TabSheet(Componente : TComponent);
var TabControle, TabAtivar, CompAux : TComponent;
begin
    CompAux := Componente;             // jogamos o componente para um auxiliar
    while CompAux.HasParent do begin   // enquanto existir ancestral...
        if (CompAux.GetParentComponent.ClassType = TTabSheet) then begin // se o tipo de classe do ancestral for TTabSheet então...
           TabAtivar   := CompAux.GetParentComponent;  // jogamos o ancestral (TabSheet) para variável de trabalho
           TabControle := CompAux.GetParentComponent.GetParentComponent; // jogamos o PageControl para variável de trabalho
           TPageControl(TabControle).ActivePage := TTabSheet(TabAtivar); // ativamos o tabsheet que queríamos
           break;                      // e caímos fora !
        end else if (CompAux.GetParentComponent.ClassType = TcxPageControl) then begin
           TabAtivar   := CompAux.GetParentComponent;
           if not (TabAtivar.HasParent) then begin
             TcxPageControl(TabAtivar).ActivePage := TcxTabSheet(CompAux);
             Break;
           end else
             TcxPageControl(TabAtivar).ActivePage := TcxTabSheet(CompAux);
        end;                                                              
        CompAux := CompAux.GetParentComponent; // jogamos o ancestral para o auxiliar
    end;
end;
{$ENDIF}

(*------------------------------------------------------------------------------
  retorna uma string contendo um valor por extenso com valor na faixa de bilhoes
  onde o tipo 1-reais   - especifica valores em moeda corrente do pais
              8-metros  - medidas metricas
              9-dolares
 ------------------------------------------------------------------------------ *)
function extenso(valor:extended; Tipo:Byte) :string;
Var Uc,Dc,Cc     : byte; // centenas
    Um,Dm,Cm     : byte; // milhar
    Ul,Dl,Cl     : byte; // milhoes
    Ub,Db,Cb     : byte; // bilhoes
    sMilSingular : String;
    sMilPlural   : String;
    sDezSingular : String;
    sDezPlural   : String;
    str          : string;
    ext          : string;
const
    tabU   : array[1..19] of string = ('um',        'dois',      'tres',
                                       'quatro',    'cinco',     'seis',
                                       'sete',      'oito',      'nove',
                                       'dez',       'onze',      'doze',
                                       'treze',     'quatorze',  'quinze',
                                       'dezesseis', 'dezessete', 'dezoito',
                                       'dezenove');
    tabD   : array[1..9] of string =  ('dez',         'vinte',
                                       'trinta',      'quarenta',
                                       'cinquenta',   'sessenta',
                                       'setenta',     'oitenta',
                                       'noventa');
    tabC   : array[1..9] of string =  ('cem',         'duzentos',
                                       'trezentos',   'quatrocentos',
                                       'quinhentos',  'seiscentos',
                                       'setecentos',  'oitocentos',
                                       'novecentos');

    function Un(Pos:byte):byte;      { retorna o valor do digito na posicao }
    begin
       result := strtoint(str[pos]);
    end;

    function MaisEouV(Val:word) :string; { retorna E ou virgula conforme val }
    begin
       if Val = 0 then
          Result := ' e '
       else
          Result :=  ', ';
    end;

    function MaisE(Val:word) :string; { gera um 'e' de ligacao se val > 0 }
    begin
      result := '';
      if val > 0 then result := ' e ';
    end;

    { retorna um grupo de valores centenas, milhares etc... }
    function Grupo(C,D,U:byte) :String;
    var tmp : string;
    begin
      tmp := '';
      if C > 0 then
         if (C = 1) and ((u+d) > 0) then
            tmp := 'cento'   // apenas uma centena
         else
            tmp := tabC[C];  // uma das centenas 1..9

      if (D * 10 + U) in [1..19] then   // se unidade,  1..19, (10 inclusive)
         tmp := tmp + MaisE(C) + TabU[D * 10 + U]
      else begin
         if D > 0 then       tmp := tmp + MaisE(C) + TabD[D];    // dezena  2..9
         if U in [1..9] then tmp := tmp + MaisE(D+C) + TabU[U];  // unidade 1..9
        end;
    result := tmp;  { retorna a string temporaria ou espaços }
    end;

begin
    sMilSingular := '';   { Assume com padrao}
    sMilPlural   := '';
    sDezSingular := '';
    sDezPlural   := '';
    Case Tipo of
       0..1: begin // reais
             sMilSingular := 'real';
             sMilPlural   := 'reais';
             sDezSingular := 'centavo';
             sDezPlural   := 'centavos';
          end;
       8: begin // Medida
             sMilSingular := 'metro';
             sMilPlural   := 'metros';
             sDezSingular := 'centimetro';
             sDezPlural   := 'centimetros';
          end;
       9: begin // Dolar
             sMilSingular := 'dolar';
             sMilPlural   := 'dolares';
             sDezSingular := 'cent';
             sDezPlural   := 'cents';
          end;
       end;

    Ext := ''; // sera armazenado aqui
    try                 {  123456789012345 -> posicao das casas BBB.LLL.MMM.CCC,UU }
      Str :=  FormatFloat('000000000000.00',Valor);
    except
      exit;
    end;

    Ub := Un(03); Db := Un(02); Cb := Un(01);  //bilhoes  - uni,dez,cen
    Ul := Un(06); Dl := Un(05); Cl := Un(04);  //milhoes
    Um := Un(09); Dm := Un(08); Cm := Un(07);  //milhares
    Uc := Un(12); Dc := Un(11); Cc := Un(10);  //centenas

    if (Cb+Db+Ub) > 0 then begin      { trata do grupo BBB - bilhoes }
       ext := grupo(Cb,Db,Ub);
       if ((Cb + Db) = 0) and (Ub = 1) then
          ext := ext + ' bilhao'      { Aqui o grupo de valor }
       else
          ext := ext + ' bilhoes';    { plural }

      if (Ul+Dl+Cl+Um+Dm+Cm+Uc+Dc+Cc) = 0 then  { Se o restante do valor = 0 }
          ext := ext + ' de '
       else
          ext := ext + MaisEouV(CC);
     end;

    if (Cl+Dl+Ul) > 0 then begin                { trata o grupo LLL - milhoes  }
       ext := ext + grupo(CL,DL,UL);
       if ((Cl + Dl) = 0) and (Ul = 1) then
           ext := ext + ' milhao'
       else
           ext := ext + ' milhoes';

       if (Um+Dm+Cm+Uc+Dc+Cc) = 0 then
          ext := ext + ' de '
       else
          ext := ext + MaisEouV(CC);
    end;


    if (Cm+Dm+Um) > 0 then begin          { trata o grupo MMM - milhares }
       ext := ext + grupo(Cm,Dm,Um);
       if ((Cm + Dm) = 0) and (Um = 1) then
           ext := ext + ' mil'
       else
           ext := ext + ' mil';

       if (Uc+Dc+Cc) = 0 then
          ext := ext + ' ' + sMilPlural  // moeda
       else
          ext := ext + MaisEouV(CC);
    end;

    if Uc+Dc+Cc > 0 then begin             { trata o grupo CCC - milhoes  }
       ext := ext + Grupo(Cc,Dc,Uc);
       if ((Cc+ Dc) = 0) and (Uc = 1) then
           ext := ext + ' ' + sMilSingular   // moeda
       else
           ext := ext + ' ' + sMilPlural;    // moeda
    end;

    Uc := Un(15);    // retira campos para os centavos
    Dc := Un(14);

    if  (Valor > 0.99) and ((Uc+Dc) > 0) then ext := ext + ' e ';
    if  (dc+uc) = 1 then
        ext := ext + grupo(0,Dc,Uc) + ' ' + sDezSingular   // singular
    else
        if Uc+Dc > 1 then
           ext := ext + grupo(0,Dc,Uc) + ' ' + sDezPlural; // plural
   result := ext;

end;

// retorna uma string de estenso preenchida com * no final
function PreencheExtenso(Tam:Byte;Valor:Extended;Tipo:Byte):String;
var aux : string;
    tmp : string;
    K   : word;
begin
     aux := Extenso(Valor,Tipo);
     if (length(aux) > tam) or (length(aux) > 254) or (tam > 253) then begin
        result := aux;
        exit;
        end;
     tmp := aux+' ';
     for K:= length(aux) to tam do
         tmp := tmp + '*';
     result := tmp;
end;

{ Retorna uma string referente uma linha especifica do texto passado como parametro
  Texto:         Texto a ser comparado
  LengthLines:   Tamanho das linhas
  LengthPriLine: Tamanho da primeira linha
  LineReturn:    Linha a ser retornada
  JustUltLine:   Se justifica ou não a ultima linha do texto }
function Justifica(Texto:String; LengthLines:Integer; LengthPriLine:Integer; LineReturn:Integer; JustUltLine:Boolean):String;
Var Tmp         :String;
    Contador    :Integer;
    Contador1   :Integer;
    Controle    :Boolean;
    UltimaLinha :Boolean;
    TTLines     :Integer;
    Lines       :Array[1..99] of string;
begin
   if LineReturn > 99 then LineReturn := 99;

   while Pos('  ',Texto)> 0 do Delete(Texto, Pos('  ',Texto), 1); // tira um espaco onde tiver dois ou mais espaços

   // Primeira Linha
   Contador := LengthPriLine+1;
   while ((Copy(Texto,Contador, 1) <> ' ') and (Copy(Texto,Contador, 1) <> '') and (Contador > 0)) do Dec(Contador);
   if Contador = 0 then Contador := LengthPriLine+1;
   Lines[1] := Copy(Texto,1,Contador-1);
   Tmp      := Trim(Copy(Texto,Contador, Length(Texto)));

   // Linhas restantes
   TTLines := 1;
   while (TTLines <= 99) and (Length(Tmp) > 0) do begin

      Inc(TTLines);
      Contador1 := LengthLines+1;
      if Pos(#$D#$A,Copy(Tmp,1,Contador1-1))>0 then begin
         Contador1 := Pos(#$D#$A,Copy(Tmp,1,Contador1-1))+2;
      end else begin
         while ((Copy(Tmp,Contador1, 1) <> ' ') and (Copy(Tmp,Contador1, 1) <> '') and (Contador1 > 0)) do Dec(Contador1);
         if Contador1 = 0 then Contador1 := LengthLines+1;
      end;
      Lines[TTLines] := Copy(Tmp,1,Contador1-1);
      Tmp            := Trim(Copy(Tmp,Contador1, Length(Tmp)));

   end;

   // Justifica
   if LineReturn = 1 then Contador := LengthPriLine else Contador := LengthLines;
   Tmp         := Lines[LineReturn];
   UltimaLinha := (LineReturn=TTLines);
   if (Pos(#$D#$A,Tmp)>0) then begin
      Tmp := Copy(Tmp,1,Pos(#$D#$A, Tmp)-1);
      UltimaLinha := True;
   end;
   if (Tmp <> '') and (not UltimaLinha or JustUltLine) then begin
      Contador1 := Length(Tmp);
      Controle  := True;
      while Length(Tmp) < Contador do begin
         if Copy(Tmp, Contador1, 1) = ' ' then begin
            if Controle then Insert(' ', Tmp, Contador1);
            Controle := False;
         end else begin
            Controle := True;
         end;
         Dec(Contador1);
         if Contador1 = 0 then begin
            Contador1 := Length(Tmp);
            if Contador1 = Length(Lines[LineReturn]) then Break;
         end;
      end;
   end;
   Result := Tmp;
end;


procedure MontaExtensoCheques(const Valor: extended; var vExtenso1, vExtenso2: string);
var vStr, vExt: string;
begin
   vStr := MascaraValor(Valor, 2);
   vStr := StrRight(vStr, 11);
   vStr := '('+ AnsiReplaceText(vStr, ' ', '#') + ') ';  // 14

   vExt      := Extenso(Valor,0);
   vExtenso1 := Copy(vExt, 1, 40);

   if Length(vExt) <= 40 then begin
      vExtenso1 := vExtenso1 + Replique('*', 40 - Length(vExtenso1));
      vExtenso2 :=             Replique('*', 58);
   end else begin
      vExtenso2 := Copy(vExt, 41, 58);
      if Length(vExtenso2) < 58 then begin
         vExtenso2 := vExtenso2 + Replique('*', 58 - Length(vExtenso2));
      end;
   end;
   vExtenso1 := UpperCase(vStr + vExtenso1);
   vExtenso2 := UpperCase(vExtenso2);
end;


function JustificaTexto(Texto:String; TamLines:Integer; TabPrimLine:Integer):TStringList;
Var Posicao   :Integer;
    Paragrafo :String;
    Cont      :Integer;
begin
   Result := TStringList.Create;
   while Pos(#$D#$A, Texto)>0 do begin
      Posicao   := Pos(#$D#$A,    Texto);
      Paragrafo := Copy(Texto,0,Posicao-1);
      Cont := 1;
      while Justifica(Paragrafo, TamLines, TamLines, Cont, False)<>'' do begin
         if Cont = 1 then begin
            Result.Add(Espaco(TabPrimLine)+Justifica(Paragrafo, TamLines, TamLines-TabPrimLine, Cont, False));
         end else begin
            Result.Add(Justifica(Paragrafo, TamLines, TamLines, Cont, False));
         end;
         Inc(Cont);
      end;
      Texto   := Trim(Copy(Texto,Posicao+2, Length(Texto)));
   end;

   Result.Add(Espaco(TabPrimLine)+Justifica(Texto, TamLines, TamLines-TabPrimLine, 1, False));
// Result.Add(Justifica(Texto, TamLines, TamLines, 1, False));

   Cont := 2;
   while Justifica(Texto, TamLines, TamLines, Cont, False)<>'' do begin
      Result.Add(Justifica(Texto, TamLines, TamLines, Cont, False));
      Inc(Cont);
   end;
end;


{ Usado para codificar valores  Tipo: Custo codificado }
function Codificar(vsStr, vsCod :String) :String;
var K : Integer;
begin
  Result := '';
  if trim(vsCod) = '' then
     Result := vsStr
  else
     for K := 1 to Length(vsStr) do
        if CharInSet(vsStr[k], ['0'..'9']) then
           Result := Result + vsCod[StrToInt(vsStr[k]) + 1]
        else
          if CharInSet(vsStr[k], [',','.','-']) then
             Result :=  Result + vsStr[k];
end;

{ retorna o 1º ou o 2º parametro de acordo com o valor passado em value }
function IIF(Value : Boolean; Verdadeiro, Falso : Variant): Variant;
begin
   if Value then Result := Verdadeiro else Result := Falso;
end;

function Lower(Text:String): String;
var Ind: Integer;
const LW = 'áâãàéêíóôõúüûçñ';
      UP = 'ÁÂÃÀÉÊÍÓÔÕÚÜÛÇÑ';
begin
  Result := '';
  for Ind := 1 to Length(Text) do
      if Pos(Copy(Text, Ind, 1), UP) > 0 then
         Result := Result + Copy(LW, Pos(Copy(Text, Ind, 1), UP), 1)
      else
         Result := Result + LowerCase(Copy(Text, Ind, 1));
end;

function Upper(Text: String): String;
var Ind: Integer;
const LW = 'áâãàéêíóôõúüûçñ';
      UP = 'ÁÂÃÀÉÊÍÓÔÕÚÜÛÇÑ';
begin
  Result := '';
  for Ind := 1 to Length(Text) do
      if Pos(Copy(Text, Ind, 1), LW) > 0 then
         Result := Result + Copy(UP, Pos(Copy(Text, Ind, 1), LW), 1)
      else
         Result := Result + UpperCase(Copy(Text, Ind, 1));
end;

function BetWeen(Const Compare, Valor1, Valor2:Variant):Boolean;
begin
   Result := (Compare>=Valor1) and (Compare<=Valor2);
end;

{ Verifica se existe uma impressora instalada no windows com o nome
  igual ao parametro passado em impressora, se existir retorna o numero dela }
function FindPrinter(const Impressora: string): Integer;
var K: Integer;
begin
   Result := -1;
   with Printer do begin
     for K := 0 to (Printers.Count -1) do begin
       if UpperCase(Printers.Strings[K]) = UpperCase(Impressora) then begin
         Result := K;
         Break;
       end;
     end;
   end;
   ImprimeSemSpoolWin := Pos('LQ-570', UpperCase(Impressora)) = 0;
end;

{ procedure para definir o tamanho da pagina da impressora }
procedure SetPrinterPage(PaperSize, Width, Height : LongInt);
var
   ADevice, ADriver, APort: array[0..255] of char;
   DeviceMode: THandle;
   M: PDevMode;
   S:String;
begin
   // Força o uso de Printer. Se esta linha for removida, a primeira
   // invocação falha. Bug da VCL
   S := Printer.Printers[Printer.PrinterIndex];
   // Pega dados da impressora atual
   Printer.GetPrinter(ADevice, ADriver, APort, DeviceMode);
   // Pega um ponteiro para DEVMODE
   M := GlobalLock(DeviceMode);
   try
      if M <> nil then begin
         // Muda tamanho do papel
         M^.dmFields      := DM_PAPERSIZE;
         if PaperSize = DMPAPER_USER then M^.dmFields := M^.dmFields or DM_PAPERLENGTH or DM_PAPERWIDTH;
         M^.dmPaperLength := Height;
//         M^.dmPaperWidth  := Width;
         M^.dmPaperSize   := PaperSize;
         // Atualiza
         Printer.SetPrinter(ADevice, ADriver, APort, DeviceMode);
      end;
   finally
      GlobalUnlock(DeviceMode);
   end;
end;

procedure PrintCmd(Value: TPrintCmd; Const ImprimeDireto: Boolean=false);
begin
{   if ImprimeDireto then begin
       case value of
         tpCmd20cpi : Write(F, lg20Cpp);
         tpCmd17cpi : Write(F, lg17Cpp);
         tpCmd12cpi : Write(F, lg12Cpp);
         tpCmd10cpi : Write(F, lg10Cpp);
         tpCmd06cpi : Write(F, lg12Cpp + lgCondensado);
         tpCmd05cpi : Write(F, lg10Cpp + lgCondensado);
         tpCmdLgNeg : Write(F, lgEnfatizado);
         tpCmdDgNeg : Write(F, dgEnfatizado);
       end;
     end;
     Exit;
   end;
}
   if Value = tpCmdNeg10 then begin
      Printer.Canvas.Font.Name  := 'Courier New';
      Printer.Canvas.Font.Size  := 12;
   end else if Value = tpCmdNeg12 then begin
      Printer.Canvas.Font.Name  := 'Courier New';
      Printer.Canvas.Font.Size  := 10;
   end else begin
      if Value = tpCmd20cpi then Printer.Canvas.Font.Name  := 'Draft 20Cpi';
      if Value = tpCmd17cpi then Printer.Canvas.Font.Name  := 'Draft 17Cpi';
      if Value = tpCmd12cpi then Printer.Canvas.Font.Name  := 'Draft 12Cpi';
      if Value = tpCmd10cpi then Printer.Canvas.Font.Name  := 'Draft 10Cpi';
      if Value = tpCmd06cpi then Printer.Canvas.Font.Name  := 'Draft 6Cpi';
      if Value = tpCmd05cpi then Printer.Canvas.Font.Name  := 'Draft 5Cpi';
   end;
   { Negrito / Desnegrito}
   if Value = tpCmdLgNeg  then Printer.Canvas.Font.Style := [fsBold];
   if Value = tpCmdDgNeg  then Printer.Canvas.Font.Style := [];
end;


{ procedure para definir o tamanho da pagina da impressora }
function MudaTamPapel(const Paper: TPapel): TPapel;
var
  S: string;
  HDevMode: THandle;
  DevMode: PDeviceMode;
  Device, Driver, Port: PChar;
begin
  { Bug VCL - A Impressora selecionada às vezes falha. Com
    esta instrução funciona adequadamente. }
  S := Printer.Printers[Printer.PrinterIndex];
  { Aloca memória para as variáveis PChar }
  GetMem(Device, 255);
  GetMem(Driver, 255);
  GetMem(Port,   255);
  try
    { Obtém dados da impressora atual }
    Printer.GetPrinter(Device, Driver, Port, HDevMode);

    { Aloca ponteiro }
    DevMode := GlobalLock(HDevMode);
    try
       if DevMode <> nil then begin
          with DevMode^ do begin

             { Salva tamanho atual }
             Result.Size   := dmPaperSize;
             Result.Width  := dmPaperWidth;
             Result.Height := dmPaperLength;

             { Define o novo tamanho }
             dmPaperSize := Paper.Size;
             dmFields    := dmFields or DM_PAPERSIZE;

             { Se for tamanho personalizado... }
             if Paper.Size = DMPAPER_USER then begin

               { Define altura }
               dmPaperLength := Paper.Height;
               dmFields      := dmFields or DM_PAPERLENGTH;

               { Define largura }
               dmPaperWidth  := Paper.Width;
               dmFields      := dmFields or DM_PAPERWIDTH;
             end;
          end;

          { Aplica as modificações }
          Printer.SetPrinter(Device, Driver, Port, HDevMode);
       end else
         raise Exception.Create('Erro ao definir tamanho de papel.');
    finally
      { Desaloca ponteiro }
      GlobalUnlock(HDevMode)
    end;
  finally
    { Desaloca a memória das variáveis PChar }
    FreeMem(Device, 255);
    FreeMem(Driver, 255);
    FreeMem(Port,   255);
  end;
end;

// procedure para impressão de bloquetos e duplicatas
// a partir de arquivo texto de definições no Spool do windows
// com TPrinter
// Moncks - Julho/2002
{ Parâmetros passados para a procedure:
  Arq........--> arquivo de definições (Ex.: BB.txt)
  TipoImp....--> tipo de Impressora (Ex.: tiBloqueto)
  Query1.....--> Dados do cabeçalho
  Query2.....--> Dados de linha-detalhe
}
procedure GeraImpressao2(Arq: String; TipoImp : TImpressora; Query1, Query2: TSimpleDataSet);

   procedure PegaDef;                                           // procedure para pegar definições do arquivo texto
      var ArqDef : TextFile;                                     // declaração de variáveis auxiliares
   begin
        cLin  := 0;                                              // inicializa contador de linhas
        slDEF := TStringList.Create;                             // cria StringList de definições
        AssignFile(ArqDef, Arq);                                 // atribui direcionamento do Parâmetro para o arquivo
        Reset(ArqDef);                                           // define arquivo como de leitura
        ReadLn(ArqDef, Linha);                                   // lê primeira linha do arquivo
        if Linha <> '#INICIO' then begin                         // se não for INICIO então não é arq. de definições válido,
           raise Exception.Create(Arq +' não é um arquivo de definições válido!');
        end;
        while True do begin                                      // looping para ler todo arquivo linha a linha...
            ReadLn(ArqDef, Linha);                               // lê próxima linha...
            if Linha = '#FINAL' then Break;                      // se for FINAL... caímos fora!
            if pos('#D', Linha) > 0 then begin                      // pega definição de formulário...
               iSize        := StrToInt(copy(Linha,pos('#D', Linha)+3,3));       // pega o tipo de página
               iWidth       := StrToInt(copy(Linha,pos('#D', Linha)+7,4));       // pega a largura da página
               iHeight      := StrToInt(copy(Linha,pos('#D', Linha)+12,4));      // pega o comprimento da página
               iAlturaLinha := StrToInt(copy(Linha,pos('#D', Linha)+17,2));      // pega a altura da linha (sextos=24 ou oitavos=18)
               Titulo       := copy(Linha,pos('#D', Linha)+21,20);               // pega o título do relatório
               Tamanho      := copy(Linha,pos('#D', Linha)+44,10);               // pega o tamanho da fonte
               if copy(Linha,pos('#D', Linha)+57,11) = 'poPortrait ' then begin   // pega a orientação da página (poPortrait=vertical ou poLandscape=horizontal)
                  Orientacao := poPortrait;
               end else begin
                  Orientacao := poLandscape;
               end;
            end else begin
               slDEF.Add(Linha);                                 // adiciona linha no StringList de definições
            end;
        end;
   end;

   procedure Inicia;
   var Papel : TPapel;
   begin
       if not Printer.Printing then begin
          Papel.Size                  := iSize;
          Papel.Width                 := iWidth;
          Papel.Height                := iHeight;
          Printer.Title               := Titulo;
          PrintCmd(TipoCmd(Tamanho));
          Printer.Canvas.Font.Pitch   := fpFixed;
          Printer.Orientation         := Orientacao;
          MudaTamPapel(Papel);
          LarguraCaracter             := Printer.Canvas.TextWidth('a');
          LarguraPagina               := (Printer.PageWidth  div LarguraCaracter) - 1;
          AlturaLinha                 := iAlturaLinha;
          AlturaPagina                := (Printer.PageHeight div AlturaLinha    ) - 1;
          PCol                        := 0;
          PRow                        := 1;
          Printer.BeginDoc;
       end;
   end;

   procedure SetText(Row, Col : Integer; Texto : String);
   begin
      QuebrouPagina := False;
      if (Row >= AlturaPagina) then begin
         Printer.NewPage;
         QuebrouPagina := True;
         Row := 0;
      end;
      PCol := Col;
      PRow := Row;
      With Printer.Canvas do begin
         Col := (Col * LarguraCaracter);
         Row := (Row * AlturaLinha);
         TextOut(Col, Row, Texto);   { Texto impresso }
      end;
   end;

   procedure Finaliza;
   var Papel : TPapel;
   begin
      { Finaliza a impressao se tiver imprimindo }

      if Printer.Printing then Printer.EndDoc;

      { Volta tamanho do papel }
      Papel.Size                  := 256;
      Papel.Width                 := 2159;
      Papel.Height                := 2794;
      Printer.Canvas.Font.Name    := 'Draft 10Cpi'; // Tamanho normal
      MudaTamPapel(Papel);
      LarguraCaracter := Printer.Canvas.TextWidth('a');
      LarguraPagina   := (Printer.PageWidth  div LarguraCaracter) - 1;
      AlturaLinha     := 24;
      AlturaPagina    := (Printer.PageHeight div AlturaLinha    ) - 1;
      PCol                        := 0;
      PRow                        := 1;


   end;

var i,                                                           // declaração de variáveis auxiliares
    Lin, Col, Tam,
    LIni, LFim,
    LdtCont,
    Contador      : integer;
    sx,Campo,
    Mascara,
    Atributo      : string;
    sAux          : string;
begin
  if FindPrinter(PortaImpressora(TipoImp)) < 0 then begin
     Mensagem('Impressora não encontrada', tmInforma);
     Exit;
  end else begin
     Printer.PrinterIndex := FindPrinter(PortaImpressora(TipoImp));
  end;
  if not FileExists(Arq) then begin
     raise Exception.Create('O arquivo '+Arq +' não foi encontrado!');
  end;
  PegaDef;                                                       // chama procedure para pegar definições do arquivo texto
  Inicia;
  Query1.FindFirst;
  Contador := 0;
  while not Query1.Eof do begin
    i := 0;                                                      // inicializa contador de linhas
    LIni  := 0;
    LFim  := 0;
    while i <= slDEF.Count-1 do begin                            // enquanto existir linhas de definição...
      sx  := slDEF.Strings[i];                                   // usa variável string para trabalhar
      if pos('#A', sx) > 0 then begin                         // se encontrar caracter de atributo nesta linha...
         Atributo := copy(sx,pos('#A', sx)+3,10);               // pega atributo do campo
         PrintCmd(TipoCmd(Atributo));
      end;
      if pos('#P', sx) > 0 then begin                            // se existir posicionamento de campos...
         Lin      := StrToInt(copy(sx,pos('#P', sx)+3,3));       // pega o número da Linha do arq. de definições
         Col      := StrToInt(copy(sx,pos('#P', sx)+7,3));       // pega o número da coluna do arq. de definições
         Tam      := StrToInt(copy(sx,pos('#P', sx)+11,3));      // pega o tamanho do campo do arq. de definições
         Campo    := copy(sx,pos('<', sx)+1,(pos('>', sx)-(pos('<', sx)+1))); // pega o nome do campo do arq. de definições
         Mascara  := copy(sx,pos('[', sx)+1,(pos(']', sx)-(pos('[', sx)+1))); // pega a máscara do campo do arq. de definições
         if copy(Campo,1,1) <> '''' then begin                   // se não for uma constante
           try
             Query1.FieldByName(Campo).FieldNo;                  // testa existência do campo
           except
             ShowMessage('Nome de Campo Inválido -> '+Campo);
             exit;
           end;
         end;
         // insere campo do tamanho definido na linha e coluna correspondente
         if copy(Campo,1,1) <> '''' then begin                         // se não for uma constante...
           sAux := Copy(Query1.FieldByName(Campo).Value,1,Tam);
           if Mascara <> '' then begin
              SetText(Lin, Col, Format(Mascara, [StrToFloat(DesmascaraNumero(sAux))])); // insere mascarado
           end else begin
              SetText(Lin, Col, sAux);                               // insere sem máscara
           end;
         end else begin                                              // se não é porque é uma constante, então...
           SetText(Lin, Col, copy(Campo,2,Tam))    ;                 // insere constante sem máscara
         end;
      end else begin
     // ******************************************************************************************
     //   verifica se é definição de linha detalhe para pegar valores das linhas de início e fim
     // ******************************************************************************************
          if pos('#L', sx) > 0 then begin                        // se for definição de linha-detalhe...
             LIni   := StrToInt(copy(sx,pos('#L', sx)+3,3));     // pega o número da Linha Inicial do looping do arq. de definições
             LFim   := StrToInt(copy(sx,pos('#L', sx)+7,3));     // pega o número da Linha Final do looping do arq. de definições
          end;
     // ***********************************************
     //   verifica se é linha detalhe para fazer loop
     // ***********************************************
          if pos('#X', sx) > 0 then begin                        // se for definição de posicionamento de campo de linha-detalhe...
             LdtCont   := LIni;                                  // Carrega contador de linhas-detalhe
             Query2.First;                                       // posiciona no primeiro registro do arquivo de produtos da Nota
             while (LdtCont <= LFim) and                         // enquanto for linha-detalhe e
                   (not Query2.Eof) do begin                     // não for final de arquivo de produtos da Nota então...
                if pos('#A', sx) > 0 then begin                  // se encontrar caracter de atributo nesta linha...
                   Atributo := copy(sx,pos('#A', sx)+3,10);         // pega atributo do campo
                   PrintCmd(TipoCmd(Tamanho));
                end;
                Lin    := LdtCont;                               // Define o número da linha através do looping
                Col    := StrToInt(copy(sx,pos('#X', sx)+3,3));  // pega o número da coluna do arq. de definições
                Tam    := StrToInt(copy(sx,pos('#X', sx)+7,3));  // pega o tamanho do campo do arq. de definições
                Campo  := copy(sx,pos('<', sx)+1,(pos('>', sx)-(pos('<', sx)+1))); // pega o nome do campo do arq. de definições
                Mascara := copy(sx,pos('[', sx)+1,(pos(']', sx)-(pos('[', sx)+1))); // pega a máscara do campo do arq. de definições
                // insere campo do tamanho definido na linha e coluna correspondente
                if Mascara <> '' then begin                      // se foi definida máscara para o campo, então...
                   SetText(Lin, Col, format(mascara, [StrToFloat(copy(Query2.FieldByName(Campo).Value,1,Tam))])); // insere mascarado
                end else begin                                   // se não...
                   SetText(Lin, Col, copy(Query2.FieldByName(Campo).Value,1,Tam)); // insere sem máscara
                 end;
                Inc(LdtCont);                                    // incrementa contador de linhas-detalhe
                Query2.Next;                                     // lê próximo registro do arquivo de produtos da Nota
             end;
          end;
      end;
      Inc(i);                                                    // incrementa contador de linhas
    end;
    Query1.Next;
    Inc(Contador);
    if Contador < Query1.RecordCount then begin
       Printer.NewPage;
    end;
  end;
  Finaliza;
end;

procedure GeraImpressao2(Arq: String; TipoImp : TImpressora; Query1, Query2: TFDMemTable);

   procedure PegaDef;                                           // procedure para pegar definições do arquivo texto
      var ArqDef : TextFile;                                     // declaração de variáveis auxiliares
   begin
        cLin  := 0;                                              // inicializa contador de linhas
        slDEF := TStringList.Create;                             // cria StringList de definições
        AssignFile(ArqDef, Arq);                                 // atribui direcionamento do Parâmetro para o arquivo
        Reset(ArqDef);                                           // define arquivo como de leitura
        ReadLn(ArqDef, Linha);                                   // lê primeira linha do arquivo
        if Linha <> '#INICIO' then begin                         // se não for INICIO então não é arq. de definições válido,
           raise Exception.Create(Arq +' não é um arquivo de definições válido!');
        end;
        while True do begin                                      // looping para ler todo arquivo linha a linha...
            ReadLn(ArqDef, Linha);                               // lê próxima linha...
            if Linha = '#FINAL' then Break;                      // se for FINAL... caímos fora!
            if pos('#D', Linha) > 0 then begin                      // pega definição de formulário...
               iSize        := StrToInt(copy(Linha,pos('#D', Linha)+3,3));       // pega o tipo de página
               iWidth       := StrToInt(copy(Linha,pos('#D', Linha)+7,4));       // pega a largura da página
               iHeight      := StrToInt(copy(Linha,pos('#D', Linha)+12,4));      // pega o comprimento da página
               iAlturaLinha := StrToInt(copy(Linha,pos('#D', Linha)+17,2));      // pega a altura da linha (sextos=24 ou oitavos=18)
               Titulo       := copy(Linha,pos('#D', Linha)+21,20);               // pega o título do relatório
               Tamanho      := copy(Linha,pos('#D', Linha)+44,10);               // pega o tamanho da fonte
               if copy(Linha,pos('#D', Linha)+57,11) = 'poPortrait ' then begin   // pega a orientação da página (poPortrait=vertical ou poLandscape=horizontal)
                  Orientacao := poPortrait;
               end else begin
                  Orientacao := poLandscape;
               end;
            end else begin
               slDEF.Add(Linha);                                 // adiciona linha no StringList de definições
            end;
        end;
   end;

   procedure Inicia;
   var Papel : TPapel;
   begin
       if not Printer.Printing then begin
          Papel.Size                  := iSize;
          Papel.Width                 := iWidth;
          Papel.Height                := iHeight;
          Printer.Title               := Titulo;
          PrintCmd(TipoCmd(Tamanho));
          Printer.Canvas.Font.Pitch   := fpFixed;
          Printer.Orientation         := Orientacao;
          MudaTamPapel(Papel);
          LarguraCaracter             := Printer.Canvas.TextWidth('a');
          LarguraPagina               := (Printer.PageWidth  div LarguraCaracter) - 1;
          AlturaLinha                 := iAlturaLinha;
          AlturaPagina                := (Printer.PageHeight div AlturaLinha    ) - 1;
          PCol                        := 0;
          PRow                        := 1;
          Printer.BeginDoc;
       end;
   end;

   procedure SetText(Row, Col : Integer; Texto : String);
   begin
      QuebrouPagina := False;
      if (Row >= AlturaPagina) then begin
         Printer.NewPage;
         QuebrouPagina := True;
         Row := 0;
      end;
      PCol := Col;
      PRow := Row;
      With Printer.Canvas do begin
         Col := (Col * LarguraCaracter);
         Row := (Row * AlturaLinha);
         TextOut(Col, Row, Texto);   { Texto impresso }
      end;
   end;

   procedure Finaliza;
   var Papel : TPapel;
   begin
      { Finaliza a impressao se tiver imprimindo }

      if Printer.Printing then Printer.EndDoc;

      { Volta tamanho do papel }
      Papel.Size                  := 256;
      Papel.Width                 := 2159;
      Papel.Height                := 2794;
      Printer.Canvas.Font.Name    := 'Draft 10Cpi'; // Tamanho normal
      MudaTamPapel(Papel);
      LarguraCaracter := Printer.Canvas.TextWidth('a');
      LarguraPagina   := (Printer.PageWidth  div LarguraCaracter) - 1;
      AlturaLinha     := 24;
      AlturaPagina    := (Printer.PageHeight div AlturaLinha    ) - 1;
      PCol                        := 0;
      PRow                        := 1;


   end;

var i,                                                           // declaração de variáveis auxiliares
    Lin, Col, Tam,
    LIni, LFim,
    LdtCont,
    Contador      : integer;
    sx,Campo,
    Mascara,
    Atributo      : string;
    sAux          : string;
begin
  if FindPrinter(PortaImpressora(TipoImp)) < 0 then begin
     Mensagem('Impressora não encontrada', tmInforma);
     Exit;
  end else begin
     Printer.PrinterIndex := FindPrinter(PortaImpressora(TipoImp));
  end;
  if not FileExists(Arq) then begin
     raise Exception.Create('O arquivo '+Arq +' não foi encontrado!');
  end;
  PegaDef;                                                       // chama procedure para pegar definições do arquivo texto
  Inicia;
  Query1.FindFirst;
  Contador := 0;
  while not Query1.Eof do begin
    i := 0;                                                      // inicializa contador de linhas
    LIni  := 0;
    LFim  := 0;
    while i <= slDEF.Count-1 do begin                            // enquanto existir linhas de definição...
      sx  := slDEF.Strings[i];                                   // usa variável string para trabalhar
      if pos('#A', sx) > 0 then begin                         // se encontrar caracter de atributo nesta linha...
         Atributo := copy(sx,pos('#A', sx)+3,10);               // pega atributo do campo
         PrintCmd(TipoCmd(Atributo));
      end;
      if pos('#P', sx) > 0 then begin                            // se existir posicionamento de campos...
         Lin      := StrToInt(copy(sx,pos('#P', sx)+3,3));       // pega o número da Linha do arq. de definições
         Col      := StrToInt(copy(sx,pos('#P', sx)+7,3));       // pega o número da coluna do arq. de definições
         Tam      := StrToInt(copy(sx,pos('#P', sx)+11,3));      // pega o tamanho do campo do arq. de definições
         Campo    := copy(sx,pos('<', sx)+1,(pos('>', sx)-(pos('<', sx)+1))); // pega o nome do campo do arq. de definições
         Mascara  := copy(sx,pos('[', sx)+1,(pos(']', sx)-(pos('[', sx)+1))); // pega a máscara do campo do arq. de definições
         if copy(Campo,1,1) <> '''' then begin                   // se não for uma constante
           try
             Query1.FieldByName(Campo).FieldNo;                  // testa existência do campo
           except
             ShowMessage('Nome de Campo Inválido -> '+Campo);
             exit;
           end;
         end;
         // insere campo do tamanho definido na linha e coluna correspondente
         if copy(Campo,1,1) <> '''' then begin                         // se não for uma constante...
           sAux := Copy(Query1.FieldByName(Campo).Value,1,Tam);
           if Mascara <> '' then begin
              SetText(Lin, Col, Format(Mascara, [StrToFloat(DesmascaraNumero(sAux))])); // insere mascarado
           end else begin
              SetText(Lin, Col, sAux);                               // insere sem máscara
           end;
         end else begin                                              // se não é porque é uma constante, então...
           SetText(Lin, Col, copy(Campo,2,Tam))    ;                 // insere constante sem máscara
         end;
      end else begin
     // ******************************************************************************************
     //   verifica se é definição de linha detalhe para pegar valores das linhas de início e fim
     // ******************************************************************************************
          if pos('#L', sx) > 0 then begin                        // se for definição de linha-detalhe...
             LIni   := StrToInt(copy(sx,pos('#L', sx)+3,3));     // pega o número da Linha Inicial do looping do arq. de definições
             LFim   := StrToInt(copy(sx,pos('#L', sx)+7,3));     // pega o número da Linha Final do looping do arq. de definições
          end;
     // ***********************************************
     //   verifica se é linha detalhe para fazer loop
     // ***********************************************
          if pos('#X', sx) > 0 then begin                        // se for definição de posicionamento de campo de linha-detalhe...
             LdtCont   := LIni;                                  // Carrega contador de linhas-detalhe
             Query2.First;                                       // posiciona no primeiro registro do arquivo de produtos da Nota
             while (LdtCont <= LFim) and                         // enquanto for linha-detalhe e
                   (not Query2.Eof) do begin                     // não for final de arquivo de produtos da Nota então...
                if pos('#A', sx) > 0 then begin                  // se encontrar caracter de atributo nesta linha...
                   Atributo := copy(sx,pos('#A', sx)+3,10);         // pega atributo do campo
                   PrintCmd(TipoCmd(Tamanho));
                end;
                Lin    := LdtCont;                               // Define o número da linha através do looping
                Col    := StrToInt(copy(sx,pos('#X', sx)+3,3));  // pega o número da coluna do arq. de definições
                Tam    := StrToInt(copy(sx,pos('#X', sx)+7,3));  // pega o tamanho do campo do arq. de definições
                Campo  := copy(sx,pos('<', sx)+1,(pos('>', sx)-(pos('<', sx)+1))); // pega o nome do campo do arq. de definições
                Mascara := copy(sx,pos('[', sx)+1,(pos(']', sx)-(pos('[', sx)+1))); // pega a máscara do campo do arq. de definições
                // insere campo do tamanho definido na linha e coluna correspondente
                if Mascara <> '' then begin                      // se foi definida máscara para o campo, então...
                   SetText(Lin, Col, format(mascara, [StrToFloat(copy(Query2.FieldByName(Campo).Value,1,Tam))])); // insere mascarado
                end else begin                                   // se não...
                   SetText(Lin, Col, copy(Query2.FieldByName(Campo).Value,1,Tam)); // insere sem máscara
                 end;
                Inc(LdtCont);                                    // incrementa contador de linhas-detalhe
                Query2.Next;                                     // lê próximo registro do arquivo de produtos da Nota
             end;
          end;
      end;
      Inc(i);                                                    // incrementa contador de linhas
    end;
    Query1.Next;
    Inc(Contador);
    if Contador < Query1.RecordCount then begin
       Printer.NewPage;
    end;
  end;
  Finaliza;
end;

function TipoCmd(Comando: String): TPrintCmd;
begin
   Result := tpCmd10cpi;
   if Comando = 'tpCmd20cpi' then Result := tpCmd20cpi;
   if Comando = 'tpCmd17cpi' then Result := tpCmd17cpi;
   if Comando = 'tpCmd12cpi' then Result := tpCmd12cpi;
   if Comando = 'tpCmd10cpi' then Result := tpCmd10cpi;
   if Comando = 'tpCmd06cpi' then Result := tpCmd06cpi;
   if Comando = 'tpCmd05cpi' then Result := tpCmd05cpi;
   if Comando = 'tpCmdLgNeg' then Result := tpCmdLgNeg;
   if Comando = 'tpCmdDgNeg' then Result := tpCmdDgNeg;
   if Comando = 'tpCmdNeg10' then Result := tpCmdNeg10;
   if Comando = 'tpCmdNeg12' then Result := tpCmdNeg12;
   if Comando = 'tpCmdComum' then Result := tpCmdComum;
end;

procedure LimpaLinhaStringGrid(StringGrid :TStringGrid; Linha :Integer);
var Coluna :Integer;
begin
   with StringGrid do begin
      for Coluna := 0 to ColCount + 90 do Cells[Coluna, Linha] := '';
   end;
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure ImprimeNF(oImp: TImprimeNota);

procedure MontaCabecalho;     forward;
procedure MontaFatura;        forward;
procedure MontaProdutos;      forward;
procedure MontaServicos;      forward;
procedure MontaObservacao;    forward;
procedure MontaRodape;        forward;
procedure MontaRodapeQuebra;  forward;

var
   Sx,XX             :string;  // auxiliares
   Contador, j       :Integer;
   ContServ          :Integer;
   ContObser         :Integer;
   contProd          :Integer;
   ConPag            :Integer;
   JaImprimiuServico :Boolean;

   isValueF          :Boolean; // verdadeiro quando for um  valor ponto flutuante[%10.2F]
   isValueM          :Boolean; // verdadeiro quando for um  valor mascarado [#.##0,00]
   isConstant        :boolean; // verdadeiro quando for uma constante
   isExtenso         :boolean; // verdadeiro se for um extenso dentro do parcelamento;

   isDateTime        :boolean; // verdadeiro quando for uma data ou hora;
   isString          :boolean;

   Campo             :string;  // nome do campo do arq. de definições-deve estar declarado no ClientDataSet;
   Conteudo          :string;

   {$IFNDEF DEPURACAO}
   Msg: String;
   {$ENDIF}

   procedure PegaDef;                                            // procedure para pegar definições do arquivo texto
   var ArqDef : TextFile;                                        // declaração de variáveis auxiliares
       iPd    : integer;
   begin
     cLin  := 0;                                              // inicializa contador de linhas
     slDEF := TStringList.Create;                             // cria StringList de definições
     AssignFile(ArqDef, oImp.ArquivoNF);                      // atribui direcionamento do Parâmetro para o arquivo
     Reset(ArqDef);                                           // define arquivo como de leitura
     ReadLn(ArqDef, Linha);                                   // lê primeira linha do arquivo
     if Linha <> '#INICIO' then begin                         // se não for INICIO então não é arq. de definições válido,
        MessageDlg(oImp.ArquivoNF + ' não é um arquivo de definições válido!', mtInformation	, [mbOk], 0);
        Exit;                                                 // e caímos fora...
     end;
     while True do begin                                      // looping para ler todo arquivo linha a linha...
         ReadLn(ArqDef, Linha);                                   // lê próxima linha...
         if Linha = '#FINAL' then Break;                          // se for FINAL... caímos fora!
         iPd := Pos('#D', Linha);
         if iPd > 0 then begin                                    // pega definição de formulário...
            iSize        := StrToInt(copy(Linha, iPd + 03, 03));  // pega o tipo de página
            iWidth       := StrToInt(copy(Linha, iPd + 07, 04));  // pega a largura da página
            iHeight      := StrToInt(copy(Linha, iPd + 12, 04));  // pega o comprimento da página
            iAlturaLinha := StrToInt(copy(Linha, iPd + 17, 02));  // pega a altura da linha (sextos=24 ou oitavos=18)
            Titulo       :=          copy(Linha, iPd + 21, 20);   // pega o título do relatório
            Tamanho      :=          copy(Linha, iPd + 44, 10);   // pega o tamanho da fonte
            Orientacao   := poPortrait;
            if copy(Linha, iPd + 57, 11) = 'poLandscape' then    // pega a orientação da página (poPortrait=vertical ou poLandscape=horizontal)
               Orientacao := poLandscape;

            QtdeProdutos  := StrToInt(copy(Linha, iPd + 70,02));
            QtdeServicos  := StrToInt(copy(Linha, iPd + 73,02));
            NormalBold    :=          copy(Linha, iPd + 77,10);     // estilo de impressão normal ou Bold
         end else begin
            slDEF.Add(Linha);                                             // adiciona linha no StringList de definições
         end;
     end;
   end;

   procedure Inicia;
   var Papel : TPapel;
   begin
       ConPag := 0;
       if not Printer.Printing then begin
          Papel.Size                  := iSize;
          Papel.Width                 := iWidth;
          Papel.Height                := iHeight;
          Printer.Title               := Titulo;
          Printer.Canvas.Font.Pitch   := fpFixed;
          Printer.Orientation         := Orientacao;
          MudaTamPapel(Papel);

          case TipoCmd(Tamanho) of
              tpCmd20cpi : begin
                             Multiplicador := 0.5;
                             Multiplica    := 2;
                           end;
              tpCmd17cpi : begin
                             Multiplicador := 0.6060;
                             Multiplica    := 1.65;
                           end;
              tpCmd12cpi : begin
                             Multiplicador := 0.8333;
                             Multiplica    := 1.2;
                           end;
              tpCmd10cpi : begin
                             Multiplicador := 1.0;
                             Multiplica    := 1;
                           end;
              tpCmd06cpi : begin
                             Multiplicador := 0.6;
                             Multiplica    := 1;
                           end;
              tpCmd05cpi : begin
                             Multiplicador := 0.7;
                             Multiplica    := 1;
                           end;
          end;

          LarguraCaracter             := Printer.Canvas.TextWidth('a');
          LarguraPagina               := (Printer.PageWidth  div LarguraCaracter) - 1;
          AlturaLinha                 := iAlturaLinha;
          AlturaPagina                := (Printer.PageHeight div AlturaLinha    );
          PCol                        := 0;
          PRow                        := 1;
          if AlturaLinha = 18 then begin
             QtdeLinhas               := Trunc((iHeight/254)*8);
          end else begin
             QtdeLinhas               := Trunc((iHeight/254)*6);
          end;
          PrintCmd(TipoCmd(Tamanho));
          {$IFDEF DEPURACAO}
 //          showMessage('inicio de documento');
          {$ELSE}
          Printer.BeginDoc;
          {$ENDIF}
       end;
   end;


   // sempre defaults
   function MontaVariaveisImpressao(const Qry: TFDMemTable; sTipoVariavel:string): boolean;
   begin
     Result  := True;

     Lin     := StrToInt(copy(Sx, pos(sTipoVariavel, Sx)+4,3));         // pega o número da Linha do arq. de definições
     Col     := StrToInt(copy(Sx, pos(sTipoVariavel, Sx)+8,3));         // pega o número da coluna do arq. de definições
     Col     := Round(Col * Multiplicador);
     Col     := Round(Col * Multiplica   );
     Tam     := StrToInt(copy(sx,pos(sTipoVariavel, sx)+12,3));         // pega o tamanho do campo do arq. de definições
     Campo   := copy(sx,pos('<', sx)+1,(pos('>', sx)-(pos('<', sx)+1))); //
     sMascara:= copy(sx,pos('[', sx)+1,(pos(']', sx)-(pos('[', sx)+1))); // pega a máscara do campo do arq. de definições

     isValueF   := Pos('F]', UpperCase(SX)) > 0;
     isValueM   := Pos('#0.0', sMascara) > 0;

     isExtenso  := pos('#EX', SX)>0;
     isConstant := Copy(Campo, 1, 1) = '''';           // constande pre definida
     isString   := Pos('S]', UpperCase(sMascara)) > 0; // string mascarada
     isDateTime := Pos('/',  sMascara) > 0;

     if isConstant or isExtenso then
       Campo := AnsiReplaceText(Campo, chr($27), '')  // retira aspas,
     else
       try                               // se nao for constante, testa se o campo existe no ClientDataSet;
         Qry.FieldByName(Campo).FieldNo;
       except
         ShowMessage(Format('Nome de Campo %S Inválido -> %S ', [sTipoVariavel, Campo]));
         Result := False;
       end;
   end;


   function MontaALinhaDeImpressao(X: string): string;

       function VoltaVlr(sV: String): currency;
       begin
         if Trim(sV) = '' then
            Result := 0
         else
            Result := StrToFloat(sV);
       end;

   begin
      if (Conteudo='-1') then begin
         Conteudo := '';
         Insert(Conteudo, X, Col);
      end else begin
         if isConstant or isExtenso then
           Insert(Copy(Campo, 1, Tam), X, Col)
         else
           if isValueF then
              Insert(StrRight(Format(sMascara, [VoltaVlr(Conteudo)]), Tam), X, Col)
           else
             if isValueM then
                Insert(StrRight(FormatFloat(sMascara, VoltaVlr(Conteudo)), Tam), X, Col)
             else
               if isString then
                  Insert(Format(sMascara, [Conteudo]), X, Col)
               else
                  if isDateTime then
                     Insert(FormatDateTime(sMascara, StrToDateTime(Conteudo)), X, Col)
                  else
                     Insert(Conteudo, X, Col);  // XXXX = conteudo sem mascara <XXXX>
      end;
      Result := X;
   end;

   function MontaALinhaDeImpressao1(const Qry: TFDMemTable; X: string; iLinha:Integer): String;
   begin
     Result := X;

     if not (isConstant or isExtenso) then
        Conteudo := Copy(Qry.FieldByName(Campo).asString, 1, TAM);

     Result := MontaALinhaDeImpressao(Result);

     slNF.Delete(iLinha);
     slNF.Insert(iLinha, Result);

   end;

   // DA NOTA FISCAL
   procedure SetText(Row, Col: Integer; Texto : String);
   begin
     QuebrouPagina := False;
     {$IFDEF DEPURACAO}
     if (Row > AlturaPagina) then begin
        slNF.SaveToFile('C:\TEMP\NOTA'+oImp.TipoDeNota+IntToStr(ConPag)+'.TXT');
        QuebrouPagina := True;
        Row := 0;
     end;
     PCol := Col;
     PRow := Row;
     {$ELSE}
     if Printer.Printing then begin
        if (Row > AlturaPagina) then begin
           Printer.NewPage;
           QuebrouPagina := True;
           Row := 0;
        end;
        PCol := Col;
        PRow := Row;
        with Printer.Canvas do begin
          Col := (Col * LarguraCaracter);
          Row := (Row * AlturaLinha);
          TextOut(Col, Row, Texto);   { Texto impresso }
        end;
     end;
     {$ENDIF}
   end;

   procedure Finaliza;
   var Papel : TPapel;
   begin
      {$IFDEF DEPURACAO}
      slNF.SaveToFile('C:\TEMP\NOTA'+oImp.TipoDeNota+IntToStr(ConPag)+'.TXT');
      {$ELSE}
      if Printer.Printing then
         Printer.EndDoc;
      {$ENDIF}
      Papel.Size               := 256;
      Papel.Width              := 2159;
      Papel.Height             := 2794;
      Printer.Canvas.Font.Name := 'Draft 10Cpi'; // Tamanho normal
      MudaTamPapel(Papel);

      LarguraCaracter :=  Printer.Canvas.TextWidth('a');
      LarguraPagina   := (Printer.PageWidth  div LarguraCaracter) - 1;
      AlturaLinha     := 24;
      AlturaPagina    := (Printer.PageHeight div AlturaLinha    ) - 1;
      PCol            := 0;
      PRow            := 1;
   end;

   procedure MontaCabecalho;
   begin
     i := 0;
     Inc(Conpag);
     oImp.QueryCabecalho.First;
     while i <= slDEF.Count-1 do begin                            // enquanto existir linhas de definição...
        Sx := slDEF.Strings[i];                                   // usa variável string para trabalhar
        if (Pos('#PC', Sx) > 0) then begin
           if not MontaVariaveisImpressao(oImp.QueryCabecalho, '#PC') then
              Exit;
           MontaALinhaDeImpressao1(oImp.QueryCabecalho, slNF.Strings[Lin], Lin);
        end;
        Inc(i);
     end;
   end;

   procedure MontaFatura;
   var lPrimeiraVez : Boolean;
   begin
     i := 0;
     oImp.QueryFatura.First;
     lPrimeiraVez := True;
     if (oImp.QueryFatura.Eof and oImp.QueryFatura.Bof) then exit;

     while i <= slDEF.Count-1 do begin                         // enquanto existir linhas de definição...
        sx  := slDEF.Strings[i];                               // usa variável string para trabalhar

        if (Pos('#PF', sx) > 0) then begin

           if (pos('PARCELA', sx) > 0) then begin
              if (not lPrimeiraVez) then begin
                 oImp.QueryFatura.Next;
                 if (oImp.QueryFatura.Eof) then exit;
              end;
           end;

           if not MontaVariaveisImpressao(oImp.QueryFatura, '#PF') then
              Exit;

           MontaALinhaDeImpressao1(oImp.QueryFatura, slNF.Strings[Lin], Lin);

           lPrimeiraVez := False;
        end;
        Inc(i);
     end;
   end;

   procedure MontaExtenso;
   var i, tamExt : integer;
       VlrParcela: currency;
       sPar      : string;
       sExtenso  : string;
       sExtenso1 : string;
       sExtenso2 : string;
       sExtenso3 : string;
       Vencimento: string;
   begin
     VlrParcela := 0.00;
     oImp.QueryFatura.First;

     while not oImp.QueryFatura.eof do begin
        if (oImp.QueryFatura.FieldByName('Parcela').asString >= '001') and
           (oImp.QueryFatura.FieldByName('Parcela').asString <= '099') then begin

            Vencimento := oImp.QueryFatura.FieldByName('Vencimento').asString;
            sPar       := oImp.QueryFatura.FieldByName('Parcela').asString;
            VlrParcela := VlrParcela + oImp.QueryFatura.FieldByName('Valor').asCurrency;
        end;
        oImp.QueryFatura.Next;
     end;

     i := 0; TamExt := 0;
     while i <= slDEF.Count-1 do begin
        sx := UpperCase(slDEF.Strings[i]);
        if (Pos('<EXTENSO1>', sx) > 0) then begin
           MontaVariaveisImpressao(oImp.QueryFatura, '#EX');
           TamExt := Tam;
        end;
        inc(i);
     end;
     if tamExt = 0 then Exit;


     if (VlrParcela > 0.00) then begin
        sExtenso  := PreencheExtenso(250, VlrParcela, 1);
        sExtenso1 := Copy(sExtenso, 001, tamExt);
        sExtenso2 := Copy(sExtenso, tamExt+1, tamExt);
        sExtenso3 := Copy(sExtenso, tamExt*2+1, tamExt)
     end else begin
        sExtenso1 := StringOfChar('*', tamExt);
        sExtenso2 := StringOfChar('*', tamExt);
        sExtenso3 := StringOfChar('*', tamExt);
     end;



     // Valor da Nota - sera com #PR, impresso em outra rotina;
     // parcela / extenso - serao com #EX
     i := 0;
     while i <= slDEF.Count-1 do begin
        sx  := UpperCase(slDEF.Strings[i]);
        if (Pos('#EX', sx) > 0) then begin

           Campo := '';

           if not MontaVariaveisImpressao(oImp.QueryFatura, '#EX') then
              Exit;


           IF POS('<PARCELA>',  SX)>0 then Campo := sPar;
           IF POS('<VALOR>',    SX)>0 then Campo := StrRight(FormatFloat('#,##0.00', VlrParcela), Tam);
           IF POS('<VENCIMEN',  SX)>0 then Campo := StrRight(Vencimento, Tam);
           IF POS('<EXTENSO1>', SX)>0 then Campo := sExtenso1;
           IF POS('<EXTENSO2>', SX)>0 then Campo := sExtenso2;
           IF POS('<EXTENSO3>', SX)>0 then Campo := sExtenso3;

           MontaALinhaDeImpressao1(oImp.QueryFatura, slNF.Strings[Lin], Lin);

        end;
        Inc(i);
     end;
   end;

   procedure QuebraPagina;
   var Contador, J: Integer;
   begin
      XX := slNF.Strings[Lin+1];                     // joga linha existente para uma variável de trabalho
      Insert('CONTINUA...', XX, Col-11);             // insere constante sem máscara
      slNF.Delete(Lin+1);                            // remove linha antiga do StringList
      slNF.Insert(Lin+1,XX);                         // insere linha nova com o campo já posicionado
      MontaRodapeQuebra;
      {$IFDEF DEPURACAO}
      slNF.SaveToFile('C:\TEMP\NOTA'+oImp.TipoDeNota+IntToStr(ConPag)+'.TXT');
      {$ENDIF}
      for Contador := 1 to slNF.Count-1 do begin     // enquanto existir linhas de definição...
        SetText(Contador, 0, TrimRight(slNF.Strings[Contador]));
        if QuebrouPagina then
           Break;  // efetuou nova pagina
      end;
      {$IFDEF DEPURACAO}

      {$ELSE}
      if not QuebrouPagina then
         Printer.NewPage;
      {$ENDIF}

      slNF.Clear;
//    for j := 0 to ContLinha do begin              // looping para verificar se existe linhas suficientes
      for j := 0 to Alturapagina do begin           // looping para verificar se existe linhas suficientes
         if slNF.Count <= j then begin              // no StringList de NF para atender o posicionamento do campo
            slNF.Insert(j,Replique(' ', iLargura)); // cria linha em branco com oitenta espaços
         end;
      end;

      MontaCabecalho;
      MontaFatura;

      ContProd := 1;
   end;

   procedure MontaProdutos;
   var // contador  : Integer;
       // J         : Integer;
       iServ: Integer;
   begin
      i  := 0;
      sx := slDEF.Strings[i];                                   // usa variável string para trabalhar

      oImp.QueryProdutos.First;

      ContProd := 1;

      if (oImp.QueryProdutos.Eof and oImp.QueryProdutos.Bof) then
         exit;

      while (ContProd <= oImp.QueryProdutos.RecordCount) and not(oImp.QueryProdutos.Eof) do begin
         if ContProd > QtdeProdutos then begin
            // desvio da rotina com quebra de página
            // Imprime Serviços... (se houverem e o que der...)
            iServ := 0;
            sx    := slDEF.Strings[i]; // usa variável string para trabalhar

            if not JaImprimiuServico then
               oImp.QueryServicos.First;

            ContServ := 1;

            if not (oImp.QueryServicos.Eof and oImp.QueryServicos.Bof) then begin
               while (ContServ <= oImp.QueryServicos.RecordCount) and not(oImp.QueryServicos.Eof) do begin
                  if ContServ <= QtdeServicos then begin
                     while iServ <= slDEF.Count - 1 do begin
                        sx := slDEF.Strings[iServ];

                        if (pos('#PS', sx) > 0) then begin
                           if not MontaVariaveisImpressao(oImp.QueryServicos, '#PS') then
                              Exit;

                           Lin := Lin + ContServ;
                           MontaALinhaDeImpressao1(oImp.QueryServicos, slNF.Strings[Lin], Lin);
                        end;

                        Inc(iServ);
                     end;

                     iServ := 0;
                     Inc(ContServ);

                     JaImprimiuServico := True;

                     oImp.QueryServicos.Next;
                  end else
                     ContServ := oImp.QueryServicos.RecordCount + 1;
               end;
            end;

            QuebraPagina;
            ContProd := 1;
            i        := 0;
         end;

         // Faz para todos os produtos
         while i <= slDEF.Count - 1 do begin
            sx  := slDEF.Strings[i];

            if (pos('#PP', sx) > 0) then begin
               if not MontaVariaveisImpressao(oImp.QueryProdutos, '#PP') then
                  Exit;

               Lin := Lin + ContProd;
               MontaALinhaDeImpressao1(oImp.QueryProdutos, slNF.Strings[Lin], Lin);
            end;

            Inc(i);
         end;

         i := 0;
         Inc(ContProd);
         oImp.QueryProdutos.Next;
      end;
   end;

   procedure MontaServicos;
   begin
     i  := 0;
     sx := slDEF.Strings[i];                                   // usa variável string para trabalhar

     if not JaImprimiuServico then
        oImp.QueryServicos.First;


     ContServ := 1;

     if (oImp.QueryServicos.Eof and oImp.QueryServicos.Bof) then
        exit;

     while (ContServ <= oImp.QueryServicos.RecordCount) and not(oImp.QueryServicos.Eof) do begin
        if QtdeServicos > 0 then begin
           if ContServ > QtdeServicos then begin
              QuebraPagina;
              ContServ := 0;
              oImp.QueryServicos.Prior;
           end;
        end;

        while i <= slDEF.Count-1 do begin                            // enquanto existir linhas de definição...
           sx  := slDEF.Strings[i];                                   // usa variável string para trabalhar
           if (pos('#PS', sx) > 0) then begin
              if not MontaVariaveisImpressao(oImp.QueryServicos, '#PS') then
                 Exit;
              Lin := Lin + ContServ;
              MontaALinhaDeImpressao1(oImp.QueryServicos, slNF.Strings[Lin], Lin);
           end;
           Inc(i);
        end;
        i := 0;
        Inc(ContServ);
        oImp.QueryServicos.Next;
     end;
   end;

   procedure MontaObservacao;  // no corpo da nota - igual ao produto
   begin
     i         := 0;
     ContObser := 1;

     with oImp do begin
       QueryObservacao.First;
       if (QueryObservacao.Eof and QueryObservacao.Bof) then exit;

       while (ContObser <= QueryObservacao.RecordCount) and not(QueryObservacao.Eof) do begin
         if (ContObser + ContProd) > (QtdeProdutos - 1) then begin
             QuebraPagina;
             ContObser := 1;
             ContProd  := 0; // nao ha produtos impressos nesta pagina
         end;
         while i <= slDEF.Count-1 do begin
            sx := slDEF.Strings[i];
            if (Pos('#PO', sx) > 0) then begin
              if not MontaVariaveisImpressao(QueryObservacao, '#PO') then
                 Exit;
              Lin := Lin + (ContObser + ContProd);
              MontaALinhaDeImpressao1(QueryObservacao, slNF.Strings[Lin], Lin);
            end;
            Inc(i);
         end;
         i := 0;
         Inc(ContObser);
         QueryObservacao.Next;
       end;
     end;
   end;

   procedure MontaRodape;
   var ImpDescto : boolean;

       procedure MontaDadosRodape(sTipoBloco:string);

         procedure TestaDesconto;
         begin
           if ImpDescto then
             MontaALinhaDeImpressao1(oImp.QueryRodape, slNF.Strings[Lin], Lin);
         end;


       begin
         I := 0;
         with oImp do begin
           QueryRodape.First;
           ImpDescto := oImp.QueryRodape.FieldByName('TOTALDESCONTO').asCurrency > 0.00;
           while I <= slDEF.Count-1 do begin
             SX := slDEF.Strings[I];
             if (Pos(sTipoBloco, SX) > 0) then begin
                if not MontaVariaveisImpressao(QueryRodape, sTipoBloco) then
                   Exit;

                if Campo = 'DESCONTO ESPECIAL' then
                   TestaDesconto
                else
                   if Campo = 'TOTALDESCONTO' then
                      TestaDesconto
                   else
                       MontaALinhaDeImpressao1(QueryRodape, slNF.Strings[Lin], Lin);
             end;
             Inc(I);
           end;
         end;
       end;

   begin { MontaRodape }
     MontaDadosRodape('#PR');
     MontaDadosRodape('#PD');
   end;

   procedure MontaRodapeQuebra;   // Rodapé de quebra de página
   begin
     i := 0;
     oImp.QueryRodape.First;

     while i <= slDEF.Count-1 do begin                            // enquanto existir linhas de definição...
        sx  := slDEF.Strings[i];                                   // usa variável string para trabalhar
        if (Pos('#PR', sx) > 0) then begin

           if not MontaVariaveisImpressao(oImp.QueryRodape, '#PR') then
              Exit;

           // insere campo do tamanho definido na linha e coluna correspondente
           xx := slNF.Strings[Lin];                                   // joga linha existente para uma variável de trabalho
           if copy(Campo,1,1) <> '''' then begin                     // se não for uma constante...
             if sMascara <> '' then begin                            // se foi definida máscara para o campo, então...
                Insert(Copy('**,***.**',1,Tam), XX, Col);             // insere sem máscara
             end;
           end else begin                                            // se não é porque é uma constante, então...
             Insert(copy(Campo,2,Tam), xx, Col);                      // insere constante sem máscara
           end;
           slNF.Delete(Lin);                                         // remove linha antiga do StringList
           slNF.Insert(Lin,xx);                                       // insere linha nova com o campo já posicionado
        end;
        Inc(i);
     end;
   end;

   const S1='Impressora da nota %S não encontrada. Configure o parametro %S no GER605RA';
   var iPrinter : integer;

   function ProcuraImpressoraServico: string;
   begin
      result := '';
      iPrinter := FindPrinter(PortaImpressora(tiNotaServico));
      if iPrinter < 0 then
         Result := Format(s1, ['de serviços', '[ImpNotaServico]']);
   end;

   function ProcuraImpressoraProdutos: string;
   begin
      result := '';
      iPrinter := FindPrinter(PortaImpressora(tiNotaFiscal));
      if iPrinter < 0 then
         Result := Format(s1, ['fiscal', '[ImpNotaFiscal]']);
   end;

begin
  slNF := TStringList.Create;                               // cria StringList de definições
  PegaDef;                                                  // chama procedure para pegar definições do arquivo texto

  {$IFDEF DEPURACAO}
   // o arquivo gravado é NOTA.TXT no diretório da aplicação
  {$ELSE}
  if oIMP.TipoDeNota = 'SERVICOS' then  // depende do parametro NotaServicoSeparada=True no VEN414;
     Msg := ProcuraImpressoraServico
  else
     Msg := ProcuraImpressoraProdutos;  // PRODUTOS e TODOS -> (produtos e servicos na mesma notas);
  if msg <> '' then begin
     Mensagem(Msg, tmInforma);
     exit;
  end;
  if iPrinter >= 0 then
     Printer.PrinterIndex := iPrinter;
  {$ENDIF}

  if iAlturaLinha = 18 then
     iLinhasPolegada := 8
  else
     iLinhasPolegada := 6;

  Inicia;  // inicia impressora

  ContLinha := Trunc((iHeight/254)*iLinhasPolegada);
  iLargura  := Trunc((iWidth/254)*Trunc(Multiplica*10));

//  for j := 0 to ContLinha do begin            // looping para verificar se existe linhas suficientes
  for j := 0 to AlturaPagina do begin           // looping para verificar se existe linhas suficientes
     if slNF.Count <= j then begin              // no StringList de NF para atender o posicionamento do campo
        slNF.Insert(j,Replique(' ', iLargura)); // cria linha em branco com oitenta espaços
     end;
  end;

  JaImprimiuServico := False;

  MontaCabecalho;
  MontaFatura;
  MontaProdutos;
  MontaServicos;
  MontaObservacao;
  MontaRodape;
  MontaExtenso;

  for Contador := 1 to slNF.Count-1 do begin                            // enquanto existir linhas de definição...
    SetText(Contador, 0, TrimRight(slNF.Strings[Contador]));
  end;

  {$IFDEF DEPURACAO}
  oImp.QueryCabecalho.SaveToFile('C:\TEMP\Cabecalho.xml', sfXML);
  oImp.QueryFatura.SaveToFile('C:\TEMP\Fatura.xml', sfXML);
  oImp.QueryProdutos.SaveToFile('C:\TEMP\Produtos.xml', sfXML);
  oImp.QueryServicos.SaveToFile('C:\TEMP\Servicos.xml', sfXML);
  oImp.QueryObservacao.SaveToFile('C:\TEMP\Observacao.xml', sfXML);
  oImp.QueryRodape.SaveToFile('C:\TEMP\Rodape.xml', sfXML);
  {$ENDIF}

  Finaliza;

  if Assigned(slNF) then // teio
     slNF.Free;

end;

{$ENDIF}

function  TestaAcesso(TesPrograma :String; TipoAcesso :TTipoAcesso = TmGeral) :Boolean;
var
  Letra: string;
  T : Integer ;
  QueryPesquisa: TFDQuery;
begin
  Result := True;
  QueryPesquisa               := TFDQuery.Create(nil);
  QueryPesquisa.Connection := FDataModule.FDConnection;
  try
    with QueryPesquisa do
    begin
      Sql.Clear;
      Sql.Add(FORMAT('SELECT * FROM TGerAcesso           ' +
                   'WHERE                                ' +
                   'Empresa  =%s AND                     ' +
                   'Usuario  =%s AND                     ' +
                   'Programa =%s                         ',
                   [
                     QuotedStr(Empresa),
                     QuotedStr(Usuario),
                     QuotedStr(TesPrograma)
                   ]
                   ));

      { Se retorna FALSE, não encontrou nenhum registro }
      if  EcoOpen(QueryPesquisa) then
      begin
        if TipoAcesso = TmGeral  then
        begin
          T := Length(TesPrograma)-1;
          Letra:= Copy(TesPrograma,T,2);

          if Letra = 'CA' then
          begin
            if not ((FieldByName('Geral').asString = 'S') or
               (FieldByName('INCLUIR').asString = 'S')or
               (FieldByName('ALTERAR').asString = 'S') or
               (FieldByName('CONSULTAR').asString = 'S')or
               (FieldByName('EXCLUIR').asString = 'S'))
            then
            begin
              Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
              QueryPesquisa.Close;
              Result := False;
              Close;
            end;
          end else
          begin
            if FieldByName('Geral').asString <> 'S' then
            begin
              Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
              QueryPesquisa.Close;
              Result := False;
              Close;
            end;
          end;
        end else
        If TipoAcesso = TmIncluir  then
        begin
          if FieldByName('INCLUIR').asString <> 'S' then
          begin
            Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
            QueryPesquisa.Close;
            Result := False;
            Close;
          end;
        end else
        if TipoAcesso = TmEditar  then
        begin
          if FieldByName('ALTERAR').asString <> 'S' then
          begin
            Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
            QueryPesquisa.Close;
            Result := False;
            Close;
          end;
        end else
        if TipoAcesso = TmConsultar then
        begin
          if FieldByName('CONSULTAR').asString <> 'S' then
          begin
            Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
            QueryPesquisa.Close;
            Result := False;
            Close;
          end;
        end else
        if TipoAcesso = TmDeletar then
        begin
          if FieldByName('EXCLUIR').asString <> 'S' then
          begin
            Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
            QueryPesquisa.Close;
            Result := False;
            Close;
          end;
        end;
      end else
      begin
        { Se não encontrou nenhum registro ou não tem acesso ao programa }
        Mensagem('O usuário ' + Usuario + ' não tem direitos de acesso ao programa ' + TesPrograma + '!',tmInforma);
        QueryPesquisa.Close;
        Result := False;
        Close;
      end;
    end;
  finally
   QueryPesquisa.Close;
   QueryPesquisa.Free;
  end;
end;

// Formata uma data para usar em clausulas
function DataDMA(Data: TDateTime):string;
begin
   try
     Result := FormatDateTime('dd/mm/yyyy',Data);
   except
     Result := '';
   end;
end;

function DataMDA(Data: TDateTime):string;
begin
   try
     Result := FormatDateTime('mm/dd/yyyy',Data);
   except
     Result := '';
   end;
end;

// Formata uma data para usar em clausulas
function DataTimeSQL(Data: TDateTime):string;
begin
   try
     Result := FormatDateTime('dd.mm.yyyy hh:mm:ss',Data);
   except
     Result := '';
   end;
end;

// - Substitui strings dentro de um texto : Retorna texto alterado
function Substitui(TextoOld :String ; TextoNew: String; TextoPai: String) :String;
begin
    Result  := AnsiReplaceText(TextoPai, TextoOld, TextoNew);
end;

function  TestaUsuario(TesUsuario :String) :Boolean; // verifica se existe um determidado usuário
var
   QueryPesquisa : TFDQuery;
begin
   Result := True;
   QueryPesquisa               := TFDQuery.Create(nil);
   QueryPesquisa.Connection := FDataModule.FDConnection;

   try
     with QueryPesquisa do
     begin
        Sql.Clear;
        Sql.Add('SELECT Usuario FROM TGerUsuario  ' +
                     'WHERE                       ' +
                     'Usuario  =:Usuario          ');
        ParamByName('Usuario').asString   := TesUsuario;

        if EcoOpen(QueryPesquisa) then
        begin
           if (QueryPesquisa.FieldByName('Usuario').asString) <> TesUsuario then
           begin
               Mensagem('O usuário não cadastrado!',tmInforma);
               QueryPesquisa.Close;
               Result := False;
           end;
        end;
     end;
   finally
      QueryPesquisa.Close;
     QueryPesquisa.Free;
   end;
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
function SolicitaSenha(Tipo:Word):Boolean;
begin
   Result := False;
   Application.CreateForm(TFVEN601RF, FVEN601RF);
   try
      with FVEN601RF do
      begin
         if (ShowModal=mrOK) then
         begin
            if (Tipo=1) and (FVEN601RF.QueryPesquisa.FieldByName('I').AsString='S') then Result := True;
            if (Tipo=2) and (FVEN601RF.QueryPesquisa.FieldByName('J').AsString='S') then Result := True;
            if (Tipo=3) and (FVEN601RF.QueryPesquisa.FieldByName('L').AsString='S') then Result := True;
            if not Result then begin
               Mensagem('Usuário com autonomia restrita.', tmInforma);
               Exit;
            end;
         end else
         begin
            Exit;
         end;
      end;
   finally
      FreeAndNil(FVEN601RF);
   end;
end;

{-----------------------------------------------------------------------------}
{ atenção as rotinas abaixo deve estar em conformidade com o prg TSC-Filizola }
{ testa se e um produto pesado de uma etiqueta impressa na balanca filizola   }
{-----------------------------------------------------------------------------}
function IsProdutoPesado(sEAN13: string): boolean;
begin
  if (Trim(sEAN13) <> '') then
    Result := (sEAN13[1] = '2') and (Length(sEAN13)=13)
  else
    Result := false;
end;
{-----------------------------------------------------------------------------}
{ extrai o valor da vda de um prod. pesado em etiqueta impressa na filizola   }
{-----------------------------------------------------------------------------}
function ExtraiVendaDoEAN(sEAN13: string): Extended;
var Aux :string;
begin
  Result := 0; // nao ha'valor ou nao e'um ean13
  try
    if TestaEAN13(sEAN13) then
    begin
      Aux := sEAN13[8]+sEan13[9]+sEAN13[10]+sEan13[11]+sEAN13[12];
      Result := (StrToFloat(Copy(sEan13,8,5)) / 100);
    end;
  except
    Result := -1; // ilegal
  end;
end;

{$ENDIF}

function RetornaHoraMinuto(Tempo:Currency):String;
Var Horas   :Real;
    Minutos :Real;
begin
   Result := '00:00';
   if (Tempo>0) then begin
      Tempo   := Round(Tempo * 1440); { 24 x 60 = 1440 'Total de minutos no dia' }
      Horas   := Trunca(Divide(Tempo,60),0);
      Minutos := Round(Tempo-(Horas*60));
      if Horas>99 then begin
         Result  := FloatToStr(Horas)+':'+StrZero(Minutos,2);
      end else begin
         Result  := StrZero(Horas,2)+':'+StrZero(Minutos,2);
      end;
   end
end;

function CalculaPercentual(Valor1,Valor2:Extended):Extended;
begin
    Result := 0;
    if Valor1 = 0 then exit;
    if Valor2 > 0 then
       Result := ((Valor2/Valor1)-1)*100;
end;

function MontaGrupoEmpresas(sGrupo, sTabela, sComeco:String):String;
var Filtro        :String;
    QueryPesquisa : TSQLQuery;
begin
   Result := sComeco+' ('+sTabela+'.Empresa = '+QuotedStr(Empresa)+') ';
   QueryPesquisa               := TSQLQuery.Create(nil);
   QueryPesquisa.SQLConnection := FDataModule.sConnection;

   try
      with QueryPesquisa do begin
         Sql.Clear;
         Sql.Add('SELECT Empresa            ' +
                 'FROM TGerGrupoEmpresas    ' +
                 'WHERE                     ' +
                 'Grupo  =:Grupo            ');
         ParamByName('Grupo').asString   := sGrupo;
         Open;
         if (Eof and Bof) then exit;
         Filtro := sComeco+'(';
         while not Eof do begin
            Filtro := Filtro + iif(Filtro<>sComeco+'(',' OR ','')+sTabela+'.Empresa = '''+FieldByName('Empresa').asString+'''';
            Next;
         end;
         Result := Filtro+') ';
         Close;
      end;
   finally
      QueryPesquisa.free;
   end;
end;

// - Retorna o peso do produto principal dividido pela quantidade da embalagem do produto principal
//   Produto  - Codigo do produto principal
//   TipoPeso - L (PesoLiquido) _  B (PesoBruto)
function PesoPrincipal(Produto :String; TipoPeso :String) :Real;
var
   QueryPesquisaPeso: TFDQuery;
begin
   Result := 0;

   QueryPesquisaPeso := TFDQuery.Create(nil);
   try
      QueryPesquisaPeso.Connection := FDataModule.FDConnection;

      QueryPesquisaPeso.Sql.Clear;
      QueryPesquisaPeso.Sql.Add(Format('SELECT PesoBruto,      ' +
                                       '       PesoLiquido,    ' +
                                       '       QtdeEmbalagem   ' +
                                       'FROM TEstProdutoGeral  ' +
                                       'WHERE Codigo = ''%s''  ',
                                       [Produto]));

      if EcoOpen(QueryPesquisaPeso) then
      begin
         if TipoPeso = 'B' then
            Result := QueryPesquisaPeso.FieldByName('PesoBruto').asFloat;
         if TipoPeso = 'L' then
            Result := QueryPesquisaPeso.FieldByName('PesoLiquido').asFloat;
      end;
   finally
      QueryPesquisaPeso.Close;
      QueryPesquisaPeso.Free;
   end;
end;

procedure PesoPrincipal(const Produto: string; var pesoBruto, pesoLiquido: Real); overload;
var
   QueryPesquisaPeso: TFDQuery;
begin
   pesoBruto := 0;
   pesoLiquido := 0;

   QueryPesquisaPeso := TFDQuery.Create(nil);
   try
      QueryPesquisaPeso.Connection := FDataModule.FDConnection;

      QueryPesquisaPeso.SQL.Clear;
      QueryPesquisaPeso.SQL.Add(Format('SELECT PesoBruto, ' +
                                       '       PesoLiquido, ' +
                                       '       QtdeEmbalagem ' +
                                       'FROM TEstProdutoGeral ' +
                                       'WHERE Codigo = %s',
                                       [QuotedStr(Produto)]));

      if EcoOpen(QueryPesquisaPeso) then
      begin
         pesoBruto := QueryPesquisaPeso.FieldByName('PesoBruto').asFloat;
         pesoLiquido := QueryPesquisaPeso.FieldByName('PesoLiquido').asFloat;
      end;
   finally
      QueryPesquisaPeso.Close;
      QueryPesquisaPeso.Free;
   end;
end;

procedure CalculaMargemLucro(Var PrecoProduto:TPrecoProduto);
var ValorBase   :Currency;
    IcmsApurado :Currency;
begin
   with PrecoProduto do begin
      IcmsApurado := 0;
      if (ClassificacaoEstadual=0) then begin
         if (AliqVendaIcms1<>0) and (AliqCompraIcms<>0) then begin
            CalculaIcmsCompra(PrecoProduto);
            CalculaIcmsVenda(PrecoProduto,2);
            IcmsApurado := (VlrIcmsVenda - VlrIcmsCompra);
         end;
      end;
      ValorBase   := CustoBaseCalculo(PrecoProduto) + IcmsApurado + (PrecoVenda * CargaPS / 100);
      MargemLucro := (1 - Divide(ValorBase, PrecoVenda)) * 100;
   end;
end;

procedure CalculaCustoFinal(Var PrecoProduto:TPrecoProduto);
begin
   with PrecoProduto do begin
      CustoFinal := 0;

      if CustoReposicao>0 then begin
         if (ClassificacaoEstadual = 0) and ((VendaCSF1 = '000') or (VendaCSF1 = '020') or (VendaCSF1 = '090')) then begin
            CustoFinal := CustoReposicao + (VlrIcmsVenda - VlrIcmsCompra);
         end else begin
            CustoFinal := CustoReposicao;
         end;
         if (CustoReposicao <> 0) and (CargaPS <> 0) then begin
            CustoFinal := CustoFinal + (PrecoVenda * CargaPS/100) ;
         end;
      end;
   end;
end;

procedure CalculaPrecoVenda(Var PrecoProduto:TPrecoProduto);
var ValorBase    :Currency;
begin
   with PrecoProduto do begin
      ValorBase    := CustoBaseCalculo(PrecoProduto);
      if (ClassificacaoEstadual=0) then begin
         if (AliqVendaIcms1<>0) and (AliqCompraIcms<>0) then begin
            CalculaIcmsCompra(PrecoProduto);
            CalculaIcmsVenda(PrecoProduto);
            ValorBase    := ValorBase + (VlrIcmsVenda - VlrIcmsCompra);
         end;
      end;
      PrecoVenda := Divide(ValorBase, (1 - ((MargemLucro+CargaPS) / 100) ));
   end;
end;

procedure CalculaIcmsCompra(Var PrecoProduto:TPrecoProduto);
var ValorBase :Currency;
begin
   with PrecoProduto do begin
      VlrIcmsCompra := 0;
      if ClassificacaoEstadual<=1 then begin
         if AliqCompraIcms<>0 then begin
            ValorBase     := iif(PautaCompra>CustoFabrica, PautaCompra, CustoFabrica);
            ValorBase     := ValorBase - (ValorBase * AliqCompraReducao / 100);
            VlrIcmsCompra := (ValorBase * AliqCompraIcms / 100);
         end;
      end;
   end;
end;

procedure CalculaIcmsGarantidoIntegral(Var PrecoProduto:TPrecoProduto);
begin
   with PrecoProduto do begin
      VlrIcmsGarantidoIntegral := 0;
      if (ClassificacaoEstadual=1) and (EstadoOrigem<>'MT') then begin
         if AliqVendaIcms1<>0 then begin
            VlrIcmsGarantidoIntegral := ((CustoOrigem+(CustoOrigem*MargemLucroGarantido/100))*AliqVendaIcms1/100) - VlrIcmsCompra;
         end;
      end;
   end;
end;

function CustoBaseCalculo(var PrecoProduto:TPrecoProduto):Currency;
begin
   with PrecoProduto do begin
      Result := iif(PdtCustoUtilizado=0, CustoReposicao, CustoMedioReposicao);
   end;
end;

function RetornaBaseCalculo(var PrecoProduto:TPrecoProduto):Currency;
var ValorBase    :Currency;
    PercReduzido :Currency;
begin
   with PrecoProduto do begin
      ValorBase    := CustoBaseCalculo(PrecoProduto);
      PercReduzido := 0;
      if ClassificacaoEstadual=0 then begin
         if (AliqVendaIcms1<>0) then begin
            CalculaIcmsCompra(PrecoProduto);
            PercReduzido := AliqVendaIcms1 - (AliqVendaIcms1 * AliqVendaReducao / 100);
            ValorBase    := ValorBase - VlrIcmsCompra;
         end;
      end;
      Result := Divide(ValorBase, (1 - ((MargemLucro+CargaPS+PercReduzido)/100)));
   end;
end;

procedure CalculaIcmsVenda(Var PrecoProduto:TPrecoProduto;Opcao:Integer=1);
var ValorBase    :Currency;
begin
   with PrecoProduto do begin
      VlrIcmsVenda := 0;

      if Opcao=1 then begin { Quando ainda nao se tem o preco de venda }
         if ClassificacaoEstadual=0 then begin
            if (AliqVendaIcms1<>0) then begin
               ValorBase    := RetornaBaseCalculo(PrecoProduto);
               ValorBase    := iif(ValorBase>PautaVenda, ValorBase, PautaVenda);
               ValorBase    := ValorBase - (ValorBase * AliqVendaReducao / 100);
               VlrIcmsVenda := ValorBase * AliqVendaIcms1 / 100;
            end;
         end;
      end;

      if Opcao=2 then begin {Quando ja se tem o preco de venda }
         if ClassificacaoEstadual=0 then begin
            if (AliqVendaIcms1<>0) then begin
               ValorBase    := iif(PrecoVenda>PautaVenda, PrecoVenda, PautaVenda);
               ValorBase    := ValorBase - (ValorBase * AliqVendaReducao / 100);
               VlrIcmsVenda := ValorBase * AliqVendaIcms1 / 100;
            end;
         end;
      end;

   end;
end;


{//-------------------------------------------------------------------------------------
   TotalDeDiasPeriodo - Retorna o número da diferença em dias num período

   DataInicial  - Primeira data (informada)
   DataFim      - Ultima data   (informada)
   DiaInicial01 - Determina se o dia da primeira data deve começar no dia 1 (true), ou
                se ultiliza o dia informado na data (false) - DEFAULT TRUE
   DiaFinal31   - Determina se o dia da ultima data deve ser o ultimo dia do mês da data
                informada(true), ou se ultiza o dia informado na data(false)

   OBS: Se DiaFinal31 igual a false, ele verifica se o dia da data informado é maior que
        o dia da data atual, se for ele utiliza o dia da data atual.
}//-------------------------------------------------------------------------------------
function TotalDeDiasPeriodo(DataInicio, DataFim :TDateTime;DiaInicial01:Boolean = true;DiaFinal31:Boolean = false) :Integer;
begin
   if DiaInicial01 then
      DataInicio := EncodeDate(YearOf(DataInicio), MonthOf(DataInicio), 01);

   if DiaFinal31 then
      DataFim  := EncodeDate(YearOf(DataFim), MonthOf(DataFim), DaysInAMonth(YearOf(DataFim),MonthOf(DataFim)))
   else if DataFim > Date then
      DataFim  := EncodeDate(YearOf(DataFim), MonthOf(DataFim), DayOf(Date));

   Result := DaysBetween(DataInicio,DataFim);

end;

function VerificaVcto(Dia, Mes, Ano:Integer): TDateTime;
begin
  if sUsaDiaUtil then
    result:= ProximoDiaUtil(EncodeDate(Ano, Mes, Dia))
  else
    result:= EncodeDate(Ano, Mes, Dia);
end;

{ para facilitar a leitura dos parametros }
function GetParmEco(Const ModuloEco: TModulosEco; Parametro: string; TipoCampo: TFieldType=ftBoolean): Variant;
begin
  Case ModuloEco of
    meCaixa     : Result:= ParametroAtivo('TCXAPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    meBanco     : Result:= ParametroAtivo('TBANPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    mePagar     : Result:= ParametroAtivo('TPAGPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    meReceber   : Result:= ParametroAtivo('TRECPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    meVenda     : Result:= ParametroAtivo('TVENPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    meEstoque   : Result:= ParametroAtivo('TESTPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    meExportacao: Result:= ParametroAtivo('TEXPPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
    meMadeireira: Result:= ParametroAtivo('TMADPARAMETRO', Parametro, Geral.Empresa, TipoCampo);
  end;
end;

{ para facilitar a leitura dos parametros, assume default ftBoolean }
function ParamAtivo(const ecPar: TParmsEco): Variant;
begin
  case ecPar of
    peAutenticaDocumento     :Result:=GetParmEco(meReceber, 'ImpAutenticacao');

    peUsaRegistradora        :Result:=GetParmEco(meVenda, 'UsaRegistradora');
    peUsaDeptoCrediario      :Result:=GetParmEco(meVenda, 'UsaDeptoCrediario');
    peEmiteBoleto            :Result:=GetParmEco(meVenda, 'EmiteBoleto');
    peEmitePedido            :Result:=GetParmEco(meVenda, 'EmitePedido');
    peEmitePedidoRecibo      :Result:=GetParmEco(meVenda, 'EmitePedidoRecibo');
    peEmiteContrato          :Result:=GetParmEco(meVenda, 'EmiteContrato');
    peEmiteNotaPromissoria   :Result:=GetParmEco(meVenda, 'EmiteNotaPromissoria');
    peEmiteEntrega           :Result:=GetParmEco(meVenda, 'EmiteEntrega');
    peEmiteListaSeparacao    :Result:=GetParmEco(meVenda, 'EmiteListaSeparacao');
    peEmiteRecibo            :Result:=GetParmEco(meVenda, 'EmiteRecibo');
    peEmiteDuplicata         :Result:=GetParmEco(meVenda, 'EmiteDuplicata');
    peEmiteCarne             :Result:=GetParmEco(meVenda, 'EmiteCarne');
    peEmiteNotaFiscal        :Result:=GetParmEco(meVenda, 'EmiteNotaFiscal');
    peEmiteCupomFiscal       :Result:=GetParmEco(meVenda, 'EmiteCupomFiscal');
    peEmiteCupomVinculado    :Result:=GetParmEco(meVenda, 'EmiteCupomVinculado');

    peEditaObservacaoNota    :Result:=GetParmEco(meVenda, 'EditaObservacaoNota');
    peConfirmaBoleto         :Result:=GetParmEco(meVenda, 'ConfirmaBoleto');
    peConfirmaCarne          :Result:=GetParmEco(meVenda, 'ConfirmaCarne');
    peConfirmaDuplicata      :Result:=GetParmEco(meVenda, 'ConfirmaDuplicata');
    peConfirmaPromissoria    :Result:=GetParmEco(meVenda, 'ConfirmaPromissoria');
    peConfirmaContrato       :Result:=GetParmEco(meVenda, 'ConfirmaContrato');
    peConfirmaPedido         :Result:=GetParmEco(meVenda, 'ConfirmaPedido');
    peConfirmaPedidoRecibo   :Result:=GetParmEco(meVenda, 'ConfirmaPedidoRecibo');
    peConfirmaEntrega        :Result:=GetParmEco(meVenda, 'ConfirmaEntrega');
    peConfirmaSeparacao      :Result:=GetParmEco(meVenda, 'ConfirmaSeparacao');
    peConfirmaRecibo         :Result:=GetParmEco(meVenda, 'ConfirmaRecibo');

    peBoletoNaVenda          :Result:=GetParmEco(meVenda, 'BoletoNaVenda');
    peCarneNaVenda           :Result:=GetParmEco(meVenda, 'CarneNaVenda');
    peDuplicataNaVenda       :Result:=GetParmEco(meVenda, 'DuplicataNaVenda');
    pePromissoriaNaVenda     :Result:=GetParmEco(meVenda, 'PromissoriaNaVenda');
    peContratoNaVenda        :Result:=GetParmEco(meVenda, 'ContratoNaVenda');
    pePedidoNaVenda          :Result:=GetParmEco(meVenda, 'PedidoNaVenda');
    peReciboNaVenda          :Result:=GetParmEco(meVenda, 'ReciboNaVenda');
    peComprovanteNaVenda     :Result:=GetParmEco(meVenda, 'ComprovanteNaVenda');
    peListaNaVenda           :Result:=GetParmEco(meVenda, 'ListaNaVenda');

    peApresentaTelaDetalhes  :Result:=GetParmEco(meVenda, 'ApresentaTelaDetalhes');

    peImpNomeCliente         :Result:=GetParmEco(meVenda, 'ImpNomeCliente     ');
    peImpEnderecoCliente     :Result:=GetParmEco(meVenda, 'ImpEnderecoCliente ');
    peImpCidadeCliente       :Result:=GetParmEco(meVenda, 'ImpCidadeCliente   ');
    peImpCnpjCpfCliente      :Result:=GetParmEco(meVenda, 'ImpCnpjCpfCliente  ');
    peImpDocumentoOrigem     :Result:=GetParmEco(meVenda, 'ImpDocumentoOrigem ');
    peImpComprador           :Result:=GetParmEco(meVenda, 'ImpComprador       ');
    peImpParcelas            :Result:=GetParmEco(meVenda, 'ImpParcelas        ');
    peImpMensagem            :Result:=GetParmEco(meVenda, 'ImpMensagem        ');

    peMensagem1              :Result:=GetParmEco(meVenda, 'Mensagem1          ', ftString);
    peMensagem2              :Result:=GetParmEco(meVenda, 'Mensagem2          ', ftString);
    peImpMediaConsumo        :Result:=GetParmEco(meVenda, 'ImpMediaConsumo    ');

    // estoques
    peEstoqueNegativo        :Result:=GetParmEco(meEstoque, 'EstoqueNegativo'    ,ftBoolean);
    peUsaOrdemProducao       :Result:=GetParmEco(meEstoque, 'UsaOrdemProducao'   ,ftBoolean);

    // madeireira
    peUsaRomaneio            :Result:=GetParmEco(meMadeireira, 'UsaRomaneio'     ,ftBoolean);
    peEditaRomaneio          :Result:=GetParmEco(meMadeireira, 'EditaRomaneio'   ,ftBoolean);
    peImprimeRomaneio        :Result:=GetParmEco(meMadeireira, 'ImprimeRomaneio' ,ftBoolean);
    peEditaPauta             :Result:=GetParmEco(meMadeireira, 'EditaPauta'      ,ftBoolean);
    peAlteraPauta            :Result:=GetParmEco(meMadeireira, 'AlteraPauta'     ,ftBoolean);
    peIncentivoFiscal        :Result:=GetParmEco(meMadeireira, 'IncentivoFiscal' ,ftCurrency);
    peFundei                 :Result:=GetParmEco(meMadeireira, 'Fundei'          ,ftCurrency);
    peFethab                 :Result:=GetParmEco(meMadeireira, 'Fethab'          ,ftCurrency);
  end;

end;

{ le um parametro da tabela - assume default ftBoolean se nao for passado }
function ParametroAtivo(Tabela, Parametro, Empresa: string; TipoCampo: TFieldType=ftBoolean): Variant;
var
   QueryPesquisa: TFDQuery;
   Aux: string;
begin
   case TipoCampo of
     ftBoolean  : Result := False;
     ftString   : Result := '';
     ftInteger  : Result := 0;
     ftCurrency : Result := 0.0;
   end;

   QueryPesquisa := TFDQuery.Create(nil);
   QueryPesquisa.Connection := FDataModule.FDConnection;

   try

     with QueryPesquisa do
     begin
       Aux := TrimDup(Format('SELECT %S as Param FROM %S WHERE Empresa = %S',
                     [Parametro, Tabela, QuotedStr(Empresa)]));
       Sql.Clear;
       Sql.Add(Aux);
       if not EcoOpen(QueryPesquisa) then
         Mensagem('Parâmetro não encontrado ('+Aux+')', tmErro)
       else
         case TipoCampo of
           ftBoolean  :Result:=iif(FieldByName('Param').AsString='S', True, False);
           ftString   :Result:=    FieldByName('Param').AsString;
           ftInteger  :Result:=    FieldByName('Param').AsInteger;
           ftCurrency :Result:=    FieldByName('Param').asCurrency;
           ftFloat    :Result:=    FieldByName('Param').AsFloat;
           ftDate     :Result:=    FieldByName('Param').asDateTime;
         end;
     end;
   finally
      QueryPesquisa.Close;
     QueryPesquisa.Free;
   end;
end;

// usado para retornar a proxima sequencia q sera inserida nos lactos do caixa
function PegaAProximaSequenciaDoCaixa(var QueryPesquisa: TSQLQuery; Caixa: string; DataCaixa:TDateTime): Integer;
begin
  with QueryPesquisa do
  begin
     Sql.Clear;
     Sql.Add('SELECT MAX(Numero) As Numero  ' +
             'FROM   TCxaLancamento         ' +
             'WHERE  Empresa = :Empresa AND ' +
             '       Caixa   = :Caixa   AND ' +
             '       Data    = :Data        ');
     ParamByName('Empresa').asString        := Empresa;
     ParamByName('Caixa').asString          := Caixa;
     ParamByName('Data').asDate             := DataCaixa;
     EcoOpen(QueryPesquisa);
     Result := StrToIntDef(FieldByName('Numero').asString, 0) + 1;
     Close;
  end;
end;

// usado para retornar a proxima sequencia q sera inserida nos lactos do caixa
function PegaAProximaSequenciaDoCaixa(var QueryPesquisa: TFDQuery; Caixa: string; DataCaixa:TDateTime): Integer;
const
  CSql = 'SELECT MAX(Numero) As Numero  ' +
         'FROM   TCxaLancamento         ' +
         'WHERE  Empresa = :Empresa AND ' +
         '       Caixa   = :Caixa   AND ' +
         '       Data    = :Data        ' ;

begin
  Result := 1;
  try
    QueryPesquisa.Close;
    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add(CSql);
    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Caixa').AsString   := Caixa;;
    QueryPesquisa.ParamByName('Data').asDate      := DataCaixa;
    if EcoOpen(QueryPesquisa) then
    begin
      Result := StrToIntDef(QueryPesquisa.FieldByName('Numero').asString, 0) + 1;
    end;
  except
    raise;
  end;
end;

// usado para retornar a proxima sequencia q sera inserida nos lactos do banco
function PegaAProximaSequenciaDoBanco(var QueryPesquisa: TSQLQuery; Banco: string; DataBanco:TDateTime): Integer;
begin
  with QueryPesquisa do begin
     { Seleciona o numero do movimento no banco + 1 }
     Sql.Clear;
     Sql.Add('SELECT Max(Numero) As Numero  ' +
             'FROM   TBanMovimento          ' +
             'WHERE  Empresa = :Empresa AND ' +
             '       Conta   = :Conta   AND ' +
             '       Data    = :Data        ');
     ParamByName('Empresa').AsString := Empresa;
     ParamByName('Conta').AsString   := Banco;
     ParamByName('Data').AsDate      := DataBanco;
     EcoOpen(QueryPesquisa);
     Result := StrToIntDef(FieldByName('Numero').asString, 0) + 1;
     Close;
  end;
end;

function PegaAProximaSequenciaDoBanco(var QueryPesquisa: TFDQuery; Banco: string; DataBanco:TDateTime): Integer;
begin
  try
    with QueryPesquisa do
    begin
       { Seleciona o numero do movimento no banco + 1 }
       Sql.Clear;
       Sql.Add('SELECT Max(Numero) As Numero  ' +
               'FROM   TBanMovimento          ' +
               'WHERE  Empresa = :Empresa AND ' +
               '       Conta   = :Conta   AND ' +
               '       Data    = :Data        ');
       ParamByName('Empresa').AsString := Empresa;
       ParamByName('Conta').AsString   := Banco;
       ParamByName('Data').AsDate      := DataBanco;
       EcoOpen(QueryPesquisa);
       Result := StrToIntDef(FieldByName('Numero').asString, 0) + 1;
    end;
  finally
    QueryPesquisa.Close;
  end;
end;


function TrimDup(const Str :string): string;
begin
  Result := TrimLeft(TrimRight(Str));
  while Pos('  ',Result) > 0 do Delete(Result, Pos('  ', Result), 1);
end;

procedure SetPrinterPage1(Width, Height : LongInt);
var
   Device : array[0..255] of char;
   Driver : array[0..255] of char;
   Port   : array[0..255] of char;
   hDMode : THandle;
   PDMode : PDEVMODE;
begin
     Printer.GetPrinter(Device, Driver, Port, hDMode);
     If hDMode <> 0 then
     begin
          pDMode := GlobalLock( hDMode );
          If pDMode <> nil then
          begin
               pDMode^.dmPaperSize   := DMPAPER_USER;
               pDMode^.dmPaperWidth  := Width;
               pDMode^.dmPaperLength := Height;
               pDMode^.dmFields := pDMode^.dmFields or DM_PAPERSIZE;
               GlobalUnlock( hDMode );
          end;
     end;
end;

// retorna um codigo formatado com o ano AnnnnnD;
function MontaCodigoChegada(Cod, Dig: string): string;
var b: Word;
begin
   if not CharInSet(Dig[1], ['0','1']) then Dig := '0';
   B      := YearOf(DataCaixa) - 2000;
   Result := InttoStr(B) + ZerosLeft(Cod, 5) + Dig;
end;

function StrToByte(sStr: string): byte;
var C1,C2:char;
begin
  C1 := UpperCase(sStr[1])[1];
  C2 := UpperCase(sStr[2])[1];
  Result := GeralHexChar(C1) shl 4 + GeralHexChar(C2);
end;

function GeralTimeSQL(const Data: TDateTime):string;
begin
  try
    Result := FormatDateTime('hh:mm:ss',Data);
  except
    Result := 'Null';
  end;
end;

function  GeralQTimeSQL(const Data: TDateTime):string;
begin
 Result := QuotedStr(GeralTimeSQL(Data));
end;

function  GeralDataSQL(const Data: TDateTime):string;
begin
  try
    Result := FormatDateTime('dd.mm.yyyy',Data);
  except
    Result := 'Null';
  end;
end;

function  GeralQDataSQL(const Data: TDateTime):string;
begin
    Result := QuotedStr(GeralDataSQL(Data));
end;

function GeralHexChar(const c: Char): Byte;
begin
  case c of
    '0'..'9': Result :=  Byte(c) - Byte('0');
    'a'..'f': Result := (Byte(c) - Byte('a')) + 10;
    'A'..'F': Result := (Byte(c) - Byte('A')) + 10;
  else
    Result := 0;
  end;
end;
//Pega o nome do computador
function  NomeDeComputador : string;
var Buf : array [0..256] of Char; Nsize : Dword;
begin
  Nsize := 256;
  GetComputerName(Buf, Nsize);
  Result := String(Buf);
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
function ChamaTelaDeExportar(NomeDaLista:String;NomeDaQuery:TSQLQuery):boolean;
begin
   Result := False;
   Form_SalvaListaEmArquivo := TForm_SalvaListaEmArquivo.Create(Application);
  try
    Form_SalvaListaEmArquivo.Titulo := NomeDaLista;
    Form_SalvaListaEmArquivo.Query  := NomeDaQuery;
    Form_SalvaListaEmArquivo.ShowModal;
    finally
      Form_SalvaListaEmArquivo.Free;
      Form_SalvaListaEmArquivo := nil;
    end;
end;
{$ENDIF}

{ ================================================================================== }
{ retorna em um string as unidades de rede  D6-WIN2000-S                             }
{ ================================================================================== }
function DrivesDaRede: string;
var
    i, dwResult: DWORD;
    hEnum      : THANDLE;
    lpnrDrv, lpnrDrvLoc: PNETRESOURCE;
    s: string;
// const
    cbBuffer: DWORD; // = 16384;
    cEntries: DWORD; // = $FFFFFFFF;
begin
  dwResult := WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0, nil, hEnum);
  if dwResult <> NO_ERROR then begin
     ShowMessage('Não é possível enumerar as conexões da rede.');
     Exit;
  end;
  S := '';
  repeat
    lpnrDrv   := PNETRESOURCE(GlobalAlloc(GPTR, cbBuffer));
    dwResult  := WNetEnumResource(hEnum, cEntries, lpnrDrv, cbBuffer);
    if dwResult = NO_ERROR then begin
      s := 'Unidades de red:'#13#10;
      lpnrDrvLoc := lpnrDrv;
      for i := 0 to cEntries - 1 do begin
        if lpnrDrvLoc^.lpLocalName <> nil then
           s := s + lpnrDrvLoc^.lpLocalName + #9 + lpnrDrvLoc^.lpRemoteName + #13#10;
        Inc(lpnrDrvLoc);
      end;
    end else if dwResult <> ERROR_NO_MORE_ITEMS then begin
      s := s + 'Não é possível completar a numeração dos drives de rede.';
      GlobalFree(HGLOBAL(lpnrDrv));
      break;
    end;
    GlobalFree(HGLOBAL(lpnrDrv));
  until dwResult = ERROR_NO_MORE_ITEMS;
  WNetCloseEnum(hEnum);

  if s = '' then
     s := 'Não existem conexões na rede.';

  Result := s;
end;

function DataEstaNoPeriodo(DataINI, DataFim, Data: TDateTime): boolean;
begin
  if (Data < DataINI) or (Data > DataFIM) then
    Result := False
  else
    Result := True;
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure SalvaPadraoRelatorio(Classe:TForm;Query:TSQLQuery;Usuario,Programa,sNome,sObs:String);
var Contador:Integer;
    Componente  :TComponent;
    IDSeq       :Integer;
    function ProxID(Tabela,Campo:String):Integer;
    begin
      With Query do begin
        Sql.Clear;
        Sql.Add(Format('Select Max(%s) as ID From %s ',
                [Campo,
                 Tabela]));
        if EcoOpen(Query) then
           Result := StrToIntDef(FieldByName('ID').asString, 0) + 1
        else
           Result := 1;   
      end;
    end;

    procedure GravaPadrao(Modulo,NomeComponente,Valor:String);
    begin
      With Query do begin
        Sql.Clear;
        Sql.Add('Insert Into TGerPadraoTelaItens(IDTela,         '+
                '                                Programa,       '+
                '                                Usuario,        '+
                '                                Modulo,         '+
                '                                ComponentName,  '+
                '                                Valor )         '+
                'Values                                          '+
                '                               (:IDTela,        '+
                '                                :Programa,      '+
                '                                :Usuario,       '+
                '                                :Modulo,        '+
                '                                :ComponentName, '+
                '                                :Valor )        ');
        ParamByName('IDTela').AsInteger       := IDSeq;
        ParamByName('Programa').AsString      := Programa;
        ParamByName('Usuario').AsString       := Usuario;
        ParamByName('Modulo').AsString        := Modulo;
        ParamByName('ComponentName').AsString := NomeComponente;
        ParamByName('Valor').AsString         := Valor;
        EcoExecSql(Query);
      end;
    end;
begin
  IDSeq := ProxID('TGERPADRAOTELA','ID');
  With Query do begin
    Sql.Clear;
    Sql.Add(Format('Insert Into TGerPadraoTela (ID,          '+
                   '                            Programa,    '+
                   '                            Usuario,     '+
                   '                            Nome,        '+
                   '                            Observacao)  '+
                   'Values                                   '+
                   '                           (%d,          '+
                   '                            %s,          '+
                   '                            %s,          '+
                   '                            %s,          '+
                   '                            %s         ) ',
            [IDSeq,
             QuotedStr(Programa),
             QuotedStr(Usuario),
             QuotedStr(sNome),
             QuotedStr(sObs)]));
    EcoExecSql(Query);
  end;

   with Classe do begin
     for Contador := 0 to ComponentCount - 1 do begin
       Componente := Components[Contador];
       if (Componente is TCheckBox) then begin
         GravaPadrao(Name,TCheckBox(Componente).Name,iif(TCheckBox(Componente).Checked,'1','0'));
       end;
       if (Componente is TRadioButton) then begin
         GravaPadrao(Name, TRadioButton(Componente).Name, iif(TRadioButton(Componente).Checked=True,'1','0'));
       end;
       if (Componente is TNumEdit) then begin
         GravaPadrao(Name,TNumEdit(Componente).Name,  TNumEdit(Componente).Text);
       end;
       if (Componente is TEditBox) then begin
         if TEditBox(Componente).Text<>'' then begin
           GravaPadrao(Name,TEditBox(Componente).Name,  TEditBox(Componente).Text);
         end;
       end;
       if (Componente is TComboBox) then begin
          GravaPadrao(Name,TComboBox(Componente).Name,  intToStr(TComboBox(Componente).ItemIndex));
       end;
     end;
   end;
end;

procedure SalvaPadraoRelatorio(Classe:TForm;Query:TFDQuery;Usuario,Programa,sNome,sObs:String);
var Contador:Integer;
    Componente  :TComponent;
    IDSeq       :Integer;
    function ProxID(Tabela,Campo:String):Integer;
    begin
      With Query do begin
        Sql.Clear;
        Sql.Add(Format('Select Max(%s) as ID From %s ',
                [Campo,
                 Tabela]));
        if EcoOpen(Query) then
           Result := StrToIntDef(FieldByName('ID').asString, 0) + 1
        else
           Result := 1;
      end;
    end;

    procedure GravaPadrao(Modulo,NomeComponente,Valor:String);
    begin
      With Query do begin
        Sql.Clear;
        Sql.Add('Insert Into TGerPadraoTelaItens(IDTela,         '+
                '                                Programa,       '+
                '                                Usuario,        '+
                '                                Modulo,         '+
                '                                ComponentName,  '+
                '                                Valor )         '+
                'Values                                          '+
                '                               (:IDTela,        '+
                '                                :Programa,      '+
                '                                :Usuario,       '+
                '                                :Modulo,        '+
                '                                :ComponentName, '+
                '                                :Valor )        ');
        ParamByName('IDTela').AsInteger       := IDSeq;
        ParamByName('Programa').AsString      := Programa;
        ParamByName('Usuario').AsString       := Usuario;
        ParamByName('Modulo').AsString        := Modulo;
        ParamByName('ComponentName').AsString := NomeComponente;
        ParamByName('Valor').AsString         := Valor;
        EcoExecSql(Query);
      end;
    end;
begin
  IDSeq := ProxID('TGERPADRAOTELA','ID');
  With Query do begin
    Sql.Clear;
    Sql.Add(Format('Insert Into TGerPadraoTela (ID,          '+
                   '                            Programa,    '+
                   '                            Usuario,     '+
                   '                            Nome,        '+
                   '                            Observacao)  '+
                   'Values                                   '+
                   '                           (%d,          '+
                   '                            %s,          '+
                   '                            %s,          '+
                   '                            %s,          '+
                   '                            %s         ) ',
            [IDSeq,
             QuotedStr(Programa),
             QuotedStr(Usuario),
             QuotedStr(sNome),
             QuotedStr(sObs)]));
    EcoExecSql(Query);
  end;

   with Classe do begin
     for Contador := 0 to ComponentCount - 1 do begin
       Componente := Components[Contador];
       if (Componente is TCheckBox) then begin
         GravaPadrao(Name,TCheckBox(Componente).Name,iif(TCheckBox(Componente).Checked,'1','0'));
       end;
       if (Componente is TRadioButton) then begin
         GravaPadrao(Name, TRadioButton(Componente).Name, iif(TRadioButton(Componente).Checked=True,'1','0'));
       end;
       if (Componente is TNumEdit) then begin
         GravaPadrao(Name,TNumEdit(Componente).Name,  TNumEdit(Componente).Text);
       end;
       if (Componente is TEditBox) then begin
         if TEditBox(Componente).Text<>'' then begin
           GravaPadrao(Name,TEditBox(Componente).Name,  TEditBox(Componente).Text);
         end;
       end;
       if (Componente is TComboBox) then begin
          GravaPadrao(Name,TComboBox(Componente).Name,  intToStr(TComboBox(Componente).ItemIndex));
       end;
     end;
   end;
end;


procedure LePadraoDeRelatorio(Classe:TForm;Query:TSQLQuery;IDRelatorio:Integer);
var Contador, indice :Integer;
    Componente  :TComponent;
    Aux         :String;

    function LeValor(sModulo,sUsuario,NomeComponente:String):String;
    begin
      Result := '';
      With Query do begin
        Sql.Clear;
        Sql.Add(Format('Select Valor From TGerPadraoTelaItens '+
                       'Where Modulo      = %s                '+
                       'And   Usuario     = %s                '+
                       'And ComponentName = %s                '+
                       'And IDTela        = %d                ',
                [QuotedStr(sModulo),
                 QuotedStr(sUsuario),
                 QuotedStr(NomeComponente),
                 IDRelatorio]));
        if EcoOpen(Query) then
          Result := FieldByName('Valor').AsString;
      end;
    end;

begin
   with Classe do begin
     for Contador := 0 to ComponentCount - 1 do begin
       Componente := Components[Contador];
       if (Componente is TCheckBox) then begin
          TCheckBox(Componente).Checked := LeValor(Name,Usuario,TCheckBox(Componente).Name)='1';
       end;
       if (Componente is TRadioButton) then begin
          TRadioButton(Componente).Checked := LeValor(Name,Usuario,TCheckBox(Componente).Name)='1';
       end;
       if (Componente is TNumEdit) then begin
          TNumEdit(Componente).Text := LeValor(Name,Usuario,TCheckBox(Componente).Name);
       end;
       if (Componente is TEditBox) then begin
          TEditBox(Componente).Text := LeValor(Name,Usuario,TCheckBox(Componente).Name);
       end;
       if (Componente is TComboBox) then begin
          Aux  := LeValor(Name,Usuario,TComboBox(Componente).Name);
          if(Aux<>'') then begin
             indice := StrToInt(aux);
             if(indice >= TComboBox(Componente).Items.Count) then
                indice:=0;
             TComboBox(Componente).ItemIndex := indice;
          end;
       end;
     end;
   end;
end;

procedure LePadraoDeRelatorio(Classe:TForm;Query:TFDQuery;IDRelatorio:Integer);  overload;
var Contador, indice :Integer;
    Componente  :TComponent;
    Aux         :String;

    function LeValor(sModulo,sUsuario,NomeComponente:String):String;
    begin
      Result := '';
      With Query do begin
        Sql.Clear;
        Sql.Add(Format('Select Valor From TGerPadraoTelaItens '+
                       'Where Modulo      = %s                '+
                       'And   Usuario     = %s                '+
                       'And ComponentName = %s                '+
                       'And IDTela        = %d                ',
                [QuotedStr(sModulo),
                 QuotedStr(sUsuario),
                 QuotedStr(NomeComponente),
                 IDRelatorio]));
        if EcoOpen(Query) then
          Result := FieldByName('Valor').AsString;
      end;
    end;

begin
   with Classe do begin
     for Contador := 0 to ComponentCount - 1 do begin
       Componente := Components[Contador];
       if (Componente is TCheckBox) then begin
          TCheckBox(Componente).Checked := LeValor(Name,Usuario,TCheckBox(Componente).Name)='1';
       end;
       if (Componente is TRadioButton) then begin
          TRadioButton(Componente).Checked := LeValor(Name,Usuario,TCheckBox(Componente).Name)='1';
       end;
       if (Componente is TNumEdit) then begin
          TNumEdit(Componente).Text := LeValor(Name,Usuario,TCheckBox(Componente).Name);
       end;
       if (Componente is TEditBox) then begin
          TEditBox(Componente).Text := LeValor(Name,Usuario,TCheckBox(Componente).Name);
       end;
       if (Componente is TComboBox) then begin
          Aux  := LeValor(Name,Usuario,TComboBox(Componente).Name);
          if(Aux<>'') then begin
             indice := StrToInt(aux);
             if(indice >= TComboBox(Componente).Items.Count) then
                indice:=0;
             TComboBox(Componente).ItemIndex := indice;
          end;
       end;
     end;
   end;
end;
{$ENDIF}

// grava um arquivo de log no diretorio apontado pela variavel "ServidorLog" do Eco.INI
// use: GravaArquivoDeLog("nome do arquivo de log", "e string de log");
// ex:GravaArquivoDeLog('GER408LA.LOG', FormatDateTime('dd/mm/yy hh:mm:ss "- "', Now) + Aux);

procedure GravaArquivoDeLog(const Arquivo, StrLog : string);
var F : TextFile;
    PathAndFile : string;
begin
  {$i-} // desabilita erros em disco
  try
    PathAndFile := ExtractFilePath(ServidorLog);
    PathAndFile := IncludeTrailingPathDelimiter(PathAndFile) + Arquivo;
    AssignFile(F, PathAndFile);
    if not FileExists(PathAndFile) then
       Rewrite(F);
    Append(F);
    WriteLn(F, StrLog);
  finally
    CloseFile(F);
  end;
  {$i+} // habilita erros em disco
end;

function ValorEntre(ValorInicio,ValorFim,ValorPraTeste:Currency):Boolean; 
begin
  Result := (ValorPraTeste >= ValorInicio) and (ValorPraTeste <= ValorFim);
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure ImprimirRelatorioCrystal(const CaminhoRelatorio: string; const Parametros: array of variant; const EmVideo: Boolean = True;
                                   const EscolheImpressora: Boolean = True; const NomeImpressora: string = ''; const QuantidadeVias: Integer = 1;
                                   const OrientacaoPapel: string = 'R');
begin
   Application.CreateForm(TFPreviewCrystal, FPreviewCrystal);
   try
      FPreviewCrystal.Enter(CaminhoRelatorio, Parametros, EmVideo, EscolheImpressora, NomeImpressora, QuantidadeVias, OrientacaoPapel);
   finally
      FPreviewCrystal.Destroy;
   end;
end;

procedure ChamaFormCrystal(Relatorio: string; Parametros: array of variant; EmVideo: Boolean=True;
                           EscolheImpressora: Boolean=True; NomeImpressora:String=''; QuantidadeVias: Integer=1;
                           OrientacaoPapel: String = 'R');

var
   FSR: TSearchRec;
   FRelatorio: string;
   sRelatorio: string;
begin
   if (FindFirst(DiretorioDeRelatorios + Relatorio, SysUtils.faAnyFile, FSR) = 0) then
   begin
      Application.CreateForm(TFPesquisaRelatorioCrystal, FPesquisaRelatorioCrystal);
      try
      with FPesquisaRelatorioCrystal do
      begin
         try
            Repeat
               CDSRelatorios.Append;
               CDSRelatoriosNOME.AsString := FSR.Name;
               CDSRelatorios.Post;
            until FindNext(FSR) <> 0;

            if (CDSRelatorios.RecordCount > 1) then
            begin
               CDSRelatorios.First;
               ShowModal;

               if ModalResult = mrOk then
                  FRelatorio := DiretorioDeRelatorios + CDSRelatoriosNOME.AsString
               else
                  Exit;
            end;
            FRelatorio := DiretorioDeRelatorios + CDSRelatoriosNOME.AsString;
         finally
            SysUtils.FindClose(FSR);
            CDSRelatorios.Close;
            FreeAndNil(FPesquisaRelatorioCrystal);
         end;
      end;
      finally
         FreeAndNil(FPesquisaRelatorioCrystal);
      end;
   end
   else
   begin
      sRelatorio := Relatorio;

      if (Pos('*', sRelatorio) > 0) then
         sRelatorio := Copy(sRelatorio, 1, (Length(sRelatorio) - 1));

      FRelatorio := DiretorioDeRelatorios + sRelatorio + '.rpt';
   end;

   Application.CreateForm(TFPreviewCrystal, FPreviewCrystal);
   try
      FPreviewCrystal.Enter(FRelatorio, Parametros, EmVideo, EscolheImpressora, NomeImpressora, QuantidadeVias, OrientacaoPapel);
   finally
      FPreviewCrystal.Destroy;
   end;
end;

{$ENDIF}

procedure CreateODBCDriver;

type
  TSQLConfigDataSource = function( hwndParent: HWND; fRequest: WORD; lpszDriver: MarshaledAString; lpszAttributes: PMarshaledAString ): BOOL; stdcall;

const
  ODBC_ADD_DSN        = 1; // Adiciona uma fonte de dados (data source)
  ODBC_CONFIG_DSN     = 2; // Configura a fonte de dados (data source)
  ODBC_REMOVE_DSN     = 3; // Remove a fonte de dados (data source)
  ODBC_ADD_SYS_DSN    = 4; // Adiciona um DSN no sistema
  ODBC_CONFIG_SYS_DSN = 5; // Configura o DSN do sistema
  ODBC_REMOVE_SYS_DSN = 6; // Remove o DSN do sistema

  fbDriver: AnsiString = 'Firebird/InterBase(r) driver';

var
  pFn: TSQLConfigDataSource;
  hLib: HMODULE;
  strFile, strClient, strAttr: UnicodeString;
  AnsiAttr: AnsiString;

  function TryWithRequest(const ARequest: Cardinal): Boolean;
  begin
    Result := pFn(0, ARequest, Pointer(fbDriver), Pointer(AnsiAttr));
  end;

begin
  strFile := FDataModule.sConnection.Params.Values['DataBase'];
  strClient := FDataModule.sConnection.VendorLib;

  hLib := LoadLibrary('ODBCCP32'); // Carregando a biblioteca do driver ODBC
  if (hLib = 0) then
  begin
    ShowMessage('O sistema não pôde carregar a DLL: ''ODBCCP32''');
    Exit;
  end;

  try
    pFn := GetProcAddress(hLib, 'SQLConfigDataSource');
    if Assigned(pFn) then
    begin
      FmtStr(strAttr,
        'DSN        =EcoSistemas'#0 +
        'DATABASE   =%s'#0 +
        'USER       =SYSDBA'#0 +
        'PWD        =masterkey'#0 +
        'CLIENT     =%s'#0 +
        'DIALECT    =3'#0 +
        'READONLY   =1'#0 +
        'Exclusive  =1'#0 +
        'Description=Banco de dados Ecocentauro sistemas'#0#0, [strFile, strClient]);

      AnsiAttr := AnsiString(strAttr);

      if (TryWithRequest(ODBC_ADD_SYS_DSN) or TryWithRequest(ODBC_ADD_DSN)) then
        ShowMessage('A conexão DSN EcoSistemas foi criada com sucesso.')
      else
        ShowMessage('Falha ao tentar criar EcoSistemas.');
    end else
    begin
      ShowMessageFmt('Não foi encontrado o ponto de entrada do método: ''SQLConfigDataSource'' na DLL: ''%s''', [GetModuleName(hLib)]);
    end;
  finally
    FreeLibrary(hLib);
  end;
end;

function DescricaoDaGrade(Query:TFDQuery;sProduto:String):String;
begin
   With Query do begin 
      Sql.Clear;
      Sql.Add(Format('Select                                                          '+
                     '   Pdg.Descricao||'' ''||                                       '+
                     '   (SELECT TEstGradeLinha.Descricao                             '+
                     '      FROM TEstGradeCelula                                      '+
                     '     INNER JOIN TEstGradeLinha                                  '+
                     '        ON TEstGradeCelula.IDGrade = TEstGradeLinha.IDGrade     '+
                     '       AND TEstGradeCelula.IDLinha = TEstGradeLinha.IDLinha     '+
                     '     WHERE TEstGradeCelula.IDCelula = Pdg.IDCelula)             '+
                     '   ||'' / ''||                                                  '+
                     '   (SELECT TEstGradeColuna.Descricao                            '+
                     '       FROM TEstGradeCelula                                     '+
                     '      INNER JOIN TEstGradeColuna                                '+
                     '         ON TEstGradeCelula.IDGrade = TEstGradeColuna.IDGrade   '+
                     '        AND TEstGradeCelula.IDColuna = TEstGradeColuna.IDColuna '+
                     '      WHERE TEstGradeCelula.IDCelula = Pdg.IDCelula) as Descr   '+
                     'From TEstProdutoGeral Pdg                                       '+
                     'Where Pdg.Codigo = %s                                           ',
              [QuotedStr(sProduto)]));
      if EcoOpen(Query) then
         Result := FieldByName('Descr').AsString;
   end;
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
procedure ImprimeCheque(Numero: string; Banco: string; Valor: Double; Data: TDate; Cidade: string; Favorecido: string);
var IniFile: TIniFile;
    FACBrCHQ: TACBrCHQ;
begin
   IniFile  := TIniFile.Create(UPPERCASE(ChangeFileExt(Application.ExeName, '.ini')));
   FACBrCHQ := TACBrCHQ.Create(nil);
   try
      with IniFile do begin
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'BEMATECH'         then FACBrCHQ.Modelo := chqBematech;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'CHRONOS'          then FACBrCHQ.Modelo := chqChronos;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'IMPRESSORA COMUM' then FACBrCHQ.Modelo := chqImpressoraComum;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'IMPRESSORA ECF'   then FACBrCHQ.Modelo := chqImpressoraECF;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'PERTO'            then FACBrCHQ.Modelo := chqPerto;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'SCHALTER'         then FACBrCHQ.Modelo := chqSchalter;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'SOTOMAQ'          then FACBrCHQ.Modelo := chqSotomaq;
         if ReadString('CHEQUE', 'MARCA', 'MARCA') = 'URANO'            then FACBrCHQ.Modelo := chqUrano;

         FACBrCHQ.Porta := ReadString('CHEQUE', 'COM', 'COM');

         if (ReadString('CHEQUE', 'MARCA', 'MARCA') <> '') and (ReadString('CHEQUE', 'MARCA', 'MARCA') <> 'NENHUMA') then begin
            if ReadString('CHEQUE', 'COM', 'COM') <> '' then begin
               FACBrCHQ.Banco      := Banco;
               FACBrCHQ.Valor      := Valor;
               FACBrCHQ.Data       := Data;
               FACBrCHQ.Cidade     := Cidade;
//               FACBrCHQ.Favorecido := NomeEmpresa;
               try
                  FACBrCHQ.ImprimirCheque;
               except
                  on e: Exception do begin
                     Mensagem('Ocorreu um erro ao tentar imprimir o cheque, motivo: ' + e.Message+'.', tmInforma);
                  end;
               end;
            end else begin
               Mensagem('Informe a porta que a impressora esta conectada!', tmInforma);
               Exit;
            end;      
         end else begin
            Mensagem('Não existe impressora de cheques definida!', tmInforma);
            Exit;
         end;
      end;
   finally
      FreeAndNil(FACBrCHQ);
      FreeAndNil(IniFile);
   end;
end;


{$ENDIF}

function DivideValor(Valor1, Valor2: Extended): Extended;
begin
  Result := 0;
  if Valor2=0 then
    Exit;
    
  Result := Valor1/Valor2;
end;


{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }

{$IFDEF DEPURACAO}
function MensagemList(msg : String; TipoMsg : TTipoMsg; StringList : TStringList; qtdeCaracteresLinha : Integer = 0):Boolean;
begin
   application.createform(TfdtMensagemList_Nova, fdtMensagemList_Nova);
   try
      fdtMensagemList_Nova.Mensagem := Msg;
      fdtMensagemList_Nova.TipoMsg  := TipoMsg;
      fdtMensagemList_Nova.ListMensagem := StringList;
      fdtMensagemList_Nova.ShowModal;
      Result := fdtMensagemList_Nova.Tag = 1;
   finally
      fdtMensagemList_Nova.Free;
   end;
end;
{$ELSE}
function MensagemList(msg : String; TipoMsg : TTipoMsg; StringList : TStringList; qtdeCaracteresLinha : Integer = 0):Boolean;
begin
   application.createform(TfdtMensagemList, fdtMensagemList);
   try
      fdtMensagemList.Mensagem := Msg;
      fdtMensagemList.NumeroCaracteresPorLinha := qtdeCaracteresLinha;
      fdtMensagemList.TipoMsg  := TipoMsg;
      fdtMensagemList.ListMensagem :=  StringList;
      fdtMensagemList.ShowModal;
      Result := fdtMensagemList.Tag = 1;
   finally
      fdtMensagemList.Free;
   end;
end;
{$ENDIF}

{$ENDIF}

function ValidaNCMDigitado(const NcmInformado : String):Boolean;
var aux : Integer;
begin
   Result := False;

   if TryStrToInt(NcmInformado,aux) then
      Result := True;

   if Result then
      Result := Length(NcmInformado) = 8;

end;


function TipoTributacaoPeloCSF(CSF : String):Byte;
begin
   Result := StrToIntDef(Copy(CSF, 2, 2), 0);
end;

{$IFNDEF PAFECOBPL}  { Escondendo código desnecessário para o módulo PafEco.BPL (Projeto PAF) }
function BuscaComponentTela(Agrupador : TWinControl; Nome : String) : TControl;
var AuxCont : Integer;
begin
   Result := nil;
   with Agrupador do
   begin
      for AuxCont := 0 to ControlCount - 1 do
      begin
         Result := Controls[AuxCont];
         if SameText(Result.Name, Nome) then Exit;

         if Result is TDnBox then
         begin
            Result := BuscaComponentTela(TWinControl(Result), Nome);
            if SameText(Result.Name, Nome) then Exit;
         end;
         if Result is TTabSheet then
         begin
            Result := BuscaComponentTela(TWinControl(Result), Nome);
            if SameText(Result.Name, Nome) then Exit;
         end;
         if Result is TDnPanel then
         begin
            Result := BuscaComponentTela(TWinControl(Result), Nome);
            if SameText(Result.Name, Nome) then Exit;
         end;
         if Result is TcxTabSheet then
         begin
            Result := BuscaComponentTela(TWinControl(Result), Nome);
            if SameText(Result.Name, Nome) then Exit;
         end;
         if Result is TcxPageControl then
         begin
            Result := BuscaComponentTela(TWinControl(Result), Nome);
            if SameText(Result.Name, Nome) then Exit;
         end;
      end;
   end;
end;
{$ENDIF}

procedure AbrePeriodoAgrupamentoSemFinanceiro(var Registradora: string; var Periodo: Integer);
var
   QueryPesquisa: TFDQuery;
   IdPeriodo: Integer;
   NovoIdPeriodo: Integer;
   UsuarioPadraoPeriodo: string;
   PeriodoTransaction: TDBXTransaction;
   NovoPeriodoAberto: Boolean;

   procedure FecharPeriodoAgrupamentoSemFinanceiro;
   var
      QueryFechamento: TFDQuery;
      IdPeriodoAberto: Integer;
      CaixaPadrao: string;
      HistoricoEntradaReg: string;
      QueryAux: TFDQuery;
      IdLancamento: Integer;
      DataPeriodo: TDateTime;

      function NovoIdLancamento: Integer;
      begin
         Result := 1;

         with QueryAux do
         begin
            SQL.Clear;
            SQL.Add('SELECT MAX(Numero) As Numero ' +
                    '  FROM TCxaLancamento ' +
                    ' WHERE Empresa = :Empresa ' +
                    '   AND Caixa = :Caixa ' +
                    '   AND Data = ''TODAY''');

            ParamByName('Empresa').asString := Empresa;
            ParamByName('Caixa').AsString := CaixaPadrao;

            if EcoOpen(QueryAux) then
            begin
               Result := StrToIntDef(FieldByName('Numero').asString, 0) + 1;
               Close;
            end;
         end;
      end;
   begin
      QueryPesquisa := TFDQuery.Create(nil);
      QueryFechamento := TFDQuery.Create(nil);
      QueryAux := TFDQuery.Create(nil);
      try
         QueryPesquisa.Connection := FDataModule.FDConnection;
         QueryFechamento.Connection := FDataModule.FDConnection;
         QueryAux.Connection := FDataModule.FDConnection;

         with QueryPesquisa do
         begin
            CaixaPadrao := '';

            SQL.Clear;
            SQL.Add('SELECT CaixaPadrao, ' +
                    '       HistoricoEntradaReg ' +
                    '  FROM TCxaParametro ' +
                    ' WHERE Empresa = :Empresa');

            ParamByName('Empresa').AsString := Empresa;

            if EcoOpen(QueryPesquisa) then
            begin
               CaixaPadrao := FieldByName('CaixaPadrao').AsString;
               HistoricoEntradaReg := FieldByName('HistoricoEntradaReg').AsString;
            end;

            if (Trim(CaixaPadrao) = '') then
               Exit;

            SQL.Clear;
            SQL.Add('  SELECT IdPeriodo, DataAbertura ' +
                    '    FROM TVenPeriodo ' +
                    '   WHERE Empresa = :Empresa ' +
                    '     AND Registradora = :Registradora ' +
                    '     AND IdPeriodo <> :IdPeriodo ' +
                    '     AND DataFechamento IS NULL ' +
                    'ORDER BY DataAbertura DESC, horaAbertura DESC');

            ParamByName('Empresa').AsString := Empresa;
            ParamByName('Registradora').AsString := RegistradoraAgrupamentoSemFinanceiro;
            ParamByName('IdPeriodo').AsInteger := PeriodoAgrupamentoSemFinanceiro;

            if EcoOpen(QueryPesquisa) then
            begin
               while not Eof do
               begin
                  IdPeriodoAberto := FieldByName('IdPeriodo').AsInteger;
                  DataPeriodo := FieldByName('DataAbertura').AsDateTime;

                  if IntegradoVendaCaixa then
                  begin
                     IdLancamento := NovoIdLancamento;

                     QueryFechamento.SQL.Clear;
                     QueryFechamento.SQL.Add('INSERT INTO TCxaLancamento ( ' +
                                             'Empresa,                     ' +
                                             'Caixa,                       ' +
                                             'Data,                        ' +
                                             'Numero,                      ' +
                                             'Historico,                   ' +
                                             'Complemento,                 ' +
                                             'ValorDh,                     ' +
                                             'ValorCh,                     ' +
                                             'Registradora,                ' +
                                             'IdPeriodo,                   ' +
                                             'Origem,                      ' +
                                             'Usuario)                     ' +
                                             'VALUES(:Empresa,             ' +
                                             '       :Caixa,               ' +
                                             '       ''TODAY'',            ' +
                                             '       :IdLancamento,        ' +
                                             '       :Historico,           ' +
                                             '       :Complemento,         ' +
                                             '       :ValorDh,             ' +
                                             '       :ValorCh,             ' +
                                             '       :Registradora,        ' +
                                             '       :IdPeriodo,           ' +
                                             '       :Origem,              ' +
                                             '       :Usuario)             ' );

                     QueryFechamento.ParamByName('Empresa').asString := Empresa;
                     QueryFechamento.ParamByName('Caixa').asString := CaixaPadrao;
                     QueryFechamento.ParamByName('Historico').asString := HistoricoEntradaReg;
                     QueryFechamento.ParamByName('Complemento').asString := Copy(Copy(RegistradoraAgrupamentoSemFinanceiro, 1, 2) + '-' + Copy(DescricaoRegistradoraAgrupamentoSemFinanceiro, 1, 20) + ' - ' + UsuarioPadraoRegistradoraAgrupamentoSemFinanceiro + ' ' + DateTimeToStr(DataPeriodo), 1, 50);
                     QueryFechamento.ParamByName('ValorDh').AsCurrency := 0;
                     QueryFechamento.ParamByName('ValorCh').AsCurrency := 0;
                     QueryFechamento.ParamByName('IdLancamento').asInteger := IdLancamento;
                     QueryFechamento.ParamByName('Registradora').AsString := RegistradoraAgrupamentoSemFinanceiro;
                     QueryFechamento.ParamByName('IdPeriodo').AsInteger := IdPeriodoAberto;
                     QueryFechamento.ParamByName('Origem').asString := 'VEN';
                     QueryFechamento.ParamByName('Usuario').asString := Usuario;

                     EcoExecSql(QueryFechamento);

                     QueryFechamento.SQL.Clear;
                     QueryFechamento.SQL.Add('UPDATE TGERCCLANCAMENTO ' +
                                             '   SET CAIXACONTA = :CAIXA, ' +
                                             '       NUMERO = :NUMERO, ' +
                                             '       LIBERADO = ''S'', ' +
                                             '       DATALIBERACAO = ''TODAY'' ' +
                                             ' WHERE EMPRESA = :EMPRESA ' +
                                             '   AND REGISTRADORA = :REGISTRADORA ' +
                                             '   AND IDPERIODO = :IDPERIODO');

                     QueryFechamento.ParamByName('Empresa').asString := Empresa;
                     QueryFechamento.ParamByName('CAIXA').AsString := CaixaPadrao;
                     QueryFechamento.ParamByName('Numero').asInteger := IdLancamento;
                     QueryFechamento.ParamByName('Registradora').AsString := RegistradoraAgrupamentoSemFinanceiro;
                     QueryFechamento.ParamByName('IdPeriodo').AsInteger := IdPeriodoAberto;

                     EcoExecSql(QueryFechamento);
                  end;

                  QueryFechamento.SQL.Clear;
                  QueryFechamento.SQL.Add('UPDATE TVenPeriodo ' +
                                          '   SET SaldoDHInformado = 0, ' +
                                          '       SaldoCHInformado = 0, ' +
                                          '       SaldoDHApurado = 0, ' +
                                          '       SaldoCHApurado = 0, ' +
                                          '       CrediarioApurado = 0, ' +
                                          '       CrediarioInformado = 0, ' +
                                          '       CartaoApurado = 0, ' +
                                          '       CartaoInformado = 0, ' +
                                          '       TicketApurado = 0, ' +
                                          '       TicketInformado = 0, ' +
                                          '       ValeTrocoApurado = 0, ' +
                                          '       ValeTrocoInformado = 0, ' +
                                          '       RequisicaoApurado = 0, ' +
                                          '       RequisicaoInformado = 0,' +
                                          '       CartaFreteInformado = 0, ' +
                                          '       CartaFreteApurado = 0, ' +
                                          '       DataFechamento = ''TODAY'', ' +
                                          '       HoraFechamento = ''NOW'', ' +
                                          '       Versao = 1, ' +
                                          '       UsuarioFechamento = :Usuario ' +
                                          ' WHERE TVenPeriodo.Empresa = :Empresa ' +
                                          '   AND TVenPeriodo.Registradora = :Registradora ' +
                                          '   AND TVenPeriodo.IdPeriodo = :IdPeriodo');

                  QueryFechamento.ParamByName('Empresa').asString := Empresa;
                  QueryFechamento.ParamByName('Registradora').asString := RegistradoraAgrupamentoSemFinanceiro;
                  QueryFechamento.ParamByName('IdPeriodo').asInteger := IdPeriodoAberto;
                  QueryFechamento.ParamByName('Usuario').asString := Usuario;

                  EcoExecSql(QueryFechamento);

                  Next;
               end;
            end;
         end;
      finally
         FreeAndNil(QueryAux);
         FreeAndNil(QueryFechamento);
         FreeAndNil(QueryPesquisa);
      end;
   end;
begin
   Registradora := RegistradoraTrabalho;
   Periodo := PeriodoTrabalho;
   NovoPeriodoAberto := False;

   if (Trim(RegistradoraAgrupamentoSemFinanceiro) = '') then
      Exit;

   QueryPesquisa := TFDQuery.Create(nil);
   try
      QueryPesquisa.Connection := FDataModule.FDConnection;

      with QueryPesquisa do
      begin
         SQL.Clear;
         SQL.Add('SELECT Per.IdPeriodo ' +
                 '  FROM TVenPeriodo Per' +
                 '       INNER JOIN TVenRegistradora Reg ON (Reg.Empresa = Per.Empresa ' +
                 '                                      AND  Reg.Numero = Per.Registradora) ' +
                 ' WHERE Per.Empresa = :Empresa ' +
                 '   AND Per.Registradora = :Registradora ' +
                 '   AND Per.DataAbertura = ''TODAY'' ' +
                 '   AND Per.DataFechamento IS NULL ' +
                 '   AND Per.ControlePAF IS NULL ' +
                 '   AND Reg.AgrupamentoSemFinanceiro = ''S'' ' +
                 'ORDER BY Per.Empresa');

         ParamByName('Empresa').asString := Empresa;
         ParamByName('Registradora').AsString := RegistradoraAgrupamentoSemFinanceiro;

         if EcoOpen(QueryPesquisa) then
            PeriodoAgrupamentoSemFinanceiro := FieldByName('IdPeriodo').AsInteger
         else
         begin
            SQL.Clear;
            SQL.Add('SELECT IdPeriodo, UsuarioPadrao ' +
                    '  FROM TVenRegistradora ' +
                    ' WHERE Empresa = :Empresa ' +
                    '   AND Numero = :Registradora');

            ParamByName('Empresa').AsString := Empresa;
            ParamByName('Registradora').AsString := RegistradoraAgrupamentoSemFinanceiro;

            if EcoOpen(QueryPesquisa) then
            begin
               IdPeriodo := FieldByName('IdPeriodo').AsInteger;
               UsuarioPadraoPeriodo := FieldByName('UsuarioPadrao').AsString;

               NovoIdPeriodo := IdPeriodo + 1;

               PeriodoTransaction := FDataModule.sConnection.BeginTransaction;
               try
                  SQL.Clear;
                  SQL.Add('INSERT INTO TVenPeriodo (Empresa, ' +
                          '                         Registradora, ' +
                          '                         IdPeriodo, ' +
                          '                         DataAbertura, ' +
                          '                         HoraAbertura, ' +
                          '                         UsuarioPeriodo, ' +
                          '                         UsuarioAbertura) ' +
                          '     VALUES(:Empresa, ' +
                          '            :Registradora, ' +
                          '            :IdPeriodo, ' +
                          '            ''TODAY'', ' +
                          '            ''NOW'', ' +
                          '            :UsuarioPeriodo, ' +
                          '            :UsuarioAbertura)');

                  ParamByName('Empresa').asString := Empresa;
                  ParamByName('Registradora').asString := RegistradoraAgrupamentoSemFinanceiro;
                  ParamByName('IdPeriodo').asInteger := NovoIdPeriodo;
                  ParamByName('UsuarioPeriodo').asString := UsuarioPadraoPeriodo;
                  ParamByName('UsuarioAbertura').asString := Usuario;
                  EcoExecSql(QueryPesquisa);

                  SQL.Clear;
                  SQL.Add('UPDATE TVenRegistradora ' +
                          '   SET IdPeriodo = :IdPeriodo ' +
                          ' WHERE Empresa = :Empresa ' +
                          '   AND Numero = :Numero');

                  ParamByName('Empresa').asString := Empresa;
                  ParamByName('Numero').asString := RegistradoraAgrupamentoSemFinanceiro;
                  ParamByName('IdPeriodo').asInteger := NovoIdPeriodo;
                  EcoExecSql(QueryPesquisa);

                  FDataModule.sConnection.CommitFreeAndNil(PeriodoTransaction);

                  PeriodoAgrupamentoSemFinanceiro := NovoIdPeriodo;

                  NovoPeriodoAberto := True;
               except
                  FDataModule.sConnection.RollbackFreeAndNil(PeriodoTransaction);
                  PeriodoAgrupamentoSemFinanceiro := 0;
                  NovoPeriodoAberto := False;
               end;
            end;
         end;
      end;

      if (Trim(RegistradoraAgrupamentoSemFinanceiro) <> '') and (PeriodoAgrupamentoSemFinanceiro > 0) then
      begin
         Registradora := RegistradoraAgrupamentoSemFinanceiro;
         Periodo := PeriodoAgrupamentoSemFinanceiro;
      end
      else
         NovoPeriodoAberto := False;
   finally
      QueryPesquisa.Close;
      FreeAndNil(QueryPesquisa);
   end;

   if NovoPeriodoAberto then
      FecharPeriodoAgrupamentoSemFinanceiro;
end;

function GetSQLCancelamentoConcluido: string;
const
   CSQL = 'SELECT Con.IdOrigem ' +
          '  FROM TGerItensConcluidos Con ' +
          ' WHERE Con.Empresa = :Empresa ' +
          '   AND Con.Origem = :Origem ' +
          '   AND Con.IdOrigem = :IdOrigem ' +
          '   AND Con.TipoItem in (2, 3, 103, 104, 202)';
begin
   Result := CSQL;
end;

function GetSQLItemPendenteExcluir: string;
const
   CSQL = 'SELECT Pen.IdOrigem ' +
          '  FROM TGerItensPendentes Pen ' +
          ' WHERE Pen.Empresa = :Empresa ' +
          '   AND Pen.Origem = :Origem ' +
          '   AND Pen.IdOrigem = :IdOrigem ' +
          '   AND (((Pen.Status > 0) and (Pen.TipoItem in (2, 3, 103, 104, 202))) ' +
          '        OR (Pen.Status IN (201, 202, 203, 204, 205, 217, 218, 301, 302)))';
begin
   Result := CSQL;
end;

function GetSQLItemPendenteDenegado: string;
const
   CSQL = 'SELECT Pen.IdOrigem ' +
          '  FROM TGerItensPendentes Pen ' +
          ' WHERE Pen.Empresa = :Empresa ' +
          '   AND Pen.Origem = :Origem ' +
          '   AND Pen.IdOrigem = :IdOrigem ' +
          '   AND Pen.Status IN (110, 301, 302, 303)';
begin
   Result := CSQL;
end;

function ExistePedidoCancelamentoNotaFiscalPendente(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TFDQuery;
begin
  QueryPesquisa := TFDQuery.Create(nil);
  try
    QueryPesquisa.Connection := FDataModule.FDConnection;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT IdOrigem ' +
                          '  FROM TGerItensPendentes ' +
                          ' WHERE Empresa = :Empresa ' +
                          '   AND Origem = :Origem ' +
                          '   AND IdOrigem = :IdOrigem ' +
                          '   AND TipoItem IN (2, 3, 103, 104)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExistePedidoCancelamentoNotaFiscalPendente(const ARegistradora: string; const AIdPeriodo: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT Reg.IdRegistro ' +
                          '  FROM TVenRegistro Reg ' +
                          '       INNER JOIN TGerItensPendentes Gip ON (Gip.Empresa = Reg.Empresa ' +
                          '                                        AND  Gip.IdOrigem = CAST(Reg.IdPedido AS INTEGER)) ' +
                          ' WHERE Reg.Empresa = :Empresa ' +
                          '   AND Reg.Registradora = :Registradora ' +
                          '   AND Reg.Status = ''V'' ' +
                          '   AND Reg.IdPeriodo = :IdPeriodo ' +
                          '   AND Gip.Origem = 0 ' +
                          '   AND Gip.TipoItem IN (2, 3, 103, 104)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Registradora').AsString := ARegistradora;
    QueryPesquisa.ParamByName('IdPeriodo').AsInteger := AIdPeriodo;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExistePedidoNotaFiscalPendente(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT IdOrigem ' +
                          '  FROM TGerItensPendentes ' +
                          ' WHERE Empresa = :Empresa ' +
                          '   AND Origem = :Origem ' +
                          '   AND IdOrigem = :IdOrigem ' +
                          '   AND TipoItem IN (1, 101)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExistePedidoNotaFiscalPendente(const ARegistradora: string; const AIdPeriodo: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT Reg.IdRegistro ' +
                          '  FROM TVenRegistro Reg ' +
                          '       INNER JOIN TGerItensPendentes Gip ON (Gip.Empresa = Reg.Empresa ' +
                          '                                        AND  Gip.IdOrigem = CAST(Reg.IdPedido AS INTEGER)) ' +
                          ' WHERE Reg.Empresa = :Empresa ' +
                          '   AND Reg.Registradora = :Registradora ' +
                          '   AND Reg.Status = ''V'' ' +
                          '   AND Reg.IdPeriodo = :IdPeriodo ' +
                          '   AND Gip.Origem = 0 ' +
                          '   AND Gip.TipoItem IN (1, 101)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Registradora').AsString := ARegistradora;
    QueryPesquisa.ParamByName('IdPeriodo').AsInteger := AIdPeriodo;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteCancelamentoRegistroPendente(const AIdOrigem: Integer): Boolean;
const
  CSql = 'SELECT PED.CODIGO                                                                             ' +
         '   FROM TVENPEDIDO PED                                                                        ' +
         'INNER JOIN TVENREGISTRO REG                                                                   ' +
         '   ON (PED.EMPRESA = REG.EMPRESA                                                              ' +
         '  AND PED.CODIGO = REG.IDPEDIDO)                                                              ' +
         'INNER JOIN TGERITENSCONCLUIDOS CON                                                            ' +
         '   ON (PED.EMPRESA = CON.EMPRESA                                                              ' +
         '  AND PED.CODIGO =  reverse(substring(reverse(''0000000''|| con.idorigem) from 1 for 07)))    ' +
         'WHERE PED.EMPRESA = :Empresa                                                                  ' +
         '  AND CON.IDORIGEM = :IdOrigem                                                                ' +
         '  AND PED.NFCANCELADAAVULSA IS DISTINCT FROM ''S''                                            ' +
         '  AND REG.STATUS = ''V''                                                                      ' +
         '  AND CON.TIPOITEM IN (2, 3, 103, 104)                                                        ' ;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;
    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add(CSql);
    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;
    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteCancelamentoRegistroPendente(const ARegistradora: string; const AIdPeriodo: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT Con.IdOrigem ' +
                          '  FROM TGerItensConcluidos Con ' +
                          '       LEFT OUTER JOIN TVenRegistro Reg ON (Reg.empresa = Con.empresa ' +
                          '                                       AND  CAST(Reg.idpedido AS INTEGER) = Con.idorigem) ' +
                          '       INNER JOIN Tvenpedido Ped on (Ped.empresa = Reg.empresa ' +
                          '                                 and ped.codigo = Reg.idpedido) ' +
                          ' WHERE Con.Empresa = :Empresa ' +
                          '   AND Con.Origem = 0 ' +
                          '   AND Con.TipoItem in (2, 3, 103, 104) ' +
                          '   AND Reg.Registradora = :Registradora ' +
                          '   AND Reg.IdPeriodo = :IdPeriodo ' +
                          '   AND Reg.Status = ''V'' ' +
                          '   AND Ped.NFCanceladaAvulsa = ''N'' ' +
                          '   AND ped.TipoVenda <> ''DV'' ');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Registradora').AsString := ARegistradora;
    QueryPesquisa.ParamByName('IdPeriodo').AsInteger := AIdPeriodo;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteCancelamentoConcluido(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add(GetSQLCancelamentoConcluido);

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteItemPendenteExcluir(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add(GetSQLItemPendenteExcluir);

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteEmailPendente(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT Pen.IdOrigem ' +
                          '  FROM TGerItensPendentes Pen ' +
                          ' WHERE Pen.Empresa = :Empresa ' +
                          '   AND Pen.Origem = :Origem ' +
                          '   AND Pen.IdOrigem = :IdOrigem ' +
                          '   AND Pen.TipoItem in (1019, 3011)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteReimpressaoPendente(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT Pen.IdOrigem ' +
                          '  FROM TGerItensPendentes Pen ' +
                          ' WHERE Pen.Empresa = :Empresa ' +
                          '   AND Pen.Origem = :Origem ' +
                          '   AND Pen.IdOrigem = :IdOrigem ' +
                          '   AND Pen.TipoItem in (1037, 3018, 2037)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteItemPendenteDenegado(const AOrigem, AIdOrigem: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add(GetSQLItemPendenteDenegado);

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Origem').AsInteger := AOrigem;
    QueryPesquisa.ParamByName('IdOrigem').AsInteger := AIdOrigem;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function ExisteItemPendenteDenegado(const ARegistradora: string; const AIdPeriodo: Integer): Boolean;
var
  QueryPesquisa: TSQLQuery;
begin
  QueryPesquisa := TSQLQuery.Create(nil);
  try
    QueryPesquisa.SQLConnection := FDataModule.sConnection;
    QueryPesquisa.MaxBlobSize := -1;

    QueryPesquisa.SQL.Clear;
    QueryPesquisa.SQL.Add('SELECT Reg.IdRegistro ' +
                          '  FROM TVenRegistro Reg ' +
                          '       INNER JOIN TGerItensPendentes Gip ON (Gip.Empresa = Reg.Empresa ' +
                          '                                        AND  Gip.IdOrigem = CAST(Reg.IdPedido AS INTEGER)) ' +
                          ' WHERE Reg.Empresa = :Empresa ' +
                          '   AND Reg.Registradora = :Registradora ' +
                          '   AND Reg.Status = ''V'' ' +
                          '   AND Reg.IdPeriodo = :IdPeriodo ' +
                          '   AND Gip.Origem = 0 ' +
                          '   AND Gip.Status IN (110, 301, 302, 303)');

    QueryPesquisa.ParamByName('Empresa').AsString := Empresa;
    QueryPesquisa.ParamByName('Registradora').AsString := ARegistradora;
    QueryPesquisa.ParamByName('IdPeriodo').AsInteger := AIdPeriodo;

    Result := EcoOpen(QueryPesquisa);
  finally
    QueryPesquisa.Close;
    FreeAndNil(QueryPesquisa);
  end;
end;

function CodigoCombustivelExigeTransportador(const ACodigoCombustivel: string): Boolean;
const
   PartCodigoANP: array[0..8] of string = ('21', '22', '32', '41', '42', '51', '54', '81', '82');
var
   I: Integer;
   Value: string;
   Comparar: string;
begin
   Result := False;

   if (Trim(ACodigoCombustivel) = '') then
      Exit;

   Comparar := Trim(Copy(ACodigoCombustivel, 1, 2));

   for I := Low(PartCodigoANP) to High(PartCodigoANP) do
   begin
      Value := PartCodigoANP[I];

      Result := (Comparar = Value);
      if Result then
         Break;
   end;
end;

function GetQuery: TSQLQuery;
begin
  Result := TSQLQuery.Create(nil);
  Result.SQLConnection := FDataModule.sConnection;
end;

function GetFDQuery: TFDQuery;
begin
  Result := TFDQuery.Create(nil);
  Result.Connection := FDataModule.FDConnection;
end;

function DataAtual: TDateTime;
var
  vFDQuery: TFDQuery;
begin
  vFDQuery := TFDQuery.Create(nil);
  try
    vFDQuery.Connection := FDataModule.FDConnection;
    vFDQuery.SQL.Text := 'SELECT CURRENT_DATE AS DATASERVIDOR FROM RDB$DATABASE ';
    vFDQuery.Open;
    Result := vFDQuery.FieldByName('DATASERVIDOR').AsDateTime;
  finally
     vFDQuery.Close;
     FreeAndNil(vFDQuery);
  end;
end;

function DataHoraAtual: TDateTime;
var
  vFDQuery: TFDQuery;
begin
  vFDQuery := TFDQuery.Create(nil);
  try
    vFDQuery.Connection := FDataModule.FDConnection;
    vFDQuery.SQL.Text := 'SELECT CURRENT_TIMESTAMP AS DATAHORASERVIDOR FROM RDB$DATABASE ';
    vFDQuery.Open;
    Result := vFDQuery.FieldByName('DATAHORASERVIDOR').AsDateTime;
  finally
     vFDQuery.Close;
     FreeAndNil(vFDQuery);
  end;
end;

function IsUrl(S: string): Boolean;
//
// Testa se a String é uma url ou não
//
const
BADCHARS = ';*<>{}[]|\()^!';
var
p, x, c, count, i: Integer;
begin
   Result := False;
   if (Length(S) > 5) and (S[Length(S)] <> '.') and (Pos(S, '..') = 0) then
      begin
      for i := Length(BADCHARS) downto 1 do
          begin
          if Pos(BADCHARS[i], S) > 0 then
             begin
             exit;
             end;
          end;
      for i := 1 to Length(S) do
          begin
          if (Ord(S[i]) < 33) or (Ord(S[i]) > 126) then
             begin
             exit;
             end;
          end;
      if ((Pos('www.',LowerCase(S)) = 1) and (Pos('.', Copy(S, 5, Length(s))) > 0) and (Length(S) > 7)) or ((Pos('news:', LowerCase(S)) = 1) and (Length(S) > 7) and (Pos('.', Copy(S, 5, Length(S))) > 0)) then
          begin
          end
      else if ((Pos('mailto:', LowerCase(S)) = 1) and (Length(S) > 12) and (Pos('@', S) > 8) and (Pos('.', S) > 10) and (Pos('.', S) > (Pos('@', S) +1))) or ((Length(S) > 6) and (Pos('@', S) > 1) and (Pos('.', S) > 4) and (Pos('.', S) > (Pos('@', S) +1))) then
               begin
               Result := True;
               Exit;
               end
      else if ((Pos('http://', LowerCase(S)) = 1) and (Length(S) > 10) and (Pos('.', S) > 8)) or ((Pos('ftp://', LowerCase(S)) = 1) and (Length(S) > 9) and (Pos('.', S) > 7)) then
              begin
              Result := True;
              Exit;
              end
           else
              begin
              Result := True;
              end;
      for Count := 1 to 4 do
          begin
          p := Pos('.',S) - 1;
          if p < 0 then
             begin
             p := Length(S);
             end;
           Val(Copy(S, 1, p), x, c);
           if ((c <> 0) or (x < 0) or (x > 255) or (p>3)) then
              begin
              Result := False;
              Break;
              end;
           Delete(S, 1, p + 1);
           end;
      if (S <> '') then
          begin
          Result := False;
          end;
      end;
 end;

 function GetoFMemoIniFileGerConfiguracao: TMemInifile;
begin
  if not Assigned(FIniMemoFile) then
  begin
    FIniMemoFile := TMemInifile.Create(EmptyStr);
  end;
  Result := FIniMemoFile;
end;

function GetoFStreamConfiguracaoAdicional: TStrings;
begin
  if not Assigned(FStreamConfiguracaoAdicional) then
  begin
     FStreamConfiguracaoAdicional := TStringList.Create;
  end;

  Result := FStreamConfiguracaoAdicional;
end;

procedure GravarConfiguracaonoMemoFileConfigPostEco360(const aIpServidor, aPortaConexao, aUsuario,
                                             aSenha, aNome, aCaminhoPost, aLinkServido360: String);
begin
  {Solitiação: 134530 Data: 29/05/2017 09:26:31 Usuário: Carlos.Schroder
   Detalhes: Dados de Conexão da Base dos Relatórios em Postgres}
  GetoFMemoIniFileGerConfiguracao.EraseSection(cSecaoConectPost);
  GetoFMemoIniFileGerConfiguracao.EraseSection(cSecaoEco360);

  GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoConectPost, 'IP SERVIDOR', aIpServidor);
  GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoConectPost, 'PORTA SERVIDOR', aPortaConexao);
  GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoConectPost, 'USUARIO SERVIDOR', aUsuario);
  GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoConectPost, 'SENHA SERVIDOR', aSenha);
  GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoConectPost, 'NOME BANCO', aNome);
  GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoConectPost, 'CAMINHO BIN', aCaminhoPost);

  if aLinkServido360 <> EmptyStr then
  begin
    GetoFMemoIniFileGerConfiguracao.WriteString(cSecaoEco360, 'LINK SERVIDOR', aLinkServido360);
  end;
  GetoFStreamConfiguracaoAdicional.Clear;
end;

procedure GravarConfiguracoesAdicionais(const aIpServidor,
                                              aPortaConexao,
                                              aUsuario,
                                              aSenha,
                                              aNome,
                                              aCaminhoPost,
                                              aLinkServido360: String;
                                              const aIdTipoServidor: TTipoServidor;
                                              const aIdConfiguracao: Integer = 0);
const
  cSqlInserUpdateConfigAdicional =  Concat('UPDATE OR INSERT INTO TGERCONFIGURACAO (IDCONFIGURACAO, CONFIGURACAO, IDUSUARIOALTERACAO, DATAALTERACAO, IDTIPOSERVIDOR) ',
                                           'VALUES (:IDCONFIGURACAO, :CONFIGURACAO, :IDUSUARIOALTERACAO, :DATAALTERACAO, :IDTIPOSERVIDOR)                            ',
                                           'MATCHING (IDCONFIGURACAO)                                                                                                ');
var
  vFDQuery: TFDQuery;

begin
   vFDQuery := TFDQuery.Create(nil);
   try
     GravarConfiguracaonoMemoFileConfigPostEco360(aIpServidor, aPortaConexao, aUsuario, aSenha, aNome, aCaminhoPost, aLinkServido360);
     GetoFMemoIniFileGerConfiguracao.GetStrings(GetoFStreamConfiguracaoAdicional);

     vFDQuery.Connection := FDataModule.FDConnection;
     vFDQuery.SQL.Text := cSqlInserUpdateConfigAdicional;
     vFDQuery.Close;

     vFDQuery.ParamByName('IDUSUARIOALTERACAO').AsInteger := IfThen(IdUsuario = ZeroValue, FIdUsuarioConfiguracaoPost, IDUsuario);
     vFDQuery.ParamByName('DATAALTERACAO').AsDate := DataAtual;
     vFDQuery.ParamByName('CONFIGURACAO').AsString := GetoFStreamConfiguracaoAdicional.Text;

     if aIdTipoServidor = TTipoServidor.IdTipoServidorRelatorioLocal then
     begin
        if aIdConfiguracao > 0 then
           vFDQuery.ParamByName('IDCONFIGURACAO').AsInteger := Integer(aIdConfiguracao);

        vFDQuery.ParamByName('IDTIPOSERVIDOR').AsInteger := Integer(IdTipoServidorRelatorioLocal);
     end;

     {$IF (not DEFINED(DESENVOLVIMENTO)) AND (not DEFINED(DEPURACAO)) AND (not DEFINED(CERTIFICACAO))}
     if aIdTipoServidor = TTipoServidor.IdTipoServidorRelatorioCertificacao then
     begin
        if aIdConfiguracao > 0 then
           vFDQuery.ParamByName('IDCONFIGURACAO').AsInteger := Integer(aIdConfiguracao);

        vFDQuery.ParamByName('IDTIPOSERVIDOR').AsInteger := Integer(IdTipoServidorRelatorioCertificacao);
     end;
     {$endif}

     vFDQuery.ExecSQL;
   finally
     vFDQuery.Close;
     FreeAndNil(vFDQuery);
     FreeAndNil(FIniMemoFile);
   end;
end;

procedure CarregarConfiguracadoMemoFile(var aIpServidor, aPortaConexao, aUsuario, aSenha, aNomeBanco, aCaminhoPasta, aLinkServidor360: String);
begin
  aIpServidor := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoConectPost, 'IP SERVIDOR', EmptyStr);
  aPortaConexao := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoConectPost, 'PORTA SERVIDOR', EmptyStr);
  aUsuario := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoConectPost, 'USUARIO SERVIDOR', EmptyStr);
  aSenha := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoConectPost, 'SENHA SERVIDOR', EmptyStr);
  aNomeBanco := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoConectPost, 'NOME BANCO', EmptyStr);
  aCaminhoPasta := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoConectPost, 'CAMINHO BIN', EmptyStr);

  aLinkServidor360 := GetoFMemoIniFileGerConfiguracao.ReadString(cSecaoEco360, 'LINK SERVIDOR', EmptyStr);
end;


procedure CarregarConfiguracoesAdicionais(var aIpServidor, aPortaConexao, aUsuario, aSenha, aNomeBanco, aCaminhoPasta, aLinkServidor360: String; var aIdConfiguracao: Integer);
var
  vFDQuery: TFDQuery;
  const
    cSqlCOnfiguracaoAdicional = Concat('SELECT IDCONFIGURACAO, CONFIGURACAO, IDUSUARIOALTERACAO, DATAALTERACAO, IDTIPOSERVIDOR ',
                                       '  FROM TGERCONFIGURACAO                                                ',
                                       ' WHERE IDTIPOSERVIDOR = :IDTIPOSERVIDOR                                ');
begin
  vFDQuery := TFDQuery.Create(Nil);
  try
    vFDQuery.Close;
    vFDQuery.Connection := FDataModule.FDConnection;
    vFDQuery.SQL.Text := cSqlConfiguracaoAdicional;
    vFDQuery.ParamByName('IDTIPOSERVIDOR').AsInteger := Integer(IdTipoServidorRelatorioLocal);
    vFDQuery.Open;
    if not vFDQuery.IsEmpty then
    begin
      GetoFStreamConfiguracaoAdicional.Text := vFDQuery.FieldByName('CONFIGURACAO').AsString;
      FIdUsuarioConfiguracaoPost := vFDQuery.FieldByName('IDUSUARIOALTERACAO').AsInteger;
      GetoFMemoIniFileGerConfiguracao.SetStrings(GetoFStreamConfiguracaoAdicional);
      aIdConfiguracao := vFDQuery.FieldByName('IdConfiguracao').AsInteger;
    end;
    CarregarConfiguracadoMemoFile(aIpServidor, aPortaConexao, aUsuario, aSenha, aNomeBanco, aCaminhoPasta, aLinkServidor360);
  finally
    vFDQuery.Close;
    FreeAndNil(vFDQuery);
    FreeAndNil(FIniMemoFile);
  end;
end;

function ValidaCompatibilidadeVersoes(const aValidarSomenteVersaoRelease: Boolean = True) : Boolean;
var
  Ax1, Ax2 :String;
  vFDQuery: TFDQuery;
begin
   vFDQuery := TFDQuery.Create(Nil);
   try
     vFDQuery.Close;
     vFDQuery.Connection := FDataModule.FDConnection;
     vFDQuery.SQL.Text := 'SELECT FIRST 1 A.VERSAO FROM TGERLICENCA A';
     vFDQuery.Open;

     fVersaoBancoInt := vFDQuery.FieldByName('VERSAO').AsInteger;

     fVersaoBanco := ErtStrUtils.MaskFormatText('9.9.99.99;0; ', IntToStr(fVersaoBancoInt));
   finally
      vFDQuery.Close;
      FreeAndNil(vFDQuery);
   end;

  {$IFDEF DESENVOLVIMENTO}
   if aValidarSomenteVersaoRelease then
   begin
     Result := True;
     Exit;
   end;
   {$ENDIF}
   {$IFDEF DEPURACAO}
   if aValidarSomenteVersaoRelease then
   begin
     Result := True;
     Exit;
   end;
   {$ENDIF}

   Ax1 := Copy(Utilitarios.GetFileVersion.Versao, 6, 6);
   Ax2 := Copy(fVersaoBanco, 1, 6);
   Result := Ax1 = Ax2;

   if not Result then
   begin
     Mensagem(Concat('A versão do sistema é incompatível com ',
                     'a versão do banco de dados.            ',  sLineBreak,
                     '- Versão do sistema: ',      ax1, sLineBreak,
                     '- Versão do banco de dados: ', ax2),
            tmErro);
   end;

end;

function TemEmpresaPrincipal: Boolean;
var
  vFDQuery: TFDQuery;
const
  cSqlIndPrincipal = Concat('SELECT e.codigo, e.idnprincipal FROM tgerempresa e             ',
                            'WHERE ((e.idnprincipal <> 0) and (e.idnprincipal is not null)) ');
begin
  Result := True;

   vFDQuery := TFDQuery.Create(Nil);
   try
      vFDQuery.Close;
      vFDQuery.Connection := FDataModule.FDConnection;
      vFDQuery.SQL.Text := cSqlIndPrincipal;
      vFDQuery.Open;

      if vFDQuery.IsEmpty then
      begin
         Result := False;
      end;
   finally
     vFDQuery.Close;
     FreeAndNil(vFDQuery);
   end;
end;

function ValidarLoginEmpresaPrincipal(const AValidar: Boolean): Boolean;
var
  vFDQuery: TFDQuery;
const
  cSqlEmpresa = Concat('SELECT e.codigo, e.idnprincipal FROM tgerempresa e          ',
                       'WHERE e.codigo = :codigo                                    ',
                       'AND ((e.idnprincipal <> 0) and (e.idnprincipal is not null))');
begin
  Result := True;

  if (AValidar) and (TemEmpresaPrincipal) then
  begin
      vFDQuery := TFDQuery.Create(Nil);
      try
        vFDQuery.Close;
        vFDQuery.Connection := FDataModule.FDConnection;
        vFDQuery.SQL.Text := cSqlEmpresa;
        vFDQuery.ParamByName('codigo').AsString := Empresa;
        vFDQuery.Open;

        if vFDQuery.IsEmpty then
        begin
           Result := False;
           Mensagem('Atenção a empresa logada não é uma empresa principal. Para utilizar a tela favor logar com a empresa principal.', tmInforma);
        end;
      finally
         vFDQuery.Close;
         FreeAndNil(vFDQuery);
      end;
  end;
end;

function PermitirUtilizarFuncao(const MostrarMensagem: Boolean = True): Boolean;
begin
   Result := True;
   
//   {$IF (NOT DEFINED(DESENVOLVIMENTO)) AND (NOT DEFINED(DEPURACAO)) AND (NOT DEFINED(CERTIFICACAO))}
//      Result := False;
//      if (MostrarMensagem) then
//      begin
//         Mensagem('Disponível apenas para desenvolvimento.', tmInforma);
//         abort;
//      end;
//   {$ENDIF}
end;

procedure BuscarAlmoxPadrao;
var
  FQuery: TFDQuery;
begin
  FQuery := TFDQuery.Create(nil);
  try
    FQuery.Connection := TControllerConexao.GetInstance.GetConexaoPadrao(TAplicacao.ECO);
    FQuery.Sql.Add('SELECT ALMOXPADRAO      ' +
                   '  FROM TESTPARAMETRO    ' +
                   ' WHERE                  ' +
                   '     Empresa = :Empresa ');
    FQuery.ParamByName('Empresa').asString := Empresa;
    FQuery.Open;
    if not FQuery.IsEmpty then
    begin
      sAlmoxPadrao := FQuery.FieldByName('ALMOXPADRAO').AsString;
    end
    else
    begin
      sAlmoxPadrao := AlmoxVenda;
    end;
  finally
    FQuery.Close;
    FreeAndNil(FQuery);
  end;
end;

function IsValidRegEx(const APattern, AText: String): Boolean;
var
   RE: TRegEx;
   M: TMatch;
begin
  RE.Create(APattern);
  M := RE.Match(AText);
  Result := (M.Success and (M.Length = Length(AText)));
end;

function VerificaLetrasNumero(Placa: String): Boolean;
   const cExRePlaca = '[A-Z]{2,3}[0-9]{4}|[A-Z]{3}[0-9]{1}[A-Z]{1}[0-9]{2}';
begin
   Result := IsValidRegEx(cExRePlaca, placa);
end;

function VerificacoesPlaca(Placa: string): Boolean;
begin
   Result := True;

   if Trim(Placa) = '' then
      Exit;

   if not VerificaLetrasNumero(Placa) then
   begin
      Mensagem('A placa deve conter até 7 caracteres, sendo no máximo 4 letras com no máximo 4 números!', tmInforma);
      Result := False;
   end;
end;

function ValidarCNPJ(numCNPJ: string): boolean;
var
  cnpj: string;
  dg1, dg2: integer;
  x, total: integer;
begin
   Result := False;
   cnpj := '';
   //Analisa os formatos
   if Length(numCNPJ) = 18 then
   	  if (Copy(numCNPJ,3,1) + Copy(numCNPJ,7,1) + Copy(numCNPJ,11,1) + Copy(numCNPJ,16,1) = '../-') then
   		begin
   		   cnpj := Copy(numCNPJ,1,2) + Copy(numCNPJ,4,3) + Copy(numCNPJ,8,3) + Copy(numCNPJ,12,4) + Copy(numCNPJ,17,2);
   		   Result := True;
   		end;
   if Length(numCNPJ) = 14 then
   begin
      cnpj := numCNPJ;
      Result := True;
   end;
      //Verifica
   if Result then
   begin
      try
      		//1° digito
      		total := 0;
      		for x := 1 to 12 do
      			begin
      			if x < 5 then
      				Inc(total, StrToInt(Copy(cnpj, x, 1)) * (6 - x))
      			else
      				Inc(total, StrToInt(Copy(cnpj, x, 1)) * (14 - x));
      			end;
      		dg1 := 11 - (total mod 11);
      		if dg1 > 9 then
      			dg1 := 0;
      		//2° digito
      		total := 0;
      		for x := 1 to 13 do
      			begin
      			if x < 6 then
      				Inc(total, StrToInt(Copy(cnpj, x, 1)) * (7 - x))
      			else
      				Inc(total, StrToInt(Copy(cnpj, x, 1)) * (15 - x));
      			end;
      		dg2 := 11 - (total mod 11);
      		if dg2 > 9 then
      			dg2 := 0;
      		//Validação final
      		if (dg1 = StrToInt(Copy(cnpj, 13, 1))) and (dg2 = StrToInt(Copy(cnpj, 14, 1))) then
      			Result := True
      		else
      			Result := False;
      except
      		Result := False;
      end;
      	//Inválidos
      case AnsiIndexStr(cnpj,['00000000000000','11111111111111','22222222222222','33333333333333','44444444444444',
      						            '55555555555555','66666666666666','77777777777777','88888888888888','99999999999999']) of

      		0..9:   Result := False;

      end;
   end;
end;

function SomenteNumeros(const S: String): String;
var
  I, L: Integer;
begin
  L := 0;
  for I := 1 to Length(S) do
    if ((S[I] >= '0') and (S[I] <= '9')) then Inc(L);
  if (L = Length(S)) then
    Result := S else
  begin
    SetLength(Result, L);
    if (L > 0) then
    begin
      L := 0;
      for I := 1 to Length(S) do
        if ((S[I] >= '0') and (S[I] <= '9')) then
        begin
          Inc(L);
          Result[L] := S[I];
        end;
    end;
  end;
end;

procedure AtivarDesativarLogChangesDataSet(ACDS: TClientDataSet; const AAtivar: Boolean = True);
begin
   if ACDS.Active then
   begin
      ACDS.LogChanges := AAtivar;
   end;

   if AAtivar then
   begin
      ACDS.EnableControls;
   end
   else
   begin
      ACDS.DisableControls;
   end;
end;


initialization

finalization
   FreeAndNil(FStreamConfiguracaoAdicional);
end.

