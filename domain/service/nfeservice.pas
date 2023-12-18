unit nfeService;

interface

uses
  SysUtils,
  fpjson,
  ACBrNFe,
  pcnNFe,
  pcnSignature,
  pcnProcNFe,
  pcnConversao,
  pcnConversaoNFe,
  pcnGerador,
  mormot.core.variants;

type

  { TNfeService }

  TNfeService = class
  private
    function InfNFeToVariant(InfNFe: TInfNFe): Variant;
    function IdeToVariant(Ide: TIde): Variant;
    function EmitToVariant(Emit: TEmit): Variant;
    function AvulsaToVariant(Avulsa: TAvulsa): Variant;
    function DestToVariant(Dest: TDest): Variant;
    function RetiradaToVariant(Retirada: TRetirada): Variant;
    function EntregaToVariant(Entrega: TEntrega): Variant;
    function DetToVariant(Det: TDetCollection): Variant;
    function TotalToVariant(Total: TTotal): Variant;
    function TranspToVariant(Transp: TTransp): Variant;
    function CobrToVariant(Cobr: TCobr): Variant;
    function PagToVariant(Pag: TpagCollection): Variant;
    function InfIntermedToVariant(InfIntermed: TInfIntermed): Variant;
    function InfAdicToVariant(InfAdic: TInfAdic): Variant;
    function ExportaToVariant(Exporta: TExporta): Variant;
    function CompraToVariant(Compra: TCompra): Variant;
    function CanaToVariant(Cana: TCana): Variant;
    function SignatureToVariant(Signature: TSignature): Variant;
    function ProcNFeToVariant(ProcNFe: TProcNFe): Variant;
    function AutXMLToVariant(AutXML: TAutXMLCollection): Variant;
    function InfNFeSuplToVariant(InfNFeSupl: TInfNFeSupl): Variant;
    function InfRespTecToVariant(InfRespTec: TInfRespTec): Variant;
    function NFrefToVariant(NFref: TNFrefCollection): Variant;
    function NFrefItensToVariant(NFrefItens: TNFrefCollectionItem): Variant;
    function RefNFToVariant(RefNF: TRefNF): Variant;
    function RefECFToVariant(RefECF: TRefECF): Variant;
    function RefNFPToVariant(RefNFP: TRefNFP): Variant;
    function EnderEmitToVariant(EnderEmit: TEnderEmit): Variant;
    function EnderDestToVariant(EnderDest: TEnderDest): Variant;
    function DetItensToVariant(DetItens: TDetCollectionItem): Variant;
    function ProdToVariant(Prod: TProd): Variant;
    function ImpostoToVariant(Imposto: TImposto): Variant;
    function ObsItemToVariant(ObsItem: TObsItem): Variant;
    function ICMSToVariant(ICMS: TICMS): Variant;
    function IPIToVariant(IPI: TIPI): Variant;
    function IIToVariant(II: TII): Variant;
    function PISToVariant(PIS: TPIS): Variant;
    function PISSTToVariant(PISST: TPISST): Variant;
    function COFINSToVariant(COFINS: TCOFINS): Variant;
    function COFINSSTToVariant(COFINSST: TCOFINSST): Variant;
    function ISSQNToVariant(ISSQN: TISSQN): Variant;
    function ICMSUFDestToVariant(ICMSUFDest: TICMSUFDest): Variant;
    function DIToVariant(DI: TDICollection): Variant;
    function DIItensToVariant(DIItens: TDICollectionItem): Variant;
    function AdiToVariant(Adi: TAdiCollection): Variant;
    function AdiItensToVariant(AdiItens: TAdiCollectionItem): Variant;
    function DetExportToVariant(DetExport: TdetExportCollection): Variant;
    function DetExportItensToVariant(DetExportItens: TdetExportCollectionItem): Variant;
    function RastroToVariant(Rastro: TRastroCollection): Variant;
    function RastroItensToVariant(RastroItens: TRastroCollectionItem): Variant;
    function VeicProdToVariant(VeicProd: TVeicProd): Variant;
    function MedToVariant(Med: TMedCollection): Variant;
    function MedItensToVariant(MedItens: TMedCollectionItem): Variant;
    function ArmaToVariant(Arma: TArmaCollection): Variant;
    function ArmaItensToVariant(ArmaItens: TArmaCollectionItem): Variant;
    function CombToVariant(Comb: TComb): Variant;
    function CIDEToVariant(CIDE: TCIDE): Variant;
    function ICMSCombToVariant(ICMSComb: TICMSComb): Variant;
    function ICMSInterToVariant(ICMSInter: TICMSInter): Variant;
    function ICMSConsToVariant(ICMSCons: TICMSCons): Variant;
    function EncerranteToVariant(Encerrante: TEncerrante): Variant;
    function OrigCombToVariant(OrigComb: TorigCombCollection): Variant;
    function OrigCombItensToVariant(OrigCombItens: TorigCombCollectionItem): Variant;
    function NVEToVariant(NVE: TNVECollection): Variant;
    function NVEItensToVariant(NVEItens: TNVECollectionItem): Variant;
    function ICMSTotToVariant(ICMSTot: TICMSTot): Variant;
    function ISSQNtotToVariant(ISSQNtot: TISSQNtot): Variant;
    function RetTribToVariant(RetTrib: TretTrib): Variant;
    function TransportaToVariant(Transporta: TTransporta): Variant;
    function RetTranspToVariant(RetTransp: TRetTransp): Variant;
    function VeicTranspToVariant(VeicTransp: TVeicTransp): Variant;
    function VolToVariant(Vol: TVolCollection): Variant;
    function VolItensToVariant(VolItens: TVolCollectionItem): Variant;
    function LacresToVariant(Lacres: TLacresCollection): Variant;
    function LacresItensToVariant(LacresItens: TLacresCollectionItem): Variant;
    function ReboqueToVariant(Reboque: TReboqueCollection): Variant;
    function ReboqueItensToVariant(ReboqueItens: TReboqueCollectionItem): Variant;
    function FatToVariant(Fat: TFat): Variant;
    function DupToVariant(Dup: TDupCollection): Variant;
    function DupItensToVariant(DupItens: TDupCollectionItem): Variant;
    function PagItensToVariant(PagItens: TpagCollectionItem): Variant;
    function ObsContToVariant(ObsCont: TobsContCollection): Variant;
    function ObsContItensToVariant(ObsContItens: TobsContCollectionItem): Variant;
    function ObsFiscoToVariant(ObsFisco: TobsFiscoCollection): Variant;
    function ObsFiscoItensToVariant(ObsFiscoItens: TobsFiscoCollectionItem): Variant;
    function ProcRefToVariant(ProcRef: TprocRefCollection): Variant;
    function ProcRefItensToVariant(ProcRefItens: TprocRefCollectionItem): Variant;
    function FordiaToVariant(Fordia: TForDiaCollection): Variant;
    function FordiaItensToVariant(FordiaItens: TForDiaCollectionItem): Variant;
    function DeducToVariant(Deduc: TDeducCollection): Variant;
    function DeducItensToVariant(DeducItens: TDeducCollectionItem): Variant;
    function GeradorToVariant(Gerador: TGerador): Variant;
    function GeradorOpcoesToVariant(GeradorOpcoes: TGeradorOpcoes): Variant;
    function AutXMLItensToVariant(AutXMLItens: TautXMLCollectionItem): Variant;
  public
    function readXML(value: string): TNFe;
    function NFeToVariant(NFe: TNFe): Variant;
  end;

implementation

{ TNfeService }

function TNfeService.readXML(value: string): TNFe;
var
  NFe: TACBrNFe;
  json: Variant;
  path: string;
begin
  json := TDocVariant.NewJson(value);
  path := TDocVariantData(json).U['path'];
  NFe := TACBrNFe.Create(Nil);
  NFe.NotasFiscais.LoadFromFile(path);
  Result := NFe.NotasFiscais.Items[0].NFe;
end;

function TNfeService.NFeToVariant(NFe: TNFe): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('InfNFe', InfNFeToVariant(NFe.infNFe));
    AddValue('Ide', IdeToVariant(NFe.Ide));
    AddValue('Emit', EmitToVariant(NFe.Emit));
    AddValue('Avulsa', AvulsaToVariant(Nfe.Avulsa));
    AddValue('Dest', DestToVariant(NFe.Dest));
    AddValue('Retirada', RetiradaToVariant(Nfe.Retirada));
    AddValue('Entrega', EntregaToVariant(NFe.Entrega));
    AddValue('Det', DetToVariant(NFe.Det));
    AddValue('Total', TotalToVariant(NFe.Total));
    AddValue('Transp', TranspToVariant(NFe.Transp));
    AddValue('Cobr', CobrToVariant(NFe.Cobr));
    AddValue('Pag', PagToVariant(NFe.Pag));
    AddValue('InfIntermed', InfIntermedToVariant(NFe.InfIntermed));
    AddValue('InfAdic', InfAdicToVariant(NFe.InfAdic));
    AddValue('Exporta', ExportaToVariant(NFe.Exporta));
    AddValue('Compra', CompraToVariant(NFe.Compra));
    AddValue('Cana', CanaToVariant(NFe.Cana));
    AddValue('Signature', SignatureToVariant(NFe.Signature));
    AddValue('ProcNFe', ProcNFeToVariant(NFe.ProcNFe));
    AddValue('AutXML', AutXMLToVariant(NFe.AutXML));
    AddValue('InfNFeSupl', InfNFeSuplToVariant(NFe.InfNFeSupl));
    AddValue('InfRespTec', InfRespTecToVariant(NFe.InfRespTec));
  end;
end;

function TNfeService.InfNFeToVariant(InfNFe: TInfNFe): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('ID', InfNFe.ID);
    AddValue('Versao', InfNFe.Versao);
  end;
end;

function TNfeService.IdeToVariant(Ide: TIde): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('cUF', Ide.cUF);
    AddValue('cNF', Ide.cNF);
    AddValue('natOp', Ide.natOp);
    AddValue('indPag', IndpagToStr(Ide.indPag));
    AddValue('modelo', Ide.modelo);
    AddValue('serie', Ide.serie);
    AddValue('nNF', Ide.nNF);
    AddValue('dEmi', Ide.dEmi);
    AddValue('dSaiEnt', Ide.dSaiEnt);
    AddValue('hSaiEnt', Ide.hSaiEnt);
    AddValue('tpNF', tpNFToStr(Ide.tpNF));
    AddValue('idDest', DestinoOperacaoToStr(Ide.idDest));
    AddValue('cMunFG', Ide.cMunFG);
    AddValue('NFref', NFrefToVariant(Ide.NFref));
    AddValue('tpImp', TpImpToStr(Ide.tpImp));
    AddValue('tpEmis', TpEmisToStr(Ide.tpEmis));
    AddValue('cDV', Ide.cDV);
    AddValue('tpAmb', TpAmbToStr(Ide.tpAmb));
    AddValue('finNFe', FinNFeToStr(Ide.finNFe));
    AddValue('indFinal', ConsumidorFinalToStr(Ide.indFinal));
    AddValue('indPres', PresencaCompradorToStr(Ide.indPres));
    AddValue('indIntermed', IndIntermedToStr(Ide.indIntermed));
    AddValue('procEmi', procEmiToStr(Ide.procEmi));
    AddValue('verProc', Ide.verProc);
    AddValue('dhCont', Ide.dhCont);
    AddValue('xJust', Ide.xJust);
  end;
end;

function TNfeService.EmitToVariant(Emit: TEmit): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJCPF', Emit.CNPJCPF);
    AddValue('xNome', Emit.xNome);
    AddValue('xFant', Emit.xFant);
    AddValue('enderEmit', EnderEmitToVariant(Emit.enderEmit));
    AddValue('IE', Emit.IE);
    AddValue('IEST', Emit.IEST);
    AddValue('IM', Emit.IM);
    AddValue('CNAE', Emit.CNAE);
    AddValue('CRT', CRTToStr(Emit.CRT));
  end;
end;

function TNfeService.AvulsaToVariant(Avulsa: TAvulsa): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJ', Avulsa.CNPJ);
    AddValue('xOrgao', Avulsa.xOrgao);
    AddValue('matr', Avulsa.matr);
    AddValue('xAgente', Avulsa.xAgente);
    AddValue('fone', Avulsa.fone);
    AddValue('UF', Avulsa.UF);
    AddValue('nDAR', Avulsa.nDAR);
    AddValue('dEmi', Avulsa.dEmi);
    AddValue('vDAR', Avulsa.vDAR);
    AddValue('repEmi', Avulsa.repEmi);
    AddValue('dPag', Avulsa.dPag);
  end;
end;

function TNfeService.DestToVariant(Dest: TDest): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJCPF', Dest.CNPJCPF);
    AddValue('idEstrangeiro', Dest.idEstrangeiro);
    AddValue('xNome', Dest.xNome);
    AddValue('EnderDest', EnderDestToVariant(Dest.EnderDest));
    AddValue('indIEDest', indIEDestToStr(Dest.indIEDest));
    AddValue('IE', Dest.IE);
    AddValue('ISUF', Dest.ISUF);
    AddValue('IM', Dest.IM);
    AddValue('email', Dest.email);
  end;
end;

function TNfeService.RetiradaToVariant(Retirada: TRetirada): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJCPF', Retirada.CNPJCPF);
    AddValue('xNome', Retirada.xNome);
    AddValue('xLgr', Retirada.xLgr);
    AddValue('nro', Retirada.nro);
    AddValue('xCpl', Retirada.xCpl);
    AddValue('xBairro', Retirada.xBairro);
    AddValue('cMun', Retirada.cMun);
    AddValue('xMun', Retirada.xMun);
    AddValue('UF', Retirada.UF);
    AddValue('CEP', Retirada.CEP);
    AddValue('cPais', Retirada.cPais);
    AddValue('xPais', Retirada.xPais);
    AddValue('fone', Retirada.fone);
    AddValue('email', Retirada.email);
    AddValue('IE', Retirada.IE);
  end;
end;

function TNfeService.EntregaToVariant(Entrega: TEntrega): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJCPF', Entrega.CNPJCPF);
    AddValue('xNome', Entrega.xNome);
    AddValue('xLgr', Entrega.xLgr);
    AddValue('nro', Entrega.nro);
    AddValue('xCpl', Entrega.xCpl);
    AddValue('xBairro', Entrega.xBairro);
    AddValue('cMun', Entrega.cMun);
    AddValue('xMun', Entrega.xMun);
    AddValue('UF', Entrega.UF);
    AddValue('CEP', Entrega.CEP);
    AddValue('cPais', Entrega.cPais);
    AddValue('xPais', Entrega.xPais);
    AddValue('fone', Entrega.fone);
    AddValue('email', Entrega.email);
    AddValue('IE', Entrega.IE);
  end;
end;

function TNfeService.DetToVariant(Det: TDetCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Det.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(DetItensToVariant(Det.Items[i]));
  end;
end;

function TNfeService.TotalToVariant(Total: TTotal): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('ICMSTot', ICMSTotToVariant(Total.ICMSTot));
    AddValue('ISSQNtot', ISSQNtotToVariant(Total.ISSQNtot));
    AddValue('RetTrib', RetTribToVariant(Total.RetTrib));
  end;
end;

function TNfeService.TranspToVariant(Transp: TTransp): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('modFrete', modFreteToStr(Transp.modFrete));
    AddValue('Transporta', TransportaToVariant(Transp.Transporta));
    AddValue('retTransp', RetTranspToVariant(Transp.retTransp));
    AddValue('veicTransp', VeicTranspToVariant(Transp.veicTransp));
    AddValue('Vol', VolToVariant(Transp.Vol));
    AddValue('Reboque', ReboqueToVariant(Transp.Reboque));
    AddValue('vagao', Transp.vagao);
    AddValue('balsa', Transp.balsa);
  end;
end;

function TNfeService.CobrToVariant(Cobr: TCobr): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('Fat', FatToVariant(Cobr.Fat));
    AddValue('Dup', DupToVariant(Cobr.Dup));
  end;
end;

function TNfeService.PagToVariant(Pag: TpagCollection): Variant;
var
  i: integer;
  itens: variant;
begin
  TDocVariant.NewFast(Result);
  TDocVariant.NewFast(itens);
  TDocVariantData(Result).AddValue('vTroco', Pag.vTroco);
  for i := 0 to Pag.Count - 1 do
  begin
    TDocVariantData(itens).AddItem(PagItensToVariant(Pag.Items[i]));
  end;
  TDocVariantData(Result).AddValue('itens', itens);
end;

function TNfeService.InfIntermedToVariant(InfIntermed: TInfIntermed): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJ', InfIntermed.CNPJ);
    AddValue('idCadIntTran', InfIntermed.idCadIntTran);
  end;
end;

function TNfeService.InfAdicToVariant(InfAdic: TInfAdic): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('infAdFisco', InfAdic.infAdFisco);
    AddValue('infCpl', InfAdic.infCpl);
    AddValue('obsCont', ObsContToVariant(InfAdic.obsCont));
    AddValue('obsFisco', ObsFiscoToVariant(InfAdic.obsFisco));
    AddValue('procRef', ProcRefToVariant(InfAdic.procRef));
  end
end;

function TNfeService.ExportaToVariant(Exporta: TExporta): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('UFembarq', Exporta.UFembarq);
    AddValue('xLocEmbarq', Exporta.xLocEmbarq);
    AddValue('UFSaidaPais', Exporta.UFSaidaPais);
    AddValue('xLocExporta', Exporta.xLocExporta);
    AddValue('xLocDespacho', Exporta.xLocDespacho);
  end
end;

function TNfeService.CompraToVariant(Compra: TCompra): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xNEmp', Compra.xNEmp);
    AddValue('xPed', Compra.xPed);
    AddValue('xCont', Compra.xCont);
  end
end;

function TNfeService.CanaToVariant(Cana: TCana): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('safra', Cana.safra);
    AddValue('ref', Cana.ref);
    AddValue('fordia', FordiaToVariant(Cana.fordia));
    AddValue('qTotMes', Cana.qTotMes);
    AddValue('qTotAnt', Cana.qTotAnt);
    AddValue('qTotGer', Cana.qTotGer);
    AddValue('deduc', DeducToVariant(Cana.deduc));
    AddValue('vFor', Cana.vFor);
    AddValue('vTotDed', Cana.vTotDed);
    AddValue('vLiqFor', Cana.vLiqFor);
  end
end;

function TNfeService.SignatureToVariant(Signature: TSignature): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('Gerador', GeradorToVariant(Signature.Gerador));
    AddValue('URI', Signature.URI);
    AddValue('DigestValue', Signature.DigestValue);
    AddValue('SignatureValue', Signature.SignatureValue);
    AddValue('X509Certificate', Signature.X509Certificate);
  end
end;

function TNfeService.ProcNFeToVariant(ProcNFe: TProcNFe): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('Gerador', GeradorToVariant(ProcNFe.Gerador));
    AddValue('PathNFe', ProcNFe.PathNFe);
    AddValue('PathRetConsReciNFe', ProcNFe.PathRetConsReciNFe);
    AddValue('PathRetConsSitNFe', ProcNFe.PathRetConsSitNFe);
    AddValue('tpAmb', ProcNFe.tpAmb);
    AddValue('verAplic', ProcNFe.verAplic);
    AddValue('chNFe', ProcNFe.chNFe);
    AddValue('dhRecbto', ProcNFe.dhRecbto);
    AddValue('nProt', ProcNFe.nProt);
    AddValue('digVal', ProcNFe.digVal);
    AddValue('cStat', ProcNFe.cStat);
    AddValue('xMotivo', ProcNFe.xMotivo);
    AddValue('Versao', ProcNFe.Versao);
    AddValue('cMsg', ProcNFe.cMsg);
    AddValue('xMsg', ProcNFe.xMsg);
    AddValue('XML_NFe', ProcNFe.XML_NFe);
    AddValue('XML_prot', ProcNFe.XML_prot);
  end
end;

function TNfeService.AutXMLToVariant(AutXML: TAutXMLCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to AutXML.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(AutXMLItensToVariant(AutXML.Items[i]));
  end;
end;

function TNfeService.InfNFeSuplToVariant(InfNFeSupl: TInfNFeSupl): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('qrCode', InfNFeSupl.qrCode);
    AddValue('urlChave', InfNFeSupl.urlChave);
  end
end;

function TNfeService.InfRespTecToVariant(InfRespTec: TInfRespTec): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJ', InfRespTec.CNPJ);
    AddValue('xContato', InfRespTec.xContato);
    AddValue('email', InfRespTec.email);
    AddValue('fone', InfRespTec.fone);
    AddValue('idCSRT', InfRespTec.idCSRT);
    AddValue('hashCSRT', InfRespTec.hashCSRT);
  end
end;

function TNfeService.NFrefToVariant(NFref: TNFrefCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to NFref.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(NFrefItensToVariant(NFref.Items[i]));
  end;
end;

function TNfeService.NFrefItensToVariant(NFrefItens: TNFrefCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('refNFe', NFrefItens.refNFe);
    AddValue('refNFeSig', NFrefItens.refNFeSig);
    AddValue('refCTe', NFrefItens.refCTe);
    AddValue('RefNF', RefNFToVariant(NFrefItens.RefNF));
    AddValue('RefECF', RefECFToVariant(NFrefItens.RefECF));
    AddValue('RefNFP', RefNFPToVariant(NFrefItens.RefNFP));
  end;
end;

function TNfeService.RefNFToVariant(RefNF: TRefNF): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('cUF', RefNF.cUF);
    AddValue('AAMM', RefNF.AAMM);
    AddValue('CNPJ', RefNF.CNPJ);
    AddValue('modelo', RefNF.modelo);
    AddValue('serie', RefNF.serie);
    AddValue('nNF', RefNF.nNF);
  end;
end;

function TNfeService.RefECFToVariant(RefECF: TRefECF): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('modelo', ECFModRefToStr(RefECF.modelo));
    AddValue('nECF', RefECF.nECF);
    AddValue('nCOO', RefECF.nCOO);
  end;
end;

function TNfeService.RefNFPToVariant(RefNFP: TRefNFP): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('cUF', RefNFP.cUF);
    AddValue('AAMM', RefNFP.AAMM);
    AddValue('CNPJCPF', RefNFP.CNPJCPF);
    AddValue('IE', RefNFP.IE);
    AddValue('modelo', RefNFP.modelo);
    AddValue('serie', RefNFP.serie);
    AddValue('nNF', RefNFP.nNF);
  end;
end;

function TNfeService.EnderEmitToVariant(EnderEmit: TEnderEmit): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xLgr', EnderEmit.xLgr);
    AddValue('nro', EnderEmit.nro);
    AddValue('xCpl', EnderEmit.xCpl);
    AddValue('xBairro', EnderEmit.xBairro);
    AddValue('cMun', EnderEmit.cMun);
    AddValue('xMun', EnderEmit.xMun);
    AddValue('UF', EnderEmit.UF);
    AddValue('CEP', EnderEmit.CEP);
    AddValue('cPais', EnderEmit.cPais);
    AddValue('xPais', EnderEmit.xPais);
    AddValue('fone', EnderEmit.fone);
  end;
end;

function TNfeService.EnderDestToVariant(EnderDest: TEnderDest): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xLgr', EnderDest.xLgr);
    AddValue('nro', EnderDest.nro);
    AddValue('xCpl', EnderDest.xCpl);
    AddValue('xBairro', EnderDest.xBairro);
    AddValue('cMun', EnderDest.cMun);
    AddValue('xMun', EnderDest.xMun);
    AddValue('UF', EnderDest.UF);
    AddValue('CEP', EnderDest.CEP);
    AddValue('cPais', EnderDest.cPais);
    AddValue('xPais', EnderDest.xPais);
    AddValue('fone', EnderDest.fone);
  end;
end;

function TNfeService.DetItensToVariant(DetItens: TDetCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('Prod', ProdToVariant(DetItens.Prod));
    AddValue('Imposto', ImpostoToVariant(DetItens.Imposto));
    AddValue('pDevol', DetItens.pDevol);
    AddValue('vIPIDevol', DetItens.vIPIDevol);
    AddValue('infAdProd', DetItens.infAdProd);
    AddValue('obsCont', ObsItemToVariant(DetItens.obsCont));
    AddValue('obsFisco', ObsItemToVariant(DetItens.obsFisco));
  end;
end;

function TNfeService.ProdToVariant(Prod: TProd): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('cProd', Prod.cProd);
    AddValue('nItem', Prod.nItem);
    AddValue('cEAN', Prod.cEAN);
    AddValue('xProd', Prod.xProd);
    AddValue('NCM', Prod.NCM);
    AddValue('EXTIPI', Prod.EXTIPI);
    AddValue('CFOP', Prod.CFOP);
    AddValue('uCom', Prod.uCom);
    AddValue('qCom', Prod.qCom);
    AddValue('vUnCom', Prod.vUnCom);
    AddValue('vProd', Prod.vProd);
    AddValue('cEANTrib', Prod.cEANTrib);
    AddValue('uTrib', Prod.uTrib);
    AddValue('qTrib', Prod.qTrib);
    AddValue('vUnTrib', Prod.vUnTrib);
    AddValue('vFrete', Prod.vFrete);
    AddValue('vSeg', Prod.vSeg);
    AddValue('vDesc', Prod.vDesc);
    AddValue('vOutro', Prod.vOutro);
    AddValue('IndTot', indTotToStr(Prod.IndTot));
    AddValue('DI', DIToVariant(Prod.DI));
    AddValue('xPed', Prod.xPed);
    AddValue('nItemPed', Prod.nItemPed);
    AddValue('detExport', DetExportToVariant(Prod.detExport));
    AddValue('Rastro', RastroToVariant(Prod.Rastro));
    AddValue('veicProd', VeicProdToVariant(Prod.veicProd));
    AddValue('med', MedToVariant(Prod.med));
    AddValue('arma', ArmaToVariant(Prod.arma));
    AddValue('comb', CombToVariant(Prod.comb));
    AddValue('nRECOPI', Prod.nRECOPI);
    AddValue('nFCI', Prod.nFCI);
    AddValue('NVE', NVEToVariant(Prod.NVE));
    AddValue('CEST', Prod.CEST);
    AddValue('indEscala', Prod.indEscala);
    AddValue('CNPJFab', Prod.CNPJFab);
    AddValue('cBenef', Prod.cBenef);
    AddValue('cBarra', Prod.cBarra);
    AddValue('cBarraTrib', Prod.cBarraTrib);
  end;
end;

function TNfeService.ImpostoToVariant(Imposto: TImposto): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vTotTrib', Imposto.vTotTrib);
    AddValue('ICMS', ICMSToVariant(Imposto.ICMS));
    AddValue('IPI', IPIToVariant(Imposto.IPI));
    AddValue('II', IIToVariant(Imposto.II));
    AddValue('PIS', PISToVariant(Imposto.PIS));
    AddValue('PISST', PISSTToVariant(Imposto.PISST));
    AddValue('COFINS', COFINSToVariant(Imposto.COFINS));
    AddValue('COFINSST', COFINSSTToVariant(Imposto.COFINSST));
    AddValue('ISSQN', ISSQNToVariant(Imposto.ISSQN));
    AddValue('ICMSUFDest', ICMSUFDestToVariant(Imposto.ICMSUFDest));
  end;
end;

function TNfeService.ObsItemToVariant(ObsItem: TObsItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xCampo', ObsItem.xCampo);
    AddValue('xTexto', ObsItem.xTexto);
  end;
end;

function TNfeService.ICMSToVariant(ICMS: TICMS): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('orig', OrigToStr(ICMS.orig));
    AddValue('CST', CSTICMSToStr(ICMS.CST));
    AddValue('CSOSN', CSOSNIcmsToStr(ICMS.CSOSN));
    AddValue('modBC', modBCToStr(ICMS.modBC));
    AddValue('pRedBC', ICMS.pRedBC);
    AddValue('vBC', ICMS.vBC);
    AddValue('pICMS', ICMS.pICMS);
    AddValue('vICMS', ICMS.vICMS);
    AddValue('modBCST', modBCSTToStr(ICMS.modBCST));
    AddValue('pMVAST', ICMS.pMVAST);
    AddValue('pRedBCST', ICMS.pRedBCST);
    AddValue('vBCST', ICMS.vBCST);
    AddValue('pICMSST', ICMS.pICMSST);
    AddValue('vICMSST', ICMS.vICMSST);
    AddValue('UFST', ICMS.UFST);
    AddValue('pBCOp', ICMS.pBCOp);
    AddValue('vBCSTRet', ICMS.vBCSTRet);
    AddValue('vICMSSTRet', ICMS.vICMSSTRet);
    AddValue('motDesICMS', motDesICMSToStr(ICMS.motDesICMS));
    AddValue('pCredSN', ICMS.pCredSN);
    AddValue('vCredICMSSN', ICMS.vCredICMSSN);
    AddValue('vBCSTDest', ICMS.vBCSTDest);
    AddValue('vICMSSTDest', ICMS.vICMSSTDest);
    AddValue('vICMSDeson', ICMS.vICMSDeson);
    AddValue('vICMSOp', ICMS.vICMSOp);
    AddValue('pDif', ICMS.pDif);
    AddValue('vICMSDif', ICMS.vICMSDif);
    AddValue('vBCFCP', ICMS.vBCFCP);
    AddValue('pFCP', ICMS.pFCP);
    AddValue('vFCP', ICMS.vFCP);
    AddValue('vBCFCPST', ICMS.vBCFCPST);
    AddValue('pFCPST', ICMS.pFCPST);
    AddValue('vFCPST', ICMS.vFCPST);
    AddValue('vBCFCPSTRet', ICMS.vBCFCPSTRet);
    AddValue('pFCPSTRet', ICMS.pFCPSTRet);
    AddValue('vFCPSTRet', ICMS.vFCPSTRet);
    AddValue('pST', ICMS.pST);
    AddValue('pRedBCEfet', ICMS.pRedBCEfet);
    AddValue('vBCEfet', ICMS.vBCEfet);
    AddValue('pICMSEfet', ICMS.pICMSEfet);
    AddValue('vICMSEfet', ICMS.vICMSEfet);
    AddValue('vICMSSubstituto', ICMS.vICMSSubstituto);
    AddValue('vICMSSTDeson', ICMS.vICMSSTDeson);
    AddValue('motDesICMSST', motDesICMSToStr(ICMS.motDesICMSST));
    AddValue('vFCPDif', ICMS.vFCPDif);
    AddValue('vFCPEfet', ICMS.vFCPEfet);
    AddValue('pFCPDif', ICMS.pFCPDif);
    AddValue('adRemICMS', ICMS.adRemICMS);
    AddValue('vICMSMono', ICMS.vICMSMono);
    AddValue('adRemICMSReten', ICMS.adRemICMSReten);
    AddValue('vICMSMonoReten', ICMS.vICMSMonoReten);
    AddValue('vICMSMonoDif', ICMS.vICMSMonoDif);
    AddValue('adRemICMSRet', ICMS.adRemICMSRet);
    AddValue('vICMSMonoRet', ICMS.vICMSMonoRet);
    AddValue('qBCMono', ICMS.qBCMono);
    AddValue('qBCMonoReten', ICMS.qBCMonoReten);
    AddValue('pRedAdRem', ICMS.pRedAdRem);
    AddValue('motRedAdRem', motRedAdRemToStr(ICMS.motRedAdRem));
    AddValue('vICMSMonoOp', ICMS.vICMSMonoOp);
    AddValue('qBCMonoRet', ICMS.qBCMonoRet);
  end;
end;

function TNfeService.IPIToVariant(IPI: TIPI): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('clEnq', IPI.clEnq);
    AddValue('CNPJProd', IPI.CNPJProd);
    AddValue('cSelo', IPI.cSelo);
    AddValue('qSelo', IPI.qSelo);
    AddValue('cEnq', IPI.cEnq);
    AddValue('CST', CSTIPIToStr(IPI.CST));
    AddValue('vBC', IPI.vBC);
    AddValue('qUnid', IPI.qUnid);
    AddValue('vUnid', IPI.vUnid);
    AddValue('pIPI', IPI.pIPI);
    AddValue('vIPI', IPI.vIPI);
  end;
end;

function TNfeService.IIToVariant(II: TII): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBc', II.vBc);
    AddValue('vDespAdu', II.vDespAdu);
    AddValue('vII', II.vII);
    AddValue('vIOF', II.vIOF);
  end;
end;

function TNfeService.PISToVariant(PIS: TPIS): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CST', CSTPISToStr(PIS.CST));
    AddValue('vBC', PIS.vBC);
    AddValue('pPIS', PIS.pPIS);
    AddValue('vPIS', PIS.vPIS);
    AddValue('qBCProd', PIS.qBCProd);
    AddValue('vAliqProd', PIS.vAliqProd);
  end;
end;

function TNfeService.PISSTToVariant(PISST: TPISST): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBc', PISST.vBc);
    AddValue('pPis', PISST.pPis);
    AddValue('qBCProd', PISST.qBCProd);
    AddValue('vAliqProd', PISST.vAliqProd);
    AddValue('vPIS', PISST.vPIS);
    AddValue('indSomaPISST', indSomaPISSTToStr(PISST.indSomaPISST));
  end;
end;

function TNfeService.COFINSToVariant(COFINS: TCOFINS): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CST', CSTCOFINSToStr(COFINS.CST));
    AddValue('vBC', COFINS.vBC);
    AddValue('pCOFINS', COFINS.pCOFINS);
    AddValue('vCOFINS', COFINS.vCOFINS);
    AddValue('vBCProd', COFINS.vBCProd);
    AddValue('vAliqProd', COFINS.vAliqProd);
    AddValue('qBCProd', COFINS.qBCProd);
  end;
end;

function TNfeService.COFINSSTToVariant(COFINSST: TCOFINSST): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBC', COFINSST.vBC);
    AddValue('pCOFINS', COFINSST.pCOFINS);
    AddValue('qBCProd', COFINSST.qBCProd);
    AddValue('vAliqProd', COFINSST.vAliqProd);
    AddValue('vCOFINS', COFINSST.vCOFINS);
    AddValue('indSomaCOFINSST', indSomaCOFINSSTToStr(COFINSST.indSomaCOFINSST));
  end;
end;

function TNfeService.ISSQNToVariant(ISSQN: TISSQN): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBC', ISSQN.vBC);
    AddValue('vAliq', ISSQN.vAliq);
    AddValue('vISSQN', ISSQN.vISSQN);
    AddValue('cMunFG', ISSQN.cMunFG);
    AddValue('cListServ', ISSQN.cListServ);
    AddValue('cSitTrib', ISSQNcSitTribToStr(ISSQN.cSitTrib));
    AddValue('vDeducao', ISSQN.vDeducao);
    AddValue('vOutro', ISSQN.vOutro);
    AddValue('vDescIncond', ISSQN.vDescIncond);
    AddValue('vDescCond', ISSQN.vDescCond);
    AddValue('indISSRet', indISSRetToStr(ISSQN.indISSRet));
    AddValue('vISSRet', ISSQN.vISSRet);
    AddValue('indISS', indISSToStr(ISSQN.indISS));
    AddValue('cServico', ISSQN.cServico);
    AddValue('cMun', ISSQN.cMun);
    AddValue('cPais', ISSQN.cPais);
    AddValue('nProcesso', ISSQN.nProcesso);
    AddValue('indIncentivo', indIncentivoToStr(ISSQN.indIncentivo));
  end;
end;

function TNfeService.ICMSUFDestToVariant(ICMSUFDest: TICMSUFDest): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBCUFDest', ICMSUFDest.vBCUFDest);
    AddValue('vBCFCPUFDest', ICMSUFDest.vBCFCPUFDest);
    AddValue('pFCPUFDest', ICMSUFDest.pFCPUFDest);
    AddValue('pICMSUFDest', ICMSUFDest.pICMSUFDest);
    AddValue('pICMSInter', ICMSUFDest.pICMSInter);
    AddValue('pICMSInterPart', ICMSUFDest.pICMSInterPart);
    AddValue('vFCPUFDest', ICMSUFDest.vFCPUFDest);
    AddValue('vICMSUFDest', ICMSUFDest.vICMSUFDest);
    AddValue('vICMSUFRemet', ICMSUFDest.vICMSUFRemet);
  end;
end;

function TNfeService.DIToVariant(DI: TDICollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to DI.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(DIItensToVariant(DI.Items[i]));
  end;
end;

function TNfeService.DIItensToVariant(DIItens: TDICollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nDi', DIItens.nDi);
    AddValue('dDi', DIItens.dDi);
    AddValue('xLocDesemb', DIItens.xLocDesemb);
    AddValue('UFDesemb', DIItens.UFDesemb);
    AddValue('dDesemb', DIItens.dDesemb);
    AddValue('tpViaTransp', TipoViaTranspToStr(DIItens.tpViaTransp));
    AddValue('vAFRMM', DIItens.vAFRMM);
    AddValue('tpIntermedio', TipoIntermedioToStr(DIItens.tpIntermedio));
    AddValue('CNPJ', DIItens.CNPJ);
    AddValue('UFTerceiro', DIItens.UFTerceiro);
    AddValue('cExportador', DIItens.cExportador);
    AddValue('adi', AdiToVariant(DIItens.adi));
  end;
end;

function TNfeService.AdiToVariant(Adi: TAdiCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Adi.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(AdiItensToVariant(Adi.Items[i]));
  end;
end;

function TNfeService.AdiItensToVariant(AdiItens: TAdiCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nAdicao', AdiItens.nAdicao);
    AddValue('nSeqAdi', AdiItens.nSeqAdi);
    AddValue('cFabricante', AdiItens.cFabricante);
    AddValue('vDescDI', AdiItens.vDescDI);
    AddValue('nDraw', AdiItens.nDraw);
  end;
end;

function TNfeService.DetExportToVariant(DetExport: TdetExportCollection
  ): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to DetExport.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(DetExportItensToVariant(DetExport.Items[i]));
  end;
end;

function TNfeService.DetExportItensToVariant(
  DetExportItens: TdetExportCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nDraw', DetExportItens.nDraw);
    AddValue('nRE', DetExportItens.nRE);
    AddValue('chNFe', DetExportItens.chNFe);
    AddValue('qExport', DetExportItens.qExport);
  end;
end;

function TNfeService.RastroToVariant(Rastro: TRastroCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Rastro.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(RastroItensToVariant(Rastro.Items[i]));
  end;
end;

function TNfeService.RastroItensToVariant(RastroItens: TRastroCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nLote', RastroItens.nLote);
    AddValue('qLote', RastroItens.qLote);
    AddValue('dFab', RastroItens.dFab);
    AddValue('dVal', RastroItens.dVal);
    AddValue('cAgreg', RastroItens.cAgreg);
  end;
end;

function TNfeService.VeicProdToVariant(VeicProd: TVeicProd): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('tpOP', tpOPToStr(VeicProd.tpOP));
    AddValue('chassi', VeicProd.chassi);
    AddValue('cCor', VeicProd.cCor);
    AddValue('xCor', VeicProd.xCor);
    AddValue('pot', VeicProd.pot);
    AddValue('Cilin', VeicProd.Cilin);
    AddValue('pesoL', VeicProd.pesoL);
    AddValue('pesoB', VeicProd.pesoB);
    AddValue('nSerie', VeicProd.nSerie);
    AddValue('tpComb', VeicProd.tpComb);
    AddValue('nMotor', VeicProd.nMotor);
    AddValue('CMT', VeicProd.CMT);
    AddValue('dist', VeicProd.dist);
    AddValue('anoMod', VeicProd.anoMod);
    AddValue('anoFab', VeicProd.anoFab);
    AddValue('tpPint', VeicProd.tpPint);
    AddValue('tpVeic', VeicProd.tpVeic);
    AddValue('espVeic', VeicProd.espVeic);
    AddValue('VIN', VeicProd.VIN);
    AddValue('condVeic', condVeicToStr(VeicProd.condVeic));
    AddValue('cMod', VeicProd.cMod);
    AddValue('cCorDENATRAN', VeicProd.cCorDENATRAN);
    AddValue('lota', VeicProd.lota);
    AddValue('tpRest', VeicProd.tpRest);
  end;
end;

function TNfeService.MedToVariant(Med: TMedCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Med.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(MedItensToVariant(Med.Items[i]));
  end;
end;

function TNfeService.MedItensToVariant(MedItens: TMedCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('cProdANVISA', MedItens.cProdANVISA);
    AddValue('xMotivoIsencao', MedItens.xMotivoIsencao);
    AddValue('nLote', MedItens.nLote);
    AddValue('qLote', MedItens.qLote);
    AddValue('dFab', MedItens.dFab);
    AddValue('dVal', MedItens.dVal);
    AddValue('vPMC', MedItens.vPMC);
  end;
end;

function TNfeService.ArmaToVariant(Arma: TArmaCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Arma.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(ArmaItensToVariant(Arma.Items[i]));
  end;
end;

function TNfeService.ArmaItensToVariant(ArmaItens: TArmaCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('tpArma', ArmaItens.tpArma);
    AddValue('nSerie', ArmaItens.nSerie);
    AddValue('nCano', ArmaItens.nCano);
    AddValue('descr', ArmaItens.descr);
  end;
end;

function TNfeService.CombToVariant(Comb: TComb): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('cProdANP', Comb.cProdANP);
    AddValue('pMixGN', Comb.pMixGN);
    AddValue('descANP', Comb.descANP);
    AddValue('pGLP', Comb.pGLP);
    AddValue('pGNn', Comb.pGNn);
    AddValue('pGNi', Comb.pGNi);
    AddValue('vPart', Comb.vPart);
    AddValue('CODIF', Comb.CODIF);
    AddValue('qTemp', Comb.qTemp);
    AddValue('UFcons', Comb.UFcons);
    AddValue('CIDE', CIDEToVariant(Comb.CIDE));
    AddValue('ICMS', ICMSCombToVariant(Comb.ICMS));
    AddValue('ICMSInter', ICMSInterToVariant(Comb.ICMSInter));
    AddValue('ICMSCons', ICMSConsToVariant(Comb.ICMSCons));
    AddValue('encerrante', EncerranteToVariant(Comb.encerrante));
    AddValue('pBio', Comb.pBio);
    AddValue('origComb', OrigCombToVariant(Comb.origComb));
  end;
end;

function TNfeService.CIDEToVariant(CIDE: TCIDE): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('qBCProd', CIDE.qBCProd);
    AddValue('vAliqProd', CIDE.vAliqProd);
    AddValue('vCIDE', CIDE.vCIDE);
  end;
end;

function TNfeService.ICMSCombToVariant(ICMSComb: TICMSComb): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBCICMS', ICMSComb.vBCICMS);
    AddValue('vICMS', ICMSComb.vICMS);
    AddValue('vBCICMSST', ICMSComb.vBCICMSST);
    AddValue('vICMSST', ICMSComb.vICMSST);
  end;
end;

function TNfeService.ICMSInterToVariant(ICMSInter: TICMSInter): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBCICMSSTDest', ICMSInter.vBCICMSSTDest);
    AddValue('vICMSSTDest', ICMSInter.vICMSSTDest);
  end;
end;

function TNfeService.ICMSConsToVariant(ICMSCons: TICMSCons): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBCICMSSTCons', ICMSCons.vBCICMSSTCons);
    AddValue('vICMSSTCons', ICMSCons.vICMSSTCons);
    AddValue('UFcons', ICMSCons.UFcons);
  end;
end;

function TNfeService.EncerranteToVariant(Encerrante: TEncerrante): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nBico', Encerrante.nBico);
    AddValue('nBomba', Encerrante.nBomba);
    AddValue('nTanque', Encerrante.nTanque);
    AddValue('vEncIni', Encerrante.vEncIni);
    AddValue('vEncFin', Encerrante.vEncFin);
  end;
end;

function TNfeService.OrigCombToVariant(OrigComb: TorigCombCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to OrigComb.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(OrigCombItensToVariant(OrigComb.Items[i]));
  end;
end;

function TNfeService.OrigCombItensToVariant(
  OrigCombItens: TorigCombCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('indImport', indImportToStr(OrigCombItens.indImport));
    AddValue('cUFOrig', OrigCombItens.cUFOrig);
    AddValue('pOrig', OrigCombItens.pOrig);
  end;
end;

function TNfeService.NVEToVariant(NVE: TNVECollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to NVE.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(NVEItensToVariant(NVE.Items[i]));
  end;
end;

function TNfeService.NVEItensToVariant(NVEItens: TNVECollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('Nve', NVEItens.Nve);
  end;
end;

function TNfeService.ICMSTotToVariant(ICMSTot: TICMSTot): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vBC', ICMSTot.vBC);
    AddValue('vICMS', ICMSTot.vICMS);
    AddValue('vICMSDeson', ICMSTot.vICMSDeson);
    AddValue('vFCPUFDest', ICMSTot.vFCPUFDest);
    AddValue('vICMSUFDest', ICMSTot.vICMSUFDest);
    AddValue('vICMSUFRemet', ICMSTot.vICMSUFRemet);
    AddValue('vFCP', ICMSTot.vFCP);
    AddValue('vBCST', ICMSTot.vBCST);
    AddValue('vST', ICMSTot.vST);
    AddValue('vFCPST', ICMSTot.vFCPST);
    AddValue('vFCPSTRet', ICMSTot.vFCPSTRet);
    AddValue('vProd', ICMSTot.vProd);
    AddValue('vFrete', ICMSTot.vFrete);
    AddValue('vSeg', ICMSTot.vSeg);
    AddValue('vDesc', ICMSTot.vDesc);
    AddValue('vII', ICMSTot.vII);
    AddValue('vIPI', ICMSTot.vIPI);
    AddValue('vIPIDevol', ICMSTot.vIPIDevol);
    AddValue('vPIS', ICMSTot.vPIS);
    AddValue('vCOFINS', ICMSTot.vCOFINS);
    AddValue('vOutro', ICMSTot.vOutro);
    AddValue('vNF', ICMSTot.vNF);
    AddValue('vTotTrib', ICMSTot.vTotTrib);
    AddValue('vICMSMono', ICMSTot.vICMSMono);
    AddValue('vICMSMonoReten', ICMSTot.vICMSMonoReten);
    AddValue('vICMSMonoRet', ICMSTot.vICMSMonoRet);
    AddValue('qBCMono', ICMSTot.qBCMono);
    AddValue('qBCMonoReten', ICMSTot.qBCMonoReten);
    AddValue('qBCMonoRet', ICMSTot.qBCMonoRet);
  end;
end;

function TNfeService.ISSQNtotToVariant(ISSQNtot: TISSQNtot): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vServ', ISSQNtot.vServ);
    AddValue('vBC', ISSQNtot.vBC);
    AddValue('vISS', ISSQNtot.vISS);
    AddValue('vPIS', ISSQNtot.vPIS);
    AddValue('vCOFINS', ISSQNtot.vCOFINS);
    AddValue('dCompet', ISSQNtot.dCompet);
    AddValue('vDeducao', ISSQNtot.vDeducao);
    AddValue('vOutro', ISSQNtot.vOutro);
    AddValue('vDescIncond', ISSQNtot.vDescIncond);
    AddValue('vDescCond', ISSQNtot.vDescCond);
    AddValue('vISSRet', ISSQNtot.vISSRet);
    AddValue('cRegTrib', RegTribISSQNToStr(ISSQNtot.cRegTrib));
  end;
end;

function TNfeService.RetTribToVariant(RetTrib: TretTrib): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vRetPIS', RetTrib.vRetPIS);
    AddValue('vRetCOFINS', RetTrib.vRetCOFINS);
    AddValue('vRetCSLL', RetTrib.vRetCSLL);
    AddValue('vBCIRRF', RetTrib.vBCIRRF);
    AddValue('vIRRF', RetTrib.vIRRF);
    AddValue('vBCRetPrev', RetTrib.vBCRetPrev);
    AddValue('vRetPrev', RetTrib.vRetPrev);
  end;
end;

function TNfeService.TransportaToVariant(Transporta: TTransporta): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJCPF', Transporta.CNPJCPF);
    AddValue('xNome', Transporta.xNome);
    AddValue('IE', Transporta.IE);
    AddValue('xEnder', Transporta.xEnder);
    AddValue('xMun', Transporta.xMun);
    AddValue('UF', Transporta.UF);
  end;
end;

function TNfeService.RetTranspToVariant(RetTransp: TRetTransp): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('vServ', RetTransp.vServ);
    AddValue('vBCRet', RetTransp.vBCRet);
    AddValue('pICMSRet', RetTransp.pICMSRet);
    AddValue('vICMSRet', RetTransp.vICMSRet);
    AddValue('CFOP', RetTransp.CFOP);
    AddValue('cMunFG', RetTransp.cMunFG);
  end;
end;

function TNfeService.VeicTranspToVariant(VeicTransp: TVeicTransp): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('placa', VeicTransp.placa);
    AddValue('UF', VeicTransp.UF);
    AddValue('RNTC', VeicTransp.RNTC);
  end;
end;

function TNfeService.VolToVariant(Vol: TVolCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Vol.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(VolItensToVariant(Vol.Items[i]));
  end;
end;

function TNfeService.VolItensToVariant(VolItens: TVolCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('qVol', VolItens.qVol);
    AddValue('esp', VolItens.esp);
    AddValue('marca', VolItens.marca);
    AddValue('nVol', VolItens.nVol);
    AddValue('pesoL', VolItens.pesoL);
    AddValue('pesoB', VolItens.pesoB);
    AddValue('Lacres', LacresToVariant(VolItens.Lacres));
  end;
end;

function TNfeService.LacresToVariant(Lacres: TLacresCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Lacres.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(LacresItensToVariant(Lacres.Items[i]));
  end;
end;

function TNfeService.LacresItensToVariant(LacresItens: TLacresCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nLacre', LacresItens.nLacre);
  end;
end;

function TNfeService.ReboqueToVariant(Reboque: TReboqueCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Reboque.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(ReboqueItensToVariant(Reboque.Items[i]));
  end;
end;

function TNfeService.ReboqueItensToVariant(ReboqueItens: TReboqueCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('placa', ReboqueItens.placa);
    AddValue('UF', ReboqueItens.UF);
    AddValue('RNTC', ReboqueItens.RNTC);
  end;
end;

function TNfeService.FatToVariant(Fat: TFat): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nFat', Fat.nFat);
    AddValue('vOrig', Fat.vOrig);
    AddValue('vDesc', Fat.vDesc);
    AddValue('vLiq', Fat.vLiq);
  end;
end;

function TNfeService.DupToVariant(Dup: TDupCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Dup.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(DupItensToVariant(Dup.Items[i]));
  end;
end;

function TNfeService.DupItensToVariant(DupItens: TDupCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nDup', DupItens.nDup);
    AddValue('dVenc', DupItens.dVenc);
    AddValue('vDup', DupItens.vDup);
  end;
end;

function TNfeService.PagItensToVariant(PagItens: TpagCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('tPag', FormaPagamentoToStr(PagItens.tPag));
    AddValue('xPag', PagItens.xPag);
    AddValue('vPag', PagItens.vPag);
    AddValue('tpIntegra', tpIntegraToStr(PagItens.tpIntegra));
    AddValue('CNPJ', PagItens.CNPJ);
    AddValue('tBand', TpcnBandeiraCartao(PagItens.tBand));
    AddValue('cAut', PagItens.cAut);
    AddValue('indPag', IndpagToStr(PagItens.indPag));
  end;
end;

function TNfeService.ObsContToVariant(ObsCont: TobsContCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to ObsCont.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(ObsContItensToVariant(ObsCont.Items[i]));
  end;
end;

function TNfeService.ObsContItensToVariant(ObsContItens: TobsContCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xCampo', ObsContItens.xCampo);
    AddValue('xTexto', ObsContItens.xTexto);
  end;
end;

function TNfeService.ObsFiscoToVariant(ObsFisco: TobsFiscoCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to ObsFisco.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(ObsFiscoItensToVariant(ObsFisco.Items[i]));
  end;
end;

function TNfeService.ObsFiscoItensToVariant(
  ObsFiscoItens: TobsFiscoCollectionItem): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xCampo', ObsFiscoItens.xCampo);
    AddValue('xTexto', ObsFiscoItens.xTexto);
  end;
end;

function TNfeService.ProcRefToVariant(ProcRef: TprocRefCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to ProcRef.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(ProcRefItensToVariant(ProcRef.Items[i]));
  end;
end;

function TNfeService.ProcRefItensToVariant(ProcRefItens: TprocRefCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('nProc', ProcRefItens.nProc);
    AddValue('indProc', indProcToStr(ProcRefItens.indProc));
    AddValue('tpAto', tpAtoToStr(ProcRefItens.tpAto));
  end;
end;

function TNfeService.FordiaToVariant(Fordia: TForDiaCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Fordia.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(FordiaItensToVariant(Fordia.Items[i]));
  end;
end;

function TNfeService.FordiaItensToVariant(FordiaItens: TForDiaCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('dia', FordiaItens.dia);
    AddValue('qtde', FordiaItens.qtde);
  end;
end;

function TNfeService.DeducToVariant(Deduc: TDeducCollection): Variant;
var
  i: integer;
begin
  TDocVariant.NewFast(Result);
  for i := 0 to Deduc.Count - 1 do
  begin
    TDocVariantData(Result).AddItem(DeducItensToVariant(Deduc.Items[i]));
  end;
end;

function TNfeService.DeducItensToVariant(DeducItens: TDeducCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('xDed', DeducItens.xDed);
    AddValue('vDed', DeducItens.vDed);
  end;
end;

function TNfeService.GeradorToVariant(Gerador: TGerador): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('ArquivoFormatoXML', Gerador.ArquivoFormatoXML);
    AddValue('ArquivoFormatoTXT', Gerador.ArquivoFormatoTXT);
    AddValue('IDNivel', Gerador.IDNivel);
    AddValue('ListaDeAlertas', Gerador.ListaDeAlertas.GetText);
    AddValue('LayoutArquivoTXT', Gerador.LayoutArquivoTXT.GetText);
    AddValue('Opcoes', GeradorOpcoesToVariant(Gerador.Opcoes));
    AddValue('Prefixo', Gerador.Prefixo);
  end;
end;

function TNfeService.GeradorOpcoesToVariant(GeradorOpcoes: TGeradorOpcoes
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('SomenteValidar', GeradorOpcoes.SomenteValidar);
    AddValue('RetirarEspacos', GeradorOpcoes.RetirarEspacos);
    AddValue('RetirarAcentos', GeradorOpcoes.RetirarAcentos);
    AddValue('IdentarXML', GeradorOpcoes.IdentarXML);
    AddValue('TamanhoIdentacao', GeradorOpcoes.TamanhoIdentacao);
    AddValue('SuprimirDecimais', GeradorOpcoes.SuprimirDecimais);
    AddValue('TagVaziaNoFormatoResumido', GeradorOpcoes.TagVaziaNoFormatoResumido);
    AddValue('FormatoAlerta', GeradorOpcoes.FormatoAlerta);
    AddValue('DecimalChar', GeradorOpcoes.DecimalChar);
    AddValue('QuebraLinha', GeradorOpcoes.QuebraLinha);
  end;
end;

function TNfeService.AutXMLItensToVariant(AutXMLItens: TautXMLCollectionItem
  ): Variant;
begin
  TDocVariant.NewFast(Result);
  with TDocVariantData(Result) do
  begin
    AddValue('CNPJCPF', AutXMLItens.CNPJCPF);
  end;
end;

end.

