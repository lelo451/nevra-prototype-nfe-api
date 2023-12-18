program RestServer;

{$MODE DELPHI}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  SysUtils,
  fpjson,
  Interfaces,
  Horse,
  Horse.Jhonson,
  Horse.Compression,
  pcnNFe,
  nfeService;

var
  NFeService: TNfeService;

const
  PORT = 9000;
  HOST = 'localhost';

procedure GetNFe(Req: THorseRequest; Res: THorseResponse);
var
  NFe: TNFe;
  body: Variant;
begin
  try
    NFe        := NFeService.readXML(Req.Body);
    body       := NFeService.NFeToVariant(NFe);
    Res.Send(body);
  finally
    NFe.Free;
  end;
end;

begin
  try
    THorse
      .Use(Compression())
      .Use(Jhonson('UTF-8'));
    NFeService := TNfeService.Create;
    THorse.Get('/NFe', GetNFe);
    THorse.Host := HOST;
    THorse.Port := PORT;
    THorse.Listen;
  finally
    NFeService.Free;
    THorse.StopListen;
  end;
end.

