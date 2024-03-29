program SymLinkCreator;

{$R *.dres}

uses
  Forms,
  SymLinkUnit in 'SymLinkUnit.pas' {SymLinkForm},
  Vcl.Themes,
  Vcl.Styles;
{Russian language file}

{$R *.res}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Carbon');
  Application.Title := 'SymLinkCreator';
  Application.CreateForm(TSymLinkForm, SymLinkForm);
  Application.Run;
end.
