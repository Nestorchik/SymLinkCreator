program SymLinkCreator;

{$R *.dres}

uses
  Forms,
  SymLinkUnit in 'SymLinkUnit.pas' {SymLinkForm},
  Vcl.Themes,
  Vcl.Styles,
  SymLinkINI in 'SymLinkINI.pas';

{Russian language file}

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'SymLinkCreator';
  Application.CreateForm(TSymLinkForm, SymLinkForm);
  Application.Run;
end.
