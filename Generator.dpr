program Generator;

uses
  Vcl.Forms,
  DAOGenerator in 'DAOGenerator.pas',
  DTOGenerator in 'DTOGenerator.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
