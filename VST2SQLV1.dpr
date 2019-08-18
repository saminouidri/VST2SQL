program VST2SQLV1;

uses
  Vcl.Forms,
  Unit1 in '..\Learn_Visio_SamiN\Unit1.pas' {VST2SQL};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TVST2SQL, VST2SQL);
  Application.Run;
end.
