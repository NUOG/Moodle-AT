program MAD;

uses
  Forms,
  Main in 'Main.pas' {MoodleAD};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Moodle Advanced Tools v.3.0';
  Application.CreateForm(TMoodleAD, MoodleAD);
  Application.Run;
end.
