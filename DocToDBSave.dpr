program DocToDBSave;

uses
  Forms,
  SaveDocFunc in 'SaveDocFunc.pas',
  MatterSearch in 'MatterSearch.pas' {frmMtrSearch},
  SavedocDetails in 'SavedocDetails.pas' {frmSaveDocDetails},
  DiffUnit in 'DiffUnit.pas',
  SaveDoc in 'SaveDoc.pas' {dmSaveDoc: TDataModule},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Iceberg Classico');
  Application.Title := 'SaveDocToDB';
  Application.CreateForm(TdmSaveDoc, dmSaveDoc);
  Application.CreateForm(TfrmSaveDocDetails, frmSaveDocDetails);
  Application.Run;
end.
