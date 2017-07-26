unit SaveDoc;

interface

uses
  SysUtils, Classes, DB, DBAccess, MemDS, OraSmart, Ora, Dialogs;

type
  TdmSaveDoc = class(TDataModule)
    orsAxiom: TOraSession;
    qryEmps: TOraQuery;
    qryMatters: TOraQuery;
    dsMatters: TOraDataSource;
    qryGetSeq: TOraQuery;
    qryMatterAttachments: TSmartQuery;
    qryGetMatter: TOraQuery;
    qryGetEntity: TOraQuery;
    qryPrecCategory: TOraQuery;
    dsPrecCategory: TOraDataSource;
    qryTmp: TOraQuery;
    qrySysFile: TOraQuery;
    qryMatterAttachmentsROWID: TStringField;
    qryMatterAttachmentsDOCUMENT: TBlobField;
    qryMatterAttachmentsIMAGEINDEX: TFloatField;
    qryMatterAttachmentsFILE_EXTENSION: TStringField;
    qryMatterAttachmentsDOC_NAME: TStringField;
    qryMatterAttachmentsSEARCH: TStringField;
    qryMatterAttachmentsDOC_CODE: TStringField;
    qryMatterAttachmentsJURIS: TStringField;
    qryMatterAttachmentsD_CREATE: TDateTimeField;
    qryMatterAttachmentsAUTH1: TStringField;
    qryMatterAttachmentsD_MODIF: TDateTimeField;
    qryMatterAttachmentsAUTH2: TStringField;
    qryMatterAttachmentsPATH: TStringField;
    qryMatterAttachmentsDESCR: TStringField;
    qryMatterAttachmentsFILEID: TStringField;
    qryMatterAttachmentsDOCID: TFloatField;
    qryMatterAttachmentsNPRECCATEGORY: TFloatField;
    qryMatterAttachmentsNMATTER: TFloatField;
    procTemp: TOraStoredProc;
    qryPrecClassification: TOraQuery;
    dsPrecClassification: TOraDataSource;
    dsEmployee: TOraDataSource;
    qryEmployee: TOraQuery;
    qryMatterAttachmentsPRECEDENT_DETAILS: TStringField;
    qryMatterAttachmentsNPRECCLASSIFICATION: TFloatField;
    qryMatterAttachmentsKEYWORDS: TStringField;
    qryMatterAttachmentsDISPLAY_PATH: TStringField;
    qryMatterAttachmentsEXTERNAL_ACCESS: TStringField;
    procedure qryMatterAttachmentsNewRecord(DataSet: TDataSet);
    procedure orsAxiomError(Sender: TObject; E: EDAError;
      var Fail: Boolean);
  private
    { Private declarations }
    FUserID : string;
    FEntity : string;
    FDocID   : string;
  public
    { Public declarations }
    property UserID : string read FUserID write FUserID;
    property Entity : string read FEntity write FEntity;
    property DocID  : string read FDocID write FDocID;
  end;

var
  dmSaveDoc: TdmSaveDoc;

implementation

{$R *.dfm}


procedure TdmSaveDoc.qryMatterAttachmentsNewRecord(DataSet: TDataSet);
begin
   dmSaveDoc.qryGetSeq.ExecSQL;
   FDocID := dmSaveDoc.qryGetSeq.FieldByName('nextdoc').AsString;
   dmSaveDoc.qryMatterAttachments.FieldByName('docid').AsString := FDocID;
end;

procedure TdmSaveDoc.orsAxiomError(Sender: TObject; E: EDAError;
  var Fail: Boolean);
begin
   case E.ErrorCode of
      1005: Fail := False;
   else
      MessageDlg('Insight Database Error:'#13#10 + e.Message, mtError, [mbOK], 0);
   end;
end;

end.
