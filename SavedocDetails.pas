unit SavedocDetails;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxLabel, cxControls, cxContainer, cxEdit, cxTextEdit, ComObj,
  cxMaskEdit, cxButtonEdit, cxLookAndFeelPainters, StdCtrls, cxButtons,
  cxGroupBox, cxRadioGroup, Menus, LMDCustomComponent, LMDBrowseDlg, DB,
  cxGraphics, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, cxCheckBox, dbtables, DiffUnit, dxStatusBar,
  cxMemo, ActnList, ActnMan, cxLookAndFeels;

const
     CUSTOMPROPS: array[0..10] of string = ('MatterNo','DocID','Prec_Category','Prec_Classification','Doc_Keywords','Doc_Precedent','Doc_FileName','Doc_Author','Saved_in_DB', 'Doc_Title','Portal_Access');

type
  TfrmSaveDocDetails = class(TForm)
    btnEditMatter: TcxButtonEdit;
    lblMatter: TcxLabel;
    txtDocName: TcxTextEdit;
    cxLabel1: TcxLabel;
    rgStorage: TcxRadioGroup;
    cmbCategory: TcxLookupComboBox;
    cxLabel2: TcxLabel;
    cbOverwriteDoc: TcxCheckBox;
    cbLeaveDocOpen: TcxCheckBox;
    BrowseDlg: TLMDBrowseDlg;
    cxLabel3: TcxLabel;
    StatusBar: TdxStatusBar;
    cmbClassification: TcxLookupComboBox;
    cxLabel4: TcxLabel;
    cxLabel5: TcxLabel;
    edKeywords: TcxTextEdit;
    memoPrecDetails: TcxMemo;
    cxLabel6: TcxLabel;
    cxLabel7: TcxLabel;
    cmbAuthor: TcxLookupComboBox;
    cbPortalAccess: TcxCheckBox;
    cbNewCopy: TcxCheckBox;
    btnTxtDocPath: TcxButtonEdit;
    btnSave: TcxButton;
    btnClose: TcxButton;
    procedure btnEditMatterPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnCloseClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure rgStorageClick(Sender: TObject);
    procedure btnEditMatterPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnTxtDocPathPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure cmbCategoryPropertiesInitPopup(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    nMatter: integer;
    tmpFileName: string;
    tmpdir: string;
    FFileID: string;
    FPrec_Category: string;
    FEditing: boolean;
    FSavedInDB: string;
    FDocName: string;
    FPrec_Classification: string;
    FDoc_Keywords: string;
    FDoc_Precedent: string;
    FDoc_FileName: string;
    FDoc_Author: string;
    FOldFileID: string;
    function SaveDocument(DocSequence: string): boolean;
    procedure GetDetails;
  public
    { Public declarations }
    property DocName: string read FDocName;
  end;

var
  frmSaveDocDetails: TfrmSaveDocDetails;

function ShowDocSave: Integer; StdCall;

implementation

uses
    MatterSearch, SaveDocFunc, Word2000, Office2000, ActiveX, savedoc;

{$R *.dfm}

function ShowDocSave:integer;
var
   frmSaveDocDetails: TfrmSaveDocDetails;
begin
//   Application.Handle := AHandle;
   frmSaveDocDetails := TfrmSaveDocDetails.Create(Application);
   try
      frmSaveDocDetails.ShowModal;
      Result := frmSaveDocDetails.nMatter;
   finally
      frmSaveDocDetails.Free;
   end;
end;

procedure TfrmSaveDocDetails.btnEditMatterPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
   frmMtrSearch :=TfrmMtrSearch.Create(nil);
   try
      frmMtrSearch.MakeSql;
      if (frmMtrSearch.ShowModal = mrOK) then
      begin
         btnEditMatter.Text := frmMtrSearch.vMattersFILEID.EditValue;   //  dmSaveDoc.qryMatters.FieldByName('fileid').AsString;
         nMatter := frmMtrSearch.vMattersNMATTER.EditValue;
         cmbAuthor.EditValue := TableString('MATTER','NMATTER',nMatter,'AUTHOR');
         FFileID := btnEditMatter.Text;
         cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
         cbNewCopy.Visible := ((FOldFileID <> FFileID) and (FOldFileID <> ''));
      end;
   finally
      frmMtrSearch.Free;
   end;
end;

procedure TfrmSaveDocDetails.btnCloseClick(Sender: TObject);
var
   MSWord: _Application;
  MSDoc: _Document;
  Unknown: IUnknown;
  OLEResult: HResult;
  AMacro : string;
begin
   if Systemstring('SaveAsOnCancel') = 'Y' then
   begin
      OLEResult := GetActiveObject(CLASS_WordApplication, nil, Unknown);
      if (OLEResult = MK_E_UNAVAILABLE) then
         MSWord := CoWordApplication.Create          //get MS Word running
      else
      begin
         OleCheck(OLEResult);                           //check for errors
         OleCheck(Unknown.QueryInterface(_Application, MSWord));
      end;

      if(not VarIsNull(MSWord)) then
      begin
         AMacro := 'InsightSaveAs';
         MSWord.Run(AMacro, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                 EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam );
         MSWord.Activewindow.WindowState := wdWindowStateMaximize;
      end;
   end;
//   dmSaveDoc.orsAxiom.Disconnect;
//   dmSaveDoc.Free;
   Close;
end;

procedure TfrmSaveDocDetails.btnSaveClick(Sender: TObject);
var
   DocSequence: string;
//   bUsePath: boolean;
begin
   if btnEditMatter.Text = '' then
   begin
      with Application do
      begin
         NormalizeTopMosts;
         MessageBox('Please enter a Matter number.','DocToDBSave',MB_OK+MB_ICONEXCLAMATION);
         RestoreTopMosts;
         exit;
      end;
   end;
   if btnTxtDocPath.Text <> '' then
   begin
      try
         if cmbAuthor.Text = '' then
         begin
            with Application do
            begin
               NormalizeTopMosts;
               MessageBox('Please enter an Author.','DocToDBSave',MB_OK+MB_ICONEXCLAMATION);
               RestoreTopMosts;
               exit;
            end;
         end;
         dmSaveDoc.orsAxiom.StartTransaction;
         dmSaveDoc.qryMatterAttachments.ParamByName('docid').AsString := dmSaveDoc.DocID;
         dmSaveDoc.qryMatterAttachments.Open;

         FEditing := False;
//         bUsePath := False;
         tmpdir := GetEnvironmentVariable('TMP')+'\';

         if ((cbOverwriteDoc.Visible)  and
            (not cbOverwriteDoc.Checked)) then
            dmSaveDoc.qryMatterAttachments.insert
         else
         if (not cbOverwriteDoc.Visible) then
            dmSaveDoc.qryMatterAttachments.Insert
         else
         if (cbOverwriteDoc.Checked) then
         begin
            dmSaveDoc.qryMatterAttachments.Edit;
            FEditing := True;
         end;

//         if bUsePath then
//         begin
//            tmpDir := btnTxtDocPath.Text + '\';
//         end;

//         if txtDocName.Text = '' then
//         begin
//            tmpFileName := tmpDir + dmSaveDoc.DocID +'.doc';
//         end
//         else
//         begin
            if btnTxtDocPath.Text = '' then
               tmpFileName := txtDocName.Text
            else
               tmpFileName := btnTxtDocPath.Text;

         try
            SaveDocument(DocSequence);
            dmSaveDoc.orsAxiom.Commit;
            if (rgStorage.ItemIndex = 0) and (not cbLeaveDocOpen.Checked) then
               DeleteFile(tmpFileName);
         except
            raise;
         end;
      except
         dmSaveDoc.orsAxiom.Rollback;
      end;
      Self.Close;
   end
   else
   with Application do
   begin
      NormalizeTopMosts;
      MessageBox('Please enter a document name.','DocToDBSave',MB_OK+MB_ICONEXCLAMATION);
      RestoreTopMosts;
  end;
end;

function TfrmSaveDocDetails.SaveDocument(DocSequence: string): boolean;
var
//  varWord, varDocs, PropName, varDoc: OleVariant;
//   PropName: OleVariant;
  DocName, SavedInDB: string;
  nCat, nClass: integer;
  ltmpdir, AMacro: string;
  MSWord: _Application;
  MSDoc: _Document;
  Unknown: IUnknown;
  OLEResult: HResult;
  OLEvar: OleVariant;
  CustomDocProps,
  Item,
  Value,
  DocProps,
  SaveFormsData: OleVariant;
  i, x: integer;
  ADocID, AKeyWords, APrecDetails, AExt: string;
  PropValues: TStrings;
  bMoveSuccess: boolean;
begin
   SaveDocument := False;
   bMoveSuccess := True;

   OLEResult := GetActiveObject(CLASS_WordApplication, nil, Unknown);
   if (OLEResult = MK_E_UNAVAILABLE) then
      MSWord := CoWordApplication.Create          //get MS Word running
   else
   begin
      OleCheck(OLEResult);                           //check for errors
      OleCheck(Unknown.QueryInterface(_Application, MSWord));
   end;


   if(not VarIsNull(MSWord)) then
   begin
      try
         case rgStorage.ItemIndex of
           0:  begin
                  ltmpdir := ParseMacros(tmpFileName,TableInteger('MATTER','FILEID',FFileID,'NMATTER'));
                  ltmpDir := tmpdir+ExtractFileName(ltmpdir);  // copy(ltmpDir, 1,length(ltmpdir) - 1);
//                 if not DirectoryExists(ltmpdir) then
//                    ForceDirectories(ltmpdir);

                  if ExtractFileExt(ltmpdir) = '' then
                     ltmpdir := ltmpdir + '.doc';

                  Value := ltmpdir;
                  Item := CustomDocProps.Item[7];
                  Item.Value := Value;

                  MSWord.ActiveDocument.SaveAs(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  tmpFileName := ltmpdir;
               end;
           1:  begin
                  if (FOldFileID <> FFileID) and (FOldFileID <> '') and (not cbNewCopy.Checked) then
                  begin
                     tmpFileName := SystemString('DRAG_DEFAULT_DIRECTORY');
                     tmpFileName := tmpFileName + '\' + ExtractFileName(btnTxtDocPath.Text);

                     AExt := ExtractFileExt(tmpFileName);
                     tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
                     tmpFileName := tmpFileName + '_' + '[DOCSEQUENCE]';
                     tmpFileName := tmpFileName + AExt;

                     tmpFileName := ParseMacros(tmpFileName,TableInteger('MATTER','FILEID',uppercase(FFileID),'NMATTER'));

                     if FOldFileID <> '' then
//                        bMoveSuccess := MoveMatterDoc(tmpFileName, btnTxtDocPath.Text);
                  end
                  else
                  if (not FEditing) then
                  begin
                     if btnTxtDocPath.Text = '' then
                     begin
                        btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
//                        btnTxtDocPath.Text
                     end;
                     tmpFileName := btnTxtDocPath.Text;

                     AExt := ExtractFileExt(tmpFileName);
                     tmpFileName := Copy (tmpFileName,1, Length(tmpFileName)- Length(AExt));
                     tmpFileName := tmpFileName + '_[DOCSEQUENCE]';
                     tmpFileName := tmpFileName + AExt;

                     tmpFileName := ParseMacros(tmpFileName,TableInteger('MATTER','FILEID',FFileID,'NMATTER'));
                  end
                  else
                  begin
                     tmpFileName := tmpFileName;
                  end;

                  if ExtractFileName(tmpFileName) = '' then
                     tmpFileName  := tmpFileName + FFileID;

                  if ExtractFileExt(tmpFileName) = '' then
                     tmpFileName := tmpFileName + '.' + SystemString('default_doc_ext');  //'.doc';
               {
                  if ((DocName = '') or (pos('Document', DocName) > 0) or
                     (ExtractFileName(btnTxtDocPath.Text) <> DocName)) and (not cbOverwriteDoc.Checked) then
                  begin
                   }
                     if not DirectoryExists(ExtractFileDir(tmpFileName)) then
                        ForceDirectories(ExtractFileDir(tmpFileName));
                     Value := tmpFileName;
                     SaveFormsData := True;
                     MSWord.ActiveDocument.SaveAs(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                  EmptyParam, EmptyParam, EmptyParam, SaveFormsData, EmptyParam);
          {
                  end
                  else
                  begin
                     MSWord.ActiveDocument.Save;
                  end;      }

                  AMacro := SystemString('WORD_SAVE_MACRO');
                  if AMacro <> '' then MSWord.Run(AMacro, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
                                                EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam );
               end;
         end;

//            Value := False;
//            MSWord.ActiveDocument.Close(Value,EmptyParam, EmptyParam);
            MSWord.Activewindow.WindowState := wdWindowStateMaximize;

         try
            if bMoveSuccess then
            begin
            // write doc properties
               PropValues := TStringList.Create;
               MSWord.Visible := True;
//             varDocs := varWord.Documents;
               MSDoc := MSWord.ActiveDocument;

                  DocName := MSDoc.Name;

                  if FFileID = '' then
                     FFileID := btnEditMatter.Text
                  else
                  begin
                     if (FOldFileID <> FFileID) and (FOldFileID <> '') then
//                     if FileID <> btnEditMatter.Text then
                        FFileID := btnEditMatter.Text;
                  end;
                  PropValues.Add(FFileID);

                  ADocID := dmSaveDoc.DocID;
                  PropValues.Add(ADocID);

                  if varIsNull(cmbCategory.EditValue) or
                     (VarToStr(cmbCategory.EditValue) = '') then
                     nCat := -1
                  else
                  begin
                     try
                        nCat := cmbCategory.EditValue;
                        FPrec_Category := IntToStr(nCat);
                     except
                        nCat := -1;
                     end;
                  end;
                  PropValues.Add(IntToStr(nCat));

                  if varIsNull(cmbClassification.EditValue) or
                     (VarToStr(cmbClassification.EditValue) = '') then
                     nClass := -1
                  else
                  begin
                     try
                        nClass := cmbClassification.EditValue;
                        FPrec_Classification := IntToStr(nClass);
                     except
                        nClass := -1;
                     end;
                  end;
                  PropValues.Add(IntToStr(nClass));

                  AKeyWords := edKeywords.Text;
                  PropValues.Add(AKeyWords);

                  APrecDetails := memoPrecDetails.Text;
                  PropValues.Add(APrecDetails);

                   // empty value for file name.  file name is generated and saved later
                  PropValues.Add('');

                 // add author to array
                  PropValues.Add(cmbAuthor.EditValue);

                  case rgStorage.ItemIndex of
                     0: SavedInDB := 'Y';
                     1: SavedInDB := 'N';
                  end;
                  PropValues.Add(SavedInDB);

                  // document description - title
                  PropValues.Add(txtDocName.Text);

                  if cbPortalAccess.Checked then
                     PropValues.Add('Y')
                  else
                     PropValues.Add('N');

                  CustomDocProps := MSDoc.CustomDocumentProperties;
                  DocProps := MSDoc.BuiltInDocumentProperties;

                  for x := 0 to (length(CUSTOMPROPS) - 1) do
                  begin
                     OLEvar := CUSTOMPROPS[x];
                     Value := PropValues.Strings[x];
                     try
                        for I := 1 to length(CUSTOMPROPS) {CustomDocProps.Count} do // Iterate
                        begin
                           try
                              if CustomDocProps.Count <= x then
                              begin
                                 CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                                 break;
                              end
                              else
                              begin
                                 try
                                    if i > CustomDocProps.Count then
                                    begin
                                       CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                                       break;
                                    end
                                    else
                                    begin
                                       Item := CustomDocProps.Item[i];
                                       if (Item.Name = OLEVar) then
                                       begin
                                          Item.Value := Value;
                                          break;
                                       end;
                                    end;
                                 except
                                    CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                                 end;
                              end;
                           except
                              CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, Value ,'');
                           end;
                        end; // for
                     except
                        CustomDocProps.Add(OLEVar, False, msoPropertyTypeString, PropValues.Strings[x] ,EmptyParam);
                     end;
                  end;

                  // set document title property
                  Value := txtDocName.Text;
                  Item := DocProps.Item[1];
                  Item.Value := Value;

                  // add doc name to custom properties
                  Value := tmpFileName;
                  Item := CustomDocProps.Item[7];
                  Item.Value := Value;

                  MSWord.ActiveDocument.Fields.Update;
                  MSWord.ActiveDocument.Save();
               try
                  dmSaveDoc.qryMatterAttachments.FieldByName('docid').AsString := dmSaveDoc.DocID;
                  dmSaveDoc.qryMatterAttachments.FieldByName('fileid').AsString := btnEditMatter.Text;
                  dmSaveDoc.qryMatterAttachments.FieldByName('nmatter').AsInteger := nMatter;
                  dmSaveDoc.qryMatterAttachments.FieldByName('auth1').AsString := cmbAuthor.EditValue;  //  dmSaveDoc.UserID;
                  if not FEditing then
                     dmSaveDoc.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := Now;

                  dmSaveDoc.qryMatterAttachments.FieldByName('IMAGEINDEX').AsInteger := 2;
                  if rgStorage.ItemIndex = 0 then
                     dmSaveDoc.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName)  //txtDocName.Text + '.doc'
                  else
                     dmSaveDoc.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(tmpFileName);
                  dmSaveDoc.qryMatterAttachments.FieldByName('DESCR').AsString := txtDocName.Text;   // ExtractFileName(tmpFileName);
                  dmSaveDoc.qryMatterAttachments.FieldByName('FILE_EXTENSION').AsString := Copy(ExtractFileExt(tmpFileName),2, Length(ExtractFileExt(tmpFileName)));
                  dmSaveDoc.qryMatterAttachments.FieldByName('NPRECCATEGORY').AsString := FPrec_Category;
                  dmSaveDoc.qryMatterAttachments.FieldByName('precedent_details').AsString := memoPrecDetails.Text;
                  dmSaveDoc.qryMatterAttachments.FieldByName('KEYWORDS').AsString := edKeywords.Text;
                  dmSaveDoc.qryMatterAttachments.FieldByName('NPRECCLASSIFICATION').AsString := FPrec_Classification;
                  if cbPortalAccess.Checked then
                     dmSaveDoc.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'Y'
                  else
                     dmSaveDoc.qryMatterAttachments.FieldByName('EXTERNAL_ACCESS').AsString := 'N';

                  if FEditing then
                  begin
                     dmSaveDoc.qryMatterAttachments.FieldByName('D_MODIF').AsDateTime := Now;
                     dmSaveDoc.qryMatterAttachments.FieldByName('auth2').AsString := dmSaveDoc.UserID;
                  end;
                  if rgStorage.ItemIndex = 0 then
                  begin
                    { if FEditing then
                     begin
                        dmSaveDoc.qryTmp.SQL.Text := 'select doc.document, doc.file_extension from doc where docid = :docid';
                        dmSaveDoc.qryTmp.ParamByName('docid').AsString := DocSequence;
                        dmSavedoc.qryTmp.Open;
                        bStream := dmSavedoc.qryTmp.CreateBlobStream(dmSavedoc.qryTmp.FieldByName('DOCUMENT'),bmRead);
                        try
                           if (pos('.doc', dmSaveDoc.qryTmp.FieldByName('DOC_NAME').AsString) = 0) then
                              diff_tmpFileName := GetEnvironmentVariable('TMP')+'\gat_'+ dmSaveDoc.qryTmp.FieldByName('DOC_NAME').AsString +'.'+ dmSaveDoc.qryTmp.FieldByName('file_extension').AsString
                           else
                              diff_tmpFileName := GetEnvironmentVariable('TMP')+'\gat_'+ dmSaveDoc.qryTmp.FieldByName('DOC_NAME').AsString;
                           // in case file still sitting in tmp directory
                           DeleteFile(diff_tmpFileName);

                           bStream.Seek(0, soFromBeginning);

                           with TFileStream.Create(diff_tmpFileName, fmCreate) do
                           try
                              CopyFrom(bStream, bStream.Size)
                           finally
                              Free
                           end;
                        finally
                           bStream.Free
                        end;
                        Diff := TDiff.Create(Self);
//                        Diff.Execute(tmpFileName, diff_tmpFileName )
                        TBlobField(dmSaveDoc.qryMatterAttachments.fieldByname('DOCUMENT')).Clear;
                     end;     }
                     TBlobField(dmSaveDoc.qryMatterAttachments.fieldByname('DOCUMENT')).LoadFromFile(tmpFileName);
                  end
                  else
                  begin
                     dmSaveDoc.qryMatterAttachments.FieldByName('PATH').AsString := IndexPath(tmpFileName, 'DOC_SHARE_PATH');
                     dmSaveDoc.qryMatterAttachments.FieldByName('display_PATH').AsString := tmpFileName;
                  end;

                  dmSaveDoc.qryMatterAttachments.Post;
                  dmSaveDoc.qryMatterAttachments.ApplyUpdates;
                  dmSaveDoc.orsAxiom.Commit;

               except
                  dmSaveDoc.orsAxiom.Rollback;
               end;

               SaveDocument := True;
               if (not cbLeaveDocOpen.Checked) then
               begin
                  Value := False;
                  MSWord.ActiveDocument.Close(Value,EmptyParam, EmptyParam);
//                  Value := tmpFileName;
//                  MSWord.Documents.Open(Value, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
//                                         EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
               end;
               PropValues.Free;
               MSDoc := nil;
               MSWord := nil;
               ModalResult := mrOk;
            end;
         except
           on E: Exception do
             begin
                Application.MessageBox(PChar('Error during saving document: ' + E.Message), PChar('Insight'), MB_ICONERROR);
//                MessageDlg('Error during saving document: ' + E.Message, mtError, [mbOK], 0);
                SaveDocument := False;
             end;
         end;
      except
         on E: Exception do
          begin
             Application.MessageBox(PChar('Error during saving document (trying to establish active document): ' + E.Message), PChar('Insight'), MB_ICONERROR);
//             MessageDlg('Error during saving document: ' + E.Message, mtError, [mbOK], 0);
             SaveDocument := False;
          end;
      end;
   end;
end;

procedure TfrmSaveDocDetails.rgStorageClick(Sender: TObject);
begin
   case rgStorage.ItemIndex of
      0: begin
            btnTxtDocPath.Visible := False;
            Self.Height := 275;
         end;
      1: begin
            btnTxtDocPath.Visible := True;
            Self.Height := 307;
         end;
   end;
end;

procedure TfrmSaveDocDetails.btnEditMatterPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
   if string(DisplayValue) <> '' then
   begin
      dmSaveDoc.qryGetMatter.Close;
      dmSaveDoc.qryGetMatter.ParamByName('FILEID').AsString := string(DisplayValue);
      dmSaveDoc.qryGetMatter.Open;
      if dmSavedoc.qryGetMatter.Eof then
         MessageDlg('Invalid Matter Number', mtWarning, [mbOk], 0)
      else
      begin
         nMatter := dmSaveDoc.qryGetMatter.FieldByName('NMATTER').AsInteger;
         FFileID := string(DisplayValue);
         cmbAuthor.EditValue := TableString('MATTER','NMATTER',nMatter,'AUTHOR');
         cbOverwriteDoc.Enabled := (FOldFileID = FFileID);
         cbNewCopy.Visible := ((FOldFileID <> FFileID) and (FOldFileID <> ''));
      end;
   end;
end;

procedure TfrmSaveDocDetails.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   try
      if dmSavedoc.qryPrecCategory.Active then
         dmSavedoc.qryPrecCategory.close;
      if dmSaveDoc.orsAxiom.Connected then
         dmSaveDoc.orsAxiom.Disconnect;
   finally
      dmSaveDoc.Free;
      Action := caFree;
   end;
end;

procedure TfrmSaveDocDetails.btnTxtDocPathPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
begin
   case AButtonIndex of
      0: begin
            if BrowseDlg.Execute then
               btnTxtDocPath.Text := BrowseDlg.SelectedFolder;
         end;
      1: btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
   end;
end;

procedure TfrmSaveDocDetails.cmbCategoryPropertiesInitPopup(
  Sender: TObject);
begin
//   dmSavedoc.qryPrecCategory.Close;
//   dmSavedoc.qryPrecCategory.Open;
end;

procedure TfrmSaveDocDetails.GetDetails;
var
  varWord, varDoc, PropName : OleVariant;
begin
   try
       varWord := GetActiveOleObject('Word.Application');
   except
      on EOleSysError do
      begin
         try
            varWord := CreateOleObject('Word.Application');
         except
            on e: Exception do
            begin
               MessageDlg('Error Starting MS Word: ' + E.Message, mtError, [mbOK], 0);
               varWord := null;
            end;
         end;
      end;
   end; 

   if(not VarIsNull(varWord)) then
   begin
      try
         PropName := 'MatterNo';
         varDoc := varWord.ActiveDocument;
         FFileID := varDoc.CustomDocumentProperties[PropName].Value;
         FOldFileID := FFileID;
         btnEditMatter.EditValue := FFileID;
         btnEditMatter.Text := FFileID;
         nMatter := TableInteger('MATTER','FILEID',FFileID,'NMATTER');
         if btnEditMatter.Text <> '' then
            btnEditMatter.ValidateEdit(False);

         PropName := 'DocID';
         dmSaveDoc.DocID := varDoc.CustomDocumentProperties[PropName].Value;
//         application.MessageBox(pchar(FDocID),'help',MB_OK);
         FDocName := TableString('DOC','DOCID', dmSaveDoc.DocID, 'DOC_NAME');
         if FDocName = '' then
            FDocName := varWord.ActiveDocument.Name;

         cbOverWriteDoc.Visible := True;
         PropName := 'Prec_Category';
         try
            FPrec_Category := varDoc.CustomDocumentProperties[PropName].Value;
            cmbCategory.EditValue := FPrec_Category;
         except
            ;// in case of errors
         end;

         PropName := 'Prec_Classification';
         try
            FPrec_Classification := varDoc.CustomDocumentProperties[PropName].Value;
            cmbClassification.EditValue := FPrec_Classification;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Keywords';
         try
            FDoc_Keywords := varDoc.CustomDocumentProperties[PropName].Value;
            edKeywords.Text := FDoc_Keywords;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Precedent';
         try
            FDoc_Precedent := varDoc.CustomDocumentProperties[PropName].Value;
            memoPrecDetails.Text := FDoc_Precedent;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_FileName';
         try
            FDoc_FileName := varDoc.CustomDocumentProperties[PropName].Value;
            btnTxtDocPath.Text := FDoc_FileName;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Title';
         try
            TxtDocName.Text := varDoc.CustomDocumentProperties[PropName].Value;
         except
            ;// in case it doesnt exist
         end;

         PropName := 'Doc_Author';
         try
            FDoc_Author := varDoc.CustomDocumentProperties[PropName].Value;
            cmbAuthor.EditValue := FDoc_Author;
         except
            cmbAuthor.EditValue := dmSaveDoc.UserID;// in case it doesnt exist
         end;

         PropName := 'Saved_in_DB';
         FSavedInDB := varDoc.CustomDocumentProperties[PropName].Value;
         if FSavedInDB = 'Y' then
         begin
            rgStorage.ItemIndex := 0;
            btnTxtDocPath.Text := FDocName;
         end;
//         varWord.ActiveDocument.BuiltinDocumentProperties('Category') := IntToStr(nMatter);
         if txtDocName.Text = '' then
            txtDocName.Text := TableString('DOC','DOCID', dmSaveDoc.DocID, 'DESCR'); //  DocName;

         PropName := 'Portal_Access';
         cbPortalAccess.Checked := (varDoc.CustomDocumentProperties[PropName].Value = 'Y');

      except
         // in case of errors
      end;
   end;
end;

procedure TfrmSaveDocDetails.FormShow(Sender: TObject);
begin
//   Application.CreateForm(TdmSaveDoc, dmSaveDoc);
   try
      GetUserID;
      cbOverWriteDoc.Visible := False;
      if (FSavedInDB = 'N') or (FSavedInDB = '')  then
      begin
         rgStorage.ItemIndex := SystemInteger('DFLT_DOC_SAVE_OPTION');
//         btnTxtDocPath.Text := SystemString('DOC_DEFAULT_DIRECTORY');
      end;
      try
         GetDetails;
      except
         //
      end;
      dmSaveDoc.qryPrecCategory.Open;
      dmSaveDoc.qryPrecClassification.Open;
      dmSaveDoc.qryEmployee.Open;
//      dmSaveDoc.qryMatters.Active := True;
      StatusBar.Panels[0].Text := 'Ver: '+ReportVersion + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(Application.ExeName)))+')';
      rgStorage.Enabled := (SystemString('DISABLE_SAVE_MODE') = 'N');
   except
      Application.Terminate;
   end;
end;

end.
