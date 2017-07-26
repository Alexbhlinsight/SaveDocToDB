object dmSaveDoc: TdmSaveDoc
  OldCreateOrder = False
  Left = 659
  Top = 313
  Height = 437
  Width = 361
  object orsAxiom: TOraSession
    Options.Direct = True
    Username = 'axiom'
    Password = 'axiom'
    Server = '192.168.100.22:1521:marketng'
    LoginPrompt = False
    OnError = orsAxiomError
    HomeName = 'OraDb10g_home1'
    Left = 25
    Top = 8
  end
  object qryEmps: TOraQuery
    Session = orsAxiom
    Left = 164
    Top = 9
  end
  object qryMatters: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'select * from matter'
      'where closed = 0 and entity = :P_Entity')
    Left = 32
    Top = 65
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'P_Entity'
        Value = Null
      end>
  end
  object dsMatters: TOraDataSource
    DataSet = qryMatters
    Left = 99
    Top = 71
  end
  object qryGetSeq: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'select DOC_DOCID.nextval as nextdoc from dual')
    Left = 80
    Top = 16
  end
  object qryMatterAttachments: TSmartQuery
    Session = orsAxiom
    SQL.Strings = (
      'SELECT'
      '  DOC.DOCUMENT,'
      '  DOC.IMAGEINDEX,'
      '  DOC.FILE_EXTENSION,'
      '  DOC.DOC_NAME,'
      '  DOC.SEARCH,'
      '  DOC.DOC_CODE,'
      '  DOC.JURIS,'
      '  DOC.D_CREATE,'
      '  DOC.AUTH1,'
      '  DOC.D_MODIF,'
      '  DOC.AUTH2,'
      '  DOC.PATH,'
      '  DOC.DESCR,'
      '  DOC.FILEID,'
      '  DOC.DOCID,'
      '  DOC.NPRECCATEGORY,'
      '  DOC.NMATTER,'
      '  DOC.PRECEDENT_DETAILS,'
      '  DOC.NPRECCLASSIFICATION,'
      '  DOC.KEYWORDS,'
      '  DOC.DISPLAY_PATH,'
      '  DOC.EXTERNAL_ACCESS,'
      '  DOC.ROWID'
      'FROM'
      '  DOC'
      'where'
      '  DOCID = :DOCID')
    CachedUpdates = True
    OnNewRecord = qryMatterAttachmentsNewRecord
    Left = 34
    Top = 119
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'DOCID'
      end>
    object qryMatterAttachmentsDOCUMENT: TBlobField
      FieldName = 'DOCUMENT'
      BlobType = ftOraBlob
    end
    object qryMatterAttachmentsIMAGEINDEX: TFloatField
      FieldName = 'IMAGEINDEX'
    end
    object qryMatterAttachmentsFILE_EXTENSION: TStringField
      FieldName = 'FILE_EXTENSION'
      Size = 5
    end
    object qryMatterAttachmentsDOC_NAME: TStringField
      FieldName = 'DOC_NAME'
      Size = 60
    end
    object qryMatterAttachmentsSEARCH: TStringField
      FieldName = 'SEARCH'
      Size = 85
    end
    object qryMatterAttachmentsDOC_CODE: TStringField
      FieldName = 'DOC_CODE'
      Size = 50
    end
    object qryMatterAttachmentsJURIS: TStringField
      FieldName = 'JURIS'
      Size = 50
    end
    object qryMatterAttachmentsD_CREATE: TDateTimeField
      FieldName = 'D_CREATE'
    end
    object qryMatterAttachmentsAUTH1: TStringField
      FieldName = 'AUTH1'
      Size = 3
    end
    object qryMatterAttachmentsD_MODIF: TDateTimeField
      FieldName = 'D_MODIF'
    end
    object qryMatterAttachmentsAUTH2: TStringField
      FieldName = 'AUTH2'
      Size = 3
    end
    object qryMatterAttachmentsPATH: TStringField
      FieldName = 'PATH'
      Size = 255
    end
    object qryMatterAttachmentsDESCR: TStringField
      FieldName = 'DESCR'
      Size = 400
    end
    object qryMatterAttachmentsFILEID: TStringField
      FieldName = 'FILEID'
    end
    object qryMatterAttachmentsDOCID: TFloatField
      FieldName = 'DOCID'
      Required = True
    end
    object qryMatterAttachmentsNPRECCATEGORY: TFloatField
      FieldName = 'NPRECCATEGORY'
    end
    object qryMatterAttachmentsNMATTER: TFloatField
      FieldName = 'NMATTER'
    end
    object qryMatterAttachmentsPRECEDENT_DETAILS: TStringField
      FieldName = 'PRECEDENT_DETAILS'
      Size = 2048
    end
    object qryMatterAttachmentsNPRECCLASSIFICATION: TFloatField
      FieldName = 'NPRECCLASSIFICATION'
    end
    object qryMatterAttachmentsKEYWORDS: TStringField
      FieldName = 'KEYWORDS'
      Size = 2048
    end
    object qryMatterAttachmentsDISPLAY_PATH: TStringField
      FieldName = 'DISPLAY_PATH'
      Size = 255
    end
    object qryMatterAttachmentsEXTERNAL_ACCESS: TStringField
      FieldName = 'EXTERNAL_ACCESS'
      Size = 1
    end
    object qryMatterAttachmentsROWID: TStringField
      FieldName = 'ROWID'
      ReadOnly = True
      Size = 18
    end
  end
  object qryGetMatter: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'select fileid, nmatter '
      'from'
      'matter'
      'where'
      'fileid = :fileid')
    Left = 181
    Top = 67
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'fileid'
      end>
  end
  object qryGetEntity: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'SELECT VALUE,INTVALUE'
      'FROM SETTINGS '
      'WHERE EMP = :Emp'
      '  AND OWNER = :Owner'
      '  AND ITEM = :Item')
    Left = 25
    Top = 183
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Emp'
      end
      item
        DataType = ftUnknown
        Name = 'Owner'
      end
      item
        DataType = ftUnknown
        Name = 'Item'
      end>
  end
  object qryPrecCategory: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'select * from PRECCATEGORY')
    Left = 115
    Top = 135
  end
  object dsPrecCategory: TOraDataSource
    DataSet = qryPrecCategory
    Left = 191
    Top = 128
  end
  object qryTmp: TOraQuery
    Session = orsAxiom
    Left = 184
    Top = 186
  end
  object qrySysFile: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'SELECT * FROM SYSTEMFILE')
    Left = 116
    Top = 188
  end
  object procTemp: TOraStoredProc
    Session = orsAxiom
    Left = 28
    Top = 239
  end
  object qryPrecClassification: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'select * from PRECCLASSIFICATION')
    Left = 116
    Top = 249
  end
  object dsPrecClassification: TOraDataSource
    DataSet = qryPrecClassification
    Left = 223
    Top = 242
  end
  object dsEmployee: TOraDataSource
    DataSet = qryEmployee
    Left = 246
    Top = 159
  end
  object qryEmployee: TOraQuery
    Session = orsAxiom
    SQL.Strings = (
      'select code, name from employee where active = '#39'Y'#39' order by code')
    Left = 247
    Top = 95
  end
end
