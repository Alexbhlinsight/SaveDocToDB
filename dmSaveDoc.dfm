object DataModule2: TDataModule2
  OldCreateOrder = False
  Left = 500
  Top = 411
  Height = 148
  Width = 321
  object orsAxiom: TOraSession
    Options.KeepDesignConnected = False
    Username = 'axiom'
    Password = 'axiom'
    Server = 'AXIOMNW'
    LoginPrompt = False
    Left = 29
    Top = 9
  end
  object OraQuery1: TOraQuery
    Session = orsAxiom
    Left = 61
    Top = 57
  end
  object OraDataSource1: TOraDataSource
    Left = 144
    Top = 22
  end
end
