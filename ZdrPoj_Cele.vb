'Option Explicit
Dim DbUdajeR As Database
Dim PathDtaR As String
Dim DbUdjR As String
Dim TestnaZdr205 As Integer
Dim FNameT As String
Dim F As Integer
Dim X As Integer
Dim Kontrolor As Integer
Dim PocSestav As Integer
Dim Zac As Integer
Dim ZdrKodHl As String
Dim FirmaNazev As String
Dim FirmaDIC As String
Dim RsPojjSum As Recordset
Dim MaxRsPojjSum As Integer
Dim FirmaIco As String
Dim FirmaUlice As String
Dim FirmaAdrCislo As String
Dim FirmaPSC As String
Dim FirmaAdrObec As String
Dim FirmaUcet As String
Dim FirmaTelefon As String
Dim MzdUct As String
Dim NactiFormular As String
Dim NactiPublic As String
Dim NazevSestavy As String
Dim Dohoda As String
Dim Tisk As String
Dim KodZdrPoj As String
Dim ZamTypP  As String
Dim ZPocZam(30) As String
Dim ZVymZakl(30) As String
Dim ZPoj(30) As String
Dim Dulezite As Integer
Dim Sum272(20) As Double
Dim Sum471(20) As Double
Dim Sum472(20) As Double
Dim Sum476(20) As Double
Dim Sum274(20) As Double
Dim Sum277(20) As Double
Dim Sum354(20) As Double
Dim KodPojjMesic(20) As String
'Dim PojNazevZkratka(20) As String
Dim PocLidi(20) As Integer
Dim RsPoj As Recordset
Dim ZdrNazev(30) As String
Dim ZdrZkratka(30) As String
Dim ZdrUlice(30) As String
Dim ZdrCisloDomu(30) As String
Dim ZdrObec(30) As String
Dim ZdrPsc(30) As String
Dim ZdrTelefon(30) As String
Dim ZdrCisloUctu(30) As String
Dim ZdrKodBanky(30) As String
Dim ZdrKs(30) As String
Dim DohodaCin As String
Dim ZdrVs(30) As String
Dim ZdrSs(30) As String
Dim ZdrKod(30) As String
Dim ZdrNazevTisk(30) As String
Dim ZdrZkratkaTisk(30) As String
Dim ZdrUliceTisk(30) As String
Dim ZdrCisloDomuTisk(30) As String
Dim ZdrObecTisk(30) As String
Dim ZdrPscTisk(30) As String
Dim ZdrTelefonTisk(30) As String
Dim ZdrCisloUctuTisk(30) As String
Dim ZdrKodBankyTisk(30) As String
Dim ZdrKsTisk(30) As String
Dim ZdrVsTisk(30) As String
Dim ZdrSsTisk(30) As String
Dim ZdrKodTisk(30) As String
Dim ZdrKodXML(30) As String
Dim ZdrZkratkaXML(30) As String
Dim PocPodrizena As Integer
Dim RSRidici As Recordset
Dim MaxRsRidici As Integer
Dim PocPrvku As Integer
Dim MaxPocPodrizena As Integer
Dim maxRsZamSt As Integer
Dim MaxZaznam As Integer
Dim RidiciVyber As String
Dim PodrizenaPol As Integer
Dim PodrizenaVyber(200) As String
Dim PocStr As Integer
Dim PocRad As Integer
Dim RsZamSt As Recordset
Dim i As Integer
Dim a As Integer
Dim c As Integer
Dim Cmd As Integer
Dim RsVyber As Recordset
Dim MaxRsVyber As Integer
Dim TrVyberVse As String
Dim RidiciPol As Integer
Dim OsCisloZ As String
Dim Sestava As String
Dim veta As String
Dim RsZamS As Recordset
Dim MaxZamS As Integer
Dim Strana As Integer
Dim ZamHM  As String
Dim ZamZaklad  As String
Dim ZamVratka  As String
Dim NazevSestavyHl As String
Dim Ucetni As String
Dim OldMes As Integer
Dim OldRok As Integer
Dim OldRokOd As Integer
Dim OldRokDo As Integer
Dim OdMes As Integer
Dim DoMes As Integer
Dim KodZ(100) As String
Dim UliceZam(100) As String
Dim CpZam(100) As String
Dim PSCZam(100) As String
Dim ObecZam(100) As String
Dim rcZam(100) As String
Dim DatZm(100) As String
Dim PrijmeniZam(100) As String
Dim JmenoZam(100) As String
Dim OrigMes As Integer
Dim OrigRok As Integer
Dim PocRadek As Integer
Dim PocZamMala As Integer
Dim MaxrsMM As Integer
Dim ZamVeta(560) As String
Dim VymZaklZam As Single
Dim SumaVymZakl As Double
Dim PrumPocZamSoc As Single
Dim rc As String
Dim ZamVetaMala(50) As String
Dim JmenoZamestnance As String
Dim DatZmeny As String
Dim DruhCin As String
Dim Omluveno As Integer
Dim Vylouc As Integer
Dim UhrnPojC As Single
Dim UhrnPojOrg As Single
Dim UhrnPojZam As Single
Dim UhrnPocetZam As Single
Dim UhrnVymZakl As Double
Dim UhrnVratka As Double
Dim C202 As Double
Dim C272 As Double
Dim C476 As Double
Dim C499 As Single
Dim C9095 As Single
Dim C277 As Single
Dim C274 As Single
Dim C274277 As Single
Dim COON As Single
Dim C195 As Single
Dim C230 As Single
Dim C409 As Single
Dim C381 As Single
Dim C291 As Single
Dim DohC As Single
Dim Doh As Single
Dim C201 As Single
Dim C354 As Single
Dim CAno As Single
Dim CNe As Single
Dim PocZamZP As Single
Dim T(500) As String
Dim PocDil As Single
Dim PrgDil As Single
Dim PocDil1 As Single
Dim PrgDil1 As Single

Private Sub NactiZdrUdaje()

Dim DMes As Integer
Dim DRok As Integer
Dim Zam As Integer
Dim Dorov As Integer
Dim OON As Single
Dim Castka As Single
Dim Mpol As Integer
Dim PocTrid As Integer
Dim OldPoj As String
Dim Neuv As String
Dim DPS As String

If Dulezite = 0 Then: ProgressBar1.Visible = True


For OldMes = OdMes To DoMes ' OdKdy  To  DoKdy  - čísla měsíců
If Sestava <> "Zdr.poj. - součtová rekapitulace" Then
 SumaVymZakl = 0: PrumPocZamSoc = 0: UhrnPojC = 0: UhrnPocetZam = 0: UhrnVymZakl = 0: UhrnVratka = 0: PocRadek = 0:
 UhrnPojZam = 0: UhrnPojOrg = 0:
 C202 = 0: C272 = 0: C476 = 0: C499 = 0: C9095 = 0: C277 = 0: C274 = 0: C274277 = 0:  CAno = 0: CNe = 0: PocZamZP = 0
 C476 = 0: COON = 0: C195 = 0: C230 = 0: C409 = 0: C381 = 0: DohC = 0: C201 = 0: C354 = 0: C291 = 0
End If

Command16.Visible = True
Command16.Refresh

For i = 2 To 95: T(i) = "": Next i

If Kontrolor = 0 Then: T(1) = ""
   If RidiciPol = 2 Then: RidiciPol = 0: PodrizenaVyber(PocPodrizena) = " "
 
 
   a = TridZam(RidiciPol, RidiciVyber, PodrizenaPol, Trim(PodrizenaVyber(PocPodrizena)), OldMes, OldRok)

Command16.Visible = False
   If MaxRsTrideni > 0 Then
    RsTrideni.MoveFirst
    
    If Dulezite = 0 Then: PrgDil = 0: PocDil = Fix(10000 / MaxRsTrideni)
    
    If MaxRsTrideni = 0 Then: GoTo Vypadni
    
     For PocTrid = 1 To MaxRsTrideni
      OsCisloZ = RsTrideni.Fields("TrOsCislo")
'MsgBox (PocTrid & "  " & MaxRsTrideni)
       a = NahrajMzd(OsCisloZ, OldMes, OldRok)
        
        JmenoZamestnance = RsTrideni.Fields("TrPrijmeni") + " " + RsTrideni.Fields("TrJmeno")
        JmenoZamestnance = Mid$(JmenoZamestnance, 1, 24)
             

     If NahrajOsPam(OsCisloZ, OldMes, OldRok) = 1 Then
        For i = 1 To PocPrvku + 1
       
          If ZdrZkratka(i) = OldP(49) Then
                ZdrNazevTisk(PocPodrizena) = ZdrNazev(i)
                ZdrZkratkaTisk(PocPodrizena) = ZdrZkratka(i)
                ZdrUliceTisk(PocPodrizena) = ZdrUlice(i)
                ZdrObecTisk(PocPodrizena) = ZdrObec(i)
                ZdrCisloDomuTisk(PocPodrizena) = ZdrCisloDomu(i)
                ZdrPscTisk(PocPodrizena) = ZdrPsc(i)
                ZdrTelefonTisk(PocPodrizena) = ZdrTelefon(i)
                ZdrCisloUctuTisk(PocPodrizena) = ZdrCisloUctu(i)
                ZdrKodBankyTisk(PocPodrizena) = ZdrKodBanky(i)
                ZdrKsTisk(PocPodrizena) = ZdrKs(i)
                ZdrVsTisk(PocPodrizena) = ZdrVs(i)
                ZdrSsTisk(PocPodrizena) = ZdrSs(i)
                ZdrKodTisk(PocPodrizena) = ZdrKod(i)
               
              Exit For
          End If
       Next i
     End If
       
       If NahrajOsPam(OsCisloZ, OldMes, OldRok) = 1 Then
        If Sestava = "Hlášení pro zdr. poj." Then
         KodyZdrPoj
         rc = Trim(OldO(100))
         KodZdrPoj = Trim(OldP(55))
        
        If KodZdrPoj <> "" Then
                                               ' Přehled jednotlivců
             PocRadek = PocRadek + 3
              ZamVeta(PocRadek - 2) = KodZdrPoj + " " + rc
              ZamVeta(PocRadek - 1) = RsTrideni.Fields("TrPrijmeni") + Space(25 - Len(RsTrideni.Fields("TrPrijmeni")))
              ZamVeta(PocRadek - 1) = ZamVeta(PocRadek - 1) + RsTrideni.Fields("TrJmeno") + Space(25 - Len(RsTrideni.Fields("TrJmeno")))
              
             If Trim(OldO(158)) <> "" Then
              ZamVeta(PocRadek) = OldO(157)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + " " + OldO(164)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + " " + OldO(158)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + " " + OldO(159)
             
             Else
              ZamVeta(PocRadek) = OldO(150)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + " " + OldO(156)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + " " + OldO(151)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + " " + OldO(152)
            End If
            If PocRadek > 52 Then:  NactiZdrHlaseni: PocRadek = 0: Strana = Strana + 1
         End If
       End If
        ZamTypP = Mid$(OldP(51), 1, 12)
         'MsgBox (ZamTypP)
       Dohoda = "": Neuv = "": Uvol = ""
       If Mid$(OldO(320), 1, 6) = "Dohoda" Then: Dohoda = "Dohoda"
       If Mid$(OldO(320), 1, 12) = "Dohoda o čin" Then: DohodaCin = "DohodaCin"
       If Mid$(OldO(320), 1, 6) = "Neuvol" Then: Neuv = "Neuvol"
       If Mid$(OldO(320), 1, 4) = "Uvol" Then: Uvol = "Uvol"
       If Mid$(OldO(320), 1, 6) = "Souběž" Then: Dohoda = "Soub"
       If Mid$(OldP(39), 1, 5) = "Jiným" Then: Dohoda = "Soub"
       If Mid$(OldP(46), 1, 8) = "Nedaněno" Then: Dohoda = "Soub"
       If Mid$(OldO(269), 1, 1) = "O" Then: Dohoda = "Družst"
       Dorov = 0
       If Mid$(OldP(51), 1, 5) = "Dorov" Then: Dorov = 1
      
     End If
      
      
     If NahrajMzdOld(OsCisloZ, OldMes, OldRok) = 1 Then                                        ' načti MZDY
              If Dohoda = "Družst" And OldM(272) = 0 Then: Dohoda = "Soub"
              
                  If Dohoda <> "Soub" Then
                     If OldM(492) > 0 Or OldM(493) > 0 Or OldM(77) > 0 Or OldM(616) > 0 Or DohodaCin = "DohodaCin" Or Uvol = "Uvol" Or Mid$(OldP(51), 1, 5) <> "Nepla" Then: UhrnPocetZam = UhrnPocetZam + 1
                    If ZamTypP = "Činnost malé" Then
                     
                      If OldM(202) <= 2000 Then: UhrnPocetZam = UhrnPocetZam - 1
                    End If
                  End If
                   
                  '  MsgBox (RsTrideni.Fields("TrPrijmeni") & " x " & UhrnPocetZam)
                   'If UhrnPocetZam <> 0 Then
                    SumaVymZakl = SumaVymZakl + OldM(272)
                    VymZaklZam = OldM(272)
                    UhrnVymZakl = UhrnVymZakl + OldM(272)
                    UhrnVratka = UhrnVratka + OldM(476)
                    UhrnPojZam = UhrnPojZam + OldM(277)
                    UhrnPojOrg = UhrnPojOrg + OldM(274)
                    'MsgBox (UhrnPojOrg & "  " & UhrnPojZam)
                   'End If
                    
                    'Omluveno = Int(AbsencePocet(OldKalVeta(1), "X") / 2)
                    'Omluveno = Omluveno  + Int(AbsencePocet(OldKalVeta(1), "I") / 2) + Int(AbsencePocet(OldKalVeta(1), "Y") / 2)
 
           
       If Sestava = "Zdr.poj. - zaměstnanci" Then                                                   ' Přehled jednotlivců
         '  Command7.Caption = Str(PocRadek)
           
           If PocRadek / 70 = Int(PocRadek / 70) Then
 
               Strana = Strana + 1:
                 If PocRad > 0 Then
                        PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(212) + String$(9, "-") + Chr(247) + String$(25, "-") + Chr(247) + String$(11, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(10, "-") + Chr(182) + String$(10, "-") + Chr(182) + String$(6, "-") + Chr(166)
                        If Len(Trim(Str(C202))) > 7 And Len(Trim(Str(C476))) < 8 Then
                                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + "SUMA " + Space(42) + Chr(124) + Trim(Str(C202)) + Chr(124) + Chr(124) + Space(6 - Len(Str(C499))) + Str(C499) + Chr(124) + Space(7 - Len(Str(C9095))) + Str(C9095) + Chr(124) + Space(7 - Len(Str(C277))) + Str(C277) + Chr(124) + Space(10 - Len(Str(C274))) + Str(C274) + Chr(124) + Space(10 - Len(Str(C274277))) + Str(C274277) + Chr(124) + Space(6 - Len(Str(PocZamZP))) + Str(PocZamZP) + Chr(124)
                        Else
                       '     If Len(Trim(Str(C476))) = 8 Then
                       '             PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + "SUMA " + Space(42) + Chr(124) + Space(20 - Len(Str(C202))) + Str(C202) + Chr(124) + Space(8 - Len(Str(C272))) + Str(C272) + Chr(124) + Trim(Str(C476)) + Chr(124) + Space(6 - Len(Str(C499))) + Str(C499) + Chr(124) + Space(7 - Len(Str(C9095))) + Str(C9095) + Chr(124) + Space(7 - Len(Str(C277))) + Str(C277) + Chr(124) + Space(10 - Len(Str(C274))) + Str(C274) + Chr(124) + Space(10 - Len(Str(C274277))) + Str(C274277) + Chr(124) + Space(6 - Len(Str(PocZamZP))) + Str(PocZamZP) + Chr(124)
                       '     Else
                       '             PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + "SUMA " + Space(42) + Chr(124) + Space(20 - Len(Str(C202))) + Str(C202) + Chr(124) + Space(8 - Len(Str(C272))) + Str(C272) + Chr(124) + Space(8 - Len(Trim(Str(C476)))) + Trim(Str(C476)) + Chr(124) + Space(6 - Len(Str(C499))) + Str(C499) + Chr(124) + Space(7 - Len(Str(C9095))) + Str(C9095) + Chr(124) + Space(7 - Len(Str(C277))) + Str(C277) + Chr(124) + Space(10 - Len(Str(C274))) + Str(C274) + Chr(124) + Space(10 - Len(Str(C274277))) + Str(C274277) + Chr(124) + Space(6 - Len(Str(PocZamZP))) + Str(PocZamZP) + Chr(124)
    
                        '    End If
                         End If
                     PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(212) + String$(47, "-") + Chr(247) + String$(8, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(10, "-") + Chr(247) + String$(10, "-") + Chr(247) + String$(6, "-") + Chr(166)
                     PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(199) + String$(56, "-") + Chr(91) + "Salary s.r.o.  (Licence - " + Space(20 - Len(LicenceCis)) + LicenceCis + ")" + Chr(93) + String$(3, "-") + Chr(206)
                 End If
                If Strana > 1 Then: PocRadek = PocRadek + 4
                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = "měsíc: " + mesicStr(OldMes) + " " + Str(OldRok) + Space(75) + " strana - " + Str(Strana)
                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = "Přehled jednotlivých zaměstnanců zdravotní pojištovny :" + " " + RidiciVyber + " / " + Trim(PodrizenaVyber(PocPodrizena))
                If Trim(PodrizenaVyber(PocPodrizena)) <> "" Then
                    PocRadek = PocRadek + 1: ZamVeta(PocRadek) = ZdrNazevTisk(PocPodrizena)
                    PocRadek = PocRadek + 1: ZamVeta(PocRadek) = ZdrUliceTisk(PocPodrizena) + ", " + ZdrCisloDomuTisk(PocPodrizena) + ", " + ZdrPscTisk(PocPodrizena) + " " + ZdrObecTisk(PocPodrizena)
                End If
                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(194) + String$(9, "-") + Chr(172) + String$(25, "-") + Chr(172) + String$(8, "-") + Chr(172) + String$(8, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(10, "-") + Chr(172) + String$(10, "-") + Chr(172) + String$(6, "-") + Chr(184)
                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + " Osobní  " + Chr(124) + "Jméno zaměstnance        " + Chr(124) + "  hrubá " + Chr(124) + "VyměřZDR" + Chr(124) + "Neplac." + Chr(124) + "Zaměst." + Chr(124) + "Organizace" + Chr(124) + "Zam.+ Org." + Chr(124) + "Počet " + Chr(124)
                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + " číslo   " + Chr(124) + "                         " + Chr(124) + " základ " + Chr(124) + " +vratka" + Chr(124) + " volno " + Chr(124) + " poj.  " + Chr(124) + "   poj.   " + Chr(124) + "   poj.   " + Chr(124) + "pojišt" + Chr(124)
                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(212) + String$(9, "-") + Chr(182) + String$(25, "-") + Chr(182) + String$(8, "-") + Chr(182) + String$(8, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(10, "-") + Chr(182) + String$(10, "-") + Chr(182) + String$(6, "-") + Chr(166)
           End If
             PocRadek = PocRadek + 1
            
              ZamVeta(PocRadek) = Chr(124) + Space(9 - Len(OsCisloZ)) + OsCisloZ
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + JmenoZamestnance + Space(25 - Len(JmenoZamestnance))
            '  ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + OldO(100) + Space(11 - Len(OldO(100)))
              If Len(Str(OldM(202))) > 7 Then
                ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Trim(Str(OldM(202))): C202 = C202 + OldM(202)
              Else
                ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(8 - Len(Str(OldM(202)))) + Str(OldM(202)): C202 = C202 + OldM(202)
                
              End If
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(8 - Len(Str(OldM(272) + OldM(476)))) + Str(OldM(272)): C272 = C272 + OldM(272)
    '         If Len(Trim(Str(OldM(476)))) = 8 Then
    '           ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Trim(Str(OldM(476))): C476 = C476 + OldM(476)
    '         Else
    '           ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(8 - Len(Trim(Str(OldM(476))))) + Trim(Str(OldM(476))): C476 = C476 + OldM(476)
    '         End If
'             MsgBox (VymZaklZam & "sss" & OldM(499))
         '     ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(6 - Len(Str(Round(OldM(499), 3)))) + Str(Round(OldM(499), 3)): C499 = C499 + OldM(499)
              
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(OldM(90) + OldM(95)))) + Str(OldM(90) + OldM(95)): C9095 = C9095 + OldM(90) + OldM(95)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(OldM(277)))) + Str(OldM(277)): C277 = C277 + OldM(277)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(10 - Len(Str(OldM(274)))) + Str(OldM(274)): C274 = C274 + OldM(274)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(10 - Len(Str(OldM(274) + OldM(277)))) + Str(OldM(274) + OldM(277)): C274277 = C274277 + OldM(274) + OldM(277)
              If Trim(OldP(46)) <> "Nedaněno" And Trim(OldP(117)) <> "Ne" And OldM(491) + OldM(493) > 0 Then
               i = 0
               PocZamZP = PocZamZP + 1
               If Trim(OldP(51)) = "Činnost malého rozsahu" And OldM(272) <= 2500 Then
                PocZamZP = PocZamZP - 1
                 ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + " ČMR  " + Chr(124)
               Else
                If OldO(320) = "Dohoda o prov.p" And OldM(272) = 0 Then
                  PocZamZP = PocZamZP - 1
                  ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + " DPP  " + Chr(124)
                Else
                  ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + "    1 " + Chr(124)
                End If
               End If
               
              Else
              DPS = "      " 'Mid(OldO(320) + "      ", 1, 6)
             
              If OldO(320) = "Dohoda o prov.p" Then: DPS = " DPP  "
              If OldO(320) = "Dohoda o čin." Then: DPS = " DPČ  "
              If OldO(320) = "Souběžný" Then: DPS = "Souběh"
            '   If Trim(OldP(46)) <> "Nedaněno" And Trim(OldP(117)) = "Ne" And Trim(DPS) = "" And OldM(491) + OldM(493) > 0 Then: PocZamZP = PocZamZP + 1: DPS = "    1 opravit v Pam P117 na Ano"
              
                      ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + DPS + Chr(124)
  
              
              End If
  
'              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(13 - Len(ZamTypP)) + ZamTypP + Chr(124)
            'If PocRadek > 70 Then:  NactiZdrZam: PocRadek = 0: Strana = Strana + 1: For I = 1 To 85: T(I) = "": Next I
       End If
              
             
             
           If Sestava = "Činnost malého rozsahu" And ZamTypP = "Činnost malé" Then
            PocRadek = PocRadek + 1
            
              ZamVeta(PocRadek) = Chr(124) + Space(9 - Len(OsCisloZ)) + OsCisloZ
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + JmenoZamestnance + Space(25 - Len(JmenoZamestnance))
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(OldM(202)))) + Str(OldM(202)) + Chr(124) + Space(7 - Len(Str(OldM(272)))) + Str(OldM(272)): C202 = C202 + OldM(202): C272 = C272 + OldM(272)
              If OldM(202) > 2000 Then
               ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + "     NE         " + Chr(124)
               CNe = CNe + 1
              Else
               ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + "     ANO        " + Chr(124)
               CAno = CAno + 1
              End If
               If PocRadek > 75 Then:  NactiZdrCMR: PocRadek = 0: Strana = Strana + 1: For i = 1 To 85: T(i) = "": Next i
           End If
               
             Castka = 0
           If Sestava = "Zdr.poj. - rozdíl (HM <> Vym. základ)" And VymZaklZam <> OldM(202) Then                        ' Rozdílná hrubá mzda od vyměřovcího základu  zdr. poj. hrubky a vymzdr
             PocRadek = PocRadek + 1
             Doh = OldM(172) + OldM(203) + OldM(174) + OldM(214) + (OldM(202) * Abs(Dohoda = "Dohoda"))
              ZamVeta(PocRadek) = Chr(124) + Space(9 - Len(OsCisloZ)) + OsCisloZ
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + JmenoZamestnance + Space(25 - Len(JmenoZamestnance))
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(OldM(202)))) + Str(OldM(202)): C202 = C202 + OldM(202)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(OldM(272)))) + Str(OldM(272)): C272 = C272 + OldM(272)
              If Dorov = 1 Then: Castka = OldM(272) - OldM(202)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(6 - Len(Str(OldM(476) + Castka))) + Str(OldM(476) + Castka): C476 = C476 + OldM(476) + Castka
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(6 - Len(Str(OldM(291) + OldM(292)))) + Str(OldM(291) + OldM(292)): C291 = C291 + OldM(291) + OldM(292)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(6 - Len(Str(OldM(195)))) + Str(OldM(195)): C195 = C195 + OldM(195)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(6 - Len(Str(OldM(230)))) + Str(OldM(230)): C230 = C230 + OldM(230)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(OldM(409) + OldM(226)))) + Str(OldM(409) + OldM(226)): C409 = C409 + OldM(409) + OldM(226)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(6 - Len(Str(OldM(381)))) + Str(OldM(381)): C381 = C381 + OldM(381)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(8 - Len(Str(OldM(354)))) + Str(OldM(354)): C354 = C354 + OldM(354)
              ZamVeta(PocRadek) = ZamVeta(PocRadek) + Chr(124) + Space(7 - Len(Str(Doh))) + Str(Doh) + Chr(124): DohC = DohC + Doh
            
             ' ZamVeta(PocRadek) = ZamVeta(PocRadek) + Space(6 - Len(Str(OldM(201)))) + Str(OldM(201)) + Chr(124): C201 = C201 + OldM(201)
          
            If PocRadek > 75 Then:  NactiRozdil: PocRadek = 0: Strana = Strana + 1: For i = 1 To 85: T(i) = "": Next i
           End If
          
      End If
        If Dulezite = 0 Then: PrgDil = PrgDil + PocDil: ProgressBar1.Value = PrgDil
    On Error GoTo Vypadni
        RsTrideni.MoveNext
     Next PocTrid
   End If
   
  If PocRadek > 0 And Sestava = "Zdr.poj. - zaměstnanci" Then
                    'MsgBox (C272 & "  " & PocRadek)
                             PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(212) + String$(9, "-") + Chr(247) + String$(25, "-") + Chr(247) + String$(8, "-") + Chr(182) + String$(8, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(10, "-") + Chr(182) + String$(10, "-") + Chr(182) + String$(6, "-") + Chr(166)

                        If Len(Trim(Str(C202))) > 7 And Len(Trim(Str(C476))) < 8 Then
                                PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + "SUMA " + Space(30) + Chr(124) + Trim(Str(C202)) + Chr(124) + Str(C272) + Chr(124) + Space(7 - Len(Str(C9095))) + Str(C9095) + Chr(124) + Space(7 - Len(Str(C277))) + Str(C277) + Chr(124) + Space(10 - Len(Str(C274))) + Str(C274) + Chr(124) + Space(10 - Len(Str(C274277))) + Str(C274277) + Chr(124) + Space(5 - Len(Str(PocZamZP))) + Str(PocZamZP) + " " + Chr(124)
                        Else
                            If Len(Trim(Str(C476))) = 8 Then
                                    PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + "SUMA " + Space(30) + Chr(124) + Space(8 - Len(Str(C202))) + Str(C202) + Chr(124) + Space(8 - Len(Str(C272))) + Str(C272) + Chr(124) + Space(7 - Len(Str(C9095))) + Str(C9095) + Chr(124) + Space(7 - Len(Str(C277))) + Str(C277) + Chr(124) + Space(10 - Len(Str(C274))) + Str(C274) + Chr(124) + Space(10 - Len(Str(C274277))) + Str(C274277) + Chr(124) + Space(5 - Len(Str(PocZamZP))) + Str(PocZamZP) + " " + Chr(124)
                            Else
                             C272 = Round(C272 + 0.4999)
                                    PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(124) + "SUMA " + Space(30) + Chr(124) + Space(8 - Len(Str(C202))) + Str(C202) + Chr(124) + Space(8 - Len(Str(C272))) + Str(C272) + Chr(124) + Space(7 - Len(Str(C9095))) + Str(C9095) + Chr(124) + Space(7 - Len(Str(C277))) + Str(C277) + Chr(124) + Space(10 - Len(Str(C274))) + Str(C274) + Chr(124) + Space(10 - Len(Str(C274277))) + Str(C274277) + Chr(124) + Space(5 - Len(Str(PocZamZP))) + Str(PocZamZP) + " " + Chr(124)
    
                            End If
                        End If
                  PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(212) + String$(35, "-") + Chr(247) + String$(8, "-") + Chr(247) + String$(8, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(10, "-") + Chr(247) + String$(10, "-") + Chr(247) + String$(6, "-") + Chr(166)
                  PocRadek = PocRadek + 1: ZamVeta(PocRadek) = Chr(199) + String$(46, "-") + Chr(91) + "Salary s.r.o.  (Licence - " + Space(20 - Len(LicenceCis)) + LicenceCis + ")" + Chr(93) + String$(3, "-") + Chr(206)

      NactiZdrZam
  End If
  
 ' If PocRadek > 0 And Sestava = "Zdr.poj. - rozdíl (HM <> Vym. základ)" Then: NactiRozdil

   UhrnPojC = UhrnPojZam + UhrnPojOrg
        
Vypadni:
Next OldMes

OldMes = OldMes - 1
If Dulezite = 0 Then: ProgressBar1.Visible = False
End Sub
Public Sub NactiRozdil()
Dim veta As String


veta = Mid$(FirmaNazev, 1, 40) + " IČO: " + FirmaIco
T(1) = "měsíc: " + mesicStr(OldMes) + "/" + Str(SysRok) + " " + veta + Space(60 - Len(veta)) + " strana - " + Str(Strana)
T(2) = "Přehled jednotlivých zaměstnanců s odlišnou hrubou mzdou a vyměřovacím základem  na zdravotního pojištění "

T(3) = Chr(194) + String$(9, "-") + Chr(172) + String$(25, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(8, "-") + Chr(172) + String$(7, "-") + Chr(184)
T(4) = Chr(124) + " Osobní  " + Chr(124) + "Jméno zaměstnance        " + Chr(124) + " Hrubá " + Chr(124) + "Vyměř. " + Chr(124) + "Vratka" + Chr(124) + "Manaž." + Chr(124) + "Odměna" + Chr(124) + "Refun." + Chr(124) + " Odeč. " + Chr(124) + " Náhr." + Chr(124) + "  MVZ   " + Chr(124) + " Dohoda" + Chr(124)
T(5) = Chr(124) + " číslo   " + Chr(124) + "                         " + Chr(124) + " mzda  " + Chr(124) + "základ " + Chr(124) + "Dorov." + Chr(124) + "Rizik." + Chr(124) + "z fond" + Chr(124) + " mzda " + Chr(124) + " auto  " + Chr(124) + " úraz " + Chr(124) + "        " + Chr(124) + "       " + Chr(124)
T(6) = Chr(212) + String$(9, "-") + Chr(182) + String$(25, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(8, "-") + Chr(182) + String$(7, "-") + Chr(166)

For i = 1 To PocRadek
  T(6 + i) = ZamVeta(i)
Next i

T(7 + PocRadek) = Chr(212) + String$(9, "-") + Chr(247) + String$(25, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(8, "-") + Chr(182) + String$(7, "-") + Chr(166)
T(8 + PocRadek) = Chr(124) + "CELKEM    " + Space(25)
If Len(Str(C202)) > 6 Then
   T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Str(C202)
Else
   T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(7 - Len(Str(C202))) + Str(C202)
End If

If Len(Str(C272)) > 6 Then
   T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Str(C272)
Else
   T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(7 - Len(Str(C272))) + Str(C272)
End If

T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(6 - Len(Str(C476))) + Str(C476)
T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(6 - Len(Str(C291))) + Str(C291)
If Len(Str(C195)) > 5 Then
  T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Trim(Str(C195))
Else
  T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(6 - Len(Str(C195))) + Str(C195)
End If


T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(6 - Len(Str(C230))) + Str(C230)
T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(7 - Len(Str(C409))) + Str(C409)
T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(6 - Len(Str(C381))) + Str(C381)
T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(8 - Len(Str(C354))) + Str(C354)
T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(7 - Len(Str(DohC))) + Str(DohC) + Chr(124)
'T(8 + PocRadek) = T(8 + PocRadek) + Chr(124) + Space(7 - Len(Str(C201))) + Str(C201) + Chr(124)

T(9 + PocRadek) = Chr(212) + String$(9, "-") + "-" + String$(25, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(8, "-") + Chr(247) + String$(7, "-") + Chr(166)
T(10 + PocRadek) = Chr(199) + String$(57, "-") + Chr(91) + "Salary s.r.o. (Licence - " + Space(20 - Len(LicenceCis)) + LicenceCis + ")" + Chr(93) + String$(6, "-") + Chr(206)
'    MsgBox (Tisk & "sss")

 If Tisk = "" Then
  Text2.Font.Size = 7 * sizeFactor
   For i = 1 To PocRadek + 10
         Text2.Text = Text2.Text & T(i) & vbCrLf
          Print #1, T(i)
   Next i
 Else

     Printer.FontSize = 8
   For i = 1 To 95
     Printer.Print T(i)
   Next i
     Printer.EndDoc
 End If


End Sub
Private Sub NactiZdrCMR()
Dim veta As String
Dim Trideni As String


Trideni = PodrizenaVyber(PocPodrizena)
If PodrizenaPol = 0 Then: Trideni = "Celý podnik"
veta = Mid$(FirmaNazev, 1, 40) + " IČO: " + FirmaIco
T(1) = "měsíc: " + mesicStr(OldMes) + "/" + Str(SysRok) + " " + veta + Space(74 - Len(veta)) + " strana - " + Str(Strana)
T(2) = "Přehled jednotlivých zaměstnanců - Činnost malého rozsahu " + Trideni
T(3) = Chr(194) + String$(9, "-") + Chr(172) + String$(25, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(7, "-") + Chr(172) + String$(16, "-") + Chr(184)
T(4) = Chr(124) + " Osobní  " + Chr(124) + "Jméno zaměstnance        " + Chr(124) + " hrubá " + Chr(124) + "Vyměř. " + Chr(124) + "    Činnost     " + Chr(124)
T(5) = Chr(124) + " číslo   " + Chr(124) + "                         " + Chr(124) + " mzda  " + Chr(124) + "základ " + Chr(124) + " malého rozsahu " + Chr(124)
T(6) = Chr(212) + String$(9, "-") + Chr(182) + String$(25, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(16, "-") + Chr(166)

For i = 1 To PocRadek
  T(6 + i) = ZamVeta(i)
Next i

T(7 + PocRadek) = Chr(212) + String$(9, "-") + Chr(247) + String$(25, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(7, "-") + Chr(182) + String$(16, "-") + Chr(166)
T(8 + PocRadek) = Chr(124) + "CELKEM  " + Space(27) + Chr(124) + Space(7 - Len(Str(C202))) + Str(C202) + Chr(124) + Space(7 - Len(Str(C272))) + Str(C272) + Chr(124) + "Rozdíl " + Space(9 - Len(Str(C202 - C272))) + Str(C202 - C272) + Chr(124)
T(9 + PocRadek) = Chr(212) + String$(9, "-") + "-" + String$(25, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(7, "-") + Chr(247) + String$(16, "-") + Chr(166)
T(10 + PocRadek) = Chr(124) + "Počty zaměstnanců                   ANO  = " + Space(4 - Len(Str(CAno))) + Str(CAno) + "   NE  = " + Space(4 - Len(Str(CAno))) + Str(CAno) + Space(8) + Chr(124)
T(11 + PocRadek) = Chr(212) + String$(43, "-") + "-" + String$(24, "-") + Chr(166)
T(12 + PocRadek) = Chr(199) + String$(17, "-") + Chr(91) + "Salary s.r.o. (Licence - " + Space(20 - Len(LicenceCis)) + LicenceCis + ")" + Chr(93) + String$(3, "-") + Chr(206)


 If Tisk = "" Then
  Text2.Font.Size = 8 * sizeFactor
   For i = 1 To PocRadek + 12
         Text2.Text = Text2.Text & T(i) & vbCrLf
          Print #1, T(i)
   Next i

 Else
     Printer.FontSize = 12
   For i = 1 To 95
     Printer.Print T(i)
   Next i
     Printer.EndDoc
 End If

End Sub

Public Sub NactiZdrRekap()
Dim Odvod As Single

T(1) = Sestava + " " + RidiciVyber + " / " + ZdrNazevTisk(PocPodrizena) 'Trim(PodrizenaVyber(PocPodrizena)) 'ZdrNazevTisk(PocPodrizena)

If Trim(Label2(1)) = "Celá organizace" And Trim(Label3(1)) = "Celá organizace" Then: T(1) = Sestava + " / Celkem pojištovny"

veta = Mid$(FirmaNazev, 1, 40) + " IČO: "
T(2) = "měsíc: " + mesicStr(OldMes) + "/" + Str(SysRok) + " " + veta + FirmaIco
T(3) = ""
If Trim(Label3(1)) = "Zdravotní pojištovny" Then
   T(4) = Space(8) + ZdrKodTisk(PocPodrizena) + " " + ZdrNazevTisk(PocPodrizena)
   T(5) = Space(8) + ZdrUliceTisk(PocPodrizena) + " " + ZdrCisloDomuTisk(PocPodrizena)
   T(6) = ZdrPscTisk(PocPodrizena) + Space(8 - Len(ZdrPscTisk(PocPodrizena))) + " " + ZdrObecTisk(PocPodrizena)
End If

T(7) = Chr(194) + String$(20, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(20, "-") + Chr(172) + String$(6, "-") + Chr(172) + String$(20, "-") + Chr(184)
T(8) = Chr(124) + " Číslo účtu         " + Chr(124) + " Kód  " + Chr(124) + " Variabilní symbol  " + Chr(124) + "Konst." + Chr(124) + "Specifický symbol   " + Chr(124)
T(9) = Chr(212) + String$(20, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(20, "-") + Chr(182) + String$(6, "-") + Chr(182) + String$(20, "-") + Chr(166)
T(10) = Chr(124) + ZdrCisloUctuTisk(PocPodrizena) + Space(20 - Len(ZdrCisloUctuTisk(PocPodrizena))) + Chr(124) + ZdrKodBankyTisk(PocPodrizena) + Space(5 - Len(ZdrKodBankyTisk(PocPodrizena))) + " " + Chr(124) + ZdrVsTisk(PocPodrizena) + Space(20 - Len(ZdrVsTisk(PocPodrizena))) + Chr(124) + ZdrKsTisk(PocPodrizena) + Space(5 - Len(ZdrKsTisk(PocPodrizena))) + " " + Chr(124) + ZdrSsTisk(PocPodrizena) + Space(20 - Len(ZdrSsTisk(PocPodrizena))) + Chr(124)
T(11) = Chr(212) + String$(20, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(20, "-") + Chr(247) + String$(6, "-") + Chr(247) + String$(20, "-") + Chr(166)
T(12) = Chr(212) + String$(25, "-") + Chr(91) + "Salary s.r.o. (Licence - " + Space(20 - Len(LicenceCis)) + LicenceCis + ")" + Chr(93) + String$(3, "-") + Chr(166)
T(13) = Chr(212) + String$(18, "-") + Chr(172) + String$(14, "-") + Chr(172) + String$(11, "-") + Chr(172) + String$(12, "-") + Chr(172) + String$(17, "-") + Chr(166)
T(14) = Chr(124) + "Vyměřovací základ " + Chr(124) + "Základ vratky " + Chr(124) + "Zaměstnanec" + Chr(124) + " Organizace " + Chr(124) + "      Odvod      " + Chr(124)
T(15) = Chr(212) + String$(18, "-") + Chr(182) + String$(14, "-") + Chr(182) + String$(11, "-") + Chr(182) + String$(12, "-") + Chr(182) + String$(17, "-") + Chr(166)
Odvod = UhrnPojC
T(16) = Chr(124) + Space(18 - Len(Str(UhrnVymZakl))) + Str(UhrnVymZakl) + Chr(124) + Space(14 - Len(Str(UhrnVratka))) + Str(UhrnVratka) + Chr(124) + Space(11 - Len(Str(UhrnPojZam))) + Str(UhrnPojZam) + Chr(124) + Space(12 - Len(Str(UhrnPojOrg))) + Str(UhrnPojOrg) + Chr(124) + Space(15 - Len(Str(Odvod))) + Str(Odvod) + Space(2) + Chr(124)
T(17) = Chr(199) + String$(18, "-") + Chr(247) + String$(14, "-") + Chr(247) + String$(11, "-") + Chr(247) + String$(12, "-") + Chr(247) + String$(17, "-") + Chr(206)

 If Tisk = "" Then
  Text2.Font.Size = 8 * sizeFactor
   For i = 1 To 17
         Text2.Text = Text2.Text & T(i) & vbCrLf
          Print #1, T(i)
   Next i
 Else
     Printer.FontSize = 12
     For i = 1 To 17
       Printer.Print T(i)
     Next i
     Printer.EndDoc
 End If

End Sub
Public Sub NactiZdrZam()

  Text2.Font.Size = 7 * sizeFactor
   For i = 1 To PocRadek
         Text2.Text = Text2.Text & ZamVeta(i) & vbCrLf
          Print #1, ZamVeta(i)
   Next i
 
End Sub

Private Sub NactiZdrHlaseni()
Dim d As Integer
Dim PoLis As Integer
Dim MaxList As Integer ' Počet listů
Dim PocZaz As Integer
Dim Cy As Long
Dim RsHlaseni As Recordset
Dim MaxRsHlaseni As Integer
Dim MesRok As String
Dim MesRok2 As String
Dim PocetPrihlasek As Integer
Dim OcH As String
Dim VPoj As String
Dim HOcis(100) As String
Dim DHl As String
Dim MHl As String
Dim RHl As String
Dim ZHl As Recordset
    

MesRok = Trim(Str(OldMes)) & "." & Trim(Str(OldRok))
MesRok2 = Trim(Str(OldMes - 1)) & "." & Trim(Str(OldRok))
If OldMes = 1 Then: MesRok2 = Trim(Str(12)) & "." & Trim(Str(OldRok - 1))

If Len(MesRok) = 6 Then: MesRok = "0" & MesRok
If Len(MesRok2) = 6 Then: MesRok2 = "0" & MesRok2

VPoj = Trim(PodrizenaVyber(PocPodrizena))
MaxRsHlaseni = 0
PocetPrihlasek = 0
a = 1
For PocZaz = 1 To 100
    HOcis(PocZaz) = " "
Next
' vybere všechny co mají OldMesic, OldRok a jsou Aktivni
  Set RsHlaseni = DbUdaje.OpenRecordset("SELECT Zamestnanec.OsCislo, Zamestnanec.Prijmeni, Zamestnanec.Jmeno, Zamestnanec.Aktivni, DataPam.PoCislo, DataPam.PoNapln FROM (Zamestnanec INNER JOIN DataPam ON Zamestnanec.OsCislo = DataPam.OsCislo) WHERE  DataPam.PoCislo = 250 and trim(DataPam.PoNapln) = ""Ano"" and Zamestnanec.Aktivni = 1 ")
      On Error GoTo NikdoNeni
      
NikdoNeni:
 If err = 3021 Then: MsgBox ("Nikoho jsem nenašel"): Exit Sub
 
      RsHlaseni.MoveLast
      RsHlaseni.MoveFirst
      MaxRsHlaseni = RsHlaseni.RecordCount
      
      PocetPrihlasek = 0
      RsHlaseni.MoveFirst
      Close #3
         For PocZaz = 1 To MaxRsHlaseni
         
                OcH = RsHlaseni.Fields("OsCislo")
                 NahrajZam (OcH)
                    h = NahrajOsPam(OcH, OldMes, OldRok)
                    If Trim(OldP(49)) = Trim(VPoj) Then
                      
                      For d = 1 To Len(Trim(OldP(55)))
                     ' MsgBox (OldP(55)): Exit Sub
                        PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), d, 1): HOcis(PocetPrihlasek) = OcH
                            Printer.CurrentX = 720: Printer.CurrentY = Cy: Printer.Print Trim(KodZ(PocetPrihlasek)) 'kod změny
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy: Printer.Print Trim(OldO(100)) 'rodny číslo
                            Printer.CurrentX = 9000: Printer.CurrentY = Cy: Printer.Print Trim(OldP(59)) ' datum změny
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 600: Printer.Print Trim(ZPrijmeni) 'Trim(ZHl.Fields("Prijmeni"))
                            Printer.CurrentX = 8000: Printer.CurrentY = Cy + 600: Printer.Print Trim(ZJmeno) ' Trim(ZHl.Fields("jmeno"))
                      
                        If Trim(OldO(158)) <> "" Then
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(157))
                            Printer.CurrentX = 9300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(164))
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(159))
                            Printer.CurrentX = 3000: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(158))
                        Else
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(150))
                            Printer.CurrentX = 9300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(156))
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(152))
                            Printer.CurrentX = 3000: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(151))
                       End If
                      Next d
                     
                    End If
                      
                    
                  '    If Len(Trim(OldP(55))) = 1 Then
                  '        PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), 1, 1): HOcis(PocetPrihlasek) = OcH
                  '    End If
                  '    If Len(Trim(OldP(55))) = 2 Then
                  '       PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), 1, 1): HOcis(PocetPrihlasek) = OcH
                  '       PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), 2, 1): HOcis(PocetPrihlasek) = OcH
                  '    End If
                  '    If Len(Trim(OldP(55))) = 3 Then
                  '      PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), 1, 1): HOcis(PocetPrihlasek) = OcH
                  '      PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), 2, 1): HOcis(PocetPrihlasek) = OcH
                  '      PocetPrihlasek = PocetPrihlasek + 1: KodZ(PocetPrihlasek) = Mid(Trim(OldP(55)), 3, 1): HOcis(PocetPrihlasek) = OcH
                  '    End If
                                         ' Osob číslo vybraných
             RsHlaseni.MoveNext
         Next PocZaz
        
     
 
                ' 111 - VZP
                ' 201 - Vojenská
                ' 205 - Prumyslova
                ' 207 - OZP
                ' 209 - Zaměstnanecká
                ' 211 - Minister. Vnitra
                ' 213 - Bratrská
                ' 217 - Metal Alians
                ' 228 - Media
                

                     Set cdl = CommonDialog1 ' nutné před inicializací - připojuje common dialog
                     cdl.PrinterDefault = False
                     cdl.CancelError = True
                     InitLp
                      If PrintAno = "N" Then: Exit Sub
                      Printer.FontSize = 11
                 
                        
               OcH = " "
               MaxList = PocetPrihlasek / 4 + 0.99
               Lomitko = InStr(1, Trim(FirmaAdrCislo), "/")
               
               
               j = 0
             For PoLis = 1 To MaxList ' tolikrát kolik je počet listů
                       

                  If ZdrKodHl = "111" Then: Printer.PaintPicture Pct_VZP_111.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 111: j = j + 1
                  If ZdrKodHl = "201" Then: Printer.PaintPicture Pct_VOZP_201.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 201: j = j + 1
                  If ZdrKodHl = "205" Then: Printer.PaintPicture Pct_CPZP_205.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 205: j = j + 1
                  If ZdrKodHl = "207" Then: Printer.PaintPicture Pct_OZP_207.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 207: j = j + 1
                  If ZdrKodHl = "209" Then: Printer.PaintPicture Pct_ZPS_209.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 209: j = j + 1
                  If ZdrKodHl = "211" Then: Printer.PaintPicture Pct_ZPMVCR_211.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 211: j = j + 1
                  If ZdrKodHl = "213" Then: Printer.PaintPicture Pct_RBP_213.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 213: j = j + 1
                  If ZdrKodHl = "217" Then: Printer.PaintPicture Pct_Metal_217.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 217: j = j + 1
                  If ZdrKodHl = "228" Then: Printer.PaintPicture Pct_ZPMedia_228.Picture, 0, 0, 11950, 16800: NazevSestavyHl = 228: j = j + 1
                  If j = 1 Then: MsgBox (NazevSestavyHl): Utvor_HOZ_XML_StaryStyl
                                '  --- Hlavička
                  Printer.CurrentX = 5400: Printer.CurrentY = 2300: Printer.Print Mid(Command18.Caption, 7, 8)
                  
                  Printer.CurrentX = 900: Printer.CurrentY = 3250: Printer.Print Trim(FirmaNazev)
                  Printer.CurrentX = 900: Printer.CurrentY = 3800: Printer.Print Trim(FirmaUlice)
                  Printer.CurrentX = 6900: Printer.CurrentY = 3800: Printer.Print Trim(FirmaAdrCislo)
                  Printer.CurrentX = 9000: Printer.CurrentY = 3800: Printer.Print Trim(FirmaIco)
                  Printer.CurrentX = 900: Printer.CurrentY = 4350: Printer.Print Trim(FirmaPSC)
                  Printer.CurrentX = 2300: Printer.CurrentY = 4350: Printer.Print Trim(FirmaAdrObec)
                  Printer.CurrentX = 9000: Printer.CurrentY = 4350: Printer.Print Trim(FirmaTelefon)

                    ' ---- Tělo
                  Cy = 5200
                  
                   For PocZaz = 1 To 4
                        If PocetPrihlasek = 0 Then:  Printer.EndDoc: Msg = " Soubory Hromadného oznámení je vytvořeny spolu s XML. ": MsgInfo: Exit Sub ' když už nejsou lidi uteče
                          
                          If OcH <> Trim(HOcis(PocetPrihlasek)) Then
                              OcH = Trim(HOcis(PocetPrihlasek))
                                h = NahrajOsPam(OcH, OldMes, OldRok)
                          Set ZHl = DbUdaje.OpenRecordset("SELECT * FROM Zamestnanec WHERE OsCislo = """ & OcH & """ ")
                              ZHl.MoveFirst
                          End If
                           '----------- prvni radek
                           rcZam(PocetPrihlasek) = Trim(OldO(100))
                           DatZm(PocetPrihlasek) = Trim(OldP(59))
                           PrijmeniZam(PocetPrihlasek) = Trim(ZHl.Fields("Prijmeni"))
                           JmenoZam(PocetPrihlasek) = Trim(ZHl.Fields("jmeno"))
                           
                            Printer.CurrentX = 720: Printer.CurrentY = Cy: Printer.Print Trim(KodZ(PocetPrihlasek)) 'kod změny
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy: Printer.Print Trim(OldO(100)) 'rodny číslo
                            Printer.CurrentX = 9000: Printer.CurrentY = Cy: Printer.Print Trim(OldP(59)) ' datum změny
                            
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 600: Printer.Print Trim(ZHl.Fields("Prijmeni"))
                            Printer.CurrentX = 8000: Printer.CurrentY = Cy + 600: Printer.Print Trim(ZHl.Fields("jmeno"))
     
                      
                        If Trim(OldO(158)) <> "" Then
                         UliceZam(PocetPrihlasek) = Trim(OldO(157))
                         CpZam(PocetPrihlasek) = Trim(OldO(164))
                         PSCZam(PocetPrihlasek) = Trim(OldO(159))
                         ObecZam(PocetPrihlasek) = Trim(OldO(158))
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(157))
                            Printer.CurrentX = 9300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(164))
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(159))
                            Printer.CurrentX = 3000: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(158))
                            
                        Else
                         UliceZam(PocetPrihlasek) = Trim(OldO(150))
                         CpZam(PocetPrihlasek) = Trim(OldO(156))
                         PSCZam(PocetPrihlasek) = Trim(OldO(152))
                         ObecZam(PocetPrihlasek) = Trim(OldO(151))
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(150))
                            Printer.CurrentX = 9300: Printer.CurrentY = Cy + 1100: Printer.Print Trim(OldO(156))
                            Printer.CurrentX = 1300: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(152))
                            Printer.CurrentX = 3000: Printer.CurrentY = Cy + 1670: Printer.Print Trim(OldO(151))
                       End If
                     
                          Cy = Cy + 2400
                           PocetPrihlasek = PocetPrihlasek - 1
                 Utvor_HOZ_XML_StaryStylZam

             Next PocZaz
                              
           Printer.NewPage ' zavři tiskárnu
        Next PoLis  ' konec všech listů
      
     Print #3, "  </seznamZmenZamestnancu>"
  Print #3, "</hromadneOznameniZamestnavatele>"
 Close #3

End Sub


Private Sub NactiZdrPrehled()
Dim Kontr As Integer

If Trim(ZdrKodTisk(PocPodrizena)) = "" Then: ZdrKodTisk(PocPodrizena) = "000"
ZdrKodTisk(PocPodrizena) = Mid$(ZdrKodTisk(PocPodrizena), 1, 3)
'MsgBox (ZdrKodTisk(PocPodrizena))

T(0) = "***" + ZdrKodTisk(PocPodrizena)

    T(1) = Space(17) + Chr(194) + String$(59, "-") + Chr(184)
    T(2) = Space(17) + Chr(124) + "                         Přehled                           " + Chr(124)
    T(3) = Chr(194) + String$(5, "-") + Chr(172) + String$(7, "-") + Chr(184) + "  " + Chr(124) + " o platbě pojistného na zdravotní pojištění zaměstnavatele " + Chr(124)
    T(4) = Chr(124) + " Kód " + Chr(124) + " " + ZdrKodTisk(PocPodrizena) + Space(6 - Len(ZdrKodTisk(PocPodrizena))) + Chr(124) + "  " + Chr(124) + Space(59) + Chr(124)
    T(5) = Chr(199) + String$(5, "-") + Chr(247) + String$(7, "-") + Chr(206) + "  " + Chr(199) + String$(59, "-") + Chr(206)
    T(6) = Chr(194) + String$(60, "-") + Chr(184)
        
    T(7) = Chr(124) + "Název a sídlo zaměstnavatele :" + Space(30) + Chr(124) + " " + Chr(194) + String$(13, "-") + Chr(184)
    T(8) = Chr(124) + FirmaNazev + Space(60 - Len(FirmaNazev)) + Chr(124) + " " + Chr(124) + " Číslo plátce" + Chr(124)
    T(9) = Chr(124) + FirmaUlice + " " + FirmaAdrCislo + Space(60 - Len(FirmaUlice + " " + FirmaAdrCislo)) + Chr(124) + " " + Chr(212) + String$(10, "-") + Chr(172) + String$(2, "-") + Chr(166)
    T(10) = Chr(124) + FirmaPSC + "  " + FirmaAdrObec + Space(60 - Len(FirmaPSC + "  " + FirmaAdrObec)) + Chr(124) + " " + Chr(124) + Space(10 - Len(Mid(FirmaIco, 1, 9))) + Mid(FirmaIco, 1, 9) + Chr(124) + Mid(MzdUct, 1, 2) + Chr(124)
    T(11) = Chr(199) + String$(60, "-") + Chr(206) + " " + Chr(199) + String$(10, "-") + Chr(247) + String$(2, "-") + Chr(206)
    
    T(12) = Chr(194) + String$(76, "-") + Chr(184)
    T(13) = Chr(124) + "Hlášení za měsíc a rok :" + Space(48 - Len(Str(OldMes) + "/" + Str(SysRok))) + Str(OldMes) + "/" + Str(SysRok) + "    " + Chr(124)
    T(14) = Chr(124) + Space(76) + Chr(124)
    ZdrNazevTisk(PocPodrizena) = Mid$(ZdrNazevTisk(PocPodrizena), 1, 38)
    T(15) = Chr(124) + "Počet zaměstnanců pojištěných u " + ZdrNazevTisk(PocPodrizena) + Space(38 - Len(ZdrNazevTisk(PocPodrizena))) + Str(UhrnPocetZam) + Space(6 - Len(Str(UhrnPocetZam))) + Chr(124)
    T(16) = Chr(124) + Space(76) + Chr(124)
    'UhrnVratka = 0
    T(17) = Chr(124) + "Úhrn vyměřovacích základů zaměstnanců" + Space(35 - Len(Str(UhrnVymZakl))) + Str(UhrnVymZakl) + "    " + Chr(124)
    T(18) = Chr(124) + Space(76) + Chr(124)
    T(19) = Chr(124) + "Výše pojistného - součet pojistného zaměstnanců" + Space(25 - Len(Str(UhrnPojZam + UhrnPojOrg))) + Str(UhrnPojZam + UhrnPojOrg) + "    " + Chr(124)
    T(20) = Chr(124) + Space(76) + Chr(124)
    T(21) = Chr(199) + String$(76, "-") + Chr(206)
    T(22) = ""
    T(23) = " Dne : " + Str(Den) + "." + Str(Mesic) + "." + Str(Rok) + "                          Převzato dne:"
    T(24) = Chr(194) + String$(33, "-") + Chr(184) + Space(8) + Chr(194) + String$(33, "-") + Chr(184)
    T(25) = Chr(124) + Space(33) + Chr(124) + Space(8) + Chr(124) + Space(33) + Chr(124)
    T(26) = Chr(124) + Space(33) + Chr(124) + Space(8) + Chr(124) + Space(33) + Chr(124)
    T(27) = Chr(124) + Space(33) + Chr(124) + Space(8) + Chr(124) + Space(33) + Chr(124)
    T(28) = Chr(124) + Space(33) + Chr(124) + Space(8) + Chr(124) + Space(33) + Chr(124)
    T(29) = Chr(124) + Space(33) + Chr(124) + Space(8) + Chr(124) + Space(33) + Chr(124)
    T(30) = Chr(199) + String$(33, "-") + Chr(206) + Space(8) + Chr(199) + String$(33, "-") + Chr(206)
    
    T(31) = "razítko a podpis zaměstnavatele                razítko zdravotní pojištovny"

Kontr = 0
                 
                
              
   If ZdrKodTisk(PocPodrizena) = "207" Then
     ZdrKodXML(1) = "207": ZdrZkratkaXML(1) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(1) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(1) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(1) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 1: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "111" Then
     ZdrKodXML(2) = "111": ZdrZkratkaXML(2) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(2) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(2) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(2) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 2: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "201" Then
     ZdrKodXML(3) = "201": ZdrZkratkaXML(3) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(3) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(3) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(3) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 3: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "211" Then
     ZdrKodXML(4) = "211": ZdrZkratkaXML(4) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(4) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(4) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(4) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 4: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "205" Then
     ZdrKodXML(5) = "205": ZdrZkratkaXML(5) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(5) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(5) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(5) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 5: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "209" Then
     ZdrKodXML(6) = "209": ZdrZkratkaXML(6) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(6) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(6) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(6) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 6: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "213" Then
     ZdrKodXML(7) = "213": ZdrZkratkaXML(7) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(7) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(7) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(7) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     X = 7: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "217" Then
     ZdrKodXML(8) = "217": ZdrZkratkaXML(8) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(8) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(8) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(8) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
     'X = 8: Utvor_PPPZ_XML_StaryStyl
   End If
   If ZdrKodTisk(PocPodrizena) = "228" Then
     ZdrKodXML(9) = "228": ZdrZkratkaXML(9) = ZdrZkratkaTisk(PocPodrizena)
     ZPocZam(9) = Space(12 - Len(Trim(Str(UhrnPocetZam)))) + Trim(Str(UhrnPocetZam))
     ZVymZakl(9) = Space(12 - Len(Trim(Str(UhrnVymZakl)))) + Trim(Str(UhrnVymZakl))
     ZPoj(9) = Space(12 - Len(Trim(Str(UhrnPojZam + UhrnPojOrg)))) + Trim(Str(UhrnPojZam + UhrnPojOrg))
     Kontr = 1
    ' X = 9: Utvor_PPPZ_XML_StaryStyl
   End If

 If Tisk = "" Then
  Text2.Font.Size = 8 * sizeFactor
   For i = 0 To 31
         If i > 0 Then: Text2.Text = Text2.Text & T(i) & vbCrLf
          If Kontr = 0 Then: Print #1, T(i)
          If Kontr = 1 And i = 0 Then: Print #1, T(i)
   Next i

 Else
     Printer.FontSize = 10
   For i = 0 To 31
     Printer.Print T(i)
   Next i
     Printer.EndDoc
 End If

End Sub


Private Sub NactiTrideni()
     
     Grid1.Rows = 17 ' počet Grid1.Row   + 1 ( počet položek pro třídění + 1)
     
     '                                       číslo položky v osobním                název položky
     Grid1.Row = 1: Grid1.Col = 1: Grid1.Text = "000": Grid1.Col = 2: Grid1.Text = "Celá organizace               "
     Grid1.Row = 2: Grid1.Col = 1: Grid1.Text = "001": Grid1.Col = 2: Grid1.Text = "Středisko                     "
     Grid1.Row = 3: Grid1.Col = 1: Grid1.Text = "002": Grid1.Col = 2: Grid1.Text = "Zdravotní pojištovny          "
     Grid1.Row = 4: Grid1.Col = 1: Grid1.Text = "400": Grid1.Col = 2: Grid1.Text = "Výplatní místo                "
     Grid1.Row = 5: Grid1.Col = 1: Grid1.Text = "402": Grid1.Col = 2: Grid1.Text = "Nákladové středisko           "
     Grid1.Row = 6: Grid1.Col = 1: Grid1.Text = "404": Grid1.Col = 2: Grid1.Text = "Zúčtovací tříd. hled.   vyšší "
     Grid1.Row = 7: Grid1.Col = 1: Grid1.Text = "406": Grid1.Col = 2: Grid1.Text = "Zúčtovací tříd. hled.nejvyšší "
     Grid1.Row = 8: Grid1.Col = 1: Grid1.Text = "410": Grid1.Col = 2: Grid1.Text = "Rozborové tříd. hled.základní "
     Grid1.Row = 9: Grid1.Col = 1: Grid1.Text = "412": Grid1.Col = 2: Grid1.Text = "Rozborové třídící hled. vyšší "
     Grid1.Row = 10: Grid1.Col = 1: Grid1.Text = "424": Grid1.Col = 2: Grid1.Text = "Kategorie pro rozbor. sest.I. "
     Grid1.Row = 11: Grid1.Col = 1: Grid1.Text = "426": Grid1.Col = 2: Grid1.Text = "Vyšší rozborová kategorie     "
     Grid1.Row = 12: Grid1.Col = 1: Grid1.Text = "420": Grid1.Col = 2: Grid1.Text = "Zápočet pro statistiku        "
     Grid1.Row = 13: Grid1.Col = 1: Grid1.Text = "431": Grid1.Col = 2: Grid1.Text = "Kategorie pro rozbor. sest.II."
     Grid1.Row = 14: Grid1.Col = 1: Grid1.Text = "432": Grid1.Col = 2: Grid1.Text = "Vyšší rozborová kategorie     "
     Grid1.Row = 15: Grid1.Col = 1: Grid1.Text = "422": Grid1.Col = 2: Grid1.Text = "Zaměstnání vykon. v organizaci"
     Grid1.Row = 16: Grid1.Col = 1: Grid1.Text = "51": Grid1.Col = 2: Grid1.Text = "Kód zaměstnání (KZAM)         "
    
End Sub




Private Sub Command1_Click()
 Cmd = 1
  Grid1.Visible = True
  Frame3.Caption = " Výběr řídícího hlediska"
  Frame3.Visible = True
  'Grid3.Visible = False
  Frame1.Visible = False
  Frame5.Visible = False
  Text1.Visible = False
  'Command8.Visible = False
  'Command11 = False
   RidiciPol = 0
  RidiciVyber = ""
  NactiTrideni
End Sub

Private Sub Command10_Click()
FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1
If PrepocetAnoNe = True Then: ResultText(1) = "X": FrmPrepocVsem.Display
NactiFirmu
FrmZdrPoj.MousePointer = 11
Command9.Visible = True
Frame2.Visible = True
ProgressBar1.Visible = True
ProgressBar2.Visible = True
ProgressBar2.Value = 100
 SumaVymZakl = 0: PrumPocZamSoc = 0: UhrnPojC = 0:: UhrnPocetZam = 0: UhrnVymZakl = 0: UhrnVratka = 0:   PocRadek = 0: UhrnPojZam = 0: UhrnPojOrg = 0

       Text2 = "": Tisk = "": Sestava = "Zdr.poj. - součtová rekapitulace"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update
       
           PocDil1 = Fix(100 / MaxPocPodrizena)
           PrgDil1 = 0
           
                
           
          
      If Trim(Label2(1)) = "Celá organizace" And Trim(Label3(1)) = "Zdravotní pojištovny" And OldRok > 2010 Then
        For PocPodrizena = 1 To MaxPocPodrizena
          For i = 1 To PocPrvku + 1
    
            If ZdrZkratka(i) = Trim(PodrizenaVyber(PocPodrizena)) Then
               For a = 1 To PocPrvku + 1
               
                If Trim(ZdrKod(i)) = Trim(KodPojjMesic(a)) Then
                    SumaVymZakl = SumaVymZakl + Sum272(a) '+ Sum476(a)
                    VymZaklZam = VymZaklZam + Sum272(a) '+ Sum476(a)
                    UhrnVymZakl = UhrnVymZakl + Sum272(a)
                    UhrnVratka = UhrnVratka + Sum476(a)
                    UhrnPojZam = UhrnPojZam + Sum277(a)
                    UhrnPojOrg = UhrnPojOrg + Sum274(a)
                    UhrnPocetZam = UhrnPocetZam + PocLidi(a)
                    ZdrKodTisk(PocPodrizena) = KodPojjMesic(a)
                   UhrnPojC = UhrnPojZam + UhrnPojOrg
                End If
               Next a
               
             End If
         Next i
       Next PocPodrizena
     '  NactiZdrRekap
     Else

            For PocPodrizena = 1 To MaxPocPodrizena
                    Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
                    PrgDil1 = PrgDil1 + PocDil1
                    ProgressBar2.Value = PrgDil1
                NactiZdrUdaje
            Next PocPodrizena
     End If
                        
                        
                        
                    PocPodrizena = 30
                    ZdrZkratka(PocPodrizena) = ""
                    ZdrNazev(PocPodrizena) = "Součet všech pojištoven"
                    ZdrUlice(PocPodrizena) = ""
                    ZdrCisloDomu(PocPodrizena) = ""
                    ZdrObec(PocPodrizena) = ""
                    ZdrPsc(PocPodrizena) = ""
                    ZdrTelefon(PocPodrizena) = ""
                    ZdrCisloUctu(PocPodrizena) = ""
                    ZdrKodBanky(PocPodrizena) = ""
                    ZdrKs(PocPodrizena) = ""
                    ZdrVs(PocPodrizena) = ""
                    ZdrSs(PocPodrizena) = ""
                    ZdrKod(PocPodrizena) = ""

            If UhrnPocetZam + UhrnVymZakl + UhrnVratka > 0 Then: NactiZdrRekap
Close #1
ProgressBar1.Visible = False
ProgressBar2.Visible = False
Frame2.Visible = False
FrmZdrPoj.MousePointer = 0
End Sub

Private Sub Command11_Click()
 
   For i = 1 To PocPrvku
           Grid3.Row = i
           Grid3.Col = 1
         If Grid3.Text = "X" Then
           Grid3.Text = ""
         Else
           Grid3.Text = "X"
         End If

   Next i
End Sub



Private Sub Command12_Click()
 Frame6.Visible = True
 
 Combo3(0) = "Prosinec" '"Leden"
 Combo3(1) = Str(SysRok - 1)
 Combo3(2) = mesicStr(2) 'mesicStr(SysMes)
 Combo3(3) = Str(SysRok)
    

End Sub

Private Sub Command13_Click()
FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1
Strana = 1
Command9.Visible = True
FrmSocPoj.MousePointer = 11
    OdMes = mesicInt(Combo1)
    DoMes = mesicInt(Combo1)
    OldRok = Val(Combo2)
    Strana = 0
    Frame2.Visible = True
    ProgressBar1.Visible = True
 ProgressBar2.Visible = True
 ProgressBar2.Value = 100
Text2 = "": Tisk = "": Sestava = "Činnost malého rozsahu"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update

PocDil1 = Fix(100 / MaxPocPodrizena)
           PrgDil1 = 0
   For PocPodrizena = 1 To MaxPocPodrizena
         If Trim(PodrizenaVyber(PocPodrizena)) = "" Then: Frame2.Caption = " Celá organizace"
         Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
           PrgDil1 = PrgDil1 + PocDil1
           ProgressBar2.Value = PrgDil1
       NactiZdrUdaje
       If PocRadek > 0 Then: NactiZdrCMR: PocRadek = 0
   Next PocPodrizena
 Close #1
    Frame2.Visible = False
    ProgressBar1.Visible = False
    ProgressBar2.Visible = False
FrmSocPoj.MousePointer = 0

End Sub
Private Sub Command14_Click()
Dim Mesicek As Integer
Dim MesOd As Integer
Dim MesDo As Integer
Dim DatumS1 As String
Dim DatumS2 As String
 
  Frame6.Visible = False
  FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1
Kontrolor = 1
NactiFirmu
Command9.Visible = True
FrmZdrPoj.MousePointer = 11
    MesOd = mesicInt(Combo3(0))
    OldRokOd = Val(Combo3(1))
    OldRokDo = Val(Combo3(3))
    MesDo = mesicInt(Combo3(2))
    Strana = 0
    PocSestav = 0
    Frame2.Visible = True
    ProgressBar1.Visible = True
    ProgressBar2.Visible = True

Text2 = "": Tisk = "": Sestava = "Zdr.poj. - zaměstnanci"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update
          
          
    
  DatumS1 = DateStr(Datum(1, MesOd, OldRokOd))
  DatumS2 = DateStr(Datum(1, MesDo, OldRokDo))
 Mesicek = DateDiff("M", DatumS1, DatumS2) + 1

'MsgBox (DatumS1 & "    " & DatumS2 & "  " & Mesicek)

      PocDil1 = Fix(100 / Mesicek)
       PrgDil1 = 0

For OldRok = OldRokOd To OldRokDo
  MesOd = 1: MesDo = 12
   If OldRok = OldRokOd Then: MesOd = mesicInt(Combo3(0)): MesDo = 12
   If OldRok = OldRokDo Then: MesOd = 1: MesDo = mesicInt(Combo3(2))
   If OldRok = OldRokOd And OldRok = OldRokDo Then: MesOd = mesicInt(Combo3(0)): MesDo = mesicInt(Combo3(2))
   
 

 For Mesicek = MesOd To MesDo
                    PrgDil1 = PrgDil1 + PocDil1
                    ProgressBar2.Value = PrgDil1
    DoMes = Mesicek: OdMes = Mesicek
     If OldRok = OldRokDo And MesDo = Mesicek Then: Kontrolor = 100

          Command7_Click
         
 Next Mesicek
Next OldRok
   
 Close #1
    Frame2.Visible = False
    ProgressBar1.Visible = False
 ProgressBar2.Visible = False
FrmZdrPoj.MousePointer = 0
Kontrolor = 0
End Sub

Private Sub Command15_Click()

Frame6.Visible = False
Kontrolor = 0

End Sub

Private Sub Command17_Click()
FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1
                
    Print #1, Text2
                

Close #1
End Sub


Private Sub Command18_Click()
Dim mm As Integer
Dim rr As Integer
If Command18.Caption = "Období " + Trim(Str(SysMes)) + "/" + Trim(Str(SysRok)) Then
 mm = SysMes + 1: rr = SysRok
 If SysMes = 12 Then: mm = 1: rr = SysRok + 1
  Command18.Caption = "Období " + Trim(Str(mm)) + "/" + Trim(Str(rr))
Else
 Command18.Caption = "Období " + Trim(Str(SysMes)) + "/" + Trim(Str(SysRok))
End If


End Sub

Private Sub Command2_Click()
  
   Grid2.Visible = True
   Frame4.Visible = True
   'Grid3.Visible = False
   Frame1.Visible = False
   'Command8.Visible = False
   'Command11 = False
    
End Sub

Private Sub Command3_Click()
' If RidiciPol = 2 Then
  Cmd = 2
    Grid1.Visible = True
    Frame3.Caption = " Výběr podřízeného hlediska"
   Frame3.Visible = True
    Grid2.Visible = False
    Frame4.Visible = False
    'Grid3.Visible = False
    Frame1.Visible = False
    Frame5.Visible = False
   Text1.Visible = False
   ' Command8.Visible = False
   ' Command11 = False
'  End If
   PodrizenaPol = 0
   For i = 1 To MaxPocPodrizena: PodrizenaVyber(i) = "": Next i
   MaxPocPodrizena = 1
 If Trim(Label2(0)) = "002" Then
   NactiSeznamPolozky2
 Else
  NactiTrideni
 End If
End Sub

Private Sub Command4_Click()
FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1
If PrepocetAnoNe = True Then: ResultText(1) = "X": FrmPrepocVsem.Display
NactiFirmu
FrmZdrPoj.MousePointer = 11
Command9.Visible = True
Frame2.Visible = True
ProgressBar1.Visible = True
 ProgressBar2.Visible = True
 ProgressBar2.Value = 100
    OdMes = mesicInt(Combo1)
    DoMes = mesicInt(Combo1)
    OldRok = Val(Combo2)
   ' MsgBox (MaxPocPodrizena)
  SumaVymZakl = 0: PrumPocZamSoc = 0: UhrnPojC = 0: UhrnPocetZam = 0:  UhrnVymZakl = 0: UhrnVratka = 0:  PocRadek = 0: UhrnPojZam = 0: UhrnPojOrg = 0

       Text2 = "": Tisk = "": Sestava = "Zdr.poj. - rekapitulace pojištoven"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update
       
       PocDil1 = Fix(100 / MaxPocPodrizena)
       PrgDil1 = 0



     If Trim(Label2(1)) = "Celá organizace" And Trim(Label3(1)) = "Zdravotní pojištovny" And OldRok > 2010 Then
       
         For PocPodrizena = 1 To MaxPocPodrizena
          For i = 1 To PocPrvku + 1
          
         ' MsgBox (PojNazevZkratka(i))
         '   If Trim(ZdrKod(i)) = Trim(KodPojjMesic(PocPodrizena)) And Trim(KodPojjMesic(PocPodrizena)) <> "" Then
            If ZdrZkratka(i) = Trim(PodrizenaVyber(PocPodrizena)) Then
               For a = 1 To PocPrvku + 1
             '  MsgBox (Trim(ZdrKod(i)) & " = " & Trim(KodPojjMesic(a)) & "  = " & Sum272(a))
                If Trim(ZdrKod(i)) = Trim(KodPojjMesic(a)) Then
                    SumaVymZakl = Sum272(a) '+ Sum476(a)
                    VymZaklZam = Sum272(a) '+ Sum476(a)
                    UhrnVymZakl = Sum272(a)
                    UhrnVratka = Sum476(a)
                    UhrnPojZam = Sum277(a)
                    UhrnPojOrg = Sum274(a)
                    UhrnPocetZam = PocLidi(a)
                    ZdrKodTisk(PocPodrizena) = KodPojjMesic(a)
                    UhrnPojC = UhrnPojZam + UhrnPojOrg
                    Exit For
                End If
               Next a
                    ZdrNazevTisk(PocPodrizena) = ZdrNazev(i)
                    ZdrZkratkaTisk(PocPodrizena) = ZdrZkratka(i)
                    ZdrUliceTisk(PocPodrizena) = ZdrUlice(i)
                    ZdrObecTisk(PocPodrizena) = ZdrObec(i)
                    ZdrCisloDomuTisk(PocPodrizena) = ZdrCisloDomu(i)
                    ZdrPscTisk(PocPodrizena) = ZdrPsc(i)
                    ZdrTelefonTisk(PocPodrizena) = ZdrTelefon(i)
                    ZdrCisloUctuTisk(PocPodrizena) = ZdrCisloUctu(i)
                    ZdrKodBankyTisk(PocPodrizena) = ZdrKodBanky(i)
                    ZdrKsTisk(PocPodrizena) = ZdrKs(i)
                    ZdrVsTisk(PocPodrizena) = ZdrVs(i)
                    ZdrSsTisk(PocPodrizena) = ZdrSs(i)
                    ZdrKodTisk(PocPodrizena) = ZdrKod(i)
                    Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
               
              Exit For
             End If
           Next i
                PrgDil1 = PrgDil1 + PocDil1
                ProgressBar2.Value = PrgDil1
              If SumaVymZakl > 0 Then: NactiZdrRekap
         Next PocPodrizena

     Else

            For PocPodrizena = 1 To MaxPocPodrizena
                    Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
                    PrgDil1 = PrgDil1 + PocDil1
                    ProgressBar2.Value = PrgDil1
                    If Trim(PodrizenaVyber(PocPodrizena)) = "" Then: Frame2.Caption = " Celá organizace"
                NactiZdrUdaje
                If MaxRsTrideni > 0 Then: NactiZdrRekap
                
            Next PocPodrizena
     End If
 Close #1
            
            
 ProgressBar1.Visible = False
 ProgressBar2.Visible = False
Frame2.Visible = False
FrmZdrPoj.MousePointer = 0
End Sub

Private Sub Command5_Click()
FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1

NactiFirmu
Command9.Visible = True
FrmZdrPoj.MousePointer = 11
    OdMes = mesicInt(Combo1)
    DoMes = mesicInt(Combo1)
    OldRok = Val(Combo2)
 If Sestava = "" Then
  RidiciPol = 0
  RidiciVyber = " "
  PodrizenaPol = 0
  PocPodrizena = 1
  Strana = 1
  PodrizenaVyber(1) = " "
  Label2(0).Caption = "000"
  Label2(1).Caption = "Celá organizace            "
  Label3(0).Caption = "002"
  Label3(1).Caption = "Zdravotní pojištovny          "
 End If
  
    Frame2.Visible = True
    ProgressBar1.Visible = True
 ProgressBar2.Visible = True
   ProgressBar2.Value = 100
   PocDil1 = Fix(100 / MaxPocPodrizena)
           PrgDil1 = 0
          
Text2 = "": Tisk = "": Sestava = "Zdr.poj. - rozdíl (HM <> Vym. základ)"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update

Frame2.Caption = "  Načítám rozdíly mezi hrubou mzdou a vyměřovcím základem na  zdravotní pojištění "
   NactiZdrUdaje
   If PocRadek > 0 And Sestava = "Zdr.poj. - rozdíl (HM <> Vym. základ)" Then: NactiRozdil
  Close #1
 ProgressBar1.Visible = False
 ProgressBar2.Visible = False
 Frame2.Visible = False
FrmZdrPoj.MousePointer = 0
End Sub

Private Sub Command6_Click()
TestnaZdr205 = 1
  FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1
Command9.Visible = True
If PrepocetAnoNe = True Then: ResultText(1) = "X": FrmPrepocVsem.Display
NactiFirmu
FrmZdrPoj.MousePointer = 11

 '   OdMes = mesicInt(Combo1)
'    DoMes = mesicInt(Combo1)
'    OldRok = Val(Combo2)
    
 OtevriDataMesic
    
    
   Frame2.Visible = True
   ProgressBar2.Visible = True
   ProgressBar1.Visible = True
ProgressBar2.Value = 100
  Text2 = "": Tisk = "":
  
  Sestava = "Zdr.poj. - přehled o platbě"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update
   PocDil1 = Fix(100 / MaxPocPodrizena)
           PrgDil1 = 0
           
         
   
   If Trim(Label2(1)) = "Celá organizace" And Trim(Label3(1)) = "Zdravotní pojištovny" And OldRok > 2010 Then
        For PocPodrizena = 1 To MaxPocPodrizena
          For i = 1 To PocPrvku + 1
             If ZdrZkratka(i) = Trim(PodrizenaVyber(PocPodrizena)) Then
               For a = 1 To PocPrvku + 1
                    SumaVymZakl = 0
                    VymZaklZam = 0
                    UhrnVymZakl = 0
                    UhrnVratka = 0
                    UhrnPojZam = 0
                    UhrnPojOrg = 0
                    UhrnPocetZam = 0
                    UhrnPojC = 0
              
                If Trim(ZdrKod(i)) = Trim(KodPojjMesic(a)) Then
                    SumaVymZakl = Sum272(a) ' + Sum476(a)
                    VymZaklZam = Sum272(a) ' + Sum476(a)
                    UhrnVymZakl = Sum272(a)
                    UhrnVratka = Sum476(a)
                    UhrnPojZam = Sum277(a)
                    UhrnPojOrg = Sum274(a)
                    UhrnPocetZam = PocLidi(a)
                    ZdrKodTisk(PocPodrizena) = KodPojjMesic(a)
                    UhrnPojC = UhrnPojZam + UhrnPojOrg
                    Exit For
                End If
               Next a
                    ZdrNazevTisk(PocPodrizena) = ZdrNazev(i)
                    ZdrZkratkaTisk(PocPodrizena) = ZdrZkratka(i)
                    ZdrUliceTisk(PocPodrizena) = ZdrUlice(i)
                    ZdrObecTisk(PocPodrizena) = ZdrObec(i)
                    ZdrCisloDomuTisk(PocPodrizena) = ZdrCisloDomu(i)
                    ZdrPscTisk(PocPodrizena) = ZdrPsc(i)
                    ZdrTelefonTisk(PocPodrizena) = ZdrTelefon(i)
                    ZdrCisloUctuTisk(PocPodrizena) = ZdrCisloUctu(i)
                    ZdrKodBankyTisk(PocPodrizena) = ZdrKodBanky(i)
                    ZdrKsTisk(PocPodrizena) = ZdrKs(i)
                    ZdrVsTisk(PocPodrizena) = ZdrVs(i)
                    ZdrSsTisk(PocPodrizena) = ZdrSs(i)
                    ZdrKodTisk(PocPodrizena) = ZdrKod(i)
                    Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
               
              Exit For
             End If
          Next i
                PrgDil1 = PrgDil1 + PocDil1
                ProgressBar2.Value = PrgDil1
              If Trim(ZdrKodTisk(PocPodrizena)) <> "" Then: NactiZdrPrehled
         Next PocPodrizena
           MsgBox ("Soubory Přehled pojisného pro jednotlivé zdravotní pojišťovny " & PathXml & " jsou vytvořeny.")

   Else
   
         For PocPodrizena = 1 To MaxPocPodrizena
   'MsgBox (PocPodrizena)
           Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
           PrgDil1 = PrgDil1 + PocDil1
           ProgressBar2.Value = PrgDil1
          NactiZdrUdaje
           If UhrnPocetZam + UhrnVymZakl + UhrnVratka > 0 Then: NactiZdrPrehled
         Next PocPodrizena
   End If
   Close #1
 ProgressBar1.Visible = False
 ProgressBar2.Visible = False
 Frame2.Visible = False
FrmZdrPoj.MousePointer = 0



End Sub

Private Sub Command67_Click()
    ResultCislo(19) = Val(Combo2.Text)
    ResultCislo(18) = mesicInt(Combo1.Text)
    FrmExcelKontrZdr.Display

End Sub

Private Sub Command7_Click()

If Kontrolor = 0 Then
  FNameT = PathExe & "\Tisk.txt"  ' pomocnej soubor pro tisk
                Open FNameT For Output As #1 ' otevři soubor s číslem 1

   NactiFirmu
   Command9.Visible = True
   FrmZdrPoj.MousePointer = 11
    OdMes = mesicInt(Combo1)
    DoMes = mesicInt(Combo1)
    OldRok = Val(Combo2)
    PocRad = 0
    
    For i = 1 To PocetZamA: ZamVeta(i) = "": Next i
    Frame2.Visible = True
    ProgressBar1.Visible = True
    ProgressBar2.Visible = True
' ProgressBar2.Value = 100
     Text2 = "": Tisk = "": Sestava = "Zdr.poj. - zaměstnanci"
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update
End If

If Kontrolor = 0 Then: PocDil1 = Fix(100 / MaxPocPodrizena): PrgDil1 = 0
 
       
   For PocPodrizena = 1 To MaxPocPodrizena
                    Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
                   If Kontrolor = 0 Then
                    PrgDil1 = PrgDil1 + PocDil1
                    ProgressBar2.Value = PrgDil1
                   End If
         If Trim(PodrizenaVyber(PocPodrizena)) = "" Then: Frame2.Caption = " Celá organizace"
        Strana = 0
       NactiZdrUdaje
                   
       
       'MsgBox (PodrizenaVyber(PocPodrizena))
      ' If PocRadek > 0 Then: NactiZdrZam: PocRadek = 0
   Next PocPodrizena
If Kontrolor = 0 Then
 Close #1
    Frame2.Visible = False
    ProgressBar1.Visible = False
 ProgressBar2.Visible = False
FrmZdrPoj.MousePointer = 0
End If

End Sub

Private Sub Command8_Click()
'Grid3.Visible = False
  Frame1.Visible = False

'Command8.Visible = False
'Command11 = False
    UkazHlediska
End Sub

Private Sub Command9_Click()

' ULOZENI DULEZITY --------
Dim rsDulezite As Recordset
Dim TextPodrid As String
Dim TextVyber As String
Dim NaplnPodrid As String
Dim NaplnVyber As String
Dim MaxRsDulezite As Integer

MaxRsDulezite = 0

Set rsDulezite = DbUdaje.OpenRecordset("SELECT * FROM Dulezite ") ' otevři tabulku
    On Error Resume Next
    rsDulezite.MoveLast
    On Error Resume Next
    rsDulezite.MoveFirst
    On Error Resume Next
    MaxRsDulezite = rsDulezite.RecordCount

    rsDulezite.AddNew ' přidej nový záznam
        
     rsDulezite.Fields("Tisk") = "Ano"
     rsDulezite.Fields("Kopie") = 1
     rsDulezite.Fields("Poradi") = MaxRsDulezite + 1
     rsDulezite.Fields("Nazev") = Sestava + Str(Den) + Str(Mesic) + Str(Rok) + Format(Time, "General Date")
     rsDulezite.Fields("NazevOrig") = Sestava
     rsDulezite.Fields("Ridici") = Trim(Label2(1))
    
     rsDulezite.Fields("NazevFrm") = "TiskZdrSestavy"
     rsDulezite.Fields("Rezerva1") = "o"
     rsDulezite.Fields("RidiciCislo") = RidiciPol
     rsDulezite.Fields("PodridiciCislo") = PodrizenaPol
     rsDulezite.Fields("Podrizena") = Trim(Label3(1))
     For a = 1 To 50 ' podřídící hlediska a znaky označení - 1 až 50
      
       TextVyber = "Vyber" & Trim(Str(a))
       TextPodrid = "Podridici" & Trim(Str(a))
      If a <= PocPrvku Then
            Grid3.Row = a
            Grid3.Col = 1
       NaplnVyber = Grid3.Text '  když je vybraný
       
       Grid3.Col = 2
       NaplnPodrid = Grid3.Text
      Else
       NaplnPodrid = " "
       NaplnVyber = " "
      End If
      
      If Trim(NaplnVyber) = "" Then: NaplnVyber = " "
      If Trim(NaplnPodrid) = "" Then: NaplnPodrid = " "
       
      rsDulezite.Fields(TextVyber) = NaplnVyber
      rsDulezite.Fields(TextPodrid) = NaplnPodrid
     Next a
    
    Command9.Visible = False
    rsDulezite.Update   ' uložho přidaný
    rsDulezite.Close    ' zavři recordset
' KONEC UKLADANI ----------
Sestava = "Uloženo pod názvem : " + Sestava + Str(Den) + Str(Mesic) + Str(Rok) + Format(Time, "General Date")
    MsgBox (Sestava)
End Sub
Private Sub OtevriDataMesic()
Dim MesR As Integer
Dim RokR As Integer
Dim s As String

MesR = mesicInt(Combo1)
RokR = Combo2

OdMes = MesR
DoMes = MesR
OldMes = MesR
OldRok = RokR

    MaxRsPojjSum = 0
    PocSestav = 0
    
    
     DbUdjR = "U" & Format(Datum(1, MesR, RokR), "yyyymm") & ".mdb"
     PathDtaR = PathAdres + "\" + DbUdjR ' cesta k datům pam, osobní

      Set DbUdajeR = OpenDatabase(PathDtaR, True, False, ";pwd=Salary21") ' otevři databázi s údaji zaměstnanců

           On Error GoTo NeniPojR
            Set RsPojjSum = DbUdajeR.OpenRecordset("SELECT * FROM PojSum")
              On Error GoTo NeniPojR
             RsPojjSum.MoveLast
             RsPojjSum.MoveFirst
             MaxRsPojjSum = RsPojjSum.RecordCount
            For a = 1 To MaxRsPojjSum ' přes všechny pojistovny
             KodPojjMesic(a) = RsPojjSum.Fields("PojKod")
             PocLidi(a) = RsPojjSum.Fields("PocetLidi") ' počet lidí v pojištovně
             Sum272(a) = RsPojjSum.Fields("ZpVymZakl")       ' 272
             Sum471(a) = RsPojjSum.Fields("ZpZaklOrgCely")  ' 471
             Sum472(a) = RsPojjSum.Fields("ZpZaklZamCely")  ' 472
             Sum476(a) = RsPojjSum.Fields("ZaklVratka")   ' 476
             Sum274(a) = RsPojjSum.Fields("ZdrPojOrgan")   ' 274
             Sum277(a) = RsPojjSum.Fields("ZdrPojZam")       ' 277
             Sum354(a) = RsPojjSum.Fields("VymZaklNadLimit") ' 354
             
             If KodPojjMesic(a) = "205x" Then ' tvorba xml souboru
               Dim U As String
               Dim FNamePoj205 As String
               Dim VetaPoj205 As String
               
                U = Chr$(34) ' Uvozovky
                ' FNamePoj205 = PathXml & "\205Prumyslova" & FirmaIco & ".xml"



                FNamePoj205 = PathXml & "\205Prumyslova.xml" '2023
                 Open FNamePoj205 For Output As #3
                 
                  VetaPoj205 = "<?xml version=" & U & "1.0" & U & " encoding=" & U & "windows-1250" & U & "?>"
                   Print #3, VetaPoj205
                  VetaPoj205 = "<?xfa generator=" & U & "XFA2_4" & U & " APIVersion=" & U & "2.6.7120.0" & U & "?>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<xdp:xdp xmlns:xdp=" & U & "http://ns.adobe.com/xdp/" & U & " > "
                   Print #3, VetaPoj205
                 VetaPoj205 = "<xfa:datasets xmlns:xfa=" & U & "http://www.xfa.org/schema/xfa-data/1.0/" & U & ">"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<xfa:data>"
                   Print #3, VetaPoj205

                 VetaPoj205 = "<PPZ>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<Formular>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<Typ>1</Typ>"
                   Print #3, VetaPoj205

                 VetaPoj205 = "<ZamNaz>" & FirmaNazev & "</ZamNaz>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<ZamUli>" & FirmaUlice & "</ZamUli>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<ZamCpCo>" & FirmaAdrCislo & "</ZamCpCo>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<ZamPSC>" & FirmaPSC & "</ZamPSC>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<ZamObe>" & FirmaAdrObec & "</ZamObe>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<ZamIC>" & FirmaIco & "</ZamIC>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<ZamTel>" & FirmaTelefon & "</ZamTel>"
                   Print #3, VetaPoj205
                   
                   'MesR = mesicInt(Combo1)
                   s = Trim(Str(OldMes)) + "/" + Trim(Str(OldRok))
                   If OldMes < 10 Then: s = "0" + s


                 VetaPoj205 = "<ObdHla>" & s & "</ObdHla>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<PocZam>" & PocLidi(a) & "</PocZam>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<VymZak>" & Sum272(a) & "</VymZak>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<SumPoj>" & Sum274(a) + Sum277(a) & "</SumPoj>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<DatVyp>" & Format(Datum(Den, Mesic, Rok), "dd/mm/yyyy") & "</DatVyp>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "</Formular>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "</PPZ>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "</xfa:data>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "</xfa:datasets>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "<pdf href=" & U & "PPZ_2011_v30_CPZP_web.pdf" & U & " xmlns=" & U & "http://ns.adobe.com/xdp/pdf/" & U & "/>"
                   Print #3, VetaPoj205
                 VetaPoj205 = "</xdp:xdp>"
                   Print #3, VetaPoj205
                  Close #3
                   If TestnaZdr205 = 1 Then
                    MsgBox ("Soubor " & FNamePoj205 & " je utvořen.")
                   TestnaZdr205 = 0
                   End If
             End If ' If KodPojjMesic(a) = "205" Then ' tvorba xml souboru
             
             
            RsPojjSum.MoveNext
           Next a
           
 Exit Sub
           
NeniPojR:
  MsgBox (" Data nejsou k načtení.")
End Sub
Private Sub Form_Load()
Set frm = FrmZdrPoj
 Rozlis
  Tisk = ""
  TestnaZdr205 = 0
        Grid1.ColWidth(0) = 200 * sizeFactor
        Grid1.ColWidth(1) = 500 * sizeFactor
        Grid1.ColWidth(2) = 2500 * sizeFactor
        Grid1.Row = 0
        Grid1.Col = 0: Grid1.Text = " "
        Grid1.Col = 1: Grid1.Text = " Č. pol "
        Grid1.Col = 2: Grid1.Text = " Název "


        Grid2.ColWidth(0) = 10 * sizeFactor
        Grid2.ColWidth(1) = 500 * sizeFactor
        Grid2.ColWidth(2) = 2500 * sizeFactor
        Grid2.Row = 0
        Grid2.Col = 0: Grid2.Text = " "
        Grid2.Col = 1: Grid2.Text = "Ozn"
        Grid2.Col = 2: Grid2.Text = " Název "

        Grid3.ColWidth(0) = 10 * sizeFactor
        Grid3.ColWidth(1) = 500 * sizeFactor
        Grid3.ColWidth(2) = 2500 * sizeFactor
        Grid3.Row = 0
        Grid3.Col = 0: Grid2.Text = " "
        Grid3.Col = 1: Grid2.Text = "Ozn"
        Grid3.Col = 2: Grid2.Text = " Název "
        

 Text2 = "": Tisk = ""
 
 Command18.Caption = "Období " + Trim(Str(SysMes)) + "/" + Trim(Str(SysRok))
OrigMes = SysMes
 OrigRok = SysRok
 
 
OdMes = SysMes
DoMes = SysMes
OldMes = SysMes
OldRok = SysRok

Combo1 = mesicStr(OrigMes)
Combo2 = OrigRok
    
    MaxRsPojjSum = 0
    PocSestav = 0
    
OtevriDataMesic

           On Error GoTo NeniPoj
            Set RsPojjSum = DbUdaje.OpenRecordset("SELECT * FROM PojSum")
              On Error GoTo NeniPoj
             RsPojjSum.MoveLast
             RsPojjSum.MoveFirst
             MaxRsPojjSum = RsPojjSum.RecordCount
            For a = 1 To MaxRsPojjSum ' přes všechny pojistovny
             KodPojjMesic(a) = RsPojjSum.Fields("PojKod")
             PocLidi(a) = RsPojjSum.Fields("PocetLidi") ' počet lidí v pojištovně
             Sum272(a) = RsPojjSum.Fields("ZpVymZakl")       ' 272
             Sum471(a) = RsPojjSum.Fields("ZpZaklOrgCely")  ' 471
             Sum472(a) = RsPojjSum.Fields("ZpZaklZamCely")  ' 472
             Sum476(a) = RsPojjSum.Fields("ZaklVratka")   ' 476
             Sum274(a) = RsPojjSum.Fields("ZdrPojOrgan")   ' 274
             Sum277(a) = RsPojjSum.Fields("ZdrPojZam")       ' 277
             Sum354(a) = RsPojjSum.Fields("VymZaklNadLimit") ' 354
            RsPojjSum.MoveNext
           Next a
NeniPoj:
Label2(0).Caption = "000"
Label2(1).Caption = "Celá organizace            "
Label3(0).Caption = "002"
Label3(1).Caption = "Zdravotní pojištovny          "
Frame1.Visible = True
Grid2.Visible = True
Sestava = ""
RidiciPol = 0
RidiciVyber = " "
PodrizenaPol = 2
PodrizenaVyber(1) = " "
NactiSeznamPolozky2
For i = 0 To 3 Step 2
  For a = 1 To 12
    Combo3(i).AddItem mesicStr(Val(a))
  Next a
      For a = PosZprRok To (SysRok - 5) Step -1
        Combo3(i + 1).AddItem (a)
      Next a
Next i
      
End Sub
Private Sub NactiFirmu()
                      Set RsFirma = DbUdaje.OpenRecordset("SELECT * FROM Firma ")
                      RsFirma.MoveLast
                      RsFirma.MoveFirst
                      FirmaNazev = Trim(RsFirma.Fields("fnazev"))
                      FirmaIco = Trim(RsFirma.Fields("fico"))
                      FirmaDIC = Trim(RsFirma.Fields("fdic"))
                      FirmaUlice = Trim(RsFirma.Fields("Fulice"))
                      FirmaAdrCislo = Trim(RsFirma.Fields("FAdrCislo"))
                      FirmaPSC = Trim(RsFirma.Fields("FAdrPsc"))
                      FirmaAdrObec = Trim(RsFirma.Fields("FAdrObec"))
                      FirmaTelefon = Trim(RsFirma.Fields("FspojTelP"))
                      
                      FirmaUcet = Trim(RsFirma.Fields("Fbanka1Ucet")) + "/" + Trim(RsFirma.Fields("Fbanka1Kod"))
                      MzdUct = Trim(RsFirma.Fields("FMzdUctarna"))
                      If MzdUct = "" Then: MzdUct = "00"

     Set RsPoj = DbUdaje.OpenRecordset("pojistovny", dbOpenTable) ' oetvře tabulku cesta
     On Error GoTo Dall1 ' když není žádná databáze tam kam ukazuje cesta k datům
 
     RsPoj.MoveLast
     RsPoj.MoveFirst
     MaxZaznam = RsPoj.RecordCount

        If MaxZaznam <> 0 Then
              For i = 1 To MaxZaznam
          
                   If Trim(RsPoj.Fields("PojZkratka")) <> "" Then
                    PocPodrizena = PocPodrizena + 1
                    ZdrZkratka(i) = Trim(RsPoj.Fields("PojZkratka"))
                    ZdrNazev(i) = Trim(RsPoj.Fields("PojNazev"))
                    ZdrUlice(i) = Trim(RsPoj.Fields("PojUlice"))
                    ZdrCisloDomu(i) = Trim(RsPoj.Fields("PojCisloDomu"))
                    ZdrObec(i) = Trim(RsPoj.Fields("PojObec"))
                    ZdrPsc(i) = Trim(RsPoj.Fields("PojPsc"))
                    ZdrTelefon(i) = Trim(RsPoj.Fields("PojTelefon"))
                    ZdrCisloUctu(i) = Trim(RsPoj.Fields("PojCisloUctu"))
                    ZdrKodBanky(i) = Trim(RsPoj.Fields("PojKodBanky"))
                    ZdrKs(i) = Trim(RsPoj.Fields("PojKs"))
                    
                    ZdrVs(i) = Trim(RsPoj.Fields("PojVs"))
                    ZdrSs(i) = Trim(RsPoj.Fields("PojSs"))
                    ZdrKod(i) = Trim(RsPoj.Fields("PojKod"))
                 '  MsgBox (ZdrNazev(i) & " " & ZdrKod(i))     ' RsPoj.MoveFirst
                   End If
                   
                   RsPoj.MoveNext

              Next i
          '    MsgBox (ZdrKod(PocPodrizena))     ' RsPoj.MoveFirst
        End If
Dall1:
End Sub

Public Sub Display()
   Me.Show vbModal
End Sub

Private Sub Grid1_Click()
Text1.Visible = False
Frame5.Visible = False
Command9.Visible = False

If Cmd = 1 Then  ' Řídící hledisko
    Grid1.Col = 1
     Label2(0).Caption = Grid1.Text
      Grid1.Col = 2
      Label2(1).Caption = Grid1.Text
   If Trim(Label2(0).Caption) = "000" Then ' je to Celá organizacev Rídícím hledisku
      Grid1.Col = 2
      Label2(1).Caption = Grid1.Text
       Grid1.Visible = False
       Frame3.Visible = False
       Grid2.Visible = False
       Frame4.Visible = False
       Label2(2).Visible = False
       Command2.Visible = False
   Else
      NactiSeznamPolozky
   End If
   
Else    ' Podřídící hledisko
       Grid2.Visible = False
       Frame4.Visible = False
      Grid1.Col = 1
       Label3(0).Caption = Grid1.Text
      Grid1.Col = 2
       Label3(1).Caption = Grid1.Text
   If Trim(Label3(0).Caption) = "000" Then ' je to Celá organizacevR ídícím hledisku
      Grid1.Col = 2
      Label3(1).Caption = Grid1.Text
       Grid1.Visible = False
       Frame3.Visible = False
         Frame1.Visible = False
       'Grid3.Visible = False
       'Command8.Visible = False
       'Command11 = False
   Else
      Grid1.Visible = False
      Frame3.Visible = False
      NactiSeznamPolozky2
   End If

End If
   
   
   
End Sub
Private Sub NactiSeznamPolozky()
Dim RidHledej As String
  

  PocStr = 1
  If Trim(Label2(0)) = "001" Then ' je to středisko
       Set RsZamSt = DbUdaje.OpenRecordset("Select * From Zamestnanec Where val(mid(Do,7,4)) = ""sysrok"" Or Aktivni = 1 Order by Stredisko")
       RsZamSt.MoveLast
       RsZamSt.MoveFirst
       maxRsZamSt = RsZamSt.RecordCount

     For i = 1 To maxRsZamSt
            If Trim(Grid2.Text) <> Trim(RsZamSt.Fields("Stredisko")) Then
                MaxRsRidici = PocStr
                Grid2.Rows = PocStr + 1
                Grid2.Row = PocStr: Grid2.Col = 2
                Grid2.Text = RsZamSt.Fields("Stredisko")
                PocStr = PocStr + 1
            End If
      RsZamSt.MoveNext
     Next i
     Grid2.Visible = True
     Label2(2).Visible = False
     Frame4.Visible = True
  End If
     
   If Trim(Label2(0)) = "002" Then ' je to Zdr Pojištovny
        Set RsZamSt = DbUdaje.OpenRecordset("Select * From Pojistovny order by PojZkratka") ' oetvře tabulku cesta
            RsZamSt.MoveLast
            RsZamSt.MoveFirst
           maxRsZamSt = RsZamSt.RecordCount
     
     For i = 1 To maxRsZamSt
                Grid2.Rows = i + 1: Grid2.Row = i: Grid2.Col = 2
                Grid2.Text = RsZamSt.Fields("PojZkratka")
         RsZamSt.MoveNext
     Next i
      Grid2.Visible = True
      Frame4.Visible = True
 '    Label2(2) = "Všechny"
  '   Label2(0) = "001"
  '   Label2(1) = "Zdravotní pojištovny          "
  '   Label2(2).Visible = True
     Grid1.Visible = False
     Frame3.Visible = False
     Grid1.Visible = False
    End If
       

     
  If RidiciPol <> "001" And RidiciPol <> "002" Then ' třídící položky z Osobního

     RidHledej = Trim(Label2(0))
     Set RSRidici = DbUdaje.OpenRecordset("SELECT * FROM RozborStr WHERE PoCislo = """ & RidHledej & """ ") '  ORDER BY Napln")
         On Error GoTo Dal
         RSRidici.MoveLast
         RSRidici.MoveFirst
         MaxRsRidici = RSRidici.RecordCount
         
         Grid2.Rows = MaxRsRidici
         
         Grid2.Visible = True
         Frame4.Visible = True
         RSRidici.MoveFirst
         For i = 1 To MaxRsRidici
            Grid2.Row = i
            Grid2.Col = 2
            Grid2.Text = RSRidici.Fields("Nazev")
            RSRidici.MoveNext
         Next i
         
         
  End If
        Grid2.Visible = True
        Frame4.Visible = True
     
Dal:
     If err = 3021 Then: MaxRsRidici = 0
End Sub

Private Sub NactiSeznamPolozky2()
Dim RidHledej As String
  Dim PocStr As Integer
  Dim RsZamSt As Recordset

  PocStr = 1
    
  If Trim(Label3(0)) = "001" Then ' je to středisko
           Set RsZamSt = DbUdaje.OpenRecordset("Select * From Zamestnanec Where val(mid(Do,7,4)) = ""sysrok"" Or Aktivni = 1 Order by Stredisko")
            RsZamSt.MoveLast
            RsZamSt.MoveFirst
             maxRsZamSt = RsZamSt.RecordCount
            PocPrvku = 1
     For i = 1 To maxRsZamSt
            If Trim(Grid3.Text) <> Trim(RsZamSt.Fields("Stredisko")) Then
                MaxRsRidici = PocStr
                Grid3.Rows = PocStr + 1
                Grid3.Row = PocStr: Grid3.Col = 1: Grid3.Text = ""
                Grid3.Row = PocStr: Grid3.Col = 2: Grid3.Text = RsZamSt.Fields("Stredisko")
                PocStr = PocStr + 1
            End If
      RsZamSt.MoveNext
     Next i
     PocPrvku = MaxRsRidici
     'Grid3.Visible = True
       Frame1.Visible = True
       Frame1.Caption = " Výběr střediska "
     'Command8.Visible = True
     'Command11 = True
 End If
 
 If Trim(Label3(0)) = "002" Then ' je to Zdr Pojištovny
        Set RsZamSt = DbUdaje.OpenRecordset("Select * From Pojistovny order by PojZkratka") ' oetvře tabulku cesta
            RsZamSt.MoveLast
            RsZamSt.MoveFirst
           maxRsZamSt = RsZamSt.RecordCount
           PocPrvku = 1
     For i = 1 To maxRsZamSt
               If RsZamSt.Fields("PojZkratka") <> "" Then
                MaxRsRidici = PocPrvku
                Grid3.Rows = PocPrvku + 1: Grid3.Row = PocPrvku: Grid3.Col = 2
                Grid3.Text = RsZamSt.Fields("PojZkratka")
                If Trim(RsZamSt.Fields("PojZkratka")) = "" Then: PocPrvku = PocPrvku - 1
                PocPrvku = PocPrvku + 1
               End If
         RsZamSt.MoveNext
     Next i
     PocPrvku = MaxRsRidici
       Frame1.Visible = True
       Frame1.Caption = " Výběr zdravotní pojištovny"

     'Grid3.Visible = True
     'Command8.Visible = True
     'Command11 = True
     MaxRsRidici = maxRsZamSt
    
 End If


   If Trim(Label3(0)) <> "001" And Trim(Label3(0)) <> "001" And Trim(Label3(0)) <> "002" Then  ' třídící položky z Osobního
     RidHledej = Trim(Label3(0))
     Set RSRidici = DbUdaje.OpenRecordset("SELECT * FROM RozborStr WHERE PoCislo = """ & RidHledej & """ ") '  ORDER BY Napln")
         On Error GoTo Dal
         RSRidici.MoveLast
         RSRidici.MoveFirst
         MaxRsRidici = RSRidici.RecordCount
         
         PocPrvku = MaxRsRidici
         Grid3.Rows = MaxRsRidici + 1
           Frame1.Visible = True
           Frame1.Caption = " Výběr z " + RidHledej + " " + Trim(Label3(1))
         RSRidici.MoveFirst
         For i = 1 To MaxRsRidici
            Grid3.Row = i
            Grid3.Col = 1: Grid3.Text = ""
            Grid3.Col = 2: Grid3.Text = RSRidici.Fields("Nazev")
            RSRidici.MoveNext
         Next i

   End If
         
     
Dal:
     If err = 3021 Then: MaxRsRidici = 0
End Sub

Private Sub Grid2_Click()
 Label2(2).Visible = True
 Command2.Visible = True
 Grid2.Col = 2
 Label2(2).Caption = Grid2.Text
 Grid2.Visible = False
 Frame4.Visible = False
 Grid1.Visible = False
Frame3.Visible = False
End Sub

Private Sub Grid3_Click()
'Label2(2).Visible = True

    Grid3.Col = 1
    If Grid3.Text = "X" Then
       Grid3.Text = ""
       Else
       Grid3.Text = "X"
    End If
End Sub

Private Sub GurhanButton1_Click(Index As Integer)

Dim pojhl As Integer
'Dim RsPamch As Recordset

NactiFirmu
PDF = Index
MaxPojHl = PocPodrizena
FrmZdrPoj.MousePointer = 11
    OdMes = mesicInt(Combo1)
    DoMes = mesicInt(Combo1)
    OldRok = Val(Combo2)
   Frame2.Visible = True
   Command9.Visible = True
   ProgressBar1.Visible = True
   ProgressBar2.Visible = True
   ProgressBar2.Value = 100
     Strana = 1
Text2 = "": Tisk = "":
                Sestava = "Hlášení pro zdr. poj."
          RsPamch.AddNew
            RsPamch.Fields("Datum") = Format(date, "dd/mm/yyyy")
            RsPamch.Fields("Cas") = Format(Time, "hh/mm/ss")
            RsPamch.Fields("OsCislo") = "0"
            RsPamch.Fields("CisPol") = "0"
            RsPamch.Fields("Napln") = Sestava
            RsPamch.Fields("Rezerva") = " "
          RsPamch.Update
                   RidiciPol = 0
                   RidiciVyber = " "
                 '  PodrizenaPol = 2
   PocDil1 = Fix(100 / MaxPocPodrizena)
           PrgDil1 = 0
   
   For PocPodrizena = 1 To MaxPocPodrizena
           Frame2.Caption = " " + PodrizenaVyber(PocPodrizena) + " "
           PrgDil1 = PrgDil1 + PocDil1
           ProgressBar2.Value = PrgDil1
                  ZdrKodHl = " "
               For pojhl = 1 To MaxPojHl
                    If Trim(ZdrZkratka(pojhl)) = Trim(PodrizenaVyber(PocPodrizena)) Then
                      ZdrKodHl = ZdrKod(pojhl)
                      
                     
                     Exit For
                    End If
               
               Next pojhl
           
        
        NactiZdrHlaseni
        
     
   Next PocPodrizena
 
 
 ProgressBar1.Visible = False
 ProgressBar2.Visible = False
 Frame2.Visible = False
FrmZdrPoj.MousePointer = 0
End Sub


Private Sub GurhanButton2_Click(Index As Integer)


PDF = Index
PocRad = 0
                  If PDF = 2 Then
                      InitLpPDF
                         '    DeviceName = "Bullzip PDF Printer"
                         ' If cSetPrinter.SetPrinterAsDefault(DeviceName) Then:
                          'DejNazevFilePDF (NazevSestavyHl)
                  Else
                     Set cdl = CommonDialog1 ' nutné před inicializací - připojuje common dialog
                     InitLp
                      If PrintAno = "N" Then: Exit Sub
                      Printer.FontSize = 11
                  End If





'Set CDL = CommonDialog1 ' nutné před inicializací - připojuje common dialog
 '       InitLp ' funkce
       
        If PrintAno = "N" Then: Exit Sub
       ' Tisk = "Tisk"
     TestZmenPoPrepoctu ("Tisk - " + Sestava)
 If Sestava = "Zdr.poj. - součtová rekapitulace" Then: Printer.FontSize = 12
 If Sestava = "Zdr.poj. - přehled o platbě" Then: Printer.FontSize = 10
 If Sestava = "Činnost malého rozsahu" Then: Printer.FontSize = 12
 If Sestava = "Zdr.poj. - zaměstnanci" Then Printer.FontSize = 8
 If Sestava = "Zdr.poj. - rozdíl (HM <> Vym. základ)" Then: Printer.FontSize = 8
 If Sestava = "Zdr.poj. - rekapitulace pojištoven" Then: Printer.FontSize = 12
 
  If PDF = 2 Then: NazevSestavyHl = Sestava: DejNazevFilePDF (NazevSestavyHl)
 
 
 
        F = FreeFile
              
           If FileExist(FNameT) Then ' když je soubor vytvořen
           
             
        
            For PocetKopii = 1 To cdl.Copies
                 Open FNameT For Input As #F  ' otevři soubor
                Zac = 0
                  Do While Not EOF(F)
                  PocRad = PocRad + 1
                     Line Input #F, veta  ' načti řádek ze souboru
                     If InStr(1, veta, "strana") > 0 And Zac = 1 Then: Printer.EndDoc: For i = 1 To 3: Printer.Print " ": Next i
                     
                     If Zac = 0 Then: For i = 1 To 3: Printer.Print " ": Next i
                     
                      If Mid(veta, 1, 3) <> "***" Then
                       Printer.Print veta ' tiskni načtený řádek
                      Else
                       PocRad = PocRad - 1
                      End If
                     If Mid(veta, 1, 6) = "***111" Then: Prehled
                     If Mid(veta, 1, 6) = "***201" Then: Prehled
                     If Mid(veta, 1, 6) = "***205" Then: Prehled
                     If Mid(veta, 1, 6) = "***207" Then: Prehled
                     If Mid(veta, 1, 6) = "***209" Then: Prehled
                     If Mid(veta, 1, 6) = "***211" Then: Prehled
                     If Mid(veta, 1, 6) = "***213" Then: Prehled
                     If Mid(veta, 1, 6) = "***217" Then: Prehled
                     If Mid(veta, 1, 6) = "***228" Then: Prehled
                     
                     If Sestava = "Zdr.poj. - přehled o platbě" And PocRad = 62 Then: Printer.NewPage: For i = 1 To 3: Printer.Print " ": Next i: PocRad = 0
                     If Sestava = "Zdr.poj. - přehled o platbě" And PocRad = 31 Then: For i = 1 To 1: Printer.Print " ": Next i
                     Zac = 1
                       
                  Loop
                    Close F  ' zavři soubor
             If Sestava <> "Zdr.poj. - přehled o platbě" Then: Printer.EndDoc ' zavři tiskárnu
             If PocRad = 31 And Sestava = "Zdr.poj. - přehled o platbě" Then: Printer.EndDoc
             Next PocetKopii
            
           Else
           
           MsgBox ("Soubor tisku není vytvořen")
           End If
           
    ' TiskZdrSestavy

End Sub

Private Sub Image1_Click(Index As Integer)
Unload FrmZdrPoj
End Sub

Private Sub Combo2_Change()
For a = 1 To 12
 Combo1.AddItem mesicStr(Val(a))
Next a
For a = PosZprRok To (SysRok - 1) Step -1
        Combo2.AddItem (a)
      Next a
      
End Sub

Private Sub UkazHlediska()

For i = 1 To 200: PodrizenaVyber(i) = "": Next i: Text1 = ""
PocPodrizena = 0: MaxPocPodrizena = 0
Text1.Font.Size = 6 * sizeFactor
Text1.Visible = True
Frame5.Visible = True

 ' MsgBox (Trim(Label2(0).Caption))

   RidiciPol = Val(Trim(Label2(0).Caption))
   RidiciVyber = Trim(Label2(2).Caption)
   PodrizenaPol = Val(Trim(Label3(0).Caption))
   
   Text1.Text = Text1.Text & "Prvek řídícího hlediska    : " & vbCrLf
   Text1.Text = Text1.Text & "   " & RidiciVyber & vbCrLf
   Text1.Text = Text1.Text & String$(20, 151) & vbCrLf
   Text1.Text = Text1.Text & "Podřízená třídící hlediska : " & vbCrLf
   TrVyberVse = "N"
    If PocPrvku > 0 Then
        For i = 1 To PocPrvku
            Grid3.Row = i
            Grid3.Col = 1

                 If Grid3.Text = "X" Then
                    Grid3.Col = 2
                    PocPodrizena = PocPodrizena + 1
                    PodrizenaVyber(PocPodrizena) = Grid3.Text
                    Text1.Text = Text1.Text & "   " & PodrizenaVyber(PocPodrizena) & vbCrLf
                    
                End If
        Next i
      If i - 1 = PocPodrizena Then: TrVyberVse = "A"
      MaxPocPodrizena = PocPodrizena
    End If
    If MaxPocPodrizena = 0 Then: MaxPocPodrizena = 1
End Sub


Private Sub Prehled()
 Dim s As String
 
 X = 1
  ' Printer.FontName = "Courier New" '"Salary" '
   Printer.FontBold = True
   Printer.FontSize = 11
 If Mid(veta, 1, 6) = "***207" Then: Printer.PaintPicture Picture1.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0
 If Mid(veta, 1, 6) = "***111" Then: Printer.PaintPicture Picture2.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 2
 If Mid(veta, 1, 6) = "***201" Then: Printer.PaintPicture Picture3.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 3
 If Mid(veta, 1, 6) = "***211" Then: Printer.PaintPicture Picture4.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 4
 If Mid(veta, 1, 6) = "***205" Then: Printer.PaintPicture Picture5.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 5
 If Mid(veta, 1, 6) = "***209" Then: Printer.PaintPicture Picture15.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 6
 If Mid(veta, 1, 6) = "***213" Then: Printer.PaintPicture Picture16.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 7
 If Mid(veta, 1, 6) = "***217" Then: Printer.PaintPicture Picture18.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 8
 If Mid(veta, 1, 6) = "***228" Then: Printer.PaintPicture Picture17.Picture, 0, 0, 11600, 8250: OkrajLevy = 0: OkrajHorni = 0: X = 9

 
  'Printer.EndDoc: Exit Sub
  
  If Check1 = 1 Then
   Printer.CurrentX = 6450: Printer.CurrentY = 2270: Printer.Print "X"
  Else
   Printer.CurrentX = 5580: Printer.CurrentY = 2270: Printer.Print "X"
  End If
  
  Printer.CurrentX = 800: Printer.CurrentY = 3300: Printer.Print FirmaNazev
  
  Printer.CurrentX = 800: Printer.CurrentY = 3860: Printer.Print FirmaUlice
  i = InStr(1, FirmaAdrCislo, "/")
   If i > 0 Then
     Printer.CurrentX = 7000: Printer.CurrentY = 3860: Printer.Print Mid(FirmaAdrCislo, 1, i - 1)
     Printer.CurrentX = 7800: Printer.CurrentY = 3860: Printer.Print Mid(FirmaAdrCislo, i + 1, 5)
   Else
     Printer.CurrentX = 7000: Printer.CurrentY = 3860: Printer.Print FirmaAdrCislo
   End If
  Printer.CurrentX = 9300: Printer.CurrentY = 3860: Printer.Print FirmaIco + Space(2) + MzdUct
  
  Printer.CurrentX = 800: Printer.CurrentY = 4520: Printer.Print FirmaPSC
  Printer.CurrentX = 2200: Printer.CurrentY = 4520: Printer.Print FirmaAdrObec
  Printer.CurrentX = 9500: Printer.CurrentY = 4520: Printer.Print FirmaTelefon

s = Trim(Str(OldMes)) + "/" + Trim(Str(OldRok))
 If OldMes < 10 Then: s = "0" + s
  Printer.CurrentX = 3500: Printer.CurrentY = 5320: Printer.Print s
  Printer.CurrentX = 7000: Printer.CurrentY = 5320: Printer.Print ZVymZakl(X)

  Printer.CurrentX = 2500: Printer.CurrentY = 5900: Printer.Print ZPocZam(X)
  Printer.CurrentX = 7000: Printer.CurrentY = 5900: Printer.Print ZPoj(X)

  Printer.CurrentX = 4400: Printer.CurrentY = 7700: Printer.Print Format(Datum(Den, Mesic, Rok), "dd/mm/yyyy")
  Printer.EndDoc
 ' Printer.FontName = "Salary" '
   Printer.FontBold = False
   
End Sub

Private Sub PrehledVZP()
   
   Printer.FontSize = 11
   Printer.PaintPicture Picture2.Picture, 0, 0, 11500, 8250: OkrajLevy = 0: OkrajHorni = 0
 ' Printer.EndDoc: Exit Sub
  
  Printer.CurrentX = 800: Printer.CurrentY = 2300: Printer.Print FirmaNazev
  Printer.CurrentX = 800: Printer.CurrentY = 2550: Printer.Print Space(Len(FirmaPSC)) + " " + FirmaUlice + " " + FirmaAdrCislo
  Printer.CurrentX = 800: Printer.CurrentY = 2800: Printer.Print FirmaPSC + " " + FirmaAdrObec

  Printer.CurrentX = 7950: Printer.CurrentY = 2750: Printer.Print FirmaIco + Space(6) + MzdUct
  
  Printer.CurrentX = 8200: Printer.CurrentY = 3500: Printer.Print Space(5) + Str(OldMes) + "/" + Str(OldRok)
  Printer.CurrentX = 8300: Printer.CurrentY = 3900: Printer.Print ZPocZam(2)
  Printer.CurrentX = 8300: Printer.CurrentY = 4350: Printer.Print ZVymZakl(2)
  Printer.CurrentX = 8300: Printer.CurrentY = 4800: Printer.Print ZPoj(2)
  
  Printer.CurrentX = 2750: Printer.CurrentY = 5900: Printer.Print Format(Datum(Den, Mesic, Rok), "dd/mm/yyyy")
  Printer.EndDoc
  
End Sub
Private Sub PrehledVoZP()
   
   Printer.FontSize = 11
   Printer.PaintPicture Picture3.Picture, 0, 0, 11500, 8250: OkrajLevy = 0: OkrajHorni = 0
'  Printer.EndDoc: Exit Sub
  
  Printer.CurrentX = 1100: Printer.CurrentY = 2750: Printer.Print FirmaNazev
  Printer.CurrentX = 1100: Printer.CurrentY = 2950: Printer.Print Space(Len(FirmaPSC)) + " " + FirmaUlice + " " + FirmaAdrCislo
  Printer.CurrentX = 1100: Printer.CurrentY = 3200: Printer.Print FirmaPSC + " " + FirmaAdrObec

  
  Printer.CurrentX = 7950: Printer.CurrentY = 3150: Printer.Print FirmaIco + Space(6) + MzdUct
  
  Printer.CurrentX = 8200: Printer.CurrentY = 3850: Printer.Print Space(5) + Str(OldMes) + "/" + Str(OldRok)
  Printer.CurrentX = 8300: Printer.CurrentY = 4150: Printer.Print ZPocZam(3)
  Printer.CurrentX = 8300: Printer.CurrentY = 4500: Printer.Print ZVymZakl(3)
  Printer.CurrentX = 8300: Printer.CurrentY = 4900: Printer.Print ZPoj(3)
  
  Printer.CurrentX = 2200: Printer.CurrentY = 5550: Printer.Print Format(Datum(Den, Mesic, Rok), "dd/mm/yyyy")


  Printer.EndDoc

End Sub
Public Sub TiskZdrSestavy()
Dim rsDulezite As Recordset
Dim MaxRsDulezite As Integer
Dim PoradiT As Integer

    Tisk = "Tisk": Dulezite = 0
 'MsgBox (Sestava & " bbb " & RidiciVyber & " xxx " & PodrizenaPol & " xxx " & PodrizenaVyber(PocPodrizena) & " xxx " & OldMes & " xxx " & OldRok)
Sestava = ""
  If Trim(Sestava) = "" Then
    Dulezite = 1
     PoradiT = ResultCislo(1)

   MaxRsDulezite = 0
   Set rsDulezite = DbUdaje.OpenRecordset("SELECT * FROM Dulezite WHERE Poradi = " & PoradiT & " ") ' otevři tabulku
 
    On Error Resume Next
    rsDulezite.MoveFirst

    NactiFirmu
    Sestava = rsDulezite.Fields("NazevOrig")                           ' konkrétní tisková sestava
   ' MsgBox (Sestava & "semtu" & PodrizenaVyber(MaxPocPodrizena))
    RidiciPol = rsDulezite.Fields("RidiciCislo")                   ' číslo řídící položky
    RidiciVyber = rsDulezite.Fields("Ridici")                      ' řídící hledisko
     If RidiciPol = 0 Then: RidiciVyber = " "
    PodrizenaPol = rsDulezite.Fields("PodridiciCislo")                ' číslo podřídící položky
     MaxPocPodrizena = 0
     For a = 1 To 50 ' podřídící hlediska a znaky označení - 1 až 50
      If rsDulezite.Fields("Vyber" & Trim(Str(a))) = "X" Then
       MaxPocPodrizena = MaxPocPodrizena + 1
        PodrizenaVyber(MaxPocPodrizena) = rsDulezite.Fields("Podridici" & Trim(Str(a)))
      End If
      If Trim(rsDulezite.Fields("Podridici" & Trim(Str(a)))) = "" Then: Exit For
     Next a
     

      If MaxPocPodrizena = 0 Then: MaxPocPodrizena = 1
    OdMes = SysMes
    DoMes = SysMes
    OldRok = SysRok
    
  End If
   '   MsgBox (Sestava & " " & MaxRsTrideni & "  " & RidiciPol & "  " & RidiciVyber & "  " & PodrizenaPol & "  " & PodrizenaVyber(PocPodrizena))

   If Sestava = "Zdr.poj. - rekapitulace pojištoven" Then
     SumaVymZakl = 0: PrumPocZamSoc = 0: UhrnPojC = 0: UhrnPocetZam = 0:  UhrnVymZakl = 0: UhrnVratka = 0:  PocRadek = 0: UhrnPojZam = 0: UhrnPojOrg = 0
 
       For PocPodrizena = 1 To MaxPocPodrizena
         NactiZdrUdaje
          If PocRad > 52 Then: Printer.EndDoc: PocRad = 0
         NactiZdrRekap
       Next PocPodrizena
 
       If PocRad > 0 Then: Printer.EndDoc: PocRad = 0
   End If
   
   If Sestava = "Zdr.poj. - zaměstnanci" Then
       For PocPodrizena = 1 To MaxPocPodrizena

         NactiZdrUdaje

         NactiZdrZam

          If PocRad > 59 Then: Printer.EndDoc: PocRad = 0
       Next PocPodrizena
   End If
   
   
   If Sestava = "Zdr.poj. - rozdíl (HM <> Vym. základ)" Then

           NactiZdrUdaje
           NactiRozdil
   End If
   If Sestava = "Činnost malého rozsahu" Then
       For PocPodrizena = 1 To MaxPocPodrizena
              NactiZdrUdaje
       If PocRadek > 0 Then: NactiZdrCMR: PocRadek = 0
      Next PocPodrizena
   End If
    
   
   If Sestava = "Zdr.poj. - přehled o platbě" Then

        For PocPodrizena = 1 To MaxPocPodrizena
          NactiZdrUdaje
       
          'If UhrnPocetZam + UhrnVymZakl + UhrnVratka > 0 Then:
          NactiZdrPrehled
  '        MsgBox (xx & "  3" & Tisk)
        Next PocPodrizena
   End If
    
   If Sestava = "Zdr.poj. - součtová rekapitulace" Then
          SumaVymZakl = 0: PrumPocZamSoc = 0: UhrnPojC = 0:: UhrnPocetZam = 0: UhrnVymZakl = 0: UhrnVratka = 0:   PocRadek = 0: UhrnPojZam = 0: UhrnPojOrg = 0
               For PocPodrizena = 1 To MaxPocPodrizena
                NactiZdrUdaje
               Next PocPodrizena
                    PocPodrizena = 30
                    ZdrZkratka(PocPodrizena) = ""
                    ZdrNazev(PocPodrizena) = "Součet všech pojištoven"
                    ZdrUlice(PocPodrizena) = ""
                    ZdrCisloDomu(PocPodrizena) = ""
                    ZdrObec(PocPodrizena) = ""
                    ZdrPsc(PocPodrizena) = ""
                    ZdrTelefon(PocPodrizena) = ""
                    ZdrCisloUctu(PocPodrizena) = ""
                    ZdrKodBanky(PocPodrizena) = ""
                    ZdrKs(PocPodrizena) = ""
                    ZdrVs(PocPodrizena) = ""
                    ZdrSs(PocPodrizena) = ""
                    ZdrKod(PocPodrizena) = ""
      'MsgBox (UhrnPocetZam + UhrnVymZakl + UhrnVratka)
            'If UhrnPocetZam + UhrnVymZakl + UhrnVratka > 0 Then:
            NactiZdrRekap
            
                   
   End If
If Dulezite = 1 Then: Sestava = ""
     Tisk = ""
End Sub
Public Sub Utvor_HOZ_XML_StaryStyl()
    Dim U As String
    Dim Z_InterniIdentifikator As String
    
    U = Chr(34)
    FVystup = PathXml & "/" & NazevSestavyHl & "_" & NazevSestavyHl & "_Hromadne_" & Combo1 & "_" & Combo2 & ".xml"
           

     Open FVystup For Output As #3
    
     ' Otevření XML a hlavička
    Print #3, "<?xml version=" & U & "1.0" & U & " encoding=" & U & "windows-1250" & U & "?>"
    Print #3, "<hromadneOznameniZamestnavatele xmlns=" & U & "http://xmlns.vzp.cz/hromadneOznameniZamestnavatele/v1" & U & ">"
    
    Print #3, "  <identifikacePredmetuPodaniText>Hromadné oznámení zaměstnavatele 2026+</identifikacePredmetuPodaniText>"
    Print #3, "  <identifikacePredmetuPodaniKod>882b97d9-3a41-4552-887b-3942ae92c3ea</identifikacePredmetuPodaniKod>"
    Z_InterniIdentifikator = Trim(ZdrKodHl) & mesicInt(Combo1) & Combo2 & Replace(Now, ":", "")
    Z_InterniIdentifikator = Replace(Z_InterniIdentifikator, " ", "")
    Z_InterniIdentifikator = Replace(Z_InterniIdentifikator, ".", "")

    If Trim(Z_InterniIdentifikator) <> "" Then
        Print #3, "  <interniIdentifikacePodaniPodavatele>" & Trim(Z_InterniIdentifikator) & "</interniIdentifikacePodaniPodavatele>"
    End If
    
    If Trim(ZdrKodHl) <> "" Then
    
        Print #3, "  <kodZdravotniPojistovny>" & Trim(ZdrKodHl) & "</kodZdravotniPojistovny>"
    End If

    Print #3, "  <identifikaceZamestnavatele>"
    
    If Trim(FirmaIco) <> "" Then
        Print #3, "    <identifikacniCisloPlatce>" & Trim(FirmaIco) & "</identifikacniCisloPlatce>"
    End If
    
    If Trim(FirmaNazev) <> "" Then
        Print #3, "    <nazevPlatce>" & Trim(FirmaNazev) & "</nazevPlatce>"
    End If
    
    If Trim(FirmaUlice) <> "" Then
        Print #3, "    <adresaPlatceUlice>" & Trim(FirmaUlice) & "</adresaPlatceUlice>"
    End If
    
    If Trim(FirmaAdrCislo) <> "" Then
        Print #3, "    <adresaPlatceCisloPopisneOrientacni>" & Trim(FirmaAdrCislo) & "</adresaPlatceCisloPopisneOrientacni>"
    End If
    
    Print #3, "    <adresaPlatcePsc>" & Trim(FirmaPSC) & "</adresaPlatcePsc>"
    
    Print #3, "    <adresaPlatceObec>" & Trim(FirmaAdrObec) & "</adresaPlatceObec>"
    
    If Trim(FirmaTelefon) <> "" Then
        Print #3, "    <adresaPlatceTelefon>" & Trim(FirmaTelefon) & "</adresaPlatceTelefon>"
    End If
    
    Print #3, "  </identifikaceZamestnavatele>"
    Print #3, "  <seznamZmenZamestnancu>"
End Sub
Private Sub Utvor_HOZ_XML_StaryStylZam()
   ' For X = 1 To PocetPrihlasek
        Print #3, "    <zmenaZamestance>"
        Print #3, "      <kodzmeny>" & Trim(OldP(55)) & "</kodzmeny>"
        Print #3, "      <datumZmeny>" & Mid(Trim(OldP(59)), 1, 10) & "</datumZmeny>"
        Print #3, "      <cisloPojistence>" & Trim(OldO(100)) & "</cisloPojistence>"
        Print #3, "      <jmeno>" & Trim(ZJmeno) & "</jmeno>"
        Print #3, "      <prijmeni>" & Trim(ZPrijmeni) & "</prijmeni>"
        
        ' Adresa je zanořený komplexní typ
        
            Print #3, "      <adresa>"
            
             If Trim(OldO(160)) + Trim(OldO(165)) <> "" Then: Print #3, "        <ulice>" & Trim(OldO(160)) & " " & Trim(OldO(165)) & "</ulice>"
             If Trim(OldO(161)) <> "" Then: Print #3, "        <obec>" & Trim(OldO(161)) & "</obec>"
             If Trim(OldO(162)) <> "" Then: Print #3, "        <psc>" & Trim(OldO(162)) & "</psc>"
            
            Print #3, "      </adresa>"
       
        
        Print #3, "    </zmenaZamestance>"
  '  Next X
    
    
   
End Sub
Public Sub Utvor_PPPZ_XML_StaryStyl()
    Dim U As String
    U = Chr(34)
    Dim GUIDY As String
    Dim Z_InterniIdentifikator As String
    Dim CarkaTecka_Zaklad As String
    Dim CarkaTecka_Pojistne As String
    FVystup = PathXml & "\" & ZdrZkratkaXML(X) & "_" & Trim(ZdrKodXML(X)) & "_" & Combo1 & " " & Combo2 & ".xml"
    Open FVystup For Output As #3
    
'    i = i + 1: T(i) = "<            START TVORBY PPPZ XML"
'    i = i + 1: T(i) = String(156, "*")
    
    'MsgBox (Z_InterniIdentifikator)
    Print #3, "<?xml version=" & U & "1.0" & U & " encoding=" & U & "windows-1250" & U & "?>"
    Print #3, "<prehledPlatbyZamestnavatele xmlns=" & U & "http://xmlns.vzp.cz/PrehledPlatbyZamestnavatele/v1" & U & ">"
   ' GUIDY = CreateGUID
    Print #3, "  <identifikacePredmetuPodaniText>Přehled platby zaměstnavatele pro ZP 2026+</identifikacePredmetuPodaniText>"
    Print #3, "  <identifikacePredmetuPodaniKod>1079e224-84f4-46e4-99e8-6095bd282301</identifikacePredmetuPodaniKod>"
    
    Z_InterniIdentifikator = Trim(ZdrKodTisk(X)) & mesicInt(Combo1) & Combo2 & Replace(Now, ":", "")
    Z_InterniIdentifikator = Replace(Z_InterniIdentifikator, " ", "")
    Z_InterniIdentifikator = Replace(Z_InterniIdentifikator, ".", "")
  
    If Trim(Z_InterniIdentifikator) <> "" Then
        Print #3, "  <interniIdentifikacePodaniPodavatele>" & Trim(Z_InterniIdentifikator) & "</interniIdentifikacePodaniPodavatele>"
    End If
    
    P_TypPrehledu = "radny"
    If Check1.Value = 1 Then: P_TypPrehledu = "opravny"
    Print #3, "  <kodZdravotniPojistovny>" & Trim(ZdrKodXML(X)) & "</kodZdravotniPojistovny>"
    Print #3, "  <typPrehledu>" & Trim(P_TypPrehledu) & "</typPrehledu>"

 '   i = i + 1: T(i) = "<            IDENTIFIKACE ZAMĚSTNAVATELE (PPPZ)"
 '   i = i + 1: T(i) = String(156, "*")

    ' U monolitického designu se kód pro zaměstnavatele musí zkopírovat znovu
    Print #3, "  <identifikaceZamestnavatele>"
    Print #3, "    <identifikacniCisloPlatce>" & Trim(FirmaIco) & "</identifikacniCisloPlatce>"
    
    If Trim(Z_NazevPlatce) <> "" Then
        Print #3, "    <nazevPlatce>" & Trim(FirmaNazev) & "</nazevPlatce>"
    End If
    
    If Trim(Z_Ulice) <> "" Then
        Print #3, "    <adresaPlatceUlice>" & Trim(FirmaUlice) & "</adresaPlatceUlice>"
    End If
    
    If Trim(Z_CisloPopisne) <> "" Then
        Print #3, "    <adresaPlatceCisloPopisneOrientacni>" & Trim(FirmaAdrCislo) & "</adresaPlatceCisloPopisneOrientacni>"
    End If
    
    Print #3, "    <adresaPlatcePsc>" & Trim(FirmaPSC) & "</adresaPlatcePsc>"
    Print #3, "    <adresaPlatceObec>" & Trim(FirmaAdrObec) & "</adresaPlatceObec>"
    
    If Trim(Z_Telefon) <> "" Then
        Print #3, "    <adresaPlatceTelefon>" & Trim(FirmaTelefon) & "</adresaPlatceTelefon>"
    End If
    
    Print #3, "  </identifikaceZamestnavatele>"
    
  '  i = i + 1: T(i) = "<            ÚDAJE PLATBY (SOUČTY PPPZ)"
  '  i = i + 1: T(i) = String(156, "*")
       
    
    Print #3, "  <udajePlatby>"
    Print #3, "    <mesicHlaseni>" & Trim(mesicInt(Combo1)) & "</mesicHlaseni>"
    Print #3, "    <rokHlaseni>" & Trim(Combo2) & "</rokHlaseni>"
    
    If Val(Trim(P_PocetZamestnancu)) = 0 Then
    '    i = i + 1: T(i) = ">>> Upozornění: Počet zaměstnanců pro pojišťovnu je 0"
    End If
    Print #3, "    <pocetZamestnancu>" & Trim(ZPocZam(X)) & "</pocetZamestnancu>"
    
    ' Konverze desetinné čárky na tečku pro XSD schema
    'CarkaTecka_Zaklad = Replace(Trim(ZVymZakl(X)), ",", ".")
    'CarkaTecka_Pojistne = Replace(Trim(ZPoj(X)), ",", ".")
    
    Print #3, "    <soucetZakladuPojistneho>" & Trim(ZVymZakl(X)) & "</soucetZakladuPojistneho>"
    Print #3, "    <soucetPojistneho>" & Trim(ZPoj(X)) & "</soucetPojistneho>"
    
    Print #3, "  </udajePlatby>"
    
    Print #3, "</prehledPlatbyZamestnavatele>"
    
    Close #3
   
End Sub

