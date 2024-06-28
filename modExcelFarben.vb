Public Enum eExcelFarbe
    C_WEISS = 1
    C_GRAU25 = 2
    C_GRAU40 = 3
    C_GRAU50 = 4
    C_GRAU80 = 5
    C_SCHWARZ = 6
    C_INDIGOBLAU = 7
    C_BLAUGRAU = 8
    C_VIOLETT = 9
    C_PFLAUME = 10
    C_LAVENDEL = 11
    C_DUNKELBLAU = 12
    C_BLAU = 13
    C_HELLBLAU = 14
    C_HIMMELBLAU = 15
    C_BLASSBLAU = 16
    C_DUNKELBLAUGRUEN = 17
    C_BLAUGRUEN = 18
    C_AQUAMARIN = 19
    C_TUERKIS = 20
    C_HELLTUERKIS = 21
    C_DUNKELGRUEN = 22
    C_GRUEN = 23
    C_MEERESGRUEN = 24
    C_GRELLESGRUEN = 25
    C_HELLGRUEN = 26
    C_OLIVGRUEN = 27
    C_DUNKELGELB = 28
    C_GELBGRUEN = 29
    C_GELB = 30
    C_HELLGELB = 31
    C_BRAUN = 32
    C_ORANGE = 33
    C_HELLORANGE = 34
    C_GOLD = 35
    C_GELBBRAUN = 36
    C_DUNKELROT = 37
    C_ROT = 38
    C_ROSA = 39
    C_HELLROSA = 40
End Enum

Module modExcelFarben
    Friend Function GetTextFarbe(ByVal Farbe As eExcelFarbe) As Long
        Dim retVal As Long = 0
        Select Case Farbe
            Case eExcelFarbe.C_WEISS : retVal = 0
            Case eExcelFarbe.C_GRAU25 : retVal = 0
            Case eExcelFarbe.C_GRAU40 : retVal = 0
            Case eExcelFarbe.C_GRAU50 : retVal = 16777215
            Case eExcelFarbe.C_GRAU80 : retVal = 16777215
            Case eExcelFarbe.C_SCHWARZ : retVal = 16777215
            Case eExcelFarbe.C_INDIGOBLAU : retVal = 16777215
            Case eExcelFarbe.C_BLAUGRAU : retVal = 16777215
            Case eExcelFarbe.C_VIOLETT : retVal = 16777215
            Case eExcelFarbe.C_PFLAUME : retVal = 16777215
            Case eExcelFarbe.C_LAVENDEL : retVal = 0
            Case eExcelFarbe.C_DUNKELBLAU : retVal = 16777215
            Case eExcelFarbe.C_BLAU : retVal = 16777215
            Case eExcelFarbe.C_HELLBLAU : retVal = 16777215
            Case eExcelFarbe.C_HIMMELBLAU : retVal = 0
            Case eExcelFarbe.C_BLASSBLAU : retVal = 0
            Case eExcelFarbe.C_DUNKELBLAUGRUEN : retVal = 16777215
            Case eExcelFarbe.C_BLAUGRUEN : retVal = 16777215
            Case eExcelFarbe.C_AQUAMARIN : retVal = 0
            Case eExcelFarbe.C_TUERKIS : retVal = 0
            Case eExcelFarbe.C_HELLTUERKIS : retVal = 0
            Case eExcelFarbe.C_DUNKELGRUEN : retVal = 16777215
            Case eExcelFarbe.C_GRUEN : retVal = 16777215
            Case eExcelFarbe.C_MEERESGRUEN : retVal = 16777215
            Case eExcelFarbe.C_GRELLESGRUEN : retVal = 0
            Case eExcelFarbe.C_HELLGRUEN : retVal = 0
            Case eExcelFarbe.C_OLIVGRUEN : retVal = 16777215
            Case eExcelFarbe.C_DUNKELGELB : retVal = 0
            Case eExcelFarbe.C_GELBGRUEN : retVal = 0
            Case eExcelFarbe.C_GELB : retVal = 0
            Case eExcelFarbe.C_HELLGELB : retVal = 0
            Case eExcelFarbe.C_BRAUN : retVal = 16777215
            Case eExcelFarbe.C_ORANGE : retVal = 0
            Case eExcelFarbe.C_HELLORANGE : retVal = 0
            Case eExcelFarbe.C_GOLD : retVal = 0
            Case eExcelFarbe.C_GELBBRAUN : retVal = 0
            Case eExcelFarbe.C_DUNKELROT : retVal = 16777215
            Case eExcelFarbe.C_ROT : retVal = 16777215
            Case eExcelFarbe.C_ROSA : retVal = 16777215
            Case eExcelFarbe.C_HELLROSA : retVal = 0
            Case Else
                Debug.Assert(0)
        End Select
        Return (retVal)
    End Function

    Friend Function GetHintergrundFarbe(ByVal Farbe As eExcelFarbe) As Long
        Dim retVal As Long = 0
        Select Case Farbe
            Case eExcelFarbe.C_WEISS : retVal = 16777215
            Case eExcelFarbe.C_GRAU25 : retVal = 12632256
            Case eExcelFarbe.C_GRAU40 : retVal = 9868950
            Case eExcelFarbe.C_GRAU50 : retVal = 8421504
            Case eExcelFarbe.C_GRAU80 : retVal = 0
            Case eExcelFarbe.C_SCHWARZ : retVal = 0
            Case eExcelFarbe.C_INDIGOBLAU : retVal = 10040115
            Case eExcelFarbe.C_BLAUGRAU : retVal = 10053222
            Case eExcelFarbe.C_VIOLETT : retVal = 8388736
            Case eExcelFarbe.C_PFLAUME : retVal = 6697881
            Case eExcelFarbe.C_LAVENDEL : retVal = 16751052
            Case eExcelFarbe.C_DUNKELBLAU : retVal = 8388608
            Case eExcelFarbe.C_BLAU : retVal = 16711680
            Case eExcelFarbe.C_HELLBLAU : retVal = 16737843
            Case eExcelFarbe.C_HIMMELBLAU : retVal = 16763904
            Case eExcelFarbe.C_BLASSBLAU : retVal = 16764057
            Case eExcelFarbe.C_DUNKELBLAUGRUEN : retVal = 6697728
            Case eExcelFarbe.C_BLAUGRUEN : retVal = 8421376
            Case eExcelFarbe.C_AQUAMARIN : retVal = 13421619
            Case eExcelFarbe.C_TUERKIS : retVal = 16776960
            Case eExcelFarbe.C_HELLTUERKIS : retVal = 16777164
            Case eExcelFarbe.C_DUNKELGRUEN : retVal = 13056
            Case eExcelFarbe.C_GRUEN : retVal = 32768
            Case eExcelFarbe.C_MEERESGRUEN : retVal = 6723891
            Case eExcelFarbe.C_GRELLESGRUEN : retVal = 65280
            Case eExcelFarbe.C_HELLGRUEN : retVal = 13434828
            Case eExcelFarbe.C_OLIVGRUEN : retVal = 13107
            Case eExcelFarbe.C_DUNKELGELB : retVal = 32896
            Case eExcelFarbe.C_GELBGRUEN : retVal = 52377
            Case eExcelFarbe.C_GELB : retVal = 65535
            Case eExcelFarbe.C_HELLGELB : retVal = 10092543
            Case eExcelFarbe.C_BRAUN : retVal = 13209
            Case eExcelFarbe.C_ORANGE : retVal = 26367
            Case eExcelFarbe.C_HELLORANGE : retVal = 39423
            Case eExcelFarbe.C_GOLD : retVal = 52479
            Case eExcelFarbe.C_GELBBRAUN : retVal = 10079487
            Case eExcelFarbe.C_DUNKELROT : retVal = 128
            Case eExcelFarbe.C_ROT : retVal = 255
            Case eExcelFarbe.C_ROSA : retVal = 16711935
            Case eExcelFarbe.C_HELLROSA : retVal = 13408767
            Case Else
                Debug.Assert(0)
        End Select
        Return (retVal)
    End Function

End Module
