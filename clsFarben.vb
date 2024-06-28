'Public Class clsFarben
'    Public Const C_WEISS As Integer = 1
'    Public Const C_GRAU25 As Integer = 2
'    Public Const C_GRAU40 As Integer = 3
'    Public Const C_GRAU50 As Integer = 4
'    Public Const C_GRAU80 As Integer = 5
'    Public Const C_SCHWARZ As Integer = 6
'    Public Const C_INDIGOBLAU As Integer = 7
'    Public Const C_BLAUGRAU As Integer = 8
'    Public Const C_VIOLETT As Integer = 9
'    Public Const C_PFLAUME As Integer = 10
'    Public Const C_LAVENDEL As Integer = 11
'    Public Const C_DUNKELBLAU As Integer = 12
'    Public Const C_BLAU As Integer = 13
'    Public Const C_HELLBLAU As Integer = 14
'    Public Const C_HIMMELBLAU As Integer = 15
'    Public Const C_BLASSBLAU As Integer = 16
'    Public Const C_DUNKELBLAUGRUEN As Integer = 17
'    Public Const C_BLAUGRUEN As Integer = 18
'    Public Const C_AQUAMARIN As Integer = 19
'    Public Const C_TUERKIS As Integer = 20
'    Public Const C_HELLTUERKIS As Integer = 21
'    Public Const C_DUNKELGRUEN As Integer = 22
'    Public Const C_GRUEN As Integer = 23
'    Public Const C_MEERESGRUEN As Integer = 24
'    Public Const C_GRELLESGRUEN As Integer = 25
'    Public Const C_HELLGRUEN As Integer = 26
'    Public Const C_OLIVGRUEN As Integer = 27
'    Public Const C_DUNKELGELB As Integer = 28
'    Public Const C_GELBGRUEN As Integer = 29
'    Public Const C_GELB As Integer = 30
'    Public Const C_HELLGELB As Integer = 31
'    Public Const C_BRAUN As Integer = 32
'    Public Const C_ORANGE As Integer = 33
'    Public Const C_HELLORANGE As Integer = 34
'    Public Const C_GOLD As Integer = 35
'    Public Const C_GELBBRAUN As Integer = 36
'    Public Const C_DUNKELROT As Integer = 37
'    Public Const C_ROT As Integer = 38
'    Public Const C_ROSA As Integer = 39
'    Public Const C_HELLROSA As Integer = 40

'    Public Shared Function GetTextFarbe(ByVal iFarbe As Integer) As Long
'        Dim retVal As Long = 0
'        Select Case iFarbe
'            Case C_WEISS : retVal = 0
'            Case C_GRAU25 : retVal = 0
'            Case C_GRAU40 : retVal = 0
'            Case C_GRAU50 : retVal = 16777215
'            Case C_GRAU80 : retVal = 16777215
'            Case C_SCHWARZ : retVal = 16777215
'            Case C_INDIGOBLAU : retVal = 16777215
'            Case C_BLAUGRAU : retVal = 16777215
'            Case C_VIOLETT : retVal = 16777215
'            Case C_PFLAUME : retVal = 16777215
'            Case C_LAVENDEL : retVal = 0
'            Case C_DUNKELBLAU : retVal = 16777215
'            Case C_BLAU : retVal = 16777215
'            Case C_HELLBLAU : retVal = 16777215
'            Case C_HIMMELBLAU : retVal = 0
'            Case C_BLASSBLAU : retVal = 0
'            Case C_DUNKELBLAUGRUEN : retVal = 16777215
'            Case C_BLAUGRUEN : retVal = 16777215
'            Case C_AQUAMARIN : retVal = 0
'            Case C_TUERKIS : retVal = 0
'            Case C_HELLTUERKIS : retVal = 0
'            Case C_DUNKELGRUEN : retVal = 16777215
'            Case C_GRUEN : retVal = 16777215
'            Case C_MEERESGRUEN : retVal = 16777215
'            Case C_GRELLESGRUEN : retVal = 0
'            Case C_HELLGRUEN : retVal = 0
'            Case C_OLIVGRUEN : retVal = 16777215
'            Case C_DUNKELGELB : retVal = 0
'            Case C_GELBGRUEN : retVal = 0
'            Case C_GELB : retVal = 0
'            Case C_HELLGELB : retVal = 0
'            Case C_BRAUN : retVal = 16777215
'            Case C_ORANGE : retVal = 0
'            Case C_HELLORANGE : retVal = 0
'            Case C_GOLD : retVal = 0
'            Case C_GELBBRAUN : retVal = 0
'            Case C_DUNKELROT : retVal = 16777215
'            Case C_ROT : retVal = 16777215
'            Case C_ROSA : retVal = 16777215
'            Case C_HELLROSA : retVal = 0
'            Case Else
'                Debug.Assert(0)
'        End Select
'        Return (retVal)
'    End Function

'    Public Shared Function GetHintergrundFarbe(ByVal iFarbe As Integer) As Long
'        Dim retVal As Long = 0
'        Select Case iFarbe
'            Case C_WEISS : retVal = 16777215
'            Case C_GRAU25 : retVal = 12632256
'            Case C_GRAU40 : retVal = 9868950
'            Case C_GRAU50 : retVal = 8421504
'            Case C_GRAU80 : retVal = 0
'            Case C_SCHWARZ : retVal = 0
'            Case C_INDIGOBLAU : retVal = 10040115
'            Case C_BLAUGRAU : retVal = 10053222
'            Case C_VIOLETT : retVal = 8388736
'            Case C_PFLAUME : retVal = 6697881
'            Case C_LAVENDEL : retVal = 16751052
'            Case C_DUNKELBLAU : retVal = 8388608
'            Case C_BLAU : retVal = 16711680
'            Case C_HELLBLAU : retVal = 16737843
'            Case C_HIMMELBLAU : retVal = 16763904
'            Case C_BLASSBLAU : retVal = 16764057
'            Case C_DUNKELBLAUGRUEN : retVal = 6697728
'            Case C_BLAUGRUEN : retVal = 8421376
'            Case C_AQUAMARIN : retVal = 13421619
'            Case C_TUERKIS : retVal = 16776960
'            Case C_HELLTUERKIS : retVal = 16777164
'            Case C_DUNKELGRUEN : retVal = 13056
'            Case C_GRUEN : retVal = 32768
'            Case C_MEERESGRUEN : retVal = 6723891
'            Case C_GRELLESGRUEN : retVal = 65280
'            Case C_HELLGRUEN : retVal = 13434828
'            Case C_OLIVGRUEN : retVal = 13107
'            Case C_DUNKELGELB : retVal = 32896
'            Case C_GELBGRUEN : retVal = 52377
'            Case C_GELB : retVal = 65535
'            Case C_HELLGELB : retVal = 10092543
'            Case C_BRAUN : retVal = 13209
'            Case C_ORANGE : retVal = 26367
'            Case C_HELLORANGE : retVal = 39423
'            Case C_GOLD : retVal = 52479
'            Case C_GELBBRAUN : retVal = 10079487
'            Case C_DUNKELROT : retVal = 128
'            Case C_ROT : retVal = 255
'            Case C_ROSA : retVal = 16711935
'            Case C_HELLROSA : retVal = 13408767
'            Case Else
'                Debug.Assert(0)
'        End Select
'        Return (retVal)
'    End Function
'End Class
