' どうやら関数での表現は簡単にはできないみたい
' なので
' VBAで検討してみようかな、っと考えたりします


' =TEXTJOIN(
'     "",
'     TRUE,
'     IF(
'         UNICODE(MID(C2, SEQUENCE(LEN(C2)), 1)) >= 65296,
'         IF(
'             UNICODE(MID(C2, SEQUENCE(LEN(C2)), 1)) <= 65370,
'             CHAR(UNICODE(MID(C2, SEQUENCE(LEN(C2)), 1)) - 65248),
'             MID(C2, SEQUENCE(LEN(C2)), 1)
'         ),
'         MID(C2, SEQUENCE(LEN(C2)), 1)
'     )
' )
