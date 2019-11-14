Attribute VB_Name = "Translate"
Sub translate()
    
    Dim wordList
    
    
    metin = ActiveDocument.Range.Text
    metin = Replace(metin, "­", "")
    wordList = Split(metin, " ")
    newText = ""
    For i = LBound(wordList) To UBound(wordList)
        If Len(wordList(i)) > 0 Then
            newText = newText & wordTranslate(wordList(i)) & " "
            End If
            
    Next i
            
            ActiveDocument.Range.Text = newText
    
    
End Sub

Function wordTranslate(ByVal word As String) As String
    
    
    
    cyrill = ChrW(&H410) & ChrW(&H4D8) & ChrW(&H411) & ChrW(&H412) & ChrW(&H413) & ChrW(&H492) & ChrW(&H414) & ChrW(&H415) & ChrW(&H401) & ChrW(&H416) & ChrW(&H417) & ChrW(&H418) & ChrW(&H419) & ChrW(&H41A) & ChrW(&H49A) & ChrW(&H41B) & ChrW(&H41C) & ChrW(&H41D) & ChrW(&H4A2) & ChrW(&H41E) & ChrW(&H4E8) & ChrW(&H41F) & ChrW(&H420) & ChrW(&H421) & ChrW(&H422) & ChrW(&H423) & ChrW(&H4AE) & ChrW(&H4B0) & ChrW(&H424) & ChrW(&H425) & ChrW(&H4BA) & ChrW(&H426) & ChrW(&H427) & ChrW(&H428) & ChrW(&H429) & ChrW(&H42A) & ChrW(&H42B) & ChrW(&H406) & ChrW(&H42C) & ChrW(&H42D) & ChrW(&H42E) & ChrW(&H42F) _
                        & ChrW(&H430) & ChrW(&H4D9) & ChrW(&H431) & ChrW(&H432) & ChrW(&H433) & ChrW(&H493) & ChrW(&H434) & ChrW(&H435) & ChrW(&H451) & ChrW(&H436) & ChrW(&H437) & ChrW(&H438) & ChrW(&H439) & ChrW(&H43A) & ChrW(&H49B) & ChrW(&H43B) & ChrW(&H43C) & ChrW(&H43D) & ChrW(&H4A3) & ChrW(&H43E) & ChrW(&H4E9) & ChrW(&H43F) & ChrW(&H440) & ChrW(&H441) & ChrW(&H442) & ChrW(&H443) & ChrW(&H4AF) & ChrW(&H4B1) & ChrW(&H444) & ChrW(&H445) & ChrW(&H4BB) & ChrW(&H446) & ChrW(&H447) & ChrW(&H448) & ChrW(&H449) & ChrW(&H44A) & ChrW(&H44B) & ChrW(&H456) & ChrW(&H44C) & ChrW(&H44D) & ChrW(&H44E) & ChrW(&H44F)
    latin = "A" & ChrW(&HC1) & "BVG" & ChrW(&H1F4) & "DE" & ChrW(&HD3) & "JZ" & ChrW(&H130) & "YKQLMN" & ChrW(&H143) & "O" & ChrW(&HD3) & "PRSTWÚUFHHS" & ChrW(&H106) & ChrW(&H15A) & ChrW(&H15A) & "1" & ChrW(&H49) & ChrW(&H130) & "1" & "EUA" & _
    "aábvg" & ChrW(&H1F5) & "de" & ChrW(&HF3) & "jziykqlmn" & ChrW(&H144) & "o" & ChrW(&HF3) & "prstw" & ChrW(&HFA) & "ufhhs" & ChrW(&H107) & ChrW(&H15B) & ChrW(&H15B) & "2" & ChrW(&H131) & "i" & "2" & "eua"
    
    
    'Kazakça da sesli sessiz ince kalýn harfleri ayýrýyoruz çeviride bir kaç istisnai durum olduðu için
    ince = ChrW(&H4D8) & ChrW(&H415) & ChrW(&H406) & ChrW(&H4E8) & ChrW(&H4B0) & ChrW(&H418) & ChrW(&H4D9) & "ei" & ChrW(&H4E9) & ChrW(&H4B1) & ChrW(&H438)
    kalin = ChrW(&H410) & ChrW(&H42B) & ChrW(&H41E) & ChrW(&H4AE) & ChrW(&H401) & ChrW(&H42E) & ChrW(&H42F) & ChrW(&H423) & ChrW(&H430) & ChrW(&H44B) & ChrW(&H43E) & ChrW(&H4AF) & ChrW(&H451) & ChrW(&H44E) & ChrW(&H44F) & ChrW(&H443)
    sesli = ince & kalin
    sesliOnceki = sesli
    sesliSonraki = ChrW(&H410) & ChrW(&H42B) & ChrW(&H41E) & ChrW(&H4AE) & ChrW(&H423) & ChrW(&H430) & ChrW(&H44B) & ChrW(&H43E) & ChrW(&H4AF) & ChrW(&H443) & ChrW(&H4D8) & ChrW(&H415) & ChrW(&H406) & ChrW(&H4E8) & ChrW(&H4B0) & ChrW(&H418) & ChrW(&H4D9) & ChrW(&H435) & ChrW(&H456) & ChrW(&H4E9) & ChrW(&H4B1) & ChrW(&H438)
    ' sessiz = Oncekide sessizse
    sessiz = ChrW(&H411) & ChrW(&H412) & ChrW(&H413) & ChrW(&H492) & ChrW(&H414) & ChrW(&H416) & ChrW(&H417) & ChrW(&H419) & ChrW(&H41A) & ChrW(&H49A) & ChrW(&H41B) & ChrW(&H41C) & ChrW(&H41D) & ChrW(&H4A2) & ChrW(&H41F) & ChrW(&H420) & ChrW(&H421) & ChrW(&H422) & ChrW(&H423) & ChrW(&H424) & ChrW(&H425) & ChrW(&H4BA) & ChrW(&H426) & ChrW(&H427) & ChrW(&H428) & ChrW(&H42A) & ChrW(&H42C) & ChrW(&H431) & ChrW(&H432) & ChrW(&H433) & ChrW(&H493) & ChrW(&H434) & ChrW(&H436) & ChrW(&H437) & ChrW(&H439) & ChrW(&H43A) & ChrW(&H49B) & ChrW(&H43B) & ChrW(&H43C) & ChrW(&H43D) & ChrW(&H4A3) & ChrW(&H43F) & ChrW(&H440) & ChrW(&H441) & ChrW(&H442) & ChrW(&H443) & ChrW(&H444) & ChrW(&H445) & ChrW(&H4BB) & ChrW(&H446) & ChrW(&H447) & ChrW(&H448) & ChrW(&H44A) & ChrW(&H44C)
    sessizSonraki = sessiz & ChrW(&H44F) & ChrW(&H44E) & ChrW(&H451) & ChrW(&H42F) & ChrW(&H42E) & ChrW(&H401)
    
    Dim chr As String
    Dim index As Long
    
    Result = ""
    
    For i = 1 To Len(word)
        chr = Mid(word, i, 1)
        index = InStr(cyrill, chr)
        If index >= 1 Then
            'Y ye eþitse
            If chr = ChrW(&H423) Or chr = ChrW(&H443) Then
                    If i = 1 Then GoTo i1
                    
                    If i > 1 And InStr(sessiz, Mid(word, i - 1, 1)) >= 1 Then Result = Result & IIf(isBold(word, kalin), IIf(chr = ChrW(&H423), "U", "u"), IIf(chr = ChrW(&H423), "Ü", "ü"))
                    
                    If i > 1 And InStr(sesliOnceki, Mid(word, i - 1, 1)) >= 1 And InStr(sessizSonraki, Mid(word, i + 1, 1)) >= 1 And i < Len(word) Then Result = Result & IIf(chr = ChrW(&H423), "W", "w")
i1:
                    If i < Len(word) And InStr(sesliSonraki, Mid(word, i + 1, 1)) >= 1 Then Result = Result & IIf(chr = ChrW(&H423), "W", "w")
                           
            ' N eþitse
            ElseIf chr = ChrW(&H418) Or chr = ChrW(&H438) Then
                   Result = Result & IIf(isBold(word, kalin) = True, IIf(chr = ChrW(&H418), "I", ChrW(&H131)), "i")
                   
                    If i < Len(word) And InStr(sesliSonraki, Mid(word, i + 1, 1)) >= 1 Then Result = Result & IIf(chr = ChrW(&H418), "Y", "y")
            'Ýstisna olmayanlar için
            Else
                'W eþitse
                If chr = ChrW(&H429) Or chr = ChrW(&H449) Then
                    Result = Result & IIf(chr = ChrW(&H429), ChrW(&H15A) & ChrW(&H15A), ChrW(&H15B) & ChrW(&H15B))
                'YU için
                ElseIf chr = ChrW(&H42E) Or chr = ChrW(&H44E) Then
                   Result = Result & IIf(chr = ChrW(&H42E), "YU", "yu")
                ' Ya için
                ElseIf chr = ChrW(&H42F) Or chr = ChrW(&H44F) Then
                    Result = Result & IIf(chr = ChrW(&H42F), "YA", "ya")
                ' b bý için
                ElseIf chr = ChrW(&H42A) Or chr = ChrW(&H42C) Or chr = ChrW(&H44A) Or chr = ChrW(&H44C) Then
                    Result = Result & ""
                Else
                    Result = Result & Mid(latin, InStr(cyrill, chr), 1)
                        
                End If
            End If
        Else
            Result = Result & Mid(word, i, 1)
        End If
    Next i

       wordTranslate = Result
    
End Function



' Kelimede kalýn harf varmý diye kontrol eden fonksiyon
' Kalýn harf varsa kalýn latinde kalýn harf gelicek yoksa ince harf
Function isBold(ByVal word As String, ByVal kalin As String) As Boolean
    For i = 1 To Len(word)
        If InStr(kalin, Mid(word, i, 1)) >= 1 Then
            isBold = True
            Exit Function
        End If
    Next i
        isBold = False
            
        
End Function






