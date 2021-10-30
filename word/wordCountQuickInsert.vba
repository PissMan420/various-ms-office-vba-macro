Sub log(content As Variant)
    Call VBA.Interaction.MsgBox(content, vbOK Xor vbInformation, "log")
End Sub

Function getNbWord(sentence)
    splitted = VBA.Strings.Split(sentence, " ")
    totalWord = 0
    For len_idx = LBound(splitted) To UBound(splitted)
        totalWord = len_idx
    Next len_idx
    getNbWord = totalWord
End Function
 
Sub wordCountQuickInsert()
'
' wordCountQuickInsert Macro
'
    Dim wdApp As Word.Document
    Set wdApp = Word.ActiveDocument
    
    didShowWordCount = wdApp.Range.Find.Execute("Nombre de mots: *", NOT_SET, NOT_SET, True)
    
    nbMots = wdApp.ComputeStatistics(wdStatisticWords)
   
    If didShowWordCount Then
        nombreDeMotDuCompteur = getNbWord("Nombre de mots: ") + 1
        nbMots = nbMots - nombreDeMotDuCompteur
    End If
    replaceWith = vbNewLine & vbNewLine & "Nombre de mots: " & nbMots
    
    nbWordToSubstract = VBA.InputBox("Voulez vous soustraire un nombre de mots pour enlever des mots de partie qui ne doive pas être compté?" & vbNewLine & "(Entrer 0 pour  exclure aucun mots)", "Exclure des mots?", 0)
        
    If didShowWordCount Then
        For i = 0 To 2
            replaceWith = VBA.Strings.Replace(replaceWith, vbNewLine, "")
        Next
        wdApp.Range.Find.ExecuteOld "Nombre de mots: *", False, False, True, False, False, False, False, False, replaceWith, True
    Else
     wdApp.Range.InsertAfter (replaceWith)
    End If
    
End Sub