Function jexlArray(Valeur, separateur As String) As String

    jexlArray = "["
    


If InStr(Valeur, separateur) Then
    Valeur = Split(Valeur, separateur)
     For i = 0 To UBound(Valeur)
     jexlArray = jexlArray & """" & Valeur(i) & ""","
     Next
Else
     Valeur = Split(Valeur, separateur)
     For i = 0 To UBound(Valeur)
     jexlArray = jexlArray & """" & Valeur(i) & ""","
     Next
End If
    


jexlArray = Left(jexlArray, Len(jexlArray) - 1) & "]"

    
End Function

Function jexlOperators(c) 'frenchFunc
Set dictOperators = CreateObject("Scripting.Dictionary")

dictOperators.Item("et") = "and"
dictOperators.Item("ou") = "or"
dictOperators.Item("pas") = "not"
dictOperators.Item("egale") = "=="
dictOperators.Item("different") = "!="
dictOperators.Item("inferieur a") = "<"
dictOperators.Item("inferieur ou egale a") = "<="
dictOperators.Item("superieur a") = ">"
dictOperators.Item("superieur ou egale a") = ">="
dictOperators.Item("contient") = "=~"
dictOperators.Item("ne contient pas ") = "!~"

c = Replace(c, """", "")

Debug.Print dictOperators.Item(c)

jexlOperators = dictOperators.Item(c)

'For Each c In dictOperators.keys
'Debug.Print c & ":" & dictOperators.Item(c)
'Next


End Function