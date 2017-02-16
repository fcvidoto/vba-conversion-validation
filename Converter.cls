VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function cString(lNumero As Variant) As String
        
    On Error GoTo nelse
        cString = lNumero
        Exit Function
nelse:
    On Error GoTo 0
        cString = vbNullString
    End Function

'Converte valores em booleans
Function cBoolean(ByVal Valor As Variant) As Boolean
    
    'Verifica se o texto possui esses dados
    If Valor Like "*tru*" Then
        cBoolean = True
        Exit Function
    ElseIf Valor Like "*Fals*" Then
        cBoolean = False
        Exit Function
    End If
    
    'Verifica se o texto possui valores 0 ou 1
    If Valor Like "*1*" Then
        cBoolean = True
        Exit Function
    ElseIf Valor Like "*0*" Then
        cBoolean = False
        Exit Function
    End If
    
    'Se não possuir valores nem True or False, compara com vazio e preenchido
    If Trim(Valor) = vbNullString Or IsNull(Valor) = True Then
        cBoolean = False
    Else
        cBoolean = True
    End If
    
End Function

'Converte strings to Date
Function cDatas(ByVal Texto As Variant) As Date
            
    On Error GoTo nelse
    If IsDate(CDate(Texto)) = True Then
        cDatas = Texto
    Else
nelse:
    On Error GoTo 0
        cDatas = Empty
    End If
    
End Function

'Converte strings to numbers
Function cNumero(ByVal Texto As Variant) As Long
    
    If IsNumeric(Texto) = True Then
        cNumero = Texto
    Else
        cNumero = 0
    End If
    
End Function

'Limpa o CNPJ e formata o tamanho de maneira correta
Function cCNPJ(ByVal CNPJ As Variant) As String

    Dim limpa As String
    
    CNPJ = Trim(CNPJ)
    
    If CNPJ = vbNullString Then
        cCNPJ = "00000000000000"
        Exit Function
    End If
    
    limpa = Replace(CNPJ, ".", "")
    limpa = Replace(limpa, "/", "")
    limpa = Replace(limpa, "/", "")
    limpa = Replace(limpa, "\", "")
    limpa = Replace(limpa, "-", "")
    
    cCNPJ = limpa

End Function

