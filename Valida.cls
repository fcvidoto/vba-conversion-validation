VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Valida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Referencia - Microsoft Scripting Runtime

'Valida se o dado é date
Public Function EhData(ByVal Valor As Variant) As Boolean
    
    'Verifica se é data e se possui 2 barras de date
    If IsDate(Valor) = True And _
        VerificaAsBarras(Valor) = True Then
            
            EhData = True
            Exit Function
    Else
        EhData = False
        Exit Function
    End If
    
End Function
  
'Verifica se date 'Possui 2 '/' (barras)
Private Function VerificaAsBarras(ByVal palavra As String) As Boolean
    
    Dim i As Integer
    Dim contaBarras As Integer
    
    'Itera por cada letra da palavra (busca 2 barras)
    For i = 1 To Len(palavra)
    
        'Passa letra por letra
        If Mid(palavra, i, 1) = "/" Then
            contaBarras = contaBarras + 1
        End If
    Next i
    
    'Se encontrar 2 barras na string, então retorna 'True'
    If contaBarras = 2 Then
        VerificaAsBarras = True
    End If

End Function
    
'Valida se o dado é booleano
Public Function IsBoolean(ByVal Valor As Variant) As Boolean
    
    'Retorna que é booelan
    If Valor = "TRUE" Or Valor = "FALSE" Then
        IsBoolean = True
        Exit Function
    
    Else
        
        'Se o valor for 0 ou -1
        If Valor = 0 Or Valor = -1 Then
            IsBoolean = True
            Exit Function
        End If
    End If
    
End Function

Function cString(lNumero As Variant) As String
        
    On Error GoTo nelse
        cString = lNumero
        Exit Function
nelse:
    On Error GoTo 0
        cString = vbNullString
    End Function

'Verifica se o arquivo existe no sistema.
Public Function SeArquivoExiste(ByVal path As String) As Boolean

    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
    'Retorna true na existência do arquivo
    If fso.FileExists(path) Then
        SeArquivoExiste = True
    End If

End Function

'Verifica se a pasta existe no sistema.
Public Function SePastaExiste(ByVal path As String) As Boolean

    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
    'Retorna true na existência da pasta
    If fso.FolderExists(path) Then
        SePastaExiste = True
    End If

End Function

'Verifica se a Sheet existe no Workbook.
Public Function SeSheetExiste(ByVal nomeDaSheet As Variant) As Boolean
    Dim sheet As Worksheet
    
    For Each sheet In Excel.Application.Worksheets
        
        'Procura pelo nome da sheet
        If sheet.Name = nomeDaSheet Then
            SeSheetExiste = True
            Exit For
        End If
        
    Next sheet
    
End Function

'Verifica se o nome do workbook atual está de acordo
Public Function SeActiveWorkbookName(ByRef activeWorkbookName As String) As Boolean
    
    'Se o nome do workbook for diferente
    If ActiveWorkbook.Name <> activeWorkbookName Then
        SeActiveWorkbookName = False
        
    'Se o nome do workbook for igual
    ElseIf ActiveWorkbook.Name = activeWorkbookName Then
        SeActiveWorkbookName = True
    End If
    
End Function


