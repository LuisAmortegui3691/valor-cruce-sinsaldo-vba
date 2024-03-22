Attribute VB_Name = "Módulo1"
Sub cuentaSinSaldoProveedor()

    Dim documentosEntrada As String, documentosSalida As String
    Dim valorDebito, valorCredito
    Dim i As Long, j As Long
    Dim libroAux As String
    Dim ultimaFilaAux As Long
    Dim validaDebito As String, validaCredito As String
    
    documentosEntrada = ThisWorkbook.Sheets("main").Range("C2").Value
    documentosSalida = ThisWorkbook.Sheets("main").Range("C3").Value
    
    libroAux = documentosEntrada & "Auxiliar\"
    libroAux = Dir(libroAux)
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "Auxiliar\" & libroAux
    Application.DisplayAlerts = True
    
    ultimaFilaAux = Workbooks(libroAux).Sheets("aux").Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To ultimaFilaAux
        valorDebito = Workbooks(libroAux).Sheets("aux").Range("K" & i).Value
        validaDebito = Workbooks(libroAux).Sheets("aux").Range("V" & i)
        
        If validaDebito <> "ok" Then
            If valorDebito > 0 Then
                For j = 2 To ultimaFilaAux
                    valorCredito = Workbooks(libroAux).Sheets("aux").Range("L" & j)
                    validaCredito = Workbooks(libroAux).Sheets("aux").Range("V" & j)
                    
                    If validaCredito <> "ok" Then
                        If valorCredito > 0 Then
                            If valorDebito = valorCredito Then
                                Workbooks(libroAux).Activate
                                Workbooks(libroAux).Sheets("aux").Rows(i).Select
                                ' Cambia el color de fondo de la selección al color #ffb3ff
                                With Selection.Interior
                                    .Color = RGB(0, 255, 255) ' Código RGB para el color #ffb3ff
                                End With
                                ' Quita la selección
                                Application.CutCopyMode = False
                                
                                Workbooks(libroAux).Sheets("aux").Rows(j).Select
                                ' Cambia el color de fondo de la selección al color #ffb3ff
                                With Selection.Interior
                                    .Color = RGB(0, 255, 255) ' Código RGB para el color #ffb3ff
                                End With
                                ' Quita la selección
                                Application.CutCopyMode = False
                                
                                Workbooks(libroAux).Sheets("aux").Range("V" & i).Value = "ok"
                                Workbooks(libroAux).Sheets("aux").Range("V" & j).Value = "ok"
                            End If
                        End If
                    End If
                Next j
            End If
        End If
        
    Next i

End Sub
