' Cell inicial para insertar Sku = 2,4
' Mla comparar Origen = 2,1
' Cell Inicial destino para comparar MLA = 2,1
' Cell Inicial destino Copiar Sku= 2,6

Public Sub MatchearSku()
    Sheets("backup").Activate  

    Dim FilaCellOrigen As Integer 
    Dim FilaCellDestino As Integer 
    Dim ActualMLA As String 
    Dim ForeignMla As String 
    Dim Tmp as String 
    ' cell origen mla = FilaCellOrigen,1
    ' cell destino mla = FilaCellOrigen,1
    ' cell origen sku mla = FilaCellOrigen,4
    ' cell destino sku mla = FilaCellOrigen,6
    
    FilaCellOrigen = 2

    ActiveSheet.Cells(FilaCellOrigen,1).Select
    ActualMLA = Selection
    
    Do Until ActualMLA = ""
        Sheets("Todas").Activate   
        FilaCellDestino = 2
        TmpSku = ""

        ActiveSheet.Cells(FilaCellDestino,1).Select
        ForeignMla = Selection

        Do Until ForeignMla = ""
            If ActualMLA = ForeignMla Then 
                ActiveSheet.Cells(FilaCellDestino,6).Select
                TmpSku = Selection
                Exit Do 
            End If

            FilaCellDestino = FilaCellDestino+1
            ActiveSheet.Cells(FilaCellDestino,1).Select
            ForeignMla = Selection
        Loop
        Sheets("backup").Activate  
        Cells(FilaCellOrigen,4) = TmpSku
    Loop

