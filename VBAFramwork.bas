Public SheetDB as Worksheets

Public Function Getbaris ( Sht as Worksheets, KL as String) as long
	getbaris = sht.cells(sht.Rows.count, KL).End(xlup).row
End Function

Public Sub SetSheetDb(sht as Worksheets) as Worksheets
	Set SheetDb = Sht
End Function 

Public Function SimpanKesheet (Isi as variant)
Dim baris as long

baris = getbaris(SheetDB, "A")
SheetDb.Range("A" & Baris).resize(1,Ubound(isi)+1).value = isi
End Function