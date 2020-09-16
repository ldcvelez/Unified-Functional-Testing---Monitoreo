'Importo los datos del excel que contienen las URLs a probar
rutaExcel = "C:\MonitoreoBIP\Recursos\URLs_Monitoreo.xlsx"

DataTable.ImportSheet rutaExcel,"URLAcceso","MonitoreoBIP"

'Obtengo la cantidad de filas de la tabla 
cantFilasLocal = DataTable.GetSheet("MonitoreoBIP").GetRowCount

'Itero tantas veces como filas tenga
For i=1 To cantFilasLocal
    DataTable.SetCurrentRow(i)  
    RunAction "Pruebas [Pruebas_Res]", oneIteration,DataTable("URL",dtLocalSheet),i,mensajeFinal,mensajeSalida
    mensajeFinal = mensajeFInal & vbcrlf & mensajeSalida
Next
