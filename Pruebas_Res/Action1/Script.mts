'Minimiza la aplicacion
MinimizeQTPWindow()

'Utilizo siempre la primera fila
DataTable.SetCurrentRow(1)

'Si es la primera iteracion, abro la pagina de Banco Provincia y ejecuto la accion que guarda el user y pass, 
'para utilizarla en las demas pruebas sin pedirlas
If Parameter("url") = "http://10.5.15.176:8888/eBanking/login/inicio.htm" Then
	RunAction "EntradaBrowser [EntradaBrowserPagina]", oneIteration,Parameter("url")
	RunAction "LoginEncriptado [LoginLogout]", oneIteration
	Else
	Browser("BancoProvincia").Navigate Parameter("url")
	RunAction "Login [LoginLogout]", oneIteration
End If	

RunAction "PosicionConsolidada [MonitoreoBIP]", oneIteration

RunAction "Pagos [MonitoreoBIP]", oneIteration

'Junto a la llamada a la accion Reporte, se envian los parametros: URL que se prueba, numero de URL que se prueba, el mensaje acumulado para el mail
'final general, y el mensaje de salida de Reporte, que va formando el mensaje final general.
RunAction "Reporte [MonitoreoBIP]", oneIteration,Parameter("url"),Parameter("numURL"),Parameter("msjGralEnt"),Parameter("msjGralSal")

RunAction "Logout [LoginLogout]", oneIteration

If Parameter("url") = "http://10.1.15.249:8888/eBanking/login/inicio.htm" Then
	Browser("BancoProvincia").Close
End If


Sub MinimizeQTPWindow ()
	Set qtApp = getObject("","QuickTest.Application")
	qtApp.WindowState = "Minimized"
	Set qtApp = Nothing
End Sub


 



