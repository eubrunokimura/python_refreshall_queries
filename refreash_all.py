#Este código atualiza queries de planilhas que estão em uma pasta sincronizada com o Sharepoint

import win32com.client

xlapp = win32com.client.DispatchEx("Excel.Application")

#xlapp.Visible = 1 Caso queira que a planilha seja vista

path_planilha = "path"


wb = xlapp.Workbooks.Open(path_planilha)
wb.RefreshAll()
xlapp.CalculateUntilAsyncQueriesDone() #aguarda até que as queries sejam atualizadas
wb.Save()
xlapp.Quit()