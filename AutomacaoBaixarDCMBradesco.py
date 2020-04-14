from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlrd


print("Acessando o site do Bradesco ...\n")

#Lendo do excel
workbook = xlrd.open_workbook(r'C:\Users\Milton\Desktop\Bradesco - relatório de inadimplência.xlsx')
sheet = workbook.sheet_by_index(1)
print(int(sheet.cell_value(1, 0)))



drive = webdriver.Chrome(r'C:\Users\Milton\Desktop\Robos\chromedriver.exe')
drive.get("https://wwws.bradescosaude.com.br/PCBS-GerenciadorPortal/td/loginReferenciado.do")
time.sleep(1)

#Credenciais:
CpfUsuario = "CPF here"
CnpjHosp = "CNPJ here"
SenhaAcesso = "Password here"

#Acesso:
DigiteCpf = drive.find_element_by_id("cpfRefPJ")
DigiteCpf.clear()
DigiteCpf.send_keys(CpfUsuario)
DigiteCpf.send_keys(Keys.TAB)
DigiteCnpj = drive.find_element_by_id("cnpjRef")
DigiteCnpj.clear()
DigiteCnpj.send_keys(CnpjHosp)
DigiteCnpj.send_keys(Keys.TAB)
DigiteSenha = drive.find_element_by_id("senhaRef")
DigiteSenha.clear()
DigiteSenha.send_keys(SenhaAcesso)
DigiteSenha.send_keys(Keys.RETURN)
time.sleep(2)

#Pagina de Acesso aos Demonstrativos de Analise de CM:
drive.get("http://wwws.bradescosaude.com.br/PCBS-DemonstrativosTiss/filtro.do?padraoWeb=saude&site=Referenciado&tipoDemo=CM&nroControle=70548706572744277054870657274427&NroControle=70548706572744277054870657274427&numeroControle=70548706572744277054870657274427&portal=S&site=Referenciado&NOVO_PORTAL_DENTAL=S&cpfCnpj=60884855000154&cnpjCpf=60884855000154&cpf=29043728810&cnpj=60884855000154&cpf_cnpj=60884855000154&codigoReferenciado=60884855000154&cdTipoEntd=3&isPJ=true&cdTipoRefer=2&tipoReferenciado=2&s=referenciado&u=100004616&tipo=1&cfunc=347%27)")
time.sleep(1)
CodReferenciado = drive.find_element_by_name("cdReferencia")
CodReferenciado.click()
CodReferenciado.send_keys(Keys.DOWN)
CodReferenciado.send_keys(Keys.RETURN)
#Protocolo:
NumProtocolo = drive.find_element_by_id("formulario_numProtocolo")
CodReferenciado.click()
NumProtocolo.send_keys(int(sheet.cell_value(1, 0)))
#Pesquisar:
BotaoPesquisar = drive.find_element_by_id("formulario_iniciarListarProtocolos_iniciarListarProtocolos")
BotaoPesquisar.click()
"time.sleep(2)"

ProtocoloEncontrado = drive.find_element_by_tag_name("a")
ProtocoloEncontrado.click()

JanelaAberta = drive.find_element_by_tag_name("Title")
