import asyncio
from playwright.async_api import async_playwright
from datetime import datetime, timedelta
import pandas as pd
from time import sleep
def pause(t):
    sleep(t)

def Esperar_pelo_Seletor(page,xpath):
    while True:
        try:
            page.wait_for_selector(f'xpath={xpath}', timeout=30000)
            break

        except TimeoutError:
            sleep(1)
# Dicionário de mapeamento de nomes de lojas para números
LOJA_MAP = {
    'iguatemi': 3,
    'gramado': 4,
    'moinhos': 6,
    'barra': 7,
    'praia': 8,
    'wallig': 9,
    'matriz': 10,
    'atacado': 11,
    'franquias': 15,
    'outlet':5
}

login = {
'Nome':"seu nome",
'senha':"sua senha",
'cnpj':"seu cnpj"
}


# Função para ler itens do arquivo Excel
def ler_itens_excel(arquivo):
    df = pd.read_excel(arquivo)
    itens_por_fornecedor_data = {}
    for _, row in df.iterrows():
        fornecedor = row['fornecedor']
        data_pedido = row['data_pedido']
        loja = row['loja'].strip().lower()  # Converte o nome da loja para minúsculas e remove espaços extras

        if loja in LOJA_MAP:
            loja = LOJA_MAP[loja]
        else:
            print(f"Loja '{loja}' não reconhecida. Usando valor padrão 11.")
            loja = 11  # Valor padrão caso a loja não seja reconhecida

        if isinstance(data_pedido, datetime):
            data_pedido = data_pedido.date()  # Extrai apenas a data se for um objeto datetime
        else:
            data_pedido = datetime.strptime(data_pedido, "%d/%m/%Y").date()

        chave = (fornecedor, data_pedido, loja)  # Usando uma tupla (fornecedor, data_pedido, loja) como chave

        if chave not in itens_por_fornecedor_data:
            itens_por_fornecedor_data[chave] = {}

        codigo = str(row['codigo'])
        custo = str(row['custo']).replace(',', '.')
        cor = row['cor'].replace(' ', '').upper()
        tamanho = str(row['tamanho']).upper()
        quantidade = int(row['quantidade'])

        if codigo not in itens_por_fornecedor_data[chave]:
            itens_por_fornecedor_data[chave][codigo] = {
                'custo': custo,
                'cores': {}
            }

        if cor not in itens_por_fornecedor_data[chave][codigo]['cores']:
            itens_por_fornecedor_data[chave][codigo]['cores'][cor] = {}

        itens_por_fornecedor_data[chave][codigo]['cores'][cor][tamanho] = quantidade

    return itens_por_fornecedor_data

async def adicionar_pedidos_fornecedor(fornecedor, itens, data_pedido, loja):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        
        print(f'Processando itens para o fornecedor: {fornecedor} na data: {data_pedido} para a loja: {loja}')
        natalia = '//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[4]/div/div/div/ul/li[6]/label/input'
        # Calcular a data das parcelas adicionando 60 dias à data do pedido
        data_parcelas = data_pedido + timedelta(days=62)

        await page.goto('https://erp.varejonline.com.br/login/login?service=https%3A%2F%2Ferp.varejonline.com.br%2Ferp%2Fj_spring_cas_security_check%3Bjsessionid%3D82BB04E6B03E50D493AD1854F5DAEF4F')

        # Preenche os campos de login
        await page.fill('xpath=//*[@id="username"]', login['Nome'])
        await page.fill('xpath=//*[@id="password"]', login['senha'])
        await page.fill('xpath=//*[@id="cnpj"]', login['cnpj'])
        await page.click('xpath=//*[@id="login-form"]/div[3]/div/button')

        # Navega até a página desejada
        await page.click('xpath=//*[@id="home-favs"]/li/a/span')
        sleep(3)
        # Coloca Data, loja e Fornecedor.
        await page.wait_for_timeout(3000)
        await page.click('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[2]/div[2]/div/button')
        await page.click(f'xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[2]/div[2]/div/div/ul/li[{loja}]/label/input')
        await page.click('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[3]/div[1]/div/button')

        # Espera até que o campo do fornecedor esteja presente na página
        await page.wait_for_selector('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[3]/div[1]/div/div/div/input')

        # Preenche o fornecedor e pressiona Enter
        await page.fill('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[3]/div[1]/div/div/div/input', fornecedor)
        await page.keyboard.press('Space')
        await page.keyboard.press('Enter')
        sleep(5)
        await page.click('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[3]/div[1]/div/button/span')
        
        await page.click('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[1]/div[4]/div/div/button')
        await page.wait_for_timeout(3000)
        await page.click(f'xpath={natalia}')
        # Adiciona os itens ao pedido
        for codigo, item in itens.items():
            sleep(2)
            await page.click('xpath=//*[@id="codigoProduto"]')
            await page.fill('xpath=//*[@id="codigoProduto"]', codigo)            
            await page.keyboard.press('Enter')
            await page.wait_for_selector('xpath=//*[@id="loading-fullscreen"]', state='hidden',timeout=0)
            sleep(5)
            await page.click('xpath=//*[@id="divProdutoBaseOuMultiplaEntidade"]/div[1]/input')
            
            sleep(2)
            await page.click('xpath=//*[@id="divProdutoBaseOuMultiplaEntidade"]/div[1]/input')
            erp = '//*[@id="checkValorProduto"]'
            await page.wait_for_selector('xpath=//*[@id="loading-fullscreen"]', state='hidden',timeout=0)
                

            for cor, tamanhos in item['cores'].items():
                for tamanho, quantidade in tamanhos.items():
                    try:
                        await page.fill(f'xpath=//*[@id="{tamanho}_{cor}_qtde"]', str(quantidade))
                        await page.click('xpath=//*[@id="formPesquisa"]/fieldset/div[2]/div[1]')

                        custo_com_virgula = str(item['custo']).replace('.', ',')
                        await page.fill(f'xpath=//*[@id="{tamanho}_{cor}_valorUnitario"]', custo_com_virgula)
                        await page.click('xpath=//*[@id="formPesquisa"]/fieldset/div[2]/div[1]')
                    except Exception as e:
                        print(f"Erro ao processar o tamanho {tamanho} e cor {cor}: {e}")
                        continue

            await page.click('xpath=//*[@id="formPesquisa"]/div[2]/div/div/input')
            await page.click('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[2]/div[2]/div[4]/div/button')
            await page.click('xpath=//*[@id="formCriarPedidoCompra"]/fieldset[2]/div[2]/div[4]/div/div/ul/li[3]/label')
            await page.click('xpath=//*[@id="btnAdicionarProduto"]')

        # Parcelas de itens
        await page.fill('xpath=//*[@id="dataLimiteEntrega"]', data_pedido.strftime('%d/%m/%Y'))  
        await page.wait_for_timeout(3000)   
        await asyncio.sleep(4)
        await page.click('xpath=//*[@id="divPlanoPagamento"]/div/div/button')
        await page.click('xpath=//*[@id="divPlanoPagamento"]/div/div/div/ul/li[2]/label')
        await page.click('xpath=//*[@id="lnkParcelas"]')
        await asyncio.sleep(5)
        await page.fill('xpath=//*[@id="numeroParcelas"]', '3')
        await page.click('xpath=//*[@id="htmlPlanoPopup"]/div[2]/div[2]')
        await page.fill('xpath=//*[@id="dataBaseVencimento"]', data_parcelas.strftime('%d/%m/%Y'))  
        await page.wait_for_timeout(3000)
        await page.wait_for_selector('xpath=//*[@id="loading-fullscreen"]', state='hidden',timeout=0)
        await page.keyboard.press('Enter')
        await page.click('xpath=//*[@id="btnSalvarParcelas"]')
        sleep(5)
        await page.click('xpath=//*[@id="popupplanopagamento"]/div[1]/button')
        await page.click('xpath=//*[@id="btnSalvarPedido"]')
        await page.wait_for_selector('xpath=//*[@id="loading-fullscreen"]', state='hidden',timeout=0)
        await browser.close()

async def main():
    itens_por_fornecedor_data = ler_itens_excel('itens.xlsx')  # Substitua 'itens.xlsx' pelo nome do seu arquivo Excel
    
    for chave, itens in itens_por_fornecedor_data.items():
        fornecedor, data_pedido, loja = chave
        await adicionar_pedidos_fornecedor(fornecedor, itens, data_pedido, loja)

if __name__ == "__main__":
    asyncio.run(main())
