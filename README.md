# Relatório de Vendas por Loja 📊

Este projeto realiza uma análise de dados de vendas para gerar relatórios detalhados por loja, incluindo faturamento, quantidade de produtos vendidos e ticket médio por produto. Ao final, o relatório é enviado automaticamente por e-mail utilizando o Outlook.

## Funcionalidades 🚀

- **Cálculo de Faturamento por Loja:** Soma os valores de vendas agrupados por loja.  
- **Quantidade de Produtos Vendidos:** Soma a quantidade total de produtos vendidos por loja.  
- **Ticket Médio:** Calcula o valor médio gasto por produto em cada loja.  
- **Envio de E-mails Automatizado:** Gera e envia um relatório formatado em HTML via Outlook.

## Tecnologias Utilizadas 🛠️

- **Python**: Linguagem principal para análise e automação.
- **Pandas**: Para manipulação e análise de dados.
- **Win32com.client**: Para integração com o Outlook e envio de e-mails.
- **Excel**: Base de dados das vendas (`Vendas.xlsx`).

## Como Usar ⚙️

1. Clone o Repositório:
   
   Use o comando abaixo para clonar este repositório:
   
   `git clone https://github.com/SeuUsuario/RelatorioVendas.git`
   
   Navegue para o diretório do projeto:
   
   `cd RelatorioVendas`

2. Instale as Dependências:
   
   Certifique-se de que o Python está instalado e as bibliotecas necessárias estão configuradas. Você pode instalá-las com:
   
   `pip install pandas pywin32`

4. Prepare a Base de Dados:
   
   Certifique-se de que o arquivo `Vendas.xlsx` está no mesmo diretório do script Python.

5. Execute o Script:
   
   Para gerar e enviar o relatório, execute o script principal:
   
   `python relatorio_vendas.py`

## Observações ⚠️

- O envio de e-mails utiliza o Outlook e requer que o aplicativo esteja configurado corretamente no Windows.
  
- O script utiliza um endereço de e-mail padrão (`mathausprogramador@outlook.com`). Certifique-se de personalizar este endereço no código, se necessário.

## Exemplo do Relatório 💡

### Faturamento:
| ID Loja     | Valor Final |
|-------------|-------------|
| Loja A      | R$10.000,00 |
| Loja B      | R$8.500,00  |

### Quantidade Vendida:
| ID Loja     | Quantidade |
|-------------|------------|
| Loja A      | 120        |
| Loja B      | 85         |

### Ticket Médio:
| ID Loja     | Ticket Médio |
|-------------|--------------|
| Loja A      | R$83,33      |
| Loja B      | R$100,00     |
