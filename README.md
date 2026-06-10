# Automação CT-e — Prestação de Serviço em Desacordo

Automação em Python que registra em lote o evento de **Prestação de 
Serviço em Desacordo** (recusa de CT-e) no portal DFe da SVRS, 
eliminando o preenchimento manual conhecimento por conhecimento.


## Recursos

- Processamento em lote com limite configurável por execução
- Retry automático (3 tentativas por linha) com reinício do navegador
- Salvamento periódico da planilha (não perde progresso em caso de queda)
- Reinício preventivo do Chrome a cada N linhas (evita travamentos)
- Pula automaticamente linhas já processadas (com protocolo preenchido)
- Log completo em terminal e arquivo (`automacao_cte.log`)

## Tecnologias

- Python 3.10+
- Selenium + undetected-chromedriver
- openpyxl

## Como usar

1. Baixe o arquivo 'Automacao_CTE.py', deixe em um pasta esclusiva
2. Baixe a planilha "Anular-cte" e deixe na mesma pasta do arquivo anterior
4. Preencha os determinados dados
5. Execute: `python Automacao_CTE.py`
6. Selecione o **certificado digital** no navegador quando solicitado — 
   a automação continua sozinha a partir daí

>  O script não armazena certificados nem senhas: a autenticação é 
> feita manualmente pelo usuário no navegador, uma única vez por sessão.

## Autor

**Lucas Maia Aguiar** — Assistente Fiscal | Automação fiscal com Python e VBA
[LinkedIn](https://www.linkedin.com/in/lucas-maia-3b0a11204)
