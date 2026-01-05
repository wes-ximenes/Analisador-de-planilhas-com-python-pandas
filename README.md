# Análise de Planilhas de Veículos com Pandas

Script em Python desenvolvido para análise de planilhas Excel geradas pelo LibreOffice (formato HTML),
utilizando Pandas, lxml e openpyxl.

## Funcionalidades:
- Leitura de planilhas em formato HTML (.xls do LibreOffice)
- Tratamento e limpeza de dados
- Remoção de registros duplicados mantendo o mais recente
- Cálculo de dias desde a data de entrada
- Identificação de veículos com mais de 30 dias vencidos
- Exportação automática para Excel

## Tecnologias | bibliotecas utilizadas:
- Python 3
- Pandas
- lxml
- openpyxl
- datetime

## Regras de negócio:
- Considera apenas o registro mais recente por placa
- Veículos com mais de 30 dias são considerados vencidos
- Datas inválidas são descartadas automaticamente

## Saída:
- Exibição dos veículos vencidos no terminal
- Geração do arquivo `veiculos_vencidos.xlsx`

## Objetivo
Projeto prático desenvolvido para automação de análise de dados, funcionalidade criada para manipular dados de planilhas de forma prática.
