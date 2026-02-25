# Regras do Projeto NiniPeixe

Este documento serve para registar as regras de negócio e instruções de funcionamento do automatismo de faturas.

## 1. Funcionamento do Programa

- **Executável**: `NiniPeixe.exe`
- **Entrada**: Colocar faturas em PDF na pasta `1_Faturas_Entrada`.
- **Saída**: O programa guarda o Excel resultante na pasta `3_Ficheiro_Entradas`.
- **Arquivo**: Após a leitura, os PDFs são movidos automaticamente para a pasta `2_Faturas_Processadas`.

## 2. Mapeamento (tabela_mapeamento.xlsx)

- **Novos Campos de Referência**: `Nome_Fornecedor` (apenas para consulta).
- **Caracteristicas**: `Metodo_Producao`, `Zona_Mar`, `Massa_Agua_Origem`, `Origem_Aquacultura`, `Arte_Pesca`, `Zona_Pesca`, `Metodo_Pesca`, `Origem`, `Estado`.
- **Fornecedores**: Dados fixos dos fornecedores (Abreviatura, Classificação, etc.).
  - **Identificação Inteligente**: O sistema procura o fornecedor em **3 colunas**: `Nome_Fornecedor1`, `Nome_Fornecedor2` e `Nome_Fornecedor3`.
  - **Novo Campo**: `Documento_Sistema` (usado para ordenar o Excel final por tipo de entrada em mapas do sistema).
- **Regra de IDs**: O campo `Fornecedor_ID` liga as duas tabelas.

## 3. Regras de Extração (IA)

- **Lote Inteligente**: O lote final segue o padrão: `Abreviatura.DiaMesEntrada.LoteFornecedor`.
  - Exemplo: `LUB.2002.123456` ou `0AP.2002.123456`.
  - Se não houver lote do fornecedor: `Abreviatura.DataEntradaCompleta` (Ex: `LUB.20022026`).
- **Data de Entrada**: Ao iniciar, o programa pergunta a data (DDMM). Se ignorar (Enter), usa o dia atual. O lote é sempre calculado com esta data de entrada, não com a data da fatura.
- **Informação Técnica**: A IA agrupa num campo único (`Info_Tecnica`) dados como: Capturado/Aquacultura, País de Origem, Método de Pesca e Estado.
- **Ordenação Final**: O Excel é ordenado por `Doc_Entrada_sistema` e depois por `Nome_Fornecedor`.

## 4. Ordem das Colunas e Ordenação

- **Ordenação**:
  1. Primeiro por `Doc_Entrada_Sistema`.
  2. Segundo por `Nome_Fornecedor`.

  3. `Doc_Entrada_sistema` (B1)
  4. `Fornecedor_ID` (C1)
  5. `Nome_Fornecedor` (D1)
  6. `Abreviatura` (E1)
  7. `Classificacao` (F1)

- **Layout dos Itens (a partir da linha 3)**:
  1. `CODIGO`
  2. `Peso_KG`
  3. `Lote`
  4. `Info_Tecnica`
  5. `Preco`
  6. `Metodo_Producao`
  7. `Zona_Mar`
  8. `Massa_Agua_Origem`
  9. `Origem_Aquacultura`
  10. `Arte_Pesca`
  11. `Zona_Pesca`
  12. `Metodo_Pesca`
  13. `Origem`
  14. `Estado`
  15. Outros campos de suporte (Abreviatura, Arquivo Original, etc.)

## 5. Notas de Funcionamento

- **Relatório Final**: No fim de cada execução, o programa mostra um resumo de quantas faturas foram lidas com sucesso e se houve erros.
- **Fecho Manual**: A janela fica aberta no fim até carregar em "Enter", permitindo ler o relatório.
- **Fornecedores Inteligentes**: O sistema ignora espaços, vírgulas e pontos ao procurar o fornecedor (ex: identifica "Aqua Prime," como "Aquaprime"). Além disso, remove sufixos como "Lda", "SA", "Unipessoal", etc., para garantir que o mapeamento funcione mesmo com variações no nome da fatura.

---
**Dica**: Sempre que quisermos mudar o comportamento do programa, devemos registar a regra aqui primeiro.
