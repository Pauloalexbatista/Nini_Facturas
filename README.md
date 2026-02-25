# Nini Facturas - Manual de Instalação

Este projeto permite ler faturas de peixe (PDF) e gerar um ficheiro Excel formatado para importação no sistema.

## 🚀 Como instalar num novo PC

O programa foi compilado como um executável independente, por isso **não precisas de instalar o Python** no outro computador.

### 1. Copiar a Pasta do Projeto

Copia a pasta completa `PRJT_NiniPeixe` para o novo computador (ou faz download do GitHub).

### 2. Configurar a Chave da API (Crucial)

O ficheiro `.env` contém a tua chave privada do Google Gemini e **não vai para o GitHub** por segurança.

- No novo PC, cria um ficheiro chamado `.env` na pasta principal.
- Escreve o seguinte lá dentro:

  ```env
  GOOGLE_API_KEY=Tua_Chave_Privada_Aqui
  ```

### 3. Estrutura Necessária

Garante que os seguintes ficheiros/pastas estão presentes:

- `NiniPeixe.exe` (O programa)
- `tabela_mapeamento.xlsx` (Onde defines os teus códigos)
- `1_Faturas_Entrada/` (Pastas para os PDFs)
- `2_Faturas_Processadas/`
- `3_Ficheiro_Entradas/`

## 💻 Como Correr

1. Coloca os PDFs na pasta `1_Faturas_Entrada`.
2. Clica duas vezes em **`NiniPeixe.exe`**.
3. Segue as instruções na consola.
4. O resultado estará na pasta `3_Ficheiro_Entradas`.

---
*Nota: Se quiseres correr o código fonte em vez do executável, precisarás de instalar o Python e as bibliotecas listadas no script.*
