# Análise de Planilhas — Comparação MARANGUAPE × CONTROLE

Ferramenta em Python para comparar duas abas de um arquivo Excel (**MARANGUAPE** e **CONTROLE**) com base nos **nomes completos** das pessoas. O script identifica quem aparece em uma aba mas não na outra e gera uma terceira planilha com esses registros e todos os dados correspondentes, em ordem alfabética.

---

## Índice

1. [Contexto e objetivo](#contexto-e-objetivo)
2. [O que o script faz](#o-que-o-script-faz)
3. [Estrutura do projeto](#estrutura-do-projeto)
4. [Requisitos](#requisitos)
5. [Instalação](#instalação)
6. [Como usar](#como-usar)
7. [Formato do arquivo de entrada](#formato-do-arquivo-de-entrada)
8. [Formato do arquivo de saída](#formato-do-arquivo-de-saída)
9. [Como a comparação funciona](#como-a-comparação-funciona)
10. [Exemplo com o arquivo real](#exemplo-com-o-arquivo-real)
11. [Detalhes técnicos do código](#detalhes-técnicos-do-código)
12. [Possíveis erros e soluções](#possíveis-erros-e-soluções)
13. [Limitações conhecidas](#limitações-conhecidas)

---

## Contexto e objetivo

O arquivo **EIXO 2 - DADOS BANCÁRIOS** contém informações bancárias de pessoas distribuídas em várias abas. Duas delas são especialmente relevantes:

| Aba | Descrição |
|-----|-----------|
| **MARANGUAPE** | Lista principal de pessoas com dados bancários da região |
| **CONTROLE** | Lista de controle com um subconjunto dessas pessoas |

A necessidade era **cruzar os nomes** entre essas duas abas e descobrir:

- Quem está em **MARANGUAPE**, mas **não** está em **CONTROLE**
- Quem está em **CONTROLE**, mas **não** está em **MARANGUAPE**

Essas diferenças são exportadas para um novo arquivo Excel, facilitando a revisão manual ou a correção dos cadastros.

---

## O que o script faz

```
┌─────────────────────────────────────┐
│  Arquivo Excel (entrada)            │
│  ┌─────────────┐  ┌──────────────┐  │
│  │ MARANGUAPE  │  │  CONTROLE    │  │
│  │ 131 nomes   │  │  82 nomes    │  │
│  └──────┬──────┘  └──────┬───────┘  │
└─────────┼────────────────┼──────────┘
          │                │
          └───────┬────────┘
                  ▼
         comparar_planilhas.py
         (compara NOME COMPLETO)
                  │
                  ▼
┌─────────────────────────────────────┐
│  resultado_diferencas.xlsx (saída)  │
│  • Nomes só em MARANGUAPE           │
│  • Nomes só em CONTROLE             │
│  • Ordenados alfabeticamente        │
│  • Com todos os dados originais     │
└─────────────────────────────────────┘
```

Em resumo:

1. Lê o arquivo Excel informado pelo usuário
2. Carrega as abas **MARANGUAPE** e **CONTROLE**
3. Compara a coluna **NOME COMPLETO** entre as duas abas
4. Separa os registros que existem em apenas uma das abas
5. Ordena o resultado em ordem alfabética
6. Salva tudo em um novo arquivo Excel

---

## Estrutura do projeto

```
Analise-De-Planilhas/
├── comparar_planilhas.py      # Script principal
├── requirements.txt           # Dependências Python
├── README.md                  # Este arquivo
├── Cópia de EIXO 2 - DADOS BANCÁRIOS.xlsx   # Arquivo de entrada (exemplo)
└── resultado_diferencas.xlsx  # Arquivo gerado após a execução
```

---

## Requisitos

- **Python 3.10 ou superior** (recomendado)
- Bibliotecas listadas em `requirements.txt`:
  - `pandas` — leitura e manipulação de planilhas
  - `openpyxl` — suporte a arquivos `.xlsx`

---

## Instalação

Abra o terminal (PowerShell ou Prompt de Comando) na pasta do projeto e execute:

```powershell
cd "c:\Users\Danzo\Downloads\Analise-De-Planilhas"
pip install -r requirements.txt
```

Isso instalará automaticamente o `pandas` e o `openpyxl`.

---

## Como usar

### Uso básico

Passe o caminho do arquivo Excel como argumento:

```powershell
python comparar_planilhas.py "Cópia de EIXO 2 - DADOS BANCÁRIOS.xlsx"
```

O resultado será salvo como `resultado_diferencas.xlsx` na mesma pasta.

### Definir nome do arquivo de saída

```powershell
python comparar_planilhas.py "Cópia de EIXO 2 - DADOS BANCÁRIOS.xlsx" -o minha_saida.xlsx
```

### Usar outra coluna para comparação

Por padrão, a comparação usa a coluna `NOME COMPLETO`. Para usar outra coluna:

```powershell
python comparar_planilhas.py "arquivo.xlsx" -c "NOME SOCIAL"
```

### Ver todas as opções

```powershell
python comparar_planilhas.py --help
```

### Saída no terminal

Após a execução, o script exibe um resumo:

```
Comparação concluída: 49 registro(s) com diferença.
  - Somente em MARANGUAPE: 49
  - Somente em CONTROLE: 0
Arquivo gerado: C:\Users\Danzo\Downloads\Analise-De-Planilhas\resultado_diferencas.xlsx
```

---

## Formato do arquivo de entrada

O arquivo Excel deve conter **duas abas** com os nomes exatos:

- `MARANGUAPE`
- `CONTROLE`

Cada aba deve ter, no mínimo, a coluna **NOME COMPLETO**. As demais colunas são preservadas na saída.

### Colunas esperadas (exemplo do arquivo real)

| Coluna | Descrição |
|--------|-----------|
| `NOME COMPLETO` | Nome completo da pessoa *(usado na comparação)* |
| `NOME SOCIAL` | Nome social (opcional) |
| `BANCO (COD)` | Código do banco |
| `AGENCIA` | Número da agência |
| `DIGITO AGENCIA` | Dígito verificador da agência |
| `CONTA BANCARIA` | Número da conta |
| `DIGITO CONTA` | Dígito verificador da conta |
| `FORMATO CONTA` | Tipo da conta (ex.: CORRENTE) |
| `WHATSAPP` | Número de WhatsApp |

> A aba **CONTROLE** possui ainda a coluna `NOME`, que indica a origem do registro (ex.: CELINA, ARTHUR). Essa coluna também é incluída na saída quando presente.

---

## Formato do arquivo de saída

O arquivo gerado (`resultado_diferencas.xlsx`) contém uma aba chamada **Diferencas** com:

| Coluna | Descrição |
|--------|-----------|
| **Origem** | Indica de qual aba veio o registro: `Somente em MARANGUAPE` ou `Somente em CONTROLE` |
| **NOME COMPLETO** | Nome da pessoa |
| *(demais colunas)* | Todos os dados bancários originais da linha |

Os registros são ordenados **alfabeticamente** pelo nome completo.

### Exemplo de linhas na saída

| Origem | NOME COMPLETO | BANCO (COD) | AGENCIA | WHATSAPP | ... |
|--------|---------------|-------------|---------|----------|-----|
| Somente em MARANGUAPE | Acácio Victor Ferreira de Queiroz | 237 | 6240 | 859... | ... |
| Somente em MARANGUAPE | ANA BRUNA LOPES MACIEIRA | ... | ... | ... | ... |
| Somente em CONTROLE | Nome Exemplo | 260 | 1 | 859... | ... |

---

## Como a comparação funciona

### 1. Normalização dos nomes

Antes de comparar, cada nome passa por uma normalização para evitar falsas diferenças:

```python
def normalizar_nome(valor):
    # Remove espaços extras nas pontas
    # Converte para minúsculas (ignora maiúsculas/minúsculas)
    # Trata valores vazios ou nulos
```

Exemplos de nomes considerados **iguais** após normalização:

| Nome em MARANGUAPE | Nome em CONTROLE | Resultado |
|--------------------|------------------|-----------|
| `Ana Silva` | `ana silva` | Iguais |
| ` João Santos ` | `João Santos` | Iguais |
| `MARIA OLIVEIRA` | `Maria Oliveira` | Iguais |

### 2. Identificação das diferenças

```
Nomes em MARANGUAPE:  {Ana, Bruno, Carlos, Daniela, ...}
Nomes em CONTROLE:    {Ana, Bruno, Eduardo, ...}

Somente em MARANGUAPE = MARANGUAPE − CONTROLE  →  {Carlos, Daniela, ...}
Somente em CONTROLE   = CONTROLE − MARANGUAPE  →  {Eduardo, ...}
```

### 3. Montagem do resultado

- Cada registro diferente mantém **todas as colunas originais** da aba de origem
- É adicionada a coluna **Origem** no início
- Tudo é concatenado e ordenado alfabeticamente

### 4. Registros ignorados

- Linhas com **NOME COMPLETO** vazio ou nulo são ignoradas na comparação
- Nomes duplicados dentro da mesma aba aparecem **todos** na saída se não existirem na outra aba

---

## Exemplo com o arquivo real

Ao executar o script com o arquivo `Cópia de EIXO 2 - DADOS BANCÁRIOS.xlsx`:

| Métrica | Valor |
|---------|-------|
| Total de nomes em MARANGUAPE | 131 |
| Total de nomes em CONTROLE | 82 |
| Registros somente em MARANGUAPE | **49** |
| Registros somente em CONTROLE | **0** |
| Total na planilha de saída | **49** |

**Interpretação:** todos os 82 nomes de CONTROLE já existem em MARANGUAPE. Porém, 49 pessoas cadastradas em MARANGUAPE ainda **não foram incluídas** na aba CONTROLE.

---

## Detalhes técnicos do código

O script `comparar_planilhas.py` está organizado em funções com responsabilidades claras:

| Função | Responsabilidade |
|--------|------------------|
| `normalizar_nome()` | Padroniza nomes para comparação justa |
| `carregar_aba()` | Lê uma aba do Excel e valida se a coluna de nomes existe |
| `encontrar_diferencas()` | Cruza os dois DataFrames e retorna apenas os registros diferentes |
| `comparar_abas()` | Orquestra o fluxo completo: carregar → comparar → salvar |
| `main()` | Interface de linha de comando (CLI) |

### Constantes configuráveis

No início do arquivo, três constantes definem o comportamento padrão:

```python
ABA_MARANGUAPE = "MARANGUAPE"
ABA_CONTROLE = "CONTROLE"
COLUNA_NOME = "NOME COMPLETO"
```

Para comparar outras abas ou colunas de forma permanente, basta alterar esses valores.

### Fluxo de execução

```
main()
  │
  ├─► Valida se o arquivo existe
  │
  └─► comparar_abas()
        │
        ├─► carregar_aba("MARANGUAPE")
        ├─► carregar_aba("CONTROLE")
        │
        ├─► encontrar_diferencas()
        │     ├─► normalizar nomes
        │     ├─► calcular diferenças de conjuntos
        │     ├─► adicionar coluna Origem
        │     └─► ordenar alfabeticamente
        │
        └─► salvar resultado_diferencas.xlsx
```

---

## Possíveis erros e soluções

| Erro | Causa provável | Solução |
|------|----------------|---------|
| `arquivo não encontrado` | Caminho incorreto ou arquivo movido | Verifique o caminho e use aspas se houver espaços no nome |
| `Aba 'MARANGUAPE' não encontrada` | Nome da aba diferente no Excel | Confirme que as abas se chamam exatamente `MARANGUAPE` e `CONTROLE` |
| `Coluna 'NOME COMPLETO' não encontrada` | Coluna com nome diferente | Use `-c "Nome da Coluna"` ou renomeie a coluna no Excel |
| `ModuleNotFoundError: pandas` | Dependências não instaladas | Execute `pip install -r requirements.txt` |
| Arquivo de saída vazio | Todas as abas têm os mesmos nomes | Normal — significa que não há diferenças entre as abas |

---

## Limitações conhecidas

1. **Comparação por nome exato (normalizado):** nomes com pequenas diferenças de grafia (ex.: "José" vs "Jose", "Ana Maria" vs "Ana M.") serão tratados como pessoas diferentes.

2. **Duplicatas na mesma aba:** se o mesmo nome aparecer mais de uma vez em MARANGUAPE, todas as ocorrências entrarão na saída caso o nome não exista em CONTROLE.

3. **Abas fixas:** o script compara apenas `MARANGUAPE` e `CONTROLE`. Outras abas do arquivo (CELINA, ARTHUR, BARRETO) não são consideradas.

4. **Formato `.xlsx` apenas:** arquivos `.xls` (formato antigo) não são suportados diretamente. Converta para `.xlsx` antes de usar.

5. **Sem interface gráfica:** o script é executado via terminal. Não há janela ou botão para selecionar arquivos (pode ser adicionado futuramente).

---

## Próximos passos possíveis

Melhorias que podem ser implementadas conforme a necessidade:

- Comparar outras combinações de abas (ex.: CELINA × CONTROLE)
- Gerar abas separadas na saída (uma para cada direção)
- Interface gráfica para selecionar o arquivo
- Detecção de nomes similares (fuzzy matching) para pegar erros de digitação
- Relatório em PDF além do Excel

---

## Licença

Uso interno / livre para o projeto EIXO 2.
