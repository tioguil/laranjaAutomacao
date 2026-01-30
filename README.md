# laranjaAutomacao

Automação para registrar ponto a partir de um XLSX.

## Requisitos
- Node.js (LTS)
- NPM

## Instalação
```bash
npm install
```

Instalar os navegadores do Playwright:
```bash
npx playwright install
```

## Configuração
Crie um arquivo `.env` na raiz do projeto (use o `.env.example` como base):

```env
CODIGO_EMPREGADOR=1A25S  
PIN=12345
DEFAULT_ENTRY=09:00
START_DATE=12/01/2026
FILE_PATH=C:\caminho\para\relatorio.xlsx
```

Notas:
- `START_DATE` é opcional e filtra somente as datas **a partir** dessa data (inclusive).
- `DEFAULT_ENTRY` define o horário de entrada usado para calcular a saída.
- `FILE_PATH` deve apontar para o XLSX com as colunas esperadas.

## XLSX (Relatório de Apontamento da Genial)
O arquivo XLSX deve ser o **Relatório de Apontamento da Genial**.

Colunas esperadas na planilha (primeira aba):
- `WorkItemId`
- `Data da Tarefa`
- `CampoHora`
- `campoMinuto`

Observações:
- A coluna `Data da Tarefa` pode vir como data (dd/mm/aaaa) ou número serial do Excel.
- As colunas são lidas pelo nome, então mantenha os headers iguais aos do relatório.

## Executar
```bash
npm run dev
```

## Build / Run
```bash
npm run build
npm start
```
