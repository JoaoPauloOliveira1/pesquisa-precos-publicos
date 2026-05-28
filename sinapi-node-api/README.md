# SINAPI Node API

API leve em Node.js para expor consultas do SINAPI usando PostgreSQL no Neon.

## Objetivo

Este serviço foi pensado para:

- consultar `insumos` e `composicoes` pelo mesmo contrato esperado pelo app principal;
- armazenar dados em banco persistente;
- sincronizar mensalmente a partir das publicações oficiais da CAIXA;
- evitar dependência obrigatória de Docker/Kong no ambiente principal.

## Endpoints previstos

- `GET /health`
- `GET /status`
- `GET /insumos`
- `GET /composicoes`
- `POST /sync/manual`

## Rodar localmente

```powershell
cd "C:\Users\Everton Paiva\Desktop\Workspace\iluminPbRec\compras-gov-precos\sinapi-node-api"
npm install
npm run dev
```

## Banco

Crie as tabelas com o script:

```powershell
cd "C:\Users\Everton Paiva\Desktop\Workspace\iluminPbRec\compras-gov-precos\sinapi-node-api"
psql "$env:DATABASE_URL" -f .\src\db\schema.sql
```

## Integração com o app principal

No projeto principal, o backend atual já espera:

```text
GET /insumos/?q=...&uf=...&data_referencia=AAAA-MM&regime=...
GET /composicoes/?q=...&uf=...&data_referencia=AAAA-MM&regime=...
```

Quando esta API estiver online, use no app principal:

```text
ENABLE_SINAPI=1
SINAPI_API_URL=https://sua-api-sinapi
SINAPI_API_KEY=sua-chave
```

## Próxima etapa

O próximo passo é implementar o sincronizador mensal que:

1. verifica a referência mais recente publicada pela CAIXA;
2. baixa os arquivos ZIP/XLSX;
3. normaliza insumos e composições;
4. grava tudo no Neon;
5. publica a nova `data_referencia`.
