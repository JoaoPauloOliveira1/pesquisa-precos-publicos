# Pesquisa de Precos Publicos

Sistema para pesquisar e comparar precos em fontes publicas como PNCP/Compras.gov, SINAPI, SICRO, ORSE-SE e PE Integrado.

## Rodar localmente

```powershell
npm install
npm start
```

Depois acesse:

```text
http://localhost:3000
```

## Deploy recomendado no Render

Para colocar o sistema online com menos atrito, suba primeiro apenas o app Node.js.
O SINAPI deve ficar desativado no primeiro deploy, porque a versao local usa Docker/Kong/Postgres e isso aumenta bastante a complexidade em uma hospedagem simples.

O arquivo `render.yaml` ja deixa o projeto pronto para um Web Service Node no Render:

- Build Command: `npm install`
- Start Command: `npm start`
- Health Check Path: `/api/health`
- `AUTO_UPDATE_BASES=0`
- `ENABLE_SINAPI=0`

Com isso, o app sobe sem tentar baixar bases grandes automaticamente e sem depender do Docker do SINAPI.

## Variaveis de ambiente

| Variavel | Valor recomendado no Render | Uso |
| --- | --- | --- |
| `NODE_ENV` | `production` | Ambiente de producao |
| `NODE_VERSION` | `20` | Versao do Node no Render |
| `AUTO_UPDATE_BASES` | `0` | Evita downloads grandes na inicializacao |
| `ENABLE_SINAPI` | `0` | Desliga SINAPI quando nao houver API externa |
| `SINAPI_API_URL` | vazio | Use apenas se tiver uma API SINAPI online |
| `SINAPI_API_KEY` | vazio | Chave da API SINAPI, se existir |

## Sobre o SINAPI online

Para usar SINAPI em producao, o melhor caminho e transformar o `autoSINAPI_API` em um servico separado, com banco proprio, ou contratar/manter uma API SINAPI externa.

Quando essa API estiver online, configure no Render:

```text
ENABLE_SINAPI=1
SINAPI_API_URL=https://sua-api-sinapi
SINAPI_API_KEY=sua-chave
```

Sem essas variaveis, o sistema continua funcionando com as outras fontes.

## Atualizacao de bases

No Render, deixe `AUTO_UPDATE_BASES=0`. Se precisar atualizar manualmente:

```text
/api/sicro/atualizar
/api/sicro/gerar-lite
/api/orse/atualizar
```

Observacao: no plano sem disco persistente, arquivos gerados em runtime podem ser perdidos em novos deploys. Para producao, prefira fontes online ou um disco persistente/banco externo para bases locais.

## Estrategia recomendada para o SICRO no Render

O melhor caminho para ativar o SICRO online e usar uma base leve versionada no repositorio, em vez do `sicro.json` grande e dos downloads `.7z`.

O sistema agora aceita estes arquivos, nesta ordem:

```text
data/bases/sicro-lite.json
data/bases/sicro-lite.csv
data/bases/sicro.json
data/bases/sicro.csv
```

Fluxo sugerido:

1. Gerar a base SICRO completa localmente.
2. Gerar a base leve:

```text
/api/sicro/gerar-lite
```

ou apenas Pernambuco:

```text
/api/sicro/gerar-lite?uf=PE
```

3. Conferir o tamanho e o total de registros.
4. Versionar `data/bases/sicro-lite.json` no GitHub.
5. Fazer novo deploy no Render.

Assim o Render passa a pesquisar SICRO sem depender de download automatico nem de disco persistente.
