# Bases adicionais

Coloque aqui arquivos locais para ativar fontes referenciais:

- `sicro.csv` ou `sicro.json`
- `orse.csv` ou `orse.json`

Para o SICRO, o sistema confere automaticamente a página oficial do DNIT Nordeste e informa qual é a referência mais recente por UF.
Use os relatórios oficiais do DNIT:

https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro/nordeste/nordeste

Ao iniciar o servidor localmente, o sistema também verifica essa página e baixa os arquivos `.7z` mais recentes para:

```text
data/bases/sicro-downloads/
```

Os arquivos baixados ficam fora do Git pelo `.gitignore`.

Depois do download, o sistema extrai os `.7z` e gera automaticamente:

```text
data/bases/sicro.json
```

Esse é o arquivo pesquisável usado pelo botão `SICRO`.

Para a ORSE-SE, o sistema confere automaticamente a página oficial de Base de Dados da CEHOP e identifica o arquivo `.ORSE` mais recente:

https://orse-portal.cehop.se.gov.br/base-de-dados/

Ao iniciar o servidor, o sistema verifica essa página e tenta baixar o arquivo oficial mais recente para:

```text
data/bases/orse-downloads/
```

Depois do download, o sistema tenta extrair/importar formatos tabulares conhecidos e gera:

```text
data/bases/orse.json
```

Se o arquivo `.ORSE` estiver em formato proprietário não tabular, o status mostra o tipo detectado e a ORSE continua podendo ser ativada com um `orse.csv` ou `orse.json` manual no mesmo padrão abaixo.

Além do download oficial, a consulta ORSE-SE no sistema usa a pesquisa pública antiga de insumos da CEHOP:

```text
https://orse.cehop.se.gov.br/insumosargumento.asp
```

O sistema simula o formulário `POST`, usa o período mais recente disponível, lê a tabela HTML retornada e transforma os resultados em registros pesquisáveis da fonte `ORSE-SE`.

Também existem endpoints manuais:

```text
/api/sicro/atualizar
/api/sicro/importar
/api/sicro/status
/api/orse/atualizar
/api/orse/importar
/api/orse/status
```

Em produção, especialmente no Render, use `AUTO_UPDATE_BASES=0` para impedir downloads grandes na inicialização. Nesse caso, os endpoints acima continuam disponíveis para atualização manual.

Campos reconhecidos automaticamente:

- codigo, cod, referencia
- descricao, nome, servico, insumo
- unidade, un, und
- preco, valor, custo, custo_unitario
- data, referencia, mes
- uf, estado
- tipo, categoria, classe
- link, url

Exemplo CSV:

```csv
codigo;descricao;unidade;preco;uf;referencia;tipo
ABC123;ELETRODUTO PVC 25MM;M;12,34;PE;2026-03;Insumo
```
