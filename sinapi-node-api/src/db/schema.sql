create extension if not exists pg_trgm;

create table if not exists sinapi_referencias (
  id bigserial primary key,
  data_referencia varchar(7) not null unique,
  origem_url text,
  arquivo_origem text,
  status varchar(30) not null default 'disponivel',
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now()
);

create table if not exists sinapi_sync_runs (
  id bigserial primary key,
  tipo varchar(20) not null default 'manual',
  status varchar(20) not null default 'iniciado',
  referencia_encontrada varchar(7),
  mensagem text,
  iniciado_em timestamptz not null default now(),
  finalizado_em timestamptz
);

create table if not exists sinapi_insumos (
  id bigserial primary key,
  codigo varchar(40) not null,
  descricao text not null,
  descricao_normalizada text not null,
  unidade varchar(20),
  uf char(2) not null,
  data_referencia varchar(7) not null,
  regime varchar(30) not null default 'NAO_DESONERADO',
  preco_mediano numeric(18,4),
  origem_url text,
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now(),
  unique (codigo, uf, data_referencia, regime)
);

create table if not exists sinapi_composicoes (
  id bigserial primary key,
  codigo varchar(40) not null,
  descricao text not null,
  descricao_normalizada text not null,
  unidade varchar(20),
  uf char(2) not null,
  data_referencia varchar(7) not null,
  regime varchar(30) not null default 'NAO_DESONERADO',
  custo_total numeric(18,4),
  origem_url text,
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now(),
  unique (codigo, uf, data_referencia, regime)
);

create index if not exists idx_sinapi_insumos_contexto
  on sinapi_insumos (uf, data_referencia, regime);

create index if not exists idx_sinapi_composicoes_contexto
  on sinapi_composicoes (uf, data_referencia, regime);

create index if not exists idx_sinapi_insumos_descricao_trgm
  on sinapi_insumos using gin (descricao_normalizada gin_trgm_ops);

create index if not exists idx_sinapi_composicoes_descricao_trgm
  on sinapi_composicoes using gin (descricao_normalizada gin_trgm_ops);
