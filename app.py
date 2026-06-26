CREATE TABLE IF NOT EXISTS usuarios (
    id BIGSERIAL PRIMARY KEY,
    usuario TEXT NOT NULL UNIQUE,
    senha_hash TEXT NOT NULL,
    perfil TEXT NOT NULL,
    ativo INTEGER DEFAULT 1,
    criado_em TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS colaboradores (
    id BIGSERIAL PRIMARY KEY,
    matricula TEXT,
    nome TEXT NOT NULL,
    jornada_trabalho TEXT,
    cargo TEXT,
    setor TEXT,
    filial TEXT DEFAULT 'MG CGE',
    logins_jms TEXT,
    gestor_responsavel TEXT,
    ativo INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS presencas (
    id BIGSERIAL PRIMARY KEY,
    colaborador_id BIGINT NOT NULL REFERENCES colaboradores(id) ON DELETE CASCADE,
    data DATE NOT NULL,
    status TEXT NOT NULL,
    observacao TEXT,
    criado_em TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    atualizado_em TIMESTAMPTZ
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_presenca_colaborador_data
ON presencas (colaborador_id, data);
