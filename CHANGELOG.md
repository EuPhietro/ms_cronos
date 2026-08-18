# Changelog

Todas as mudancas relevantes deste projeto serao registradas neste arquivo.

## 0.1.0b1 - 2026-08-18

Primeira versao beta do Core.

### Adicionado

- API publica versionada e delimitada por `core.__all__`;
- empacotamento instalavel por `pyproject.toml`;
- upload sequencial de arvores com cache remoto por nivel;
- checkpoint JSON atomico, retomada e validacao de origem/destino;
- notificacoes de progresso e cancelamento cooperativo;
- retry com `Retry-After` e backoff para falhas transitorias;
- validacao centralizada de nomes e caminhos remotos;
- testes unitarios e integracao somente leitura opt-in.
- distribuicao publica sob a licenca MIT.

### Alterado

- upload direto limitado a 10 MiB antes da selecao de upload session;
- chunks de upload grande reduzidos para 10 MiB;
- validacoes de filesystem agora usam erros semanticos em vez de `assert`.

### Limitacoes

- uploads de arvore permanecem sequenciais;
- upload sessions nao sao retomadas por chunk entre processos;
- a representacao navegavel da arvore remota ainda nao faz parte da API.
