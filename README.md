# Revisor ABNT

Aplicativo web que recebe um arquivo `.docx` ou `.doc`, preserva o original e gera uma cópia
revisada com regras de apresentação acadêmica verificáveis e relatórios separados de ações e
pendências. O perfil padrão cobre NBR 14724, 6023, 10520, 6028, 6027 e 6024; NBR 6022 é aplicada
quando o documento é selecionado como artigo científico.

O primeiro componente versionado é a skill em `revisor-abnt-docx/`, com scripts Python para
auditoria e formatação de documentos Word. Ela será a referência funcional para a API de envio,
a fila de processamento, a revisão assistida e a entrega segura do arquivo formatado.

Para executar os scripts localmente, use Python 3.11+ e instale as dependências com
`python -m pip install -r requirements.txt`.

## Executar localmente

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
uvicorn app.main:app --reload
```

Abra `http://127.0.0.1:8000`. O servidor aceita arquivos de até 25 MB, cria um diretório de
trabalho por revisão em `data/jobs/` e disponibiliza o DOCX revisado e o relatório JSON. Para
aceitar `.doc`, instale o LibreOffice no servidor e deixe o comando `soffice` disponível no
`PATH` (ou defina `SOFFICE_BIN`).

### Executar em contêiner

O `Dockerfile` inclui LibreOffice Writer para conversão de `.doc` e pode ser usado assim:

```powershell
docker build -t revisor-abnt .
docker run --rm -p 8000:8000 -v ${PWD}/data:/var/lib/revisor-abnt revisor-abnt
```

## Publicar em host próprio

O repositório inclui `docker-compose.yml` e Caddy para expor o aplicativo com HTTPS automático.
O guia de implantação em VPS está em [`docs/deploy-vps.md`](docs/deploy-vps.md). Antes de abrir ao
público, configure domínio, e-mail de certificado, canal de privacidade, retenção de arquivos e
as medidas de segurança indicadas no guia.

Use também o [`docs/deploy-checklist.md`](docs/deploy-checklist.md) no momento da publicação e
da verificação no domínio.

## Estrutura do produto

- `app/`: interface e API FastAPI do MVP.
- `revisor-abnt-docx/`: skill e scripts que constituem o motor de revisão.
- `tests/`: teste do fluxo de upload, processamento e download.
- `docs/architecture.md`: limites atuais e evolução necessária antes de operação pública.
- `docs/product-design.md`: direção visual da interface.

## O que é realmente automatizado

- Página A4, margens, fonte, espaçamento, recuo e alinhamento do corpo textual identificado.
- Títulos, elementos sem numeração, referências, resumos/abstracts, palavras-chave, tabelas e
  citações longas já identificáveis na estrutura do Word.
- Preservação da numeração automática de seções, auditoria de citações autor-data/numéricas,
  correspondência provável citação–referência e ordenação alfabética apenas de referências
  autor-data que puderem ser lidas com segurança.
- Detecção/preservação de campo automático de sumário; inserção de campo TOC somente em um
  SUMÁRIO vazio; relatório JSON de ações e pendências.

## Limites do produto

- O aplicativo não deve prometer certificação ABNT automática.
- O manual da instituição do aluno prevalece sobre regras gerais.
- Não inventar ou alterar silenciosamente dados de referências e citações.
- O documento original deve permanecer intacto; a saída é sempre uma nova cópia.
- Paginação conforme a NBR 14724 depende de quebra de seção antes da parte textual. O produto
  audita esse requisito, mas não injeta `PAGE` em todos os cabeçalhos, pois isso pode numerar
  capa e elementos pré-textuais ou reiniciar a contagem.
- Correção textual de citações e metadados bibliográficos depende de fonte comprovada e é uma
  etapa assistida, não uma inferência automática.

## Fontes de apoio

As fontes públicas e seus links institucionais estão em
[`revisor-abnt-docx/references/fontes-publicas.md`](revisor-abnt-docx/references/fontes-publicas.md).
Os PDFs não são versionados neste repositório público.
