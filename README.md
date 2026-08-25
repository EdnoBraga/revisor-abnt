# Revisor ABNT

Base técnica para um aplicativo web que recebe um arquivo `.docx`, preserva o original e gera
uma cópia revisada com verificações de apresentação acadêmica, citações, referências, resumo,
sumário e numeração de seções.

O primeiro componente versionado é a skill em `revisor-abnt-docx/`, com scripts Python para
auditoria e formatação de documentos Word. Ela será a referência funcional para a API de envio,
a fila de processamento, a revisão assistida e a entrega segura do arquivo formatado.

Para executar os scripts localmente, use Python 3.11+ e instale as dependências com
`python -m pip install -r requirements.txt`.

## Limites do produto

- O aplicativo não deve prometer certificação ABNT automática.
- O manual da instituição do aluno prevalece sobre regras gerais.
- Não inventar ou alterar silenciosamente dados de referências e citações.
- O documento original deve permanecer intacto; a saída é sempre uma nova cópia.

## Fontes de apoio

As fontes públicas e seus links institucionais estão em
[`revisor-abnt-docx/references/fontes-publicas.md`](revisor-abnt-docx/references/fontes-publicas.md).
Os PDFs não são versionados neste repositório público.
