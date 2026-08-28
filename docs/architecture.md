# Arquitetura do MVP

## Fluxo atual

```text
Navegador -> FastAPI -> data/jobs/<uuid>/original.(docx|doc)
                       -> auditoria DOCX
                       -> formatação DOCX
                       -> novo DOCX + relatório JSON
                       -> download pelo navegador
```

Cada solicitação recebe um UUID e um diretório próprio. O original nunca é editado: a
formatação trabalha em uma cópia e gera `trabalho-revisado-abnt.docx`.

## Capacidades atuais

- `.docx`: validação como pacote Office, auditoria e formatação.
- `.doc`: conversão pelo LibreOffice, seguida do mesmo fluxo de `.docx`.
- Regras automatizáveis: página A4, margens, fonte, espaçamento, parágrafo, títulos
  reconhecíveis, preservação de numeração automática, referências como bloco, ordenação
  alfabética somente para listas autor-data verificáveis, citações longas já identificáveis,
  sumário/campos de atualização e relatório de ações e pendências.
- Negrito automático de título de referência (NBR 6023) para os três padrões validados
  (livro/monografia simples, número de norma NBR, instrumento legal); demais formatos
  (periódico, capítulo com "In:", autor institucional) são listados para negrito manual,
  não negritados às cegas.
- Normalização de caixa em citações indiretas (`apud`): converte separadamente o sobrenome
  original e o sobrenome citante, com a mesma cautela de sobrenomes curtos não confirmados
  na lista de referências usada para citações comuns.
- Criação da seção SUMÁRIO (título + campo TOC nativo do Word) quando ela não existe no
  documento, ancorada no início do texto; ativa por padrão (modo "Corrigir automaticamente"),
  com opção de trocar para "Somente auditar" na interface.
- Numeração de página conforme a NBR 14724, ativa por padrão (28/08/2026): insere uma quebra
  de seção antes do início do texto e um campo PAGE nativo do Word na nova seção — sem número
  visível nas páginas pré-textuais e sem reiniciar a contagem, sem precisar renderizar/contar
  páginas para calcular um valor inicial (o próprio Word/LibreOffice calcula o número absoluto
  a partir da ausência de cabeçalho na seção anterior). Também trocável para "Somente auditar".
- Relatório de ações e pendências entregue em PDF de leitura direta (`pdf_report.py`), além do
  JSON já existente — a pedido do usuário, que preferia um relatório legível a interpretar
  código/estrutura JSON.
- Início do texto: reconhece o título literal (INTRODUÇÃO/APRESENTAÇÃO/DESENVOLVIMENTO) e,
  na ausência dele, usa como alternativa o primeiro título com estilo Heading do documento —
  registrando a origem (`explicit`, `heading_after_toc` ou `heading_fallback`) no relatório
  para conferência.
- Perfil institucional selecionável (`generic` ou `cgaem`, por ora): ajusta quais elementos
  são obrigatórios (por exemplo, artigo do CGAEM/ESFCEx não exige SUMÁRIO separado) e aponta
  particularidades documentadas em `revisor-abnt-docx/references/perfil-cgaem.md`.
- Regras dependentes de evidência: fidelidade de citação, página da fonte, autoria, DOI, URL,
  data de acesso e dados bibliográficos. Estas devem ser revisadas com fontes confirmadas.

## Limites conscientes

O processamento usa `BackgroundTasks`, apropriado para uma instância única e tarefas de curta
duração. Antes de abrir o serviço para público, substituir por fila persistente e worker
dedicado, adicionar autenticação, banco de dados, armazenamento de objetos, expiração automática
de arquivos, limites por conta, antivírus e trilha de auditoria.

`cleanup_expired_jobs` remove jobs finalizados (`completed`/`failed`) após a retenção
configurada e também jobs presos em `uploaded`/`processing` há mais de 30 minutos (sinal de
processo morto por crash ou reinício do container) — sem fila/worker dedicados, não há outra
forma de detectar esse tipo de trava.

## Próximas entregas de produto

1. Login, consentimento e política de privacidade/LGPD.
2. Modelo de documento institucional: faculdade, curso e normas próprias do aluno.
3. Perfil institucional selecionável — primeira versão entregue (`generic`/`cgaem`, ver
   `references/perfil-cgaem.md`); negrito de título de referência (27/08/2026) e correção
   automática de paginação (28/08/2026) já entregues; falta cobrir capa e permitir que o
   próprio usuário envie o manual da sua instituição.
4. Tela de achados com proposta de correção e aprovação por item.
5. Motor assistido para correspondência citação–referência, que só aplica alteração com
   metadados comprovados.
6. Atualização de campos (TOC, PAGE) por uma instalação controlada do Word/LibreOffice no
   servidor, para que o usuário não precise abrir o arquivo no Word manualmente para ver o
   sumário e a numeração calculados — hoje o campo é inserido corretamente, mas seu valor só
   aparece depois de o Word atualizar os campos ao abrir o arquivo.
7. Fila e armazenamento seguro para produção.
