# Perfil operacional ABNT para o Revisor

> Atualizado em 25 de agosto de 2026. Este é um perfil operacional criado a partir de manuais públicos de bibliotecas universitárias e de atualizações públicas. Ele não reproduz nem substitui normas vendidas pela ABNT. O manual ou modelo da instituição do estudante prevalece sempre.

## Normas do perfil padrão

| Norma | Edição/atualização adotada | O que o motor pode fazer | O que precisa permanecer em revisão humana |
| --- | --- | --- | --- |
| NBR 14724 | 2024 | Aplicar A4; margens de 3 cm (superior/esquerda) e 2 cm (direita/inferior); tratar corpo, títulos e elementos reconhecíveis; auditar começo da paginação. | Confirmar capa, folha de rosto, ficha, folha de aprovação, elementos opcionais, verso e a quebra de seção que faz a numeração aparecer somente no início textual. |
| NBR 6023 | 2018, com atualizações públicas de 2025 | Formatar a lista: margem esquerda, espaço simples, sem recuo inicial e espaço entre entradas; ordenar somente listas autor-data que o motor conseguiu ler com segurança; apontar entradas não interpretadas. | Confirmar tipo de documento e todos os metadados (autoria, título, edição, local, editora, DOI, URL e acesso). Não há destaque fixo universal de **negrito** ou *itálico*: o recurso tipográfico deve ser uniforme e a delimitação do título depende do tipo de referência e/ou do manual institucional. |
| NBR 10520 | 2023 | Identificar candidatos a sistemas autor-data/numérico, citações diretas sem localizador e blocos já identificáveis como citação longa; aplicar a apresentação do bloco longo somente quando ele já estiver marcado por estilo ou recuo. Em chamadas autor-data, converter sobrenomes de pessoas físicas que estejam inteiramente em caixa alta para a forma maiúscula/minúscula, quando a lista final permitir essa identificação com segurança. | Confirmar se o trecho é realmente direto, autoria, ano, página/localização, `apud`, escolhas entre `et al.` e lista de autores e a fidelidade ao original. Siglas, chamadas ambíguas e trechos divididos em runs com formatação distinta não são alterados automaticamente. |
| NBR 6028 | 2021 | Identificar RESUMO, ABSTRACT e palavras-chave; formatar esses blocos sem recuo e em espaço simples quando a estrutura for reconhecida. | Conferir idioma, correspondência entre resumo e abstract, conteúdo, palavras-chave e limite específico do curso. |
| NBR 6027 | 2012 | Detectar SUMÁRIO, preservar um campo TOC existente e solicitar sua atualização; inserir campo automático somente se houver um SUMÁRIO explicitamente vazio. | Não inventar números de páginas. Um sumário estático ou inconsistente exige atualização no Word depois da paginação renderizada. |
| NBR 6024 | 2012 | Identificar a hierarquia de títulos e preservar listas multinível (`numPr`) do Word; aplicar alinhamento e recuos de títulos reconhecíveis. | Confirmar a sequência lógica dos níveis, referências cruzadas e a apresentação institucional. O motor jamais apaga a numeração automática para “formatar” um título. |
| NBR 6022 | 2018 | Aplicar o perfil de artigo quando o usuário o seleciona, sem forçar capa ou elementos de monografia. | Validar exigências do periódico (template, título, autoria, afiliação, resumo, palavras-chave e referências), que podem sobrepor a ABNT. |

## Regras objetivas do perfil TCC

- A parte textual é localizada prioritariamente pelo título `INTRODUÇÃO`; a ferramenta não aplica a formatação de corpo na capa ou nos elementos pré-textuais quando não consegue localizar essa transição.
- O texto corrido recebe fonte 12, alinhamento justificado, entrelinhas 1,5 e primeira linha de 1,25 cm. A fonte padrão é Times New Roman, mas o usuário pode selecionar Arial; a NBR não obriga uma família específica e o manual institucional pode determinar outra.
- Citação longa já identificada recebe recuo de 4 cm, fonte 10, espaço simples e sem recuo de primeira linha. A regra se aplica a citações com mais de três linhas; o DOCX não permite determinar com segurança o número de linhas visuais sem paginação renderizada, por isso blocos apenas "suspeitos" são relatados, não transformados às cegas.
- As referências ficam alinhadas à esquerda, em espaço simples, sem recuo inicial e separadas por espaço equivalente a uma linha simples. Em autor-data, a ordenação é alfabética; em sistema numérico, a ordem segue a primeira citação e **não** é alterada.
- Nas citações do sistema autor-data, a indicação de pessoa física usa maiúsculas e minúsculas: `(Silva, 2020)`, e não `(SILVA, 2020)`. Siglas permanecem em caixa alta, como `(IBGE, 2021)`. Isso não altera a entrada da lista final: `SILVA, Maria...` permanece em caixa alta conforme a NBR 6023.
- Todas as folhas a partir da folha de rosto são contadas, mas o algarismo aparece no canto superior direito apenas desde a primeira folha textual, normalmente a INTRODUÇÃO, em fonte 10. A criação dessa quebra de seção não é inferida pelo motor: ele a audita e reporta, pois aplicar `PAGE` no cabeçalho global causaria justamente o erro de numerar páginas pré-textuais ou reiniciar a contagem.

## Normas condicionais — não ativadas automaticamente para um TCC convencional

| Norma | Quando aplicar |
| --- | --- |
| NBR 6034:2004 | Índice remissivo, se o trabalho tiver índice. |
| NBR 12225:2004 | Lombada de exemplar físico encadernado. |
| NBR 15287:2025 | Projeto de pesquisa; não substitui a apresentação do TCC final. |
| NBR 10719:2015 | Relatório técnico e/ou científico, se esse for o tipo documental entregue. |

## Fontes públicas usadas para este perfil

- [Manual de Normalização da UNIFAL-MG, 2. ed. revista e atualizada, 2025](https://www.unifal-mg.edu.br/bibliotecas/wp-content/uploads/sites/125/2025/05/Manual-de-normalizacao-2-edicao-revista-e-atualizada-12-05-2025.pdf) — apresenta NBR 14724:2024, 10520:2023, 6028:2021, 6027:2012, 6024:2012 e 6022:2018; detalha margens, parágrafos, paginação, sumário, citações e referências.
- [Biblioteca da ECA/USP — Normalização](https://www.eca.usp.br/biblioteca/normalizacao) e [atualizações públicas da NBR 6023 de 2025](https://www.eca.usp.br/sites/default/files/2025-06/NBR%206023_2025.pdf) — usadas para manter a camada de referências atualizada sem reproduzir a norma proprietária.
- [Diretrizes para normalização de trabalhos acadêmicos da Faculdade de Farmácia/UFMG, 2024](https://www.farmacia.ufmg.br/wp-content/uploads/2024/07/Diretrizes-para-Normalizacao-de-Trabalhos-Academicos-na-FAFAR-UFMG.pdf) — segunda referência institucional para apresentação de citações e referências.

## Limites obrigatórios

- A ferramenta não certifica conformidade ABNT e não substitui conferência editorial ou bibliotecária.
- Não cria, completa, corrige ou exclui autores, páginas, títulos, DOI, URLs, datas de acesso, traduções, palavras-chave ou dados de capa sem evidência fornecida pelo documento/usuário.
- Nem toda diferença visual é erro ABNT: a instituição pode definir fonte, capa, destaque de seções, modelo de referência e paginação próprios.
