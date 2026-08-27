# Perfil institucional: CGAEM/ESFCEx (artigo científico)

Resumo das particularidades do roteiro institucional "Como elaborar o Trabalho de Conclusão
de Curso em formato de artigo científico" (ESFCEx/CGAEM), extraídas do modelo `.docx`
fornecido pelo usuário. Este perfil complementa — e, quando conflitar, prevalece sobre — o
perfil geral em [`normas-abnt.md`](normas-abnt.md), conforme o princípio de que o manual da
instituição tem precedência sobre a norma geral.

## Estrutura

- O TCC do CGAEM é apresentado no formato de **artigo científico** (NBR 6022), não como
  monografia tradicional. Use `document_type=article`.
- Elementos pré-textuais: folha de rosto, termo de cessão de direitos autorais, folha de
  aprovação. **Não há SUMÁRIO separado** — a estrutura segue direto da folha de aprovação
  para título, autoria, resumo e abstract, e então para a introdução.
- Resumo (NBR 6028) e abstract aparecem logo após a folha de aprovação, cada um seguido de
  3 a 5 palavras-chave separadas por ponto e vírgula.
- Elementos textuais: introdução, desenvolvimento (o autor escolhe o nome do(s)
  título(s), evitando literalmente a palavra "Desenvolvimento") e considerações finais.
- Elementos pós-textuais: referências, apêndice(s), anexo(s).

## Extensão

- Mínimo de 20 e máximo de 30 páginas, incluindo elementos pré e pós-textuais. O motor não
  verifica contagem de páginas (não é uma propriedade confiável do DOCX antes de renderizar);
  confira manualmente antes da entrega.

## Citação e referências

- Segue a NBR 10520:2023 (direta, indireta, citação da citação) e a NBR 6023:2025 para
  referências, ambas já cobertas pelo perfil geral.
- Referências em ordem alfabética, alinhamento à esquerda, entrelinhas simples, com espaço
  simples entre entradas — igual ao perfil geral.
- **Negrito de título em referências**: inspecionando o modelo `.docx` fornecido pelo
  usuário no nível de `run`/`rPr` (não só o texto visível), o manual do CGAEM aplica negrito
  de forma consistente a: título de livro/monografia; número de norma NBR (ex.: "**NBR
  6023**"); nome de instrumento legal com data (ex.: "**Lei nº 9.394, de 20 de dezembro de
  1996**"). Título de artigo de periódico e nome do periódico **não** aparecem em negrito no
  modelo. Essa convenção está implementada em `reference_title_span()` (motor, desde
  27/08/2026) e é aplicada com o mesmo critério nos dois perfis (`generic` e `cgaem`), por
  ser uma leitura direta da NBR 6023 e não uma regra exclusiva do CGAEM.

## O que este perfil ainda não automatiza

- Contagem de páginas (mínimo/máximo).
- Presença e formatação específica da folha de aprovação, do termo de cessão de direitos e
  da ficha catalográfica, quando exigida.
- Verificação de que o desenvolvimento evita literalmente a palavra "Desenvolvimento" como
  título de seção.

Essas verificações continuam exigindo conferência humana antes da entrega.
