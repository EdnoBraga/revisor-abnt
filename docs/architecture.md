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

## Capacidades

- `.docx`: validação como pacote Office, auditoria e formatação.
- `.doc`: conversão pelo LibreOffice, seguida do mesmo fluxo de `.docx`.
- Regras automatizáveis: página A4, margens, fonte, espaçamento, parágrafo, cabeçalhos,
  referências como bloco, citações longas identificáveis, sumário/campos de atualização e
  relatório de pendências.
- Regras dependentes de evidência: fidelidade de citação, página da fonte, autoria, DOI, URL,
  data de acesso e dados bibliográficos. Estas devem ser revisadas com fontes confirmadas.

## Limite consciente do MVP

O processamento usa `BackgroundTasks`, apropriado para uma instância única e tarefas de curta
duração. Antes de abrir o serviço para público, substituir por fila persistente e worker
dedicado, adicionar autenticação, banco de dados, armazenamento de objetos, expiração automática
de arquivos, limites por conta, antivírus e trilha de auditoria.

## Próximas entregas de produto

1. Login, consentimento e política de privacidade/LGPD.
2. Modelo de documento institucional: faculdade, curso e normas próprias do aluno.
3. Tela de achados com proposta de correção e aprovação por item.
4. Motor assistido para correspondência citação–referência, que só aplica alteração com
   metadados comprovados.
5. Fila e armazenamento seguro para produção.
