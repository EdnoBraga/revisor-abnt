# Checklist de publicação — Revisor ABNT

## Antes de publicar

- [ ] Testes locais e CI aprovados.
- [ ] Domínio exclusivo definido e DNS `A` apontando para a VPS.
- [ ] Portas 80 e 443 liberadas; SSH protegido por chave e firewall.
- [ ] `APP_DOMAIN` e `ACME_EMAIL` configurados no `.env` do servidor.
- [ ] Canal de privacidade e identificação do responsável pelo serviço definidos.
- [ ] Prazo de retenção revisado; padrão do MVP: 24 horas.
- [ ] Aviso de privacidade jurídico revisado antes da abertura pública.
- [ ] Backup não inclui uploads de trabalhos acadêmicos.

## Implantar e verificar

- [ ] `docker compose up -d --build` conclui sem erro.
- [ ] `https://DOMINIO/api/health` retorna `status: ok`.
- [ ] Certificado HTTPS válido no domínio.
- [ ] Upload de DOCX de teste gera novo DOCX e relatório.
- [ ] Upload de DOC de teste converte corretamente.
- [ ] Arquivo não aparece em URL pública previsível.
- [ ] Verificação visual do documento no Word confirma campos de sumário e paginação.

## Critérios de reversão

- Upload, download ou processamento falha no fluxo básico.
- O serviço expõe arquivo de um usuário a outro.
- HTTPS não é válido ou o domínio encaminha para serviço incorreto.
- Há erro recorrente de conversão ou consumo anormal de CPU/memória.
- Retenção não elimina os arquivos conforme configurado.
