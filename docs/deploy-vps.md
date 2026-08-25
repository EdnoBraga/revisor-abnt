# Publicação em VPS com HTTPS

Este guia prepara um servidor Linux com Docker para hospedar o Revisor ABNT. Ele não substitui
a política de privacidade, autenticação e monitoramento necessários para uma operação pública.

## Pré-requisitos

- VPS Linux com Docker Engine e Docker Compose Plugin instalados.
- Domínio ou subdomínio exclusivo, por exemplo `revisor.seudominio.com.br`.
- Registro DNS `A` desse domínio apontando para o IP público da VPS.
- Portas TCP 80 e 443 liberadas no firewall e no provedor.
- Um e-mail monitorado para a emissão e avisos do certificado HTTPS.

## Implantação

```bash
git clone https://github.com/EdnoBraga/revisor-abnt.git
cd revisor-abnt
cp .env.example .env
nano .env
docker compose up -d --build
docker compose ps
docker compose logs --tail=100 caddy
```

Preencha `APP_DOMAIN` com o domínio real e `ACME_EMAIL` com o e-mail monitorado. O Caddy emitirá
o certificado somente quando o DNS e as portas estiverem corretos.

## Verificação após publicar

1. Abra `https://SEU_DOMINIO/api/health` e confirme a resposta `status: ok`.
2. Envie um DOCX de teste, baixe o resultado e abra-o no Word.
3. Verifique a página `/privacidade`, o HTTPS válido e a ausência de listagem pública de jobs.
4. Depois do prazo configurado, confira que o diretório/volume de jobs não mantém arquivos
   concluídos ou com falha.

## Operação segura

- Mantenha `REVISOR_ABNT_RETENTION_HOURS=24` ou um prazo menor enquanto não houver motivo
  documentado para retenção maior.
- Não habilite acesso público definitivo antes de definir o controlador do serviço, um e-mail
  de contato e uma política de privacidade revisada.
- Faça backup apenas de configurações e banco de dados futuros; não use backups amplos de
  uploads de trabalhos acadêmicos.
- Atualize o servidor e as imagens regularmente. Para uma atualização: `git pull && docker
  compose up -d --build`.

## Reversão

Se o serviço apresentar erro após uma atualização, volte ao commit anterior conhecido e recrie
os contêineres:

```bash
git log --oneline -5
git checkout COMMIT_ESTAVEL
docker compose up -d --build
```

Registre o motivo e preserve os logs necessários para diagnóstico sem reter documentos enviados
por usuários além do prazo definido.
