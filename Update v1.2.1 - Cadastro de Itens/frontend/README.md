# Frontend Corpal Solicitações v1.12.86

Aplicativo estático, sem etapa de compilação.

## Mudança isolada

Somente `Solicitação de Chamado TI` abre a nova interface integrada ao GLPI.
Os demais cartões continuam usando os mesmos links da versão enviada pela CORPAL.

## Hospedagem recomendada

Para manter o SSO do Teams alinhado ao domínio definido em `webApplicationInfo.resource`,
a versão de produção deve ser publicada em:

```text
https://chamados.corpalincorporadora.com.br:8078/teams-app/
```

O Git continua sendo o repositório-fonte. A publicação é feita copiando os arquivos
estáticos para `/opt/corpal-teams-app` no servidor.

## Arquivos principais

- `index.html`
- `style.css`
- `app.js`
- `config.js`
- `manifest.template.json`
- `build-teams-package.sh`
- `build-teams-package.ps1`

## Segurança

Não coloque credenciais em `config.js` ou `app.js`.
O frontend envia somente o token SSO do Teams ao backend.
A antiga página estática `painel.html` e a senha JavaScript foram removidas desta versão.
