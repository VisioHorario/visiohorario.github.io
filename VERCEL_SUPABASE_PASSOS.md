# Vercel + Supabase

Arquivos locais que precisam ir para o GitHub:

- `index.html`
- `script.js`
- `vercel.json`
- `api/config.js`
- `.env.vercel.example`
- `supabase/setup_visiogestao.sql`

## 1. Subir para o GitHub

Envie para o repositorio a pasta `api` e os arquivos acima. O problema do `404 /api/config`
acontece quando a Vercel recebe uma versao sem `api/config.js`.

## 2. Configurar variaveis na Vercel

No projeto da Vercel, abra `Settings > Environment Variables` e cadastre:

- `VISIO_SUPABASE_URL`
- `VISIO_SUPABASE_ANON_KEY`
- `VISIO_IA_URL` (opcional)

Depois faca um novo deploy.

## 3. Criar a base no Supabase

No SQL Editor do Supabase, execute o arquivo `supabase/setup_visiogestao.sql`.

Esse script cria:

- `public.tenants`
- `public.app_users`
- `public.visio_snapshots`

Tambem ativa RLS para o usuario autenticado enxergar apenas o proprio tenant.

## 4. Criar o tenant e o usuario

Depois do SQL:

1. Confirme o tenant `ESC001` na tabela `public.tenants`.
2. Crie um usuario em `Authentication > Users`.
3. No usuario criado, preencha `user_metadata` com:

```json
{
  "tenant_id": "UUID_DO_TENANT"
}
```

4. Cadastre esse mesmo usuario na tabela `public.app_users`.

Use como base o exemplo comentado no final de `supabase/setup_visiogestao.sql`.

## 5. Testes apos o deploy

Teste nesta ordem:

1. `https://SEU-PROJETO.vercel.app/api/config`
2. login local provisório: `ESC001 / admin / admin`
3. login Supabase: `codigo da escola + email + senha`
4. salvar algum dado e confirmar leitura em `public.visio_snapshots`

## 6. Observacao importante

O login local `admin/admin` continua liberado no codigo de forma provisoria.
Quando voce quiser, eu posso trocar isso por uma chave unica de controle, como
`LOGIN_LOCAL_LIBERADO = true/false`, para facilitar desligar esse fallback na web.
