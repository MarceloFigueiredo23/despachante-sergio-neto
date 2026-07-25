# Despachante Sérgio Neto — site one-page

Landing page institucional para o Instagram [@despachante_sergioneto](https://www.instagram.com/despachante_sergioneto/), Cosmópolis-SP.

## Dados usados (v1)

| Campo | Valor |
|-------|--------|
| Nome | Despachante Sérgio Neto (Biju) |
| Endereço | Av. Éster, 119 — Bela Vista III, Cosmópolis-SP · CEP 13150-000 |
| Telefone | (19) 3872-1151 |
| Instagram | [@despachante_sergioneto](https://www.instagram.com/despachante_sergioneto/) |
| Maps | [Sergio Biju despachante](https://www.google.com/maps/place/Sergio+Biju+despachante/@-22.6485522,-47.1855965,17z) |

> **WhatsApp:** na v1 o botão usa o mesmo número do telefone (`551938721151`). Se o celular for outro, troque todos os links `wa.me` no `index.html`.

## Ver localmente

Abra o arquivo `index.html` no navegador (duplo clique) ou:

```bash
npx --yes serve .
```

## Publicar no GitHub Pages (grátis)

1. Faça login no GitHub CLI (uma vez):

```bash
gh auth login
```

2. Crie o repositório e publique:

```bash
cd "C:\Users\marce\OneDrive\Área de Trabalho\despachante-sergio-neto"
git init
git add .
git commit -m "Primeira versão do site one-page Despachante Sérgio Neto"
gh repo create despachante-sergio-neto --public --source=. --remote=origin --push
gh pages deploy . --branch main
```

Ou, no GitHub: **Settings → Pages → Deploy from branch → `main` / root**.

URL típica: `https://SEU-USUARIO.github.io/despachante-sergio-neto/`

3. Cole esse link na bio do Instagram.
