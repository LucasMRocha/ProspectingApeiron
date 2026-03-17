# Contributing to Prospecting Apeiron

Este documento orienta práticas de versionamento e colaboração para este projeto.

## Fluxo de trabalho Git (GitHub)

1. Sincronize sua branch local:
   - `git checkout main`
   - `git pull origin main`

2. Crie uma branch de feature ou bugfix:
   - `git checkout -b feature/<nome>`

3. Faça mudanças pequenas e commit:
   - `git add .`
   - `git commit -m "feat: descrição clara do que foi feito"`

4. Push para remoto:
   - `git push origin feature/<nome>`

5. Abra Pull Request em `main` com descrição e checklist.

6. Código deve passar no CI antes de merge (GitHub Actions `zoho-sync`).

7. Após merge, atualizar local:
   - `git checkout main`
   - `git pull origin main`

## Regras de commit

- prefixos sugeridos: `feat:`, `fix:`, `docs:`, `chore:`, `test:`
- descrição curta + verbo no presente

## Branches

- `main`: produção/histórico aplicável
- `dev`: integração contínua de features em desenvolvimento
- `release/v1`, `release/v2` etc para marcos de versão

## Tags de release

Para versão estável:
- `git checkout main`
- `git tag -a v1.0.0 -m "Release v1.0.0"`
- `git push origin v1.0.0`
