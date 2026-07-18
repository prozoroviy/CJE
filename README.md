# CX Review

AI-аудит загруженных экранов клиентского пути с классификацией UX/UI-замечаний и экспортом в PPTX.

## Локальный фронтенд

```bash
pnpm install
pnpm dev
```

Для настоящего анализа нужен serverless endpoint `api/analyze.js`. Проще всего запускать весь проект через Vercel CLI или развернуть репозиторий на Vercel.

## Развёртывание на Vercel

1. Импортируйте GitHub-репозиторий `prozoroviy/CJE` в Vercel.
2. Добавьте переменную окружения `OPENAI_API_KEY` в Project Settings → Environment Variables.
3. При необходимости задайте `OPENAI_MODEL`; по умолчанию используется `gpt-5.6-sol`.
4. Запустите Deploy.

Vercel автоматически соберёт Vite-приложение и развернёт функцию `/api/analyze`. Секретный ключ не попадает в браузерный JavaScript.

## Локальный полный запуск

Создайте `.env.local` на основе `.env.example`, затем:

```bash
pnpm dlx vercel dev
```

Не добавляйте `.env` и `.env.local` в git.
