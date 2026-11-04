# AGENTS.md — онбординг агентов

При любом подключении агента к проекту выполнить:

1. Прочитать `CLAUDE.md`.
2. Изучить структуру: `find . -maxdepth 3 -type f | head -50`.
3. Прочитать `README_RU.txt`.
4. Проверить git: `git status`, `git log --oneline -10`, `git remote -v`.
5. Проверить `SHA256SUMS.txt`.
6. Не удалять/не перезаписывать файлы без подтверждения.

## Структура

- `Docs/` — документация
- `Setup/` — установка (VBA, макросы, книги)
- `Demo/` — демо-данные
- `Tests/` — тесты и харнессы
- `Releases/` — дистрибутивы (zip, pdf)
- `Docs/notes/` — заметки и журналы

## Репозитории

- origin → GitHub: https://github.com/gmfy1969-beep/elka-v8-2
- gitlab → GitLab: https://gitlab.com/gmfy1969-beep/elka-v8-2

## Push

`git push origin main && git push gitlab main`
