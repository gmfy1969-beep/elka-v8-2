# ELKA v8.2 Project

## Обязательно перед началом работы

При любом подключении агента к проекту:
1. Прочитать этот файл полностью.
2. Изучить структуру проекта (`tree` / `find`).
3. Прочитать `README_RU.txt`.
4. Проверить `SHA256SUMS.txt` и git status.
5. Не удалять и не перезаписывать файлы без подтверждения.

## Структура проекта

```
elka-v8-2/
├── CLAUDE.md          # инструкции для ИИ-агентов
├── README_RU.txt      # описание проекта
├── .gitignore
├── 1_INSTALL_OPERATOR.cmd
├── 2_INSTALL_DEMO.cmd
├── 3_INSTALL_VIEWER.cmd
├── 4_UPDATE_MACROS.cmd
├── SHA256SUMS.txt
├── Docs/              # документация
├── Setup/             # установка (VBA, макросы, книги)
├── Demo/              # демо-данные
├── Tests/             # тесты и харнессы
└── Releases/          # дистрибутивы (zip, pdf)
```

## Репозитории

- GitHub: https://github.com/gmfy1969-beep/elka-v8-2 (origin)
- GitLab: https://gitlab.com/gmfy1969-beep/elka-v8-2 (gitlab)

## Правила коммитов

- Ветка по умолчанию `main`.
- Коммиты на русском или английском, понятные.
- Перед push — `git status`, `git diff --stat`.
- Push на оба remote: `git push origin main && git push gitlab main`.
- Не коммитить секреты, токены, личные данные.

## Работа с файлами

- Бинарные дистрибутивы хранятся в `Releases/`.
- Исходники тестов/харнесов в `Tests/Source_for_developers/`.
- Не менять файлы в `Setup/Object_Code` без согласования.

## Агентам

- При задаче сначала изучить проект (структура, README, git log).
- Любые изменения — через отдельную ветку + PR или с подтверждением.
- После изменений — запуск тестов (если есть) и обновление документации.
