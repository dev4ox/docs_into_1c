#!/bin/bash

# Устанавливаем директорию проекта
PROJECT_DIR="/www/var/docs_into_1c"

# Переходим в директорию проекта
cd $PROJECT_DIR || { echo "Не удалось перейти в директорию $PROJECT_DIR"; exit 1; }

# Выполняем git fetch для получения последних изменений
echo "Выполняем git fetch..."
git fetch || { echo "Ошибка при выполнении git fetch"; exit 1; }

# Выполняем git rebase для применения изменений
echo "Выполняем git rebase..."
git rebase || { echo "Ошибка при выполнении git rebase"; exit 1; }

# Перезагружаем службу docs_into_1c
echo "Перезагружаем процесс docs_into_1c..."
sudo systemctl restart docs_into_1c.service || { echo "Ошибка при перезагрузке службы docs_into_1c"; exit 1; }

# Выводим сообщение о завершении процесса
echo "Обновление и перезагрузка успешно завершены!"
