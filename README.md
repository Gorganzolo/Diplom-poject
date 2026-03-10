# Программа эксперимента с видеостимулами

## Установка зависимостей

Рекомендуется использовать отдельное виртуальное окружение:

```bash
python -m venv .venv
# Windows (PowerShell)
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Если используете IDE (PyCharm/VS Code), выберите интерпретатор из `.venv`,
иначе IDE может показывать ошибки вида «модуль `cv2` не найден» или
«не удаётся найти `PySide6`-классы».

## Подготовка

1. Положите видео-стимулы в папку `stimuli`.
2. Запустите программу:

```bash
python experiment_app.py
```

## Что делает запись камеры

- Пытается включить максимальное доступное разрешение (с приоритетом 4K → 1440p → 1080p и ниже).
- Пытается использовать высокий FPS, затем замеряет фактическую частоту кадров; запись в файл синхронизируется по реальному времени, чтобы видео не ускорялось.
- Запись стартует вместе с начальным 5-секундным отсчётом.
- Сохраняет файл `face_record.mp4` (на Windows сначала `mp4v` для стабильности, затем H.264-варианты).

## Результаты

Видео лица сохраняется в:

`data/<Фамилия>/attempt_XXX/face_record.mp4`

## Обработка через OpenFace (актеры и респонденты)

Скрипт `scripts/process_actor.py` теперь делает только одно: запускает `FeatureExtraction.exe`
для набора видео и сохраняет CSV OpenFace.

Поддерживаются два режима:

- `--mode actor` — обрабатывает стимулы из `stimuli`;
- `--mode respondent` — обрабатывает видео респондентов из `data`.

Если `--mode` не указывать, откроется GUI-окно выбора режима (actor/respondent).
Если GUI недоступен/отключён, скрипт выберет режим автоматически по папкам (`stimuli`/`data`) или возьмёт `actor` по умолчанию.

Результаты сохраняются в `processed_openface/<mode>/...` с сохранением структуры подпапок.
Во время работы отображается прогресс-бар в консоли (для `actor` и `respondent`).

GUI используется только для выбора режима. Отключить GUI-выбор можно флагом `--no-gui`.

Примеры (Windows):

```bash
# Актёры
python scripts/process_actor.py --mode actor ^
  --openface-exe "%USERPROFILE%\Desktop\OpenFace_2.2.0_win_x64\FeatureExtraction.exe"

# Респонденты
python scripts/process_actor.py --mode respondent ^
  --openface-exe "%USERPROFILE%\Desktop\OpenFace_2.2.0_win_x64\FeatureExtraction.exe"
```

Если `--openface-exe` не передан, скрипт попробует:
- `OPENFACE_EXE`;
- `FeatureExtraction.exe` в текущей папке;
- `%USERPROFILE%\Desktop\OpenFace_2.2.0_win_x64\FeatureExtraction.exe`;
- `%USERPROFILE%\Рабочий стол\OpenFace_2.2.0_win_x64\FeatureExtraction.exe`.

Проверка без запуска OpenFace:

```bash
python scripts/process_actor.py --mode actor --dry-run
python scripts/process_actor.py --mode respondent --dry-run
```
