# Программа эксперимента с видеостимулами

## Структура файлов

- `1_experiment_run.py` — запуск и проведение эксперимента (воспроизведение стимулов + запись лица участника).
- `2_processing_open_face.py` — запуск OpenFace (`FeatureExtraction.exe`) для обработки видео актёров или респондентов.
- `3_process_openface_csv_to_excel.py` — объединение CSV OpenFace в Excel с подсветкой AU-показателей.
- `stimuli/` — папка с видео-стимулами.
- `data/` — результаты по участникам (`<Фамилия>/attempt_XXX/face_record.mp4`).
- `processed_openface/` — результаты обработки OpenFace (CSV-файлы).

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

### Если видите ошибку про NumPy 1.x / NumPy 2.x

Если появляется сообщение вида `A module that was compiled using NumPy 1.x cannot be run in NumPy 2...`,
пересоздайте окружение и установите зависимости заново (в проекте уже зафиксирован `numpy<2`):

```bash
# в корне проекта
rm -rf .venv
python -m venv .venv
# Windows (PowerShell): .venv\Scripts\Activate.ps1
# Linux/macOS:
source .venv/bin/activate
pip install -r requirements.txt
```

## 1) Запуск эксперимента

1. Положите видео-стимулы в папку `stimuli`.
2. Запустите:

```bash
python 1_experiment_run.py
```

### Что делает запись камеры

- Записывает видео лица строго в разрешении **1280×720 (16:9)**.
- Записывает строго с частотой **30 FPS** (без автоподбора параметров камеры).
- Синхронизирует запись по реальному времени, чтобы итоговое видео не было короче/длиннее фактической длительности эксперимента.
- Запись стартует вместе с начальным 5-секундным отсчётом.
- Сохраняет файл `face_record.mp4`.

### Куда сохраняются результаты

`data/<Фамилия>/attempt_XXX/face_record.mp4`

## 2) Обработка видео через OpenFace

Скрипт `2_processing_open_face.py` запускает `FeatureExtraction.exe` и сохраняет CSV OpenFace,
сохраняя структуру подпапок.

Поддерживаются два режима:

- `--mode actor` — обработка стимулов из `stimuli`;
- `--mode respondent` — обработка видео респондентов из `data`.

Если `--mode` не указан:

- сначала предлагается GUI-выбор режима (если доступен),
- иначе режим выбирается автоматически по наличию папок,
- если определить нельзя — берётся `actor` по умолчанию.

Результаты сохраняются в `processed_openface/<mode>/...`.

### Примеры запуска

```bash
# Проверить список входных видео без запуска OpenFace
python 2_processing_open_face.py --mode actor --dry-run
python 2_processing_open_face.py --mode respondent --dry-run

# Реальный запуск
python 2_processing_open_face.py --mode actor --openface-exe "C:/OpenFace/FeatureExtraction.exe"
python 2_processing_open_face.py --mode respondent --openface-exe "C:/OpenFace/FeatureExtraction.exe"
```

### Поиск `FeatureExtraction.exe`

Если `--openface-exe` не передан, скрипт проверяет:

1. переменную окружения `OPENFACE_EXE`;
2. `FeatureExtraction.exe` в текущей папке;
3. `~/Desktop/OpenFace_2.2.0_win_x64/FeatureExtraction.exe`;
4. `~/Рабочий стол/OpenFace_2.2.0_win_x64/FeatureExtraction.exe`.

## 3) Экспорт CSV OpenFace в единый Excel

Скрипт `3_process_openface_csv_to_excel.py` объединяет CSV в Excel-файл
(по умолчанию `openface_processed.xlsx`).

### Базовый запуск

```bash
python 3_process_openface_csv_to_excel.py
```

### Полезные параметры

```bash
python 3_process_openface_csv_to_excel.py --mode actor
python 3_process_openface_csv_to_excel.py --mode respondent --respondent Иванов
python 3_process_openface_csv_to_excel.py --input-root processed_openface --output openface_processed.xlsx
python 3_process_openface_csv_to_excel.py --no-gui
```

### Что добавляется в Excel

- служебные колонки (`frame`, `face_id`, `timestamp`, `confidence`, `success`) если они есть;
- колонки AU из списка `TARGET_AUS` (`*_r`, `*_c`);
- подсветка:
  - `*_c != 0`;
  - `*_r > 0.6`;
  - AU-колонок, значимых для распознанной эмоции в имени файла.

## Быстрая проверка скриптов

```bash
python -m py_compile 1_experiment_run.py 2_processing_open_face.py 3_process_openface_csv_to_excel.py
python 2_processing_open_face.py --help
python 3_process_openface_csv_to_excel.py --help
```
