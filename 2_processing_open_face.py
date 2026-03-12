from __future__ import annotations

import argparse
import os
import subprocess
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk
except Exception:  # noqa: BLE001
    tk = None
    ttk = None

VIDEO_EXTENSIONS = {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"}
DEFAULT_OPENFACE_CANDIDATES = [
    Path("FeatureExtraction.exe"),
    Path.home() / "Desktop" / "OpenFace_2.2.0_win_x64" / "FeatureExtraction.exe",
    Path.home() / "Рабочий стол" / "OpenFace_2.2.0_win_x64" / "FeatureExtraction.exe",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Запуск OpenFace для актёров или респондентов.")
    parser.add_argument("--mode", choices=("actor", "respondent"), default=None)
    parser.add_argument("--openface-exe", type=Path, default=None)
    parser.add_argument("--actor-dir", type=Path, default=Path("stimuli"))
    parser.add_argument("--respondents-dir", type=Path, default=Path("data"))
    parser.add_argument("--output-dir", type=Path, default=Path("processed_openface"))
    parser.add_argument("--dry-run", action="store_true", help="Только показать, что будет обработано")
    parser.add_argument(
        "--no-gui",
        action="store_true",
        help="Не использовать GUI-окно выбора режима",
    )
    return parser.parse_args()


def choose_mode_gui() -> str | None:
    if tk is None:
        return None

    try:
        root = tk.Tk()
    except Exception:  # noqa: BLE001
        return None
    root.title("OpenFace: выбор режима")
    root.resizable(False, False)

    selected = tk.StringVar(value="actor")
    result: dict[str, str | None] = {"mode": None}

    frame = ttk.Frame(root, padding=16)
    frame.grid(row=0, column=0)

    ttk.Label(frame, text="Выберите режим обработки:").grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Radiobutton(frame, text="Актёры (stimuli)", variable=selected, value="actor").grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(8, 2)
    )
    ttk.Radiobutton(frame, text="Респонденты (data)", variable=selected, value="respondent").grid(
        row=2, column=0, columnspan=2, sticky="w", pady=(0, 10)
    )

    def on_start() -> None:
        result["mode"] = selected.get()
        root.destroy()

    def on_cancel() -> None:
        result["mode"] = None
        root.destroy()

    ttk.Button(frame, text="Старт", command=on_start).grid(row=3, column=0, padx=(0, 8))
    ttk.Button(frame, text="Отмена", command=on_cancel).grid(row=3, column=1)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result["mode"]


def resolve_mode(mode: str | None, actor_dir: Path, respondents_dir: Path, no_gui: bool) -> str:
    if mode:
        return mode

    if not no_gui:
        picked = choose_mode_gui()
        if picked:
            return picked

    if actor_dir.exists() and not respondents_dir.exists():
        print("[INFO] --mode не указан, выбран режим actor (найдена папка stimuli).")
        return "actor"
    if respondents_dir.exists() and not actor_dir.exists():
        print("[INFO] --mode не указан, выбран режим respondent (найдена папка data).")
        return "respondent"

    print("[INFO] --mode не указан, выбран режим actor по умолчанию.")
    return "actor"


def resolve_openface_exe(openface_exe: Path | None) -> Path:
    if openface_exe:
        return openface_exe

    env_path = os.getenv("OPENFACE_EXE")
    if env_path:
        return Path(env_path)

    for candidate in DEFAULT_OPENFACE_CANDIDATES:
        if candidate.exists():
            return candidate

    checked = "\n- ".join(str(p) for p in DEFAULT_OPENFACE_CANDIDATES)
    raise FileNotFoundError(
        "Не найден FeatureExtraction.exe. Укажите --openface-exe или OPENFACE_EXE."
        f"\nПроверены пути:\n- {checked}"
    )


def collect_videos(mode: str, actor_dir: Path, respondents_dir: Path) -> tuple[Path, list[Path]]:
    root = actor_dir if mode == "actor" else respondents_dir
    if not root.exists():
        raise FileNotFoundError(f"Папка не найдена: {root}")

    videos = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in VIDEO_EXTENSIONS]
    videos.sort()

    if not videos:
        raise FileNotFoundError(f"В папке {root} не найдено видеофайлов.")

    return root, videos


def run_openface_for_video(openface_exe: Path, video: Path, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [str(openface_exe), "-f", str(video), "-out_dir", str(out_dir)]
    completed = subprocess.run(cmd, check=False, capture_output=True, text=True)
    if completed.returncode != 0:
        raise RuntimeError(
            f"OpenFace error for {video}\nSTDOUT:\n{completed.stdout}\nSTDERR:\n{completed.stderr}"
        )


def print_progress(current: int, total: int) -> None:
    ratio = current / total if total else 1
    percent = int(ratio * 100)
    width = 32
    filled = int(width * ratio)
    bar = "█" * filled + "-" * (width - filled)
    print(f"\rПрогресс: [{bar}] {percent}% ({current}/{total})", end="", flush=True)
    if current == total:
        print()


def main() -> None:
    args = parse_args()
    mode = resolve_mode(args.mode, args.actor_dir, args.respondents_dir, args.no_gui)
    source_root, videos = collect_videos(mode, args.actor_dir, args.respondents_dir)

    openface_exe: Path | None = None
    if not args.dry_run:
        openface_exe = resolve_openface_exe(args.openface_exe)

    total = len(videos)
    processed = 0

    for video in videos:
        rel_parent = video.parent.relative_to(source_root)
        out_dir = args.output_dir / mode / rel_parent

        if args.dry_run:
            print(f"[DRY-RUN] {video} -> {out_dir}")
            processed += 1
            print_progress(processed, total)
            continue

        run_openface_for_video(openface_exe, video, out_dir)
        processed += 1
        print_progress(processed, total)
        print(f"[OK] {video} -> {out_dir}")

    print(f"Готово. Режим: {mode}. Обработано видео: {processed}")


if __name__ == "__main__":
    main()
