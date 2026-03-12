import re
import sys
import threading
import time
from pathlib import Path

import cv2
from PySide6.QtCore import QTimer, Qt, QUrl
from PySide6.QtGui import QFont
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

STIMULI_DIR = Path("stimuli")
DATA_DIR = Path("data")
SUPPORTED_EXTENSIONS = {".mp4", ".avi", ".mov", ".mkv"}

CAMERA_WIDTH = 1280
CAMERA_HEIGHT = 720
CAMERA_FPS = 30.0


class CameraRecorder:
    def __init__(self, output_file: Path) -> None:
        self.output_file = output_file
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self.capture: cv2.VideoCapture | None = None
        self.writer: cv2.VideoWriter | None = None

    @staticmethod
    def _set_capture_params(capture: cv2.VideoCapture) -> None:
        capture.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*"MJPG"))
        capture.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
        capture.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
        capture.set(cv2.CAP_PROP_FPS, CAMERA_FPS)

    def _create_writer(self) -> cv2.VideoWriter:
        frame_size = (CAMERA_WIDTH, CAMERA_HEIGHT)
        for code in ("mp4v", "avc1", "H264", "X264"):
            writer = cv2.VideoWriter(
                str(self.output_file),
                cv2.VideoWriter_fourcc(*code),
                CAMERA_FPS,
                frame_size,
            )
            if writer.isOpened():
                return writer
        raise RuntimeError("Не удалось инициализировать видеозапись.")

    def _record_loop(self) -> None:
        assert self.capture is not None
        assert self.writer is not None

        frame_interval = 1.0 / CAMERA_FPS
        next_frame_time = time.perf_counter()

        while not self._stop_event.is_set():
            ok, frame = self.capture.read()
            if not ok:
                time.sleep(0.005)
                continue

            resized = cv2.resize(frame, (CAMERA_WIDTH, CAMERA_HEIGHT), interpolation=cv2.INTER_AREA)
            self.writer.write(resized)

            next_frame_time += frame_interval
            sleep_for = next_frame_time - time.perf_counter()
            if sleep_for > 0:
                time.sleep(sleep_for)
            else:
                next_frame_time = time.perf_counter()

    def start(self) -> None:
        cap_api = cv2.CAP_DSHOW if sys.platform.startswith("win") else cv2.CAP_ANY
        self.capture = cv2.VideoCapture(0, cap_api)

        if not self.capture.isOpened():
            raise RuntimeError("Камера недоступна.")

        self._set_capture_params(self.capture)
        self.writer = self._create_writer()

        self._thread = threading.Thread(target=self._record_loop, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=3)

        if self.writer is not None:
            self.writer.release()
        if self.capture is not None:
            self.capture.release()


class ParticipantWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Эксперимент")

        layout = QVBoxLayout()

        self.label = QLabel("Введите фамилию участника:")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.surname_input = QLineEdit()
        self.surname_input.setPlaceholderText("Фамилия")

        self.start_button = QPushButton("Начать эксперимент")
        self.start_button.clicked.connect(self.start_experiment)

        layout.addWidget(self.label)
        layout.addWidget(self.surname_input)
        layout.addWidget(self.start_button)

        self.setLayout(layout)
        self.resize(420, 180)

    def _collect_stimuli(self) -> list[Path]:
        if not STIMULI_DIR.exists():
            raise RuntimeError("Папка stimuli не найдена.")

        stimuli = sorted(
            [p for p in STIMULI_DIR.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS],
            key=lambda p: p.name.lower(),
        )

        if not stimuli:
            raise RuntimeError("В папке stimuli нет поддерживаемых видеофайлов.")

        return stimuli

    def _create_attempt_folder(self, surname: str) -> Path:
        participant_dir = DATA_DIR / surname
        participant_dir.mkdir(parents=True, exist_ok=True)

        attempts: list[int] = []
        for attempt_path in participant_dir.iterdir():
            if not attempt_path.is_dir():
                continue
            match = re.fullmatch(r"attempt_(\d{3})", attempt_path.name)
            if match is None:
                continue
            attempts.append(int(match.group(1)))

        next_attempt = max(attempts, default=0) + 1
        attempt_dir = participant_dir / f"attempt_{next_attempt:03d}"
        attempt_dir.mkdir(parents=True, exist_ok=False)
        return attempt_dir

    def start_experiment(self) -> None:
        surname = self.surname_input.text().strip()
        if not surname:
            QMessageBox.warning(self, "Ошибка", "Введите фамилию участника.")
            return

        try:
            stimuli = self._collect_stimuli()
            attempt_dir = self._create_attempt_folder(surname)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))
            return

        self.experiment_window = ExperimentWindow(stimuli=stimuli, attempt_dir=attempt_dir)
        self.experiment_window.showFullScreen()
        self.experiment_window.show()
        self.close()


class ExperimentWindow(QMainWindow):
    def __init__(self, stimuli: list[Path], attempt_dir: Path) -> None:
        super().__init__()
        self.stimuli = stimuli
        self.attempt_dir = attempt_dir
        self.stimulus_index = 0
        self.current_countdown = 0
        self.countdown_next_step = None

        self.face_record_path = self.attempt_dir / "face_record.mp4"
        self.recorder = CameraRecorder(self.face_record_path)

        self.setWindowTitle("Эксперимент")
        self.setStyleSheet("background-color: black;")

        self.video_widget = QVideoWidget()
        self.player = QMediaPlayer()
        self.audio_output = QAudioOutput()
        self.audio_output.setVolume(1.0)
        self.player.setAudioOutput(self.audio_output)
        self.player.setVideoOutput(self.video_widget)
        self.player.mediaStatusChanged.connect(self._on_media_status_changed)

        self.message_label = QLabel("", self)
        self.message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.message_label.setStyleSheet("color: white;")
        self.message_label.setFont(QFont("Arial", 40, QFont.Weight.Bold))

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.video_widget)
        layout.addWidget(self.message_label)

        self.setCentralWidget(central)

        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self._countdown_tick)
        QTimer.singleShot(0, self._start_experiment_flow)

    def _start_experiment_flow(self) -> None:
        try:
            self.recorder.start()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", f"Не удалось запустить запись: {exc}")
            self.close()
            return

        self._start_countdown("Эксперимент начнётся через", 5, self._play_current_stimulus)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self._finish_experiment(early=True)
            return
        super().keyPressEvent(event)

    def _show_text(self, text: str) -> None:
        self.video_widget.hide()
        self.message_label.show()
        self.message_label.setText(text)

    def _show_video(self) -> None:
        self.message_label.hide()
        self.video_widget.show()

    def _start_countdown(self, title: str, seconds: int, next_step) -> None:
        self.current_countdown = seconds
        self.countdown_title = title
        self.countdown_next_step = next_step
        self._show_text(f"{title}\n\n{self.current_countdown}")
        self.countdown_timer.start(1000)

    def _countdown_tick(self) -> None:
        self.current_countdown -= 1
        if self.current_countdown <= 0:
            self.countdown_timer.stop()
            if self.countdown_next_step:
                self.countdown_next_step()
            return
        self._show_text(f"{self.countdown_title}\n\n{self.current_countdown}")

    def _play_current_stimulus(self) -> None:
        if self.stimulus_index >= len(self.stimuli):
            self._finish_experiment()
            return

        stimulus = self.stimuli[self.stimulus_index]
        self._show_video()
        self.player.setSource(QUrl.fromLocalFile(str(stimulus.resolve())))
        self.player.play()

    def _on_media_status_changed(self, status) -> None:
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            self.player.stop()
            self.stimulus_index += 1
            if self.stimulus_index < len(self.stimuli):
                self._start_countdown("Следующий стимул через", 5, self._play_current_stimulus)
            else:
                self._finish_experiment()

    def _finish_experiment(self, early: bool = False) -> None:
        self.countdown_timer.stop()
        self.player.stop()
        self.recorder.stop()

        if early:
            QMessageBox.information(self, "Эксперимент", "Эксперимент прерван. Данные сохранены.")
        else:
            QMessageBox.information(self, "Эксперимент", "Эксперимент завершён\nДанные сохранены")
        self.close()


def main() -> None:
    app = QApplication(sys.argv)
    window = ParticipantWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
