"""Background music while QS CORE · Walls & Rooms draws in AutoCAD or writes Excel.

Same player as the QS CORE main window (Windows MCI, no extra packages) and the same settings:
Settings > Sound in QS CORE (stored in ~/.dwg_measure_extractor.json: music, music_volume, music_file).
"""
from __future__ import annotations

import json
import os
import time

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_TRACK = os.path.join(HERE, "config", "qs_track_core.mp3")
QS_CORE_CFG = os.path.join(os.path.expanduser("~"), ".dwg_measure_extractor.json")


def settings() -> dict:
    """{'on': bool, 'volume': 0..100, 'path': str} from the QS CORE settings (defaults if missing)."""
    cfg = {}
    try:
        with open(QS_CORE_CFG, encoding="utf-8") as fh:
            cfg = json.load(fh)
    except (OSError, ValueError):
        pass
    f = str(cfg.get("music_file") or "").strip().strip('"')
    try:
        vol = int(cfg.get("music_volume", 35))
    except (TypeError, ValueError):
        vol = 35
    return {"on": bool(cfg.get("music", True)), "volume": vol,
            "path": f if f and os.path.isfile(f) else DEFAULT_TRACK}


class BgMusic:
    """Background music for long jobs (Excel export, drawing in AutoCAD).

    Windows plays the MP3 through MCI (winmm.dll, part of Windows - no extra packages); it loops,
    fades in / out and never blocks the window. On other systems it is a silent no-op.
    Call everything from the Tk (UI) thread."""
    ALIAS = "qscore_rooms_bg"
    MIN_PLAY = 4.0                 # a short job still gets a few seconds before the fade-out

    def __init__(self, root, alias=None):
        self.root = root
        self.alias = alias or self.ALIAS
        self.path = None
        self.is_open = False
        self.playing = False
        self.vol = 0.0             # current MCI volume 0..1000
        self.target = 0.0
        self.t_start = 0.0
        self._fade_job = None
        self._loop_job = None
        self.error = ""
        self._mci = None
        if os.name == "nt":
            try:
                import ctypes
                self._ct = ctypes
                self._mci = ctypes.windll.winmm.mciSendStringW
            except Exception as ex:                      # pragma: no cover - Windows only
                self.error = str(ex)

    @property
    def available(self):
        return self._mci is not None

    def _cmd(self, s):
        if self._mci is None:
            return 1, ""
        buf = self._ct.create_unicode_buffer(256)
        err = self._mci(s, buf, 255, 0)
        return err, buf.value

    def _open(self, path):
        if self.is_open and path == self.path:
            return True
        self.close()
        err, _ = self._cmd('open "%s" type mpegvideo alias %s' % (path, self.alias))
        if err:
            err, _ = self._cmd('open "%s" alias %s' % (path, self.alias))
        self.is_open = not err
        self.path = path if self.is_open else None
        if err:
            self.error = "MCI could not open the music file (error %d)" % err
        return self.is_open

    def start(self, path, volume=35):
        """Fade the track in (volume 0..100) and keep it looping until stop()."""
        if not self.available or not path or not os.path.isfile(path):
            return False
        if not self._open(path):
            return False
        old = getattr(self, "_stop_job", None)
        if old:
            try:
                self.root.after_cancel(old)
            except Exception:
                pass
            self._stop_job = None
        if not self.playing:
            self._cmd("setaudio %s volume to 0" % self.alias)
            self.vol = 0.0
            err, _ = self._cmd("play %s from 0 repeat" % self.alias)
            if err:
                self._cmd("play %s from 0" % self.alias)
            self.playing = True
            self.t_start = time.time()
            self._watch_loop()
        self.target = max(0.0, min(100.0, float(volume))) * 10.0
        self._fade()
        return True

    def set_volume(self, volume):
        if self.playing:
            self.target = max(0.0, min(100.0, float(volume))) * 10.0
            self._fade()

    def stop(self, now=False):
        """Fade out (after the minimum play time) and stop."""
        if not self.playing:
            return
        wait = 0.0 if now else max(0.0, self.MIN_PLAY - (time.time() - self.t_start))
        if wait > 0:
            self._later("_stop_job", int(wait * 1000), self.stop)
            return
        self.target = 0.0
        self._fade(step=40.0 if not now else 1000.0)

    def _later(self, name, ms, fn):
        old = getattr(self, name, None)
        if old:
            try:
                self.root.after_cancel(old)
            except Exception:
                pass
        try:
            setattr(self, name, self.root.after(ms, fn))
        except Exception:
            setattr(self, name, None)

    def _fade(self, step=None):
        self._fade_job = None
        if not self.playing:
            return
        d = self.target - self.vol
        st = step or (25.0 if d > 0 else 40.0)            # ~1.2 s in, ~1.4 s out
        self.vol = self.target if abs(d) <= st else self.vol + (st if d > 0 else -st)
        self._cmd("setaudio %s volume to %d" % (self.alias, int(self.vol)))
        if self.vol <= 0 and self.target <= 0:
            self._cmd("stop %s" % self.alias)
            self.playing = False
            return
        if self.vol != self.target:
            self._later("_fade_job", 50, lambda: self._fade(step))

    def _watch_loop(self):
        """Some MCI devices ignore 'repeat': restart the track when it reaches the end."""
        self._loop_job = None
        if not self.playing:
            return
        err, mode = self._cmd("status %s mode" % self.alias)
        if not err and mode.strip().lower() == "stopped":
            self._cmd("play %s from 0" % self.alias)
        self._later("_loop_job", 700, self._watch_loop)

    def close(self):
        if self.is_open:
            self._cmd("stop %s" % self.alias)
            self._cmd("close %s" % self.alias)
        self.is_open = False
        self.playing = False
        self.path = None


_player = None


def start(root):
    """Fade the music in (if switched on in QS CORE settings)."""
    global _player
    st = settings()
    if not st["on"]:
        return
    if _player is None:
        _player = BgMusic(root)
    _player.start(st["path"], st["volume"])


def stop():
    if _player is not None:
        _player.stop()


def close():
    if _player is not None:
        _player.close()
