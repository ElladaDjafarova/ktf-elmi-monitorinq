# -*- coding: utf-8 -*-
"""
KTF Elmi Monitorinq — masaüstü tətbiq (Windows .exe)

- Dashboard brauzerdə yox, tətbiqin öz pəncərəsində açılır (pywebview / Edge WebView2).
- Yazıla bilən fayllar:  %APPDATA%\\KTF Elmi Monitorinq  (config, məlumat, dashboard)
- Müəllim siyahısı:      Sənədlər\\KTF Elmi Monitorinq\\muellimler.xlsx  (rahat redaktə üçün)
- Python quraşdırmaq lazım deyil — hər şey .exe-nin içindədir.
"""
import json
import os
import shutil
import sys
import threading
import time
import traceback

APP_NAME = "KTF Elmi Monitorinq"


def res_dir():
    """Proqramla gələn fayllar (PyInstaller-də _MEIPASS)."""
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def home_dir():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, APP_NAME)


def documents_dir():
    if os.name == "nt":
        try:  # Windows-da "Sənədlər" qovluğunun real yolu (OneDrive-a köçürülmüş ola bilər)
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)  # CSIDL_PERSONAL
            if buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def prepare():
    res, home = res_dir(), home_dir()
    os.makedirs(os.path.join(home, "data"), exist_ok=True)

    # jurnal faylı — pəncərəli .exe-də konsol yoxdur
    log = open(os.path.join(home, "data", "log.txt"), "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stderr = log

    # müəllim siyahısı Sənədlər qovluğunda — istifadəçi onu rahat açıb dəyişə bilsin
    docs = os.path.join(documents_dir(), APP_NAME)
    os.makedirs(docs, exist_ok=True)
    xl = os.path.join(docs, "muellimler.xlsx")
    bundled = os.path.join(res, "muellimler.xlsx")
    stamp = os.path.join(home, "excel_version.txt")  # son quraşdırılan siyahının imzası

    def digest(path):
        import hashlib
        with open(path, "rb") as fh:
            return hashlib.sha256(fh.read()).hexdigest()

    new_sig = digest(bundled)
    old_sig = open(stamp).read().strip() if os.path.exists(stamp) else None
    if not os.path.exists(xl):
        shutil.copyfile(bundled, xl)
    elif new_sig != old_sig and old_sig is not None and digest(xl) == old_sig:
        # yeni versiyada siyahı yenilənib, istifadəçi öz faylını dəyişməyib → yenisi ilə əvəz et
        shutil.copyfile(bundled, xl)
    elif old_sig is None and digest(xl) != new_sig:
        # ilk versiyadan yenilənmə (imza yox idi): köhnə faylı saxla, yenisini qoy
        shutil.copyfile(xl, os.path.join(docs, "muellimler_kohne.xlsx"))
        shutil.copyfile(bundled, xl)
    with open(stamp, "w") as fh:
        fh.write(new_sig)

    # config: ilk dəfə proqramla gələni köçür, Excel yolunu Sənədlərə yönəlt
    cfg_path = os.path.join(home, "config.json")
    if not os.path.exists(cfg_path):
        shutil.copyfile(os.path.join(res, "config.json"), cfg_path)
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)
    if cfg.get("excel_file") != xl:
        cfg["excel_file"] = xl
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

    os.environ["KTF_HOME"] = home
    os.environ["KTF_RES"] = res
    return docs


class Args:
    """scopus_monitor.serve() üçün arqumentlər."""
    demo = False
    no_open = True
    no_refresh = False
    loop = 24.0


def error_box(msg):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, msg, APP_NAME, 0x10)
    except Exception:
        print(msg)


def main():
    docs = prepare()
    import scopus_monitor as sm  # KTF_HOME / KTF_RES təyin olunandan sonra

    url = sm.serve(Args(), 8765, background=True)
    if not url:
        error_box("Proqram başlaya bilmədi (port tapılmadı). Kompüteri yenidən başladıb yenə cəhd edin.")
        return

    # server cavab verənə qədər gözlə
    import urllib.request
    for _ in range(50):
        try:
            urllib.request.urlopen(url + "api/status", timeout=1).read()
            break
        except Exception:
            time.sleep(0.2)

    class Api:
        def open_excel(self):
            os.startfile(os.path.join(docs, "muellimler.xlsx"))

        def open_folder(self):
            os.startfile(docs)

    try:
        import webview
        window = webview.create_window(APP_NAME, url, width=1320, height=900, min_size=(900, 600),
                                       js_api=Api(), text_select=True)
        webview.start(private_mode=False, storage_path=os.path.join(home_dir(), "webview"))
    except Exception:
        # WebView2 yoxdursa — adi brauzerdə aç və proqramı arxada saxla
        traceback.print_exc()
        import webbrowser
        webbrowser.open(url)
        error_box("Dashboard brauzerdə açıldı.\n\nBu pəncərəni bağlayanda proqram da bağlanacaq.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        error_box("Gözlənilməz xəta baş verdi. Ətraflı məlumat:\n"
                  + os.path.join(home_dir(), "data", "log.txt"))
    os._exit(0)
