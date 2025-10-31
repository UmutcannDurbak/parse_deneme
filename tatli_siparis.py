
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
from copy import copy
import sys
import requests  # pyright: ignore[reportMissingModuleSource]
import json
import zipfile
import shutil
from pathlib import Path
import subprocess
from copy import copy
import sys
import requests  # pyright: ignore[reportMissingModuleSource]
import json
import zipfile
import shutil
from pathlib import Path

DATA_START_ROW = 3  # Verilerin başladığı satır (1-indexed)

# tkinterdnd2 desteği varsa import et
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # pyright: ignore[reportMissingImports]
    TK_DND_AVAILABLE = True
except ImportError:
    TK_DND_AVAILABLE = False

# PyInstaller ile build ederken .ico dosyasını eklemeyi unutmayın!
ICON_PATH = "appicon.ico"
VERSION = "v1.2.5"
DEVELOPER = "Developer U.D"

# Güncelleme ayarları
GITHUB_REPO = "UmutcannDurbak/parse_deneme"  # GitHub repository (owner/repo)
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"  # GitHub API endpoint
UPDATE_CHECK_INTERVAL = 24 * 60 * 60  # 24 saat (saniye cinsinden)
# Eğer güncelleme bulunduğunda otomatik indirme başlatılsın mı? (False = kullanıcı "İndir" butonuna basmalı)
# Otomatik indirmenin varsayılan davranışı: eğer uygulama PyInstaller ile paketlenmişse otomatik indir
"""AUTO_START_DOWNLOAD:
If True, the app will automatically start downloading an available update when it detects a newer release.
We enable this for testing/automation so the app immediately downloads the selected asset.
In production, you may prefer to enable this only when running a packaged exe (frozen).
"""
AUTO_START_DOWNLOAD = True

# Tercih edilen asset uzantı sıralaması — önce .exe, sonra .zip
PREFERRED_ASSET_EXTENSIONS = ['.exe', '.zip']

def resource_path(relative_path):
    import sys, os
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Güncelleme fonksiyonları
def get_latest_version():
    """GitHub'dan en son sürümü kontrol eder"""
    try:
        response = requests.get(GITHUB_API_URL, timeout=10)
        if response.status_code == 200:
            data = response.json()
            return data.get('tag_name', ''), data.get('html_url', '')
        return None, None
    except Exception as e:
        print(f"Güncelleme kontrolü hatası: {e}")
        return None, None

def download_github_update(download_url, progress_callback=None):
    """Stream-download a GitHub asset to 'update.zip'. Returns True on success."""
    try:
        resp = requests.get(download_url, stream=True, timeout=60)
        resp.raise_for_status()
        total = int(resp.headers.get('content-length', 0) or 0)
        downloaded = 0
        with open('update.zip', 'wb') as fh:
            for chunk in resp.iter_content(chunk_size=8192):
                if not chunk:
                    continue
                fh.write(chunk)
                downloaded += len(chunk)
                if progress_callback and total:
                    try:
                        progress_callback((downloaded / total) * 100)
                    except Exception:
                        pass
        return True
    except Exception as e:
        print(f"İndirme hatası: {e}")
        return False


def install_update():
    """Install update from update.zip.
    If running as a frozen exe, extracts to temp and launches a batch updater to replace the running exe.
    Otherwise extracts into current directory.
    Returns True on success (or when updater was launched).
    """
    try:
        frozen = getattr(sys, 'frozen', False) or hasattr(sys, '_MEIPASS')
        if frozen:
            import tempfile
            tmpdir = tempfile.mkdtemp()
            with zipfile.ZipFile('update.zip', 'r') as z:
                z.extractall(tmpdir)
            exe_name = 'tatli_siparis.exe'
            extracted_exe = os.path.join(tmpdir, exe_name)
            if not os.path.exists(extracted_exe):
                print('Kurulum hatası: exe bulunamadı in zip')
                return False

            bat_path = os.path.join(tmpdir, 'updater.bat')
            current_dir = os.path.abspath('.')
            bat = f'''@echo off
            timeout /t 2 /nobreak >nul
            :waitloop
            tasklist /FI "IMAGENAME eq {exe_name}" | find /I "{exe_name}" >nul
            if %ERRORLEVEL%==0 (
            timeout /t 1 /nobreak >nul
            goto waitloop
            )
            copy /Y "{extracted_exe}" "{os.path.join(current_dir, exe_name)}" >nul
            start "" "{os.path.join(current_dir, exe_name)}"
            rmdir /S /Q "{tmpdir}"
            del "%~f0" /Q
            '''
            with open(bat_path, 'w', encoding='utf-8') as f:
                f.write(bat)
            try:
                subprocess.Popen(['cmd', '/c', 'start', '/min', bat_path], shell=False)
            except Exception as e:
                print(f'Updater başlatılamadı: {e}')
                return False
            try:
                os.remove('update.zip')
            except:
                pass
            return True

        # non-frozen
        if os.path.exists('tatli_siparis.exe'):
            shutil.copy('tatli_siparis.exe', 'tatli_siparis_backup.exe')
        with zipfile.ZipFile('update.zip', 'r') as z:
            z.extractall('.')
        try:
            os.remove('update.zip')
        except:
            pass
        return True
    except Exception as e:
        print(f"Kurulum hatası: {e}")
        return False


def check_for_updates(silent=False):
    """Background check for updates. If AUTO_START_DOWNLOAD is True, will auto-download and install.
    Returns True if an update was applied/launched, False otherwise.
    """
    try:
        latest_version, _ = get_latest_version()
        if not latest_version:
            return False
        if not is_newer_version(latest_version, VERSION):
            return False
        # fetch release details
        r = requests.get(GITHUB_API_URL, timeout=10)
        if r.status_code != 200:
            return False
        data = r.json()
        assets = data.get('assets', [])
        best = select_best_asset(assets)
        if not best:
            return False
        dl = best.get('browser_download_url')
        if not dl:
            return False
        if silent and not AUTO_START_DOWNLOAD:
            return False
        ok = download_github_update(dl)
        if not ok:
            return False
        return install_update()
    except Exception:
        return False

def select_best_asset(assets: list):
    """Verilen asset listesi içinden en uygun (tercih edilen uzantıya göre) asset'i döndürür.
    Döndürür: asset dict veya None
    """
    if not assets:
        return None
    # normalize names
    assets_sorted = list(assets)
    # try preferred extensions in order
    for ext in PREFERRED_ASSET_EXTENSIONS:
        for a in assets_sorted:
            name = (a.get('name') or '').lower()
            if name.endswith(ext) or ext.strip('.') in name:
                return a
    # fallback: return first asset
    return assets_sorted[0]


def is_newer_version(latest_version, current_version):
    """Basit semantik versiyon karşılaştırması. 'v' önekini kaldırır ve noktalı int'leri karşılaştırır."""
    try:
        def to_tuple(v):
            v = str(v).lstrip('vV')
            parts = [int(p) for p in v.split('.') if p.isdigit() or p.isnumeric()]
            return tuple(parts)
        return to_tuple(latest_version) > to_tuple(current_version)
    except Exception:
        return False

# Yeni OOP koordinatör (eski fonksiyonlar geriye dönük uyum için içeride kullanılacak)
from shipment_oop import ShipmentCoordinator, clear_workbook_values, format_today_in_workbook, IZMIR_BRANCHES
'''
# Hücre formatını bozmadan sadece ana/master hücreye değer silen fonksiyon
def clear_cell_value_preserve_format(ws, row, col, clear_formulas=False):
    """
    Hücreyi içindeki değeri temizler ancak hücre biçimini/merge yapısını bozmaz.
    - Eğer (row,col) bir merged-range içindeyse, merged aralığın master hücresini temizler.
    - clear_formulas=False ise formülleri silmez (korur).
    Döner: True (bir değer temizlendi), False (zaten boş veya formül korundu).
    """
    # merged-range içinde mi bak
    for mr in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = mr.bounds
        if min_row <= row <= max_row and min_col <= col <= max_col:
            master = ws.cell(row=min_row, column=min_col)
            # formül koruması
            if not clear_formulas and isinstance(master.value, str) and str(master.value).startswith('='):
                return False
            if master.value not in (None, ""):
                master.value = None
                return True
            return False

    # merged değilse direkt hücreyi temizle
    cell = ws.cell(row=row, column=col)
    if not clear_formulas and isinstance(cell.value, str) and str(cell.value).startswith('='):
        return False
    if cell.value not in (None, ""):
        cell.value = None
        return True
    return False
'''
from openpyxl import load_workbook  # pyright: ignore[reportMissingModuleSource]
import datetime

def clear_all_records(status_label, log_widget):
    confirm = messagebox.askyesno("Onay", "Tüm kayıtları silmek/temizlemek istediğinize emin misiniz?")
    if not confirm:
        status_label.config(text="İşlem iptal edildi.")
        return
    try:
        output_path = "sevkiyat_tatlı.xlsx"
        if not os.path.exists(output_path):
            status_label.config(text="❌ Önce bir sevkiyat dosyası oluşturulmalı!")
            messagebox.showerror("Hata", "Önce bir sevkiyat dosyası oluşturulmalı!")
            return
        wb = load_workbook(output_path)
        cleared = 0
        for ws in wb.worksheets:
            # 2. satırdan şube başlıklarını oku
            subeler = {}
            for cell in ws[2][1:]:
                if cell.value:
                    sube_ad = str(cell.value).strip()
                    subeler[sube_ad] = {"tepsi": cell.column, "tepsi_2": cell.column+1, "adet": cell.column+2, "adet_2": cell.column+3}

            # Önemli: sadece gerektiğinde merged-range'i unmerge edeceğiz (tek tek)
            # İterasyona gir
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=1):
                ana_cell = row[0]
                if not ana_cell.value:
                    continue
                ana_ad = str(ana_cell.value).upper()
                skip_keywords = ["SIPARIS TARIHI", "SIPARIS ALAN", "TESLIM TARIHI", "TEYID EDEN"]
                if any(ana_ad.startswith(k) or ana_ad == k for k in skip_keywords):
                    continue
                for sube in subeler.values():
                    for col in [sube["tepsi"], sube["tepsi_2"], sube["adet"], sube["adet_2"]]:
                        # Eğer hedef cell merged bir range'in içinde ise ve master header ise skip clearing
                        was_cleared = _clear_cell_preserve_merge(ws, ana_cell.row, col)
                        if was_cleared:
                            cleared += 1

        wb.save(output_path)
        status_label.config(text=f"✅ Tüm kayıtlar temizlendi! ({cleared} hücre)")
        log_widget.insert(tk.END, f"Tüm kayıtlar temizlendi! ({cleared} hücre)\n")
        log_widget.see(tk.END)
    except Exception as e:
        status_label.config(text=f"❌ Hata: {e}")
        messagebox.showerror("Hata", f"Bir hata oluştu:\n{e}")

def _clear_cell_preserve_merge(ws, row, col, clear_formulas=False):
    """
    Tek bir hücreyi clear eder. Eğer hücre merged-range içindeyse:
    - Eğer merged master DATA_START_ROW'dan küçükse -> header master, silme (return False)
    - Aksi halde o range'i geçici unmerge et, hedef hücreyi temizle, sonra range'i merge edip master'ı restore et.
    Döner: True eğer bir hücre temizlendiyse, False aksi halde.
    """

    cell = ws.cell(row=row, column=col)

    if row == 2:
        return False
    if not clear_formulas and isinstance(cell.value, str) and cell.value.startswith('='):
        return False
    if cell.value not in (None, ""):
        cell.value = None
        return True
    return False

def run_process(csv_path, status_label, log_widget, izmir_day_var=None):
    try:
        log_lines = []
        def custom_print(*args, **kwargs):
            msg = ' '.join(str(a) for a in args)
            log_lines.append(msg)
            def append_log():
                log_widget.config(state='normal')
                log_widget.insert(tk.END, msg + '\n')
                log_widget.see(tk.END)
                log_widget.update_idletasks()
            log_widget.after(0, append_log)
        # Koordinatörü kullanarak üç sevkiyat dosyasını oluştur
        coord = ShipmentCoordinator()
        sheet_hint = izmir_day_var.get() if izmir_day_var else None
        sheet_hint = sheet_hint if sheet_hint not in ("", "Seçim yok") else None
        status_label.config(text="⏳ Başladı: CSV okunuyor...")
        log_widget.insert(tk.END, "[INFO] İşlem başladı: CSV okunuyor ve eşleştirilecek.\n")
        log_widget.see(tk.END)
        # Aşama: Çalıştır
        try:
            log_widget.insert(tk.END, "[STEP] Tatlı eşleştirme başlıyor...\n")
            log_widget.see(tk.END)
            t_match, t_unmatch = coord.process_tatli(csv_path, output_path="sevkiyat_tatlı.xlsx", sheet_hint=sheet_hint)
            status_label.config(text=f"⏳ Tatlı tamamlandı: {t_match} yazıldı. Donuk hazırlanıyor...")
            log_widget.insert(tk.END, "[STEP] Donuk eşleştirme başlıyor...\n")
            log_widget.see(tk.END)
            d_match, d_unmatch = coord.process_donuk(csv_path, output_path="sevkiyat_donuk.xlsx", sheet_hint=sheet_hint)
            status_label.config(text=f"⏳ Donuk tamamlandı: {d_match} yazıldı. Lojistik hazırlanıyor...")
            log_widget.insert(tk.END, "[STEP] Lojistik eşleştirme başlıyor...\n")
            log_widget.see(tk.END)
            l_match, l_unmatch = coord.process_lojistik(csv_path, output_path="sevkiyat_lojistik.xlsx", sheet_hint=sheet_hint)
            summary = {
                "tatli": {"matched": t_match, "unmatched": t_unmatch, "file": "sevkiyat_tatlı.xlsx"},
                "donuk": {"matched": d_match, "unmatched": d_unmatch, "file": "sevkiyat_donuk.xlsx"},
                "lojistik": {"matched": l_match, "unmatched": l_unmatch, "file": "sevkiyat_lojistik.xlsx"},
            }
        except Exception as e:
            log_widget.insert(tk.END, f"[ERR-E1] run_all aşamasında hata: {e}\n")
            status_label.config(text="❌ Hata: [E1] Koordinatör çalıştırma başarısız")
            raise
        # Tarih hücresini sadece Tatlı dosyasında güncelle
        try:
            format_today_in_workbook(summary["tatli"]["file"])
        except Exception as e:
            log_widget.insert(tk.END, f"[WARN-W1] Tarih yazılamadı ({summary['tatli']['file']}): {e}\n")
            log_widget.see(tk.END)
        status_label.config(text=(
            "✅ İşlem tamamlandı!\n"
            f"Tatlı: {summary['tatli']['matched']}/{summary['tatli']['file']}  "
            f"Donuk: {summary['donuk']['matched']}/{summary['donuk']['file']}  "
            f"Lojistik: {summary['lojistik']['matched']}/{summary['lojistik']['file']}"
        ))
        log_widget.insert(tk.END, "[DONE] Tüm eşleştirmeler tamamlandı ve dosyalar kaydedildi.\n")
        log_widget.see(tk.END)
        messagebox.showinfo("Başarılı", "Tüm sevkiyat dosyaları oluşturuldu.")
    except Exception as e:
        status_label.config(text=f"❌ Hata: {e}")
        log_widget.insert(tk.END, f"[ERR-E0] Genel hata: {e}\n")
        log_widget.see(tk.END)
        messagebox.showerror("Hata", f"Bir hata oluştu:\n{e}")

def select_file(status_label, log_widget, izmir_day_var=None):
    file_path = filedialog.askopenfilename(filetypes=[("CSV Dosyası", "*.csv")])
    if file_path:
        status_label.config(text="İşleniyor...")
        log_widget.delete(1.0, tk.END)
        threading.Thread(target=run_process, args=(file_path, status_label, log_widget, izmir_day_var)).start()

def on_drop(event, status_label, log_widget):
    file_path = event.data.strip('{}')
    if file_path.lower().endswith('.csv'):
        status_label.config(text="İşleniyor...")
        log_widget.delete(1.0, tk.END)
        threading.Thread(target=run_process, args=(file_path, status_label, log_widget)).start()
    else:
        messagebox.showerror("Hata", "Lütfen bir CSV dosyası bırakın.")


def open_file(path: str):
    try:
        if os.path.exists(path):
            os.startfile(path)  # Windows
        else:
            messagebox.showerror("Hata", f"Dosya bulunamadı: {path}")
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def show_update_window():
    """Güncelleme penceresini gösterir"""
    update_window = tk.Toplevel()
    update_window.title("Güncelleme Kontrolü")
    update_window.geometry("500x400")
    update_window.resizable(False, False)
    
    # Ana frame
    main_frame = tk.Frame(update_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    # Başlık
    title_label = tk.Label(main_frame, text="Güncelleme Kontrolü", font=("Arial", 16, "bold"))
    title_label.pack(pady=(0, 20))
    
    # Mevcut sürüm bilgisi
    current_version_frame = tk.Frame(main_frame)
    current_version_frame.pack(fill=tk.X, pady=(0, 10))
    tk.Label(current_version_frame, text="Mevcut Sürüm:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
    tk.Label(current_version_frame, text=VERSION, font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 0))
    
    # Durum etiketi
    status_label = tk.Label(main_frame, text="Kontrol ediliyor...", fg="blue")
    status_label.pack(pady=(0, 10))
    
    # Progress bar
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(main_frame, variable=progress_var, maximum=100)
    progress_bar.pack(fill=tk.X, pady=(0, 10))
    
    # Log alanı
    log_text = scrolledtext.ScrolledText(main_frame, height=15, width=50)
    log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    
    # Butonlar
    button_frame = tk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))
    
    check_button = tk.Button(button_frame, text="Kontrol Et", width=15)
    download_button = tk.Button(button_frame, text="İndir", width=15, state=tk.DISABLED)
    install_button = tk.Button(button_frame, text="Kur", width=15, state=tk.DISABLED)
    close_button = tk.Button(button_frame, text="Kapat", width=15)
    
    check_button.pack(side=tk.LEFT, padx=(0, 5))
    download_button.pack(side=tk.LEFT, padx=5)
    install_button.pack(side=tk.LEFT, padx=5)
    close_button.pack(side=tk.RIGHT)
    
    # Değişkenler
    update_info = {"latest_version": None, "download_url": None, "download_path": None}
    
    def log_message(message):
        log_text.insert(tk.END, f"{message}\n")
        log_text.see(tk.END)
        update_window.update()
    
    def check_updates():
        status_label.config(text="Kontrol ediliyor...", fg="blue")
        log_text.delete(1.0, tk.END)
        check_button.config(state=tk.DISABLED)
        
        def check_thread():
            try:
                log_message("GitHub'a bağlanılıyor...")
                latest_version, release_url = get_latest_version()
                
                if not latest_version:
                    status_label.config(text="❌ Bağlantı hatası", fg="red")
                    log_message("❌ GitHub'a bağlanılamadı!")
                    check_button.config(state=tk.NORMAL)
                    return
                
                log_message(f"✅ En son sürüm bulundu: {latest_version}")
                
                if is_newer_version(latest_version, VERSION):
                    status_label.config(text=f"🔄 Yeni sürüm mevcut: {latest_version}", fg="orange")
                    log_message(f"🔄 Yeni sürüm mevcut!")
                    log_message(f"   Mevcut: {VERSION}")
                    log_message(f"   Yeni: {latest_version}")
                    
                    # Download URL'ini al
                    try:
                        response = requests.get(GITHUB_API_URL, timeout=10)
                        if response.status_code == 200:
                            data = response.json()
                            assets = data.get('assets', [])
                            if assets:
                                best = select_best_asset(assets)
                                if best is not None:
                                    download_url = best.get('browser_download_url')
                                    update_info["latest_version"] = latest_version
                                    update_info["download_url"] = download_url
                                    download_button.config(state=tk.NORMAL)
                                    log_message(f"✅ İndirme hazır! (Seçilen: {best.get('name')})")
                                    # Eğer otomatik indirme ayarlıysa indir butonunu tetikle
                                    if AUTO_START_DOWNLOAD:
                                        try:
                                            # Güvenli GUI çağrısı: ana iş parçacığında invoke et
                                            update_window.after(0, download_button.invoke)
                                            log_message("🔁 Otomatik indirme başlatıldı...")
                                        except Exception:
                                            log_message("❗ Otomatik indirme başlatılamadı; lütfen 'İndir' butonuna basın.")
                                else:
                                    log_message("❌ İndirme dosyası bulunamadı!")
                            else:
                                log_message("❌ İndirme dosyası bulunamadı!")
                        else:
                            log_message("❌ Release bilgileri alınamadı!")
                    except Exception as e:
                        log_message(f"❌ Hata: {e}")
                else:
                    status_label.config(text="✅ Güncel sürüm", fg="green")
                    log_message("✅ Uygulamanız güncel!")
                
                check_button.config(state=tk.NORMAL)
                
            except Exception as e:
                status_label.config(text="❌ Hata oluştu", fg="red")
                log_message(f"❌ Hata: {e}")
                check_button.config(state=tk.NORMAL)
        
        threading.Thread(target=check_thread, daemon=True).start()
    
    def download_update():
        if not update_info["download_url"]:
            return
        
        download_button.config(state=tk.DISABLED)
        status_label.config(text="İndiriliyor...", fg="blue")
        log_message("📥 Güncelleme indiriliyor...")
        
        def download_thread():
            try:
                def progress_callback(progress):
                    progress_var.set(progress)
                    update_window.update()
                
                success = download_github_update(update_info["download_url"], progress_callback)
                
                if success:
                    status_label.config(text="✅ İndirme tamamlandı", fg="green")
                    log_message("✅ İndirme tamamlandı!")
                    install_button.config(state=tk.NORMAL)
                    update_info["download_path"] = "update.zip"
                else:
                    status_label.config(text="❌ İndirme hatası", fg="red")
                    log_message("❌ İndirme başarısız!")
                
                download_button.config(state=tk.NORMAL)
                
            except Exception as e:
                status_label.config(text="❌ İndirme hatası", fg="red")
                log_message(f"❌ Hata: {e}")
                download_button.config(state=tk.NORMAL)
        
        threading.Thread(target=download_thread, daemon=True).start()
    
    def install_update_now():
        if not update_info["download_path"] or not os.path.exists(update_info["download_path"]):
            return
        
        result = messagebox.askyesno(
            "Güncelleme Kurulumu", 
            "Güncelleme kurulacak ve uygulama yeniden başlatılacak.\n\nDevam etmek istiyor musunuz?"
        )
        
        if not result:
            return
        
        install_button.config(state=tk.DISABLED)
        status_label.config(text="Kuruluyor...", fg="blue")
        log_message("🔧 Güncelleme kuruluyor...")
        
        def install_thread():
            try:
                success = install_update()
                # Eğer uygulama PyInstaller ile paketlenmiş/frozen ise,
                # güvenli bir şekilde exe'yi değiştirmek için bir batch updater kullanıyoruz.
                frozen = getattr(sys, 'frozen', False) or hasattr(sys, '_MEIPASS')

                if frozen:
                    import tempfile
                    tmpdir = tempfile.mkdtemp()
                    with zipfile.ZipFile('update.zip', 'r') as zip_ref:
                        zip_ref.extractall(tmpdir)

                    # Hedef exe adı
                    exe_name = 'tatli_siparis.exe'
                    extracted_exe = os.path.join(tmpdir, exe_name)
                    if not os.path.exists(extracted_exe):
                        print('Kurulum hatası: ZIP içinde exe bulunamadı.')
                        return False

                    # Batch script oluştur ve çalıştır: uygulama kapanmasını bekleyip exe'yi kopyalayıp başlatır
                    bat_path = os.path.join(tmpdir, 'updater.bat')
                    # Use double quotes for paths
                    current_dir = os.path.abspath('.')
                    # Build batch script that waits for the main exe to exit, copies new exe and restarts it
                    bat_contents = f"""@echo off
                        REM Wait for the main exe to exit, then copy new exe and start it
                        timeout /t 2 /nobreak >nul
                        :waitloop
                        tasklist /FI "IMAGENAME eq {exe_name}" | find /I "{exe_name}" >nul
                        if %ERRORLEVEL%==0 (
                        timeout /t 1 /nobreak >nul
                        goto waitloop
                        )
                        echo Replacing exe...
                        copy /Y "{extracted_exe}" "{os.path.join(current_dir, exe_name)}" >nul
                        start "" "{os.path.join(current_dir, exe_name)}"
                        REM cleanup
                        rmdir /S /Q "{tmpdir}"
                        del "%~f0" /Q
                        """

                    with open(bat_path, 'w', encoding='utf-8') as f:
                        f.write(bat_contents)

                    # Launch the updater batch and exit
                    try:
                        # Use start to run in separate process
                        subprocess.Popen(['cmd', '/c', 'start', '/min', bat_path], shell=False)
                    except Exception as e:
                        print(f'Updater başlatılamadı: {e}')
                        return False

                    # Temizle: ZIP'i sil (bizim kopyamız)
                    try:
                        os.remove('update.zip')
                    except:
                        pass

                    status_label.config(text="✅ Kurulum başlatıldı", fg="green")
                    log_message("✅ Kurulum işlemi başlatıldı. Uygulama yeniden başlatılacak.")
                    install_button.config(state=tk.NORMAL)
                    return

                # Non-frozen case: install_update() should have handled extraction
                status_label.config(text="✅ Kurulum tamamlandı", fg="green")
                log_message("✅ Kurulum tamamlandı.")
                install_button.config(state=tk.NORMAL)
            except Exception as e:
                status_label.config(text="❌ Kurulum hatası", fg="red")
                log_message(f"❌ Hata: {e}")
                install_button.config(state=tk.NORMAL)

        threading.Thread(target=install_thread, daemon=True).start()
    
    def close_window():
        update_window.destroy()
    
    # Event handlers
    check_button.config(command=check_updates)
    download_button.config(command=download_update)
    install_button.config(command=install_update_now)
    close_button.config(command=close_window)
    
    # Otomatik kontrol başlat
    check_updates()


def main():
    # tkinterdnd2 varsa onun root'unu kullan
    if TK_DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    root.title("Bayi Sipariş -> Sevkiyat (Tatlı / Donuk / Lojistik)")
    root.geometry("800x600")
    root.minsize(600, 400)
    import sys
    if sys.platform == "win32":
        try:
            import ctypes
            myappid = u'bhu.tatli.sevkiyat.1.2.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass
    try:
        root.iconbitmap(resource_path(ICON_PATH))
    except Exception:
        pass

    root.grid_rowconfigure(3, weight=1)
    root.grid_columnconfigure(0, weight=1)

    info = (
        "1) Şubeden gelen CSV'yi seçin veya sürükleyip bırakın.\n"
        "2) Uygulama üç dosyayı üretir: sevkiyat_tatlı.xlsx, sevkiyat_donuk.xlsx, sevkiyat_lojistik.xlsx.\n"
        "3) İzmir/Kuşadası şubeleri için gün seçimi yapabilirsiniz.\n"
        "4) Aşağıdaki kısayollardan açabilir veya temizleyebilirsiniz."
    )
    label = tk.Label(root, text=info, wraplength=700, justify="left")
    label.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))

    # İzmir/Kuşadası gün seçimi (opsiyonel)
    izmir_day_var = tk.StringVar(value="Seçim yok")
    days = [
        "Seçim yok",
        "Salı Karşıyaka",
        "Salı İzmir",
        "Cuma İzmir",
        "Cumartesi KSK",
        "Kuşadası-Aydın",
        "Kuşadası Çmert",
    ]
    day_frame = tk.Frame(root)
    day_frame.grid(row=1, column=0, pady=(6, 0), sticky="w")
    tk.Label(day_frame, text="Hedef Sayfa (İzmir/Kuşadası): ").pack(side=tk.LEFT)
    tk.OptionMenu(day_frame, izmir_day_var, *days).pack(side=tk.LEFT)

    # Butonlar için yeni bir frame, ortalanmış ve infonun hemen altında
    btn_frame = tk.Frame(root)
    btn_frame.grid(row=2, column=0, pady=(10, 5))
    btn_frame.grid_columnconfigure(0, weight=1)
    btn_frame.grid_columnconfigure(1, weight=1)
    btn_frame.grid_columnconfigure(2, weight=1)
    btn_frame.grid_columnconfigure(3, weight=1)
    
    select_btn = tk.Button(btn_frame, text="CSV Seç veya Bırak", width=18, command=lambda: select_file(status_label, log_widget, izmir_day_var))
    select_btn.grid(row=0, column=0, padx=4)
    
    clear_btn = tk.Button(btn_frame, text="Tatlı Dosyasını Temizle", width=18, command=lambda: clear_all_records(status_label, log_widget))
    clear_btn.grid(row=0, column=1, padx=4)
    
    # Yeni butonlar ekle
    def clear_all_files():
        confirm = messagebox.askyesno("Onay", "Tüm sevkiyat dosyalarını temizlemek istediğinize emin misiniz?")
        if not confirm:
            status_label.config(text="İşlem iptal edildi.")
            return
        try:
            cleared_total = 0
            for file_path in ["sevkiyat_tatlı.xlsx", "sevkiyat_donuk.xlsx", "sevkiyat_lojistik.xlsx"]:
                if os.path.exists(file_path):
                    cleared = clear_workbook_values(file_path)
                    cleared_total += cleared
            status_label.config(text=f"✅ Tüm dosyalar temizlendi! ({cleared_total} hücre)")
            log_widget.insert(tk.END, f"Tüm dosyalar temizlendi! ({cleared_total} hücre)\n")
            log_widget.see(tk.END)
        except Exception as e:
            status_label.config(text=f"❌ Hata: {e}")
            messagebox.showerror("Hata", f"Bir hata oluştu:\n{e}")
    
    clear_all_btn = tk.Button(btn_frame, text="Tüm Dosyaları Temizle", width=18, command=clear_all_files)
    clear_all_btn.grid(row=0, column=2, padx=4)
    
    def refresh_status():
        files_status = []
        for file_path in ["sevkiyat_tatlı.xlsx", "sevkiyat_donuk.xlsx", "sevkiyat_lojistik.xlsx"]:
            if os.path.exists(file_path):
                files_status.append(f"✅ {file_path}")
            else:
                files_status.append(f"❌ {file_path}")
        status_label.config(text="\n".join(files_status))
    
    refresh_btn = tk.Button(btn_frame, text="Durumu Yenile", width=18, command=refresh_status)
    refresh_btn.grid(row=0, column=3, padx=4)
    
    # Güncelleme butonu ekle
    update_btn = tk.Button(btn_frame, text="🔄 Güncelleme", width=18, command=show_update_window)
    update_btn.grid(row=1, column=0, padx=4, pady=(5, 0))
    # Open buttons
    open_frame = tk.Frame(root)
    open_frame.grid(row=5, column=0, pady=(4, 8))
    def mk(btn_text, path):
        return tk.Button(open_frame, text=btn_text, width=22, command=lambda p=path: open_file(p))
    mk("Tatlı Dosyasını Aç", "sevkiyat_tatlı.xlsx").grid(row=0, column=0, padx=5)
    mk("Donuk Dosyasını Aç", "sevkiyat_donuk.xlsx").grid(row=0, column=1, padx=5)
    mk("Lojistik Dosyasını Aç", "sevkiyat_lojistik.xlsx").grid(row=0, column=2, padx=5)

    status_label = tk.Label(root, text="", fg="blue", anchor="w")
    status_label.grid(row=3, column=0, sticky="ew", padx=10, pady=5)

    log_widget = scrolledtext.ScrolledText(root, state='normal')
    log_widget.grid(row=4, column=0, sticky="nsew", padx=10, pady=10)
    log_widget.update_idletasks()

    # Sürükle-bırak desteği (tkinterdnd2 ile)
    if TK_DND_AVAILABLE:
        root.drop_target_register(DND_FILES)
        def drop_event_handler(e):
            # TkinterDnD bazen event.data'yı tuple olarak gönderebilir, string'e çevir
            file_path = e.data if isinstance(e.data, str) else str(e.data)
            file_path = file_path.strip('{}')
            if file_path.lower().endswith('.csv'):
                status_label.config(text="İşleniyor...")
                log_widget.delete(1.0, tk.END)
                threading.Thread(target=run_process, args=(file_path, status_label, log_widget, izmir_day_var)).start()
            else:
                messagebox.showerror("Hata", "Lütfen bir CSV dosyası bırakın.")
        root.dnd_bind('<<Drop>>', drop_event_handler)

    footer = tk.Label(root, text=f"{DEVELOPER} | {VERSION}", fg="gray")
    footer.grid(row=6, column=0, columnspan=2, sticky="ew", pady=5)

    # Başlangıçta durumu yenile
    refresh_status()
    
    # Otomatik güncelleme kontrolü (arka planda)
    def auto_check_updates():
        try:
            # Son kontrol zamanını kontrol et
            last_check_file = "last_update_check.txt"
            should_check = True
            
            if os.path.exists(last_check_file):
                try:
                    with open(last_check_file, 'r') as f:
                        last_check_time = float(f.read().strip())
                    current_time = os.path.getmtime(__file__)  # Dosya değişiklik zamanı
                    if current_time - last_check_time < UPDATE_CHECK_INTERVAL:
                        should_check = False
                except:
                    pass
            
            if should_check:
                # Arka planda kontrol et
                threading.Thread(target=lambda: check_for_updates(silent=True), daemon=True).start()
                
                # Son kontrol zamanını kaydet
                try:
                    with open(last_check_file, 'w') as f:
                        f.write(str(os.path.getmtime(__file__)))
                except:
                    pass
        except:
            pass
    
    # 2 saniye sonra otomatik kontrol başlat
    root.after(2000, auto_check_updates)

    root.mainloop()

if __name__ == "__main__":
    main()