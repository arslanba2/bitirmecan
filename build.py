# advanced_build.py - Gelişmiş DLL Sorunları Çözümü
import PyInstaller.__main__
import os
import sys
import shutil
import site
from pathlib import Path


def find_dll_paths():
    """Python kurulumunda ve site-packages dizinlerinde DLL'leri arar"""
    dll_paths = []

    # Python kurulum dizinindeki DLL'ler
    python_dir = Path(sys.executable).parent
    dll_paths.extend(list(python_dir.glob("*.dll")))

    # TCL/TK DLL'leri (tkinter için)
    tcl_tk_dirs = list(python_dir.glob("tcl*")) + list(python_dir.glob("tk*"))
    for tk_dir in tcl_tk_dirs:
        if tk_dir.is_dir():
            dll_paths.extend(list(tk_dir.glob("*.dll")))

    # Site-packages dizinindeki DLL'ler
    site_packages = site.getsitepackages()
    for sp in site_packages:
        sp_path = Path(sp)
        # Tüm alt dizinlerdeki DLL'leri bul
        dll_paths.extend(list(sp_path.glob("**/*.dll")))

    return dll_paths


def collect_dlls():
    """Gerekli DLL'leri toplayıp dist klasörüne kopyalar"""
    dll_paths = find_dll_paths()

    # DLL'leri hedef klasöre kopyala
    if not os.path.exists('dlls'):
        os.makedirs('dlls')

    copied_dlls = []
    for dll in dll_paths:
        dll_name = dll.name
        # Sadece belirli kritik DLL'leri kopyala
        if any(critical in dll_name.lower() for critical in [
            'openpyxl', 'tcl', 'tk', '_tkinter', 'tkcalendar', 'tkcalendar',
            'python', 'pandas', 'numpy', 'sqlite', 'vcruntime', 'msvcp'
        ]):
            target_path = os.path.join('dlls', dll_name)
            shutil.copy2(dll, target_path)
            copied_dlls.append(dll_name)

    print(f"Toplam {len(copied_dlls)} DLL kopyalandı: {', '.join(copied_dlls)}")
    return copied_dlls


# DLL'leri topla
print("DLL dosyaları toplanıyor...")
collected_dlls = collect_dlls()

# DLL'ler için ilave veri parametreleri oluştur
dll_args = []
for dll in collected_dlls:
    dll_args.extend(['--add-binary', f'dlls/{dll};.'])

# Ana derleme argumentleri
main_args = [
    'Main/MainController.py',  # Ana başlangıç dosyası
    '--name=WorkerAssignment_Full',  # Çıktı dosyasının adı
    '--onefile',  # Tek bir exe dosyası oluştur
    '--windowed',  # Konsol penceresi gösterme
    '--add-data=Models;Models',  # Models klasörünü ekle
    '--add-data=Functions;Functions',  # Functions klasörünü ekle
    '--add-data=Screens;Screens',  # Screens klasörünü ekle
    '--add-data=Main;Main',  # Main klasörünü ekle
    '--hidden-import=openpyxl',  # Önemli gizli bağımlılıklar
    '--hidden-import=tkcalendar',
    '--hidden-import=tkinter',
    '--collect-all=openpyxl',  # Tüm modül içeriğini topla
    '--collect-all=tkcalendar',
    '--collect-all=tkinter',
    '--clean',  # Derleme öncesi eski dosyaları temizle
    '--noconfirm',  # Var olan klasörlerin üzerine yaz (onay sorma)
    '--log-level=INFO',  # Detaylı günlük bilgileri göster
]

# DLL parametrelerini ekle
all_args = main_args + dll_args

# Exe'yi oluştur
print("EXE oluşturuluyor...")
PyInstaller.__main__.run(all_args)

print("Exe dosyası başarıyla oluşturuldu! dist/WorkerAssignment_Full.exe dosyasını çalıştırabilirsiniz.")