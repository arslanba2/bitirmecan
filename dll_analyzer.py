# dll_analyzer.py - EXE'nin bağımlılıklarını analiz eder
import os
import subprocess
import sys
from pathlib import Path


def analyze_dependencies(exe_path):
    """EXE dosyasının DLL bağımlılıklarını analiz eder"""
    try:
        # dumpbin kullanılabilir mi kontrol et (Visual Studio ile gelir)
        subprocess.check_output(['where', 'dumpbin'], stderr=subprocess.STDOUT, shell=True)
        use_dumpbin = True
    except subprocess.CalledProcessError:
        use_dumpbin = False
        print("dumpbin bulunamadı. Dependency Walker'ı manuel olarak kullanmanız gerekebilir.")

    if use_dumpbin:
        try:
            # DLL bağımlılıklarını dumpbin ile analiz et
            result = subprocess.check_output(
                ['dumpbin', '/DEPENDENTS', exe_path],
                stderr=subprocess.STDOUT,
                universal_newlines=True
            )

            # Sonuçları işle
            dependencies = []
            capture = False
            for line in result.splitlines():
                if "Image has the following dependencies" in line:
                    capture = True
                    continue
                if capture and ".dll" in line.lower():
                    dll_name = line.strip()
                    dependencies.append(dll_name)
                if "Summary" in line:
                    capture = False

            return dependencies

        except subprocess.CalledProcessError as e:
            print(f"Analiz hatası: {e}")
            return []

    return []


if __name__ == "__main__":
    # EXE dosyası yolu
    exe_path = "dist/WorkerAssignment_Debug.exe"
    if not os.path.exists(exe_path):
        print(f"Hata: {exe_path} bulunamadı.")
        sys.exit(1)

    print(f"{exe_path} bağımlılıkları analiz ediliyor...")
    dependencies = analyze_dependencies(exe_path)

    if dependencies:
        print("\nTespit edilen DLL Bağımlılıkları:")
        for dll in dependencies:
            print(f"  - {dll}")
    else:
        print("\nDLL bağımlılıkları tespit edilemedi veya araç bulunamadı.")
        print("Lütfen Dependency Walker gibi bir aracı manuel olarak kullanın:")
        print("https://www.dependencywalker.com/")