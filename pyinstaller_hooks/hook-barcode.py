# hook-barcode.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files, get_package_paths
import os

# Kumpulkan semua submodul dari 'barcode'
# Ini adalah upaya pertama untuk menangkap semua bagian internal
hiddenimports = collect_submodules('barcode')

# Tambahkan modul utama 'barcode' secara eksplisit
# Ini kadang membantu PyInstaller untuk melacak impor dari __init__.py
hiddenimports.append('barcode')

# Tambahkan juga modul internal yang paling mungkin mengandung definisi EAN13
# Berdasarkan pengalaman, ini bisa 'barcode.ean' (versi lama) atau 'barcode.codes' (versi baru)
# Karena Anda mendapatkan ModuleNotFoundError untuk barcode.ean, mari kita coba barcode.codes
hiddenimports.append('barcode.codes')
hiddenimports.append('barcode.ean') # Tetap sertakan, mungkin ada fallback

# Kumpulkan juga data files (seperti fonts) yang mungkin dibutuhkan oleh ImageWriter
# Ini akan mencari folder 'fonts' atau data lain di dalam paket barcode
datas = collect_data_files('barcode')

# Tambahkan secara eksplisit modul PIL (Pillow) yang digunakan oleh ImageWriter
# Ini sangat penting karena ImageWriter bergantung pada Pillow
hiddenimports += [
    'PIL.Image',
    'PIL.ImageDraw',
    'PIL.ImageFont',
    'PIL._imaging' # Modul C dasar Pillow, kadang PyInstaller melewatkannya
]

# --- Tambahan: Pastikan font default ImageWriter disertakan ---
# ImageWriter secara default menggunakan font dari barcode/fonts/DejaVuSansMono.ttf
# PyInstaller mungkin tidak selalu menyertakan file ini secara otomatis.
# Kita perlu menemukan path font ini di venv dan menambahkannya sebagai data.
try:
    package_path = get_package_paths('barcode')[0]
    barcode_fonts_path = os.path.join(package_path, 'fonts')
    if os.path.isdir(barcode_fonts_path):
        # Tambahkan seluruh folder fonts dari paket barcode ke bundle
        # Format: (lokasi_asli, lokasi_di_bundle)
        datas.append((barcode_fonts_path, 'barcode/fonts'))
        print(f"Added barcode fonts from: {barcode_fonts_path}")
    else:
        print(f"Barcode fonts directory not found at: {barcode_fonts_path}")
except Exception as e:
    print(f"Could not automatically add barcode fonts: {e}")