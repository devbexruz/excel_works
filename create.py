from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
import shutil
import os
import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Terminaldan nom olish
    file_name = input("Saqlanadigan fayl nomini kiriting (masalan: mydata.xlsx): ").strip()
    if not file_name:
        print("❌ Fayl nomi kiritilmadi!")
        sys.exit()

    # Fayl tanlash oynasi
    file_path, _ = QFileDialog.getOpenFileName(None, "Fayl tanlang", "", "Barcha fayllar (*.*)")
    if not file_path:
        print("❌ Fayl tanlanmadi!")
        sys.exit()

    # Saqlanadigan joy (papka)
    save_folder = os.path.expanduser("~/Downloads/loaded_files")
    os.makedirs(save_folder, exist_ok=True)
    save_path = os.path.join(save_folder, file_name)

    try:
        shutil.copy(file_path, save_path)
        QMessageBox.information(None, "OK", f"✅ Fayl saqlandi:\n{save_path}")
        print(f"✅ Fayl muvaffaqiyatli saqlandi: {save_path}")
    except Exception as e:
        QMessageBox.critical(None, "Xatolik", str(e))
        print("❌ Xato:", e)
