import os
import shutil
import zipfile
from datetime import datetime

src_dir = r'd:\new project\export_tools_app'
backup_root = r'd:\new project\_backups'
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
backup_folder = os.path.join(backup_root, f'backup_kode_ori_sebelum_perbaikan_excel_{timestamp}')
zip_file = os.path.join(backup_root, f'backup_kode_ori_sebelum_perbaikan_excel_{timestamp}.zip')

os.makedirs(backup_folder, exist_ok=True)

# Copy full application tree
shutil.copytree(
    src_dir,
    os.path.join(backup_folder, 'export_tools_app'),
    dirs_exist_ok=True,
    ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '*.log')
)

# Write metadata info
info_path = os.path.join(backup_folder, 'BACKUP_INFO.txt')
with open(info_path, 'w', encoding='utf-8') as f:
    f.write("COMMIT / BACKUP NOTE: Backup: kode ori sebelum perbaikan tabel excel oleh AI\n")
    f.write(f"TIMESTAMP: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    f.write("SCOPE: Full export_tools_app project snapshot (All templates, routes, static, services, models)\n")

# Create compressed ZIP file
with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(backup_folder):
        for file in files:
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, backup_folder)
            zf.write(full_path, rel_path)

print(f"SUCCESS: Snapshot folder saved to -> {backup_folder}")
print(f"SUCCESS: ZIP Archive saved to -> {zip_file}")
