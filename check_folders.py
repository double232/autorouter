from pathlib import Path

cases_folder = Path(r"C:\Users\zucku\OneDrive - Vernis and Bowling\Litigation Operations - Cases")

# Check client 272
print("=== Client 272 folders containing '90250' ===")
client_272 = cases_folder / "272"
if client_272.exists():
    for folder in client_272.iterdir():
        if folder.is_dir() and "90250" in folder.name:
            print(f"  {folder.name}")

# Check client 397
print("\n=== Client 397 folders containing '90250' ===")
client_397 = cases_folder / "397"
if client_397.exists():
    for folder in client_397.iterdir():
        if folder.is_dir() and "90250" in folder.name:
            print(f"  {folder.name}")
else:
    print("  Client 397 folder does not exist!")

# Check if Ricciardi folder exists anywhere
print("\n=== Searching for 'Ricciardi' folders ===")
for client_folder in cases_folder.iterdir():
    if not client_folder.is_dir():
        continue
    for matter_folder in client_folder.iterdir():
        if matter_folder.is_dir() and "ricciardi" in matter_folder.name.lower():
            print(f"  Found: {client_folder.name}\\{matter_folder.name}")
