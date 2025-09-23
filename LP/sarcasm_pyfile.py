import os
from pathlib import Path
from multiprocessing import Pool
import numpy as np
import pandas as pd
from tqdm import tqdm
import shutil
import datetime


from sarcasm import Structure
from sarcasm.export import Export

# --- SETTINGS ---
folder = Path(r"/Volumes/volume2/2025/D-Disease Modeling Team/C-MYBPC3/2025/00_ToAnalyse/SP/2D-staining/40x/SP25DC9-10-11_Plate_22270/w4_analysis")
n_pools = 3
pixelsize = 0.1699
# ----------------

# find all tif files in folder
tif_files = list(folder.glob("*.TIF"))

# filter: only those without existing *_scalars.xlsx
tif_files_to_process = [
    f for f in tif_files
    if not (f.parent / f"{f.stem}_scalars.xlsx").exists()
]

print(f"{len(tif_files)} tif-files found")
print(f"{len(tif_files_to_process)} need processing")

# function for analysis of single tif-file
def analyze_tif(file):
    file = Path(file)
    sarc = Structure(str(file), pixelsize=pixelsize)

    sarc.detect_sarcomeres(max_patch_size=(2048, 2048))
    sarc.full_analysis_structure(frames="all")
    sarc.remove_intermediate_tiffs()ß

    d = Export.get_structure_dict(sarc)

    # keep only scalar values
    clean = {}
    for k, v in d.items():
        try:
            arr = np.asarray(v)
            if arr.ndim == 0:
                clean[k] = arr.item()
            elif arr.ndim == 1 and arr.size == 1:
                clean[k] = arr[0].item()
        except Exception:
            continue

    df = pd.DataFrame([clean])
    out_xlsx = file.parent / f"{file.stem}_scalars.xlsx"
    df.T.to_excel(out_xlsx)
    return str(out_xlsx)


if __name__ == "__main__":
    # Run parallel processing
    if tif_files_to_process:
        results = []
        with Pool(n_pools) as pool:
            for res in tqdm(pool.imap_unordered(analyze_tif, tif_files_to_process),
                            total=len(tif_files_to_process),
                            desc="Processing files"):
                results.append(res)

        print("\n✅ Processing finished.")
        print(f"Processed {len(results)} / {len(tif_files)} total TIFs")
    else:
        print("🎉 All TIF files already processed — nothing to do.")

    # ============================
    # Combine all *_scalars.xlsx
    # ============================
    input_dir = Path(folder)
    all_dfs = []

    for file in input_dir.glob("*.*"):
        if file.suffix.lower() not in [".xls", ".xlsx"]:
            continue
        if "_platemap" in file.name:  # skip temp files
            continue
        try:
            engine = "openpyxl" if file.suffix.lower() == ".xlsx" else "xlrd"
            df = pd.read_excel(file, header=None, skiprows=1, engine=engine)
            df = df.iloc[:, :2]
            df = df.dropna(how="all")

            df_t = df.set_index(0).T
            df_t["filename"] = file.stem
            all_dfs.append(df_t)

        except Exception as e:
            print(f"⚠️ Skipped {file.name}: {e}")

    if all_dfs:
        combined = pd.concat(all_dfs, ignore_index=True)
        print(combined.head())

        out_file = input_dir / "combined_output.xlsx"
        combined.to_excel(out_file, index=False)
        print(f"✅ Saved combined dataframe to {out_file}")

        # ============================
        # Create new output folder
        # ============================
        today = datetime.datetime.today().strftime("%y%m%d")
        out_dir = input_dir.parent / f"sarcasm_analysis_{input_dir.name}_{today}"
        out_dir.mkdir(parents=True, exist_ok=True)

        # Copy combined output
        shutil.copy2(out_file, out_dir / out_file.name)

        # Copy platemap (if exists)
        platemap_files = list(input_dir.glob("*_platemap.*"))
        if platemap_files:
            for pm in platemap_files:
                shutil.copy2(pm, out_dir / pm.name)
            print(f"✅ Copied {len(platemap_files)} platemap file(s) to {out_dir}")
        else:
            print("⚠️ No platemap file found.")

        print(f"📂 All results saved to {out_dir}")
    else:
        print("No valid .xls/.xlsx files processed.")