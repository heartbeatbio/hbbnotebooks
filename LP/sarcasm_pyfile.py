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
folder = Path(r"/Users/lokesh.pimpale/Desktop/work_dir/sarcomere/1/1")
n_pools = 3
pixelsize = 0.1699
# Set to substring to filter (e.g. "xxx"), or None to analyze ALL
filter_substring = "_w4"   # change to None for all files
# ----------------

# Make folder for per-file Excel results
results_dir = folder.parent / f"{folder.name}_all_analysed_results"
results_dir.mkdir(parents=True, exist_ok=True)

# find all tif/tiff files (case insensitive)
tif_files = [
    f for f in folder.glob("*")
    if f.suffix.lower() in [".tif", ".tiff"]
]

# filter by substring if provided
if filter_substring:
    tif_files = [f for f in tif_files if filter_substring in f.name]

# filter: only those without existing results in results_dir
tif_files_to_process = [
    f for f in tif_files
    if not (results_dir / f"{f.stem}_scalars.xlsx").exists()
]

print(f"{len(tif_files)} tif-files matched filter")
print(f"{len(tif_files_to_process)} need processing")

# function for analysis of single tif-file
def analyze_tif(file):
    file = Path(file)
    sarc = Structure(str(file), pixelsize=pixelsize)

    sarc.detect_sarcomeres(max_patch_size=(2048, 2048))
    sarc.full_analysis_structure(frames="all")
    sarc.remove_intermediate_tiffs()

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
    temp_out = file.parent / f"{file.stem}_scalars.xlsx"
    df.T.to_excel(temp_out)

    # move to results_dir
    final_out = results_dir / temp_out.name
    shutil.move(str(temp_out), final_out)

    return str(final_out)


if __name__ == "__main__":
    # Run parallel processing
    if tif_files_to_process:
        results = []
        with Pool(n_pools) as pool:
            for res in tqdm(pool.imap_unordered(analyze_tif, tif_files_to_process),
                            total=len(tif_files_to_process),
                            desc="Processing files"):
                results.append(res)

        print("\n Processing finished.")
        print(f"Processed {len(results)} / {len(tif_files)} matched TIFs")
    else:
        print("All matched TIF files already processed.")

    # ============================
    # Combine all *_scalars.xlsx
    # ============================
    all_dfs = []
    for file in results_dir.glob("*.xlsx"):
        if "_platemap" in file.name or "combined_output" in file.name:
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

        combined_file = results_dir / "combined_output.xlsx"
        combined.to_excel(combined_file, index=False)
        print(f"✅ Saved combined dataframe to {combined_file}")

        # ============================
        # Create new output folder
        # ============================
        today = datetime.datetime.today().strftime("%y%m%d")
        out_dir = folder.parent / f"sarcasm_analysis_{folder.name}_{today}"
        out_dir.mkdir(parents=True, exist_ok=True)

        # Copy combined output
        shutil.move(str(combined_file), out_dir / combined_file.name)

        # Copy platemap (if exists in original folder)
        platemap_files = list(folder.glob("*_platemap.*"))
        if platemap_files:
            for pm in platemap_files:
                shutil.copy2(pm, out_dir / pm.name)
            print(f"✅ Copied {len(platemap_files)} platemap file(s) to {out_dir}")
        else:
            print("⚠️ No platemap file found.")

        print(f"📂 All results saved to {out_dir}")
    else:
        print("No valid .xls/.xlsx files processed.")