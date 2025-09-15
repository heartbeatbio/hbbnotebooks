import glob
import os
from pathlib import Path
from multiprocessing import Pool
import numpy as np
import pandas as pd

from sarcasm import Structure
from sarcasm.export import Export

# --- SETTINGS ---
folder = Path(r"W:\2025\A-Assay Development Team\B-Set Up\2025\A25B008\DS_D21\2D_replating\A25B008-2D-D27-T144h-POST_Plate_22212\w3")
n_pools = 3
pixelsize = 0.1699
# ----------------

# find all tif files in folder
tif_files = list(folder.glob("*.TIF"))
print(f"{len(tif_files)} tif-files found")

# function for analysis of single tif-file
def analyze_tif(file):
    
    file = Path(file)
    print(file)

    # initialize SarcAsM object
    sarc = Structure(str(file), pixelsize=pixelsize)

    # detect sarcomere structures
    sarc.detect_sarcomeres(max_patch_size=(2048, 2048))

    # run full structural analysis
    sarc.full_analysis_structure(frames="all")

    # remove intermediate tiff files to save storage
    sarc.remove_intermediate_tiffs()

    print(f"{file} successfully analyzed!")

    # Get full dictionary of features
    d = Export.get_structure_dict(sarc)

    # Keep only scalar entries (avoid ragged arrays)
    clean = {}
    for k, v in d.items():
        try:
            arr = np.asarray(v)
            if arr.ndim == 0:  # scalar
                clean[k] = arr.item()
            elif arr.ndim == 1 and arr.size == 1:  # single-value array
                clean[k] = arr[0].item()
            # skip longer arrays
        except Exception:
            continue

    # Make DataFrame (one row of scalars)
    df = pd.DataFrame([clean])

    # Save as transposed table (key | value)
    out_xlsx = file.parent / f"{file.stem}_scalars.xlsx"
    df.T.to_excel(out_xlsx)

    print("  → saved:", out_xlsx)
    return str(out_xlsx)


if __name__ == "__main__":
    with Pool(n_pools) as p:
        results = p.map(analyze_tif, tif_files)

    print("✅ All files processed.")
    print("Saved files:", results)