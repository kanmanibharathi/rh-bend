
import pandas as pd
import numpy as np
import sys
import os

# Add backend to path
sys.path.append(os.getcwd())

try:
    from multivariate_analysis.pca_analysis import PCAAnalyzer
except ImportError as e:
    print(f"Import Error: {e}")
    sys.exit(1)

# Dummy Data
data = {
    'Genotype': ['G1', 'G2', 'G3', 'G4', 'G5'],
    'PH': [100, 110, 105, 120, 115],
    'GY': [5.5, 6.0, 5.8, 6.5, 6.2],
    'TW': [70, 72, 71, 74, 73]
}
df = pd.DataFrame(data)

print("Running PCA Test...")
try:
    analyzer = PCAAnalyzer(df, 'Genotype', ['PH', 'GY', 'TW'])
    analyzer.validate()
    analyzer.run_pca()
    analyzer.generate_plots()
    print("PCA Success!")
    print("Eigenvalues:", analyzer.pca_res['eigenvalues'])
except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"PCA Failed: {e}")
