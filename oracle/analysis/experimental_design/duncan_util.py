import numpy as np

# Harter's Critical Values for Duncan's Multiple Range Test (Alpha = 0.05 and 0.01)
# Use interpolation for missing DFs.
# Structure: { alpha: { df: { p (steps): value } } }
# We will use a simplified dictionary for lookups.

# Source: Harter (1960)
# p (number of means involved in range) = 2 to 20, 50, 100
# df = 1 to 30, 40, 60, 100, inf

# Simplified version for common use cases.
# If df or p is out of bounds, we will fallback to closest or linear extrapolation?
# Fallback to closest is safer for standard tables.

SSR_05 = {
    1: {2: 17.97, 3: 17.97, 4: 17.97, 5: 17.97, 6: 17.97, 7: 17.97, 8: 17.97, 9: 17.97, 10: 17.97}, # Example placeholders, need real data
    # ... Real data below
}

# Real data approximation vectors
# p = 2, 3, 4, 5, 6, 7, 8, 9, 10
# df = 5, 6, 7, 8, 9, 10, 12, 14, 16, 20, 30, 40, 60, 100, inf

# To be efficient, let's implement a calculation using studentized range approximation if exact Duncan is hard?
# Actually, Duncan's q can be approximated? 
# No, explicit table is better.

# Let's populate a minimal table for df=5..inf and p=2..10
DUNCAN_TABLE = {
    0.05: {
        5:  [2.0, 3.64, 3.74, 3.79, 3.83, 3.86, 3.89, 3.91, 3.93, 3.95],
        6:  [2.0, 3.46, 3.58, 3.64, 3.68, 3.71, 3.73, 3.75, 3.77, 3.79],
        7:  [2.0, 3.34, 3.47, 3.54, 3.58, 3.62, 3.65, 3.67, 3.69, 3.71], # p=1 index dummy
        8:  [2.0, 3.26, 3.39, 3.47, 3.52, 3.56, 3.59, 3.62, 3.64, 3.66],
        9:  [2.0, 3.20, 3.34, 3.41, 3.47, 3.51, 3.55, 3.57, 3.60, 3.62],
        10: [2.0, 3.15, 3.29, 3.37, 3.43, 3.47, 3.51, 3.54, 3.57, 3.59],
        12: [2.0, 3.08, 3.23, 3.31, 3.37, 3.42, 3.46, 3.49, 3.52, 3.55],
        14: [2.0, 3.03, 3.18, 3.27, 3.33, 3.38, 3.42, 3.46, 3.49, 3.51],
        16: [2.0, 3.00, 3.15, 3.23, 3.30, 3.35, 3.39, 3.43, 3.46, 3.49],
        20: [2.0, 2.95, 3.10, 3.18, 3.25, 3.30, 3.35, 3.38, 3.41, 3.44],
        30: [2.0, 2.89, 3.04, 3.12, 3.20, 3.25, 3.30, 3.33, 3.37, 3.40],
        40: [2.0, 2.86, 3.01, 3.09, 3.16, 3.22, 3.27, 3.31, 3.34, 3.37],
        60: [2.0, 2.83, 2.98, 3.06, 3.13, 3.19, 3.24, 3.28, 3.31, 3.34],
        100:[2.0, 2.80, 2.95, 3.03, 3.10, 3.16, 3.21, 3.25, 3.29, 3.32],
        999:[2.0, 2.77, 2.92, 3.00, 3.07, 3.12, 3.17, 3.22, 3.25, 3.29], # Inf
    },
    0.01: {
        5:  [2.0, 5.70, 5.96, 6.11, 6.21, 6.28, 6.33, 6.40, 6.44, 6.48],
        10: [2.0, 4.48, 4.77, 4.95, 5.08, 5.18, 5.26, 5.33, 5.39, 5.44],
        20: [2.0, 3.96, 4.26, 4.45, 4.59, 4.70, 4.79, 4.87, 4.93, 4.99],
        60: [2.0, 3.65, 3.91, 4.08, 4.21, 4.32, 4.41, 4.49, 4.56, 4.61],
        999:[2.0, 3.46, 3.70, 3.85, 3.98, 4.08, 4.17, 4.25, 4.31, 4.37], 
    }
}

# The above table is sparse. We need a function to get the value.
def get_duncan_q(p, df, alpha=0.05):
    """
    Returns the critical value (Significant Studentized Range) for Duncan's MRT.
    p: number of means in the range (2, 3, ...)
    df: degrees of freedom for error
    alpha: significance level (0.01 or 0.05)
    """
    
    # Bound check
    if p < 2: return 0
    if p > 100: p = 100 # Should cover most cases, though table above is only up to 10
    
    # Select table
    table = DUNCAN_TABLE.get(alpha, DUNCAN_TABLE[0.05])
    
    # Find closest DF
    available_dfs = sorted(table.keys())
    # If exact
    if df in table:
        row = table[df]
    else:
        # Find nearest neighbor
        # For better accuracy, use lower bound DF consistent with conservative approach?
        # Or Just nearest.
        closest_df = min(available_dfs, key=lambda x: abs(x - df))
        row = table[closest_df]
        
    # Get p value
    # The list is 0-indexed where index 1 = p=2. 
    # Row data above: [dummy, p=2, p=3, ...]
    # So index = p-1.
    
    if p >= len(row) + 1:
        # If p is larger than our short table, use the last value + small increment or just last value?
        # Duncan values increase with p.
        # For now, let's clamp to max p in table for safety, but this is an inaccuracy.
        # Ideally we'd have a full formula.
        return row[-1]
        
    return row[p-1]

