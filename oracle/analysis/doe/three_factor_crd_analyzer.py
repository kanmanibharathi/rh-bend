"""
Three-Factor CRD (Completely Randomized Design) Analyzer
Handles A x B x C factorial experiments in CRD
"""

import pandas as pd
import numpy as np
from scipy import stats
from itertools import combinations


class ThreeFactorCRDAnalyzer:
    """
    Analyzes three-factor factorial experiments in Completely Randomized Design.
    
    Tests:
    - Main effects: A, B, C
    - Two-way interactions: AxB, AxC, BxC
    - Three-way interaction: AxBxC
    """
    
    def __init__(self, df, a_col, b_col, c_col, resp_col, rep_col=None):
        """
        Initialize the analyzer.
        
        Parameters:
        -----------
        df : pandas.DataFrame
            The experimental data
        a_col : str
            Column name for Factor A
        b_col : str
            Column name for Factor B
        c_col : str
            Column name for Factor C
        resp_col : str
            Column name for response variable
        rep_col : str, optional
            Column name for replication
        """
        self.df = df.copy()
        self.a_col = a_col
        self.b_col = b_col
        self.c_col = c_col
        self.resp_col = resp_col
        self.rep_col = rep_col
        
        # Convert factors to strings for consistency
        self.df[a_col] = self.df[a_col].astype(str)
        self.df[b_col] = self.df[b_col].astype(str)
        self.df[c_col] = self.df[c_col].astype(str)
        
        # Get factor levels
        self.levels_a = sorted(self.df[a_col].unique())
        self.levels_b = sorted(self.df[b_col].unique())
        self.levels_c = sorted(self.df[c_col].unique())
        
        self.n_a = len(self.levels_a)
        self.n_b = len(self.levels_b)
        self.n_c = len(self.levels_c)
        
        # Calculate replications
        self.n_reps = self._calculate_replications()
        self.n_total = len(self.df)
        
    def _calculate_replications(self):
        """Calculate number of replications per treatment combination."""
        combo_counts = self.df.groupby([self.a_col, self.b_col, self.c_col]).size()
        if combo_counts.nunique() > 1:
            raise ValueError("Unequal replications detected. CRD requires equal replications.")
        return combo_counts.iloc[0]
    
    def validate(self):
        """Validate the experimental design."""
        # Check for missing values
        required_cols = [self.a_col, self.b_col, self.c_col, self.resp_col]
        for col in required_cols:
            if self.df[col].isnull().any():
                raise ValueError(f"Missing values found in column: {col}")
        
        # Check minimum requirements
        if self.n_a < 2:
            raise ValueError("Factor A must have at least 2 levels")
        if self.n_b < 2:
            raise ValueError("Factor B must have at least 2 levels")
        if self.n_c < 2:
            raise ValueError("Factor C must have at least 2 levels")
        if self.n_reps < 2:
            raise ValueError("At least 2 replications required per treatment combination")
        
        # Check degrees of freedom
        n_treatments = self.n_a * self.n_b * self.n_c
        df_error = n_treatments * (self.n_reps - 1)
        
        if df_error <= 0:
            raise ValueError(f"Insufficient degrees of freedom. Need at least {n_treatments + 1} observations.")
        
        return True
    
    def run_anova(self):
        """
        Perform three-factor ANOVA.
        
        Returns:
        --------
        dict : ANOVA table with sources, DF, SS, MS, F, and P values
        """
        # Grand mean
        grand_mean = self.df[self.resp_col].mean()
        
        # Total sum of squares
        ss_total = np.sum((self.df[self.resp_col] - grand_mean) ** 2)
        df_total = self.n_total - 1
        
        # Calculate treatment means
        means_a = self.df.groupby(self.a_col)[self.resp_col].mean()
        means_b = self.df.groupby(self.b_col)[self.resp_col].mean()
        means_c = self.df.groupby(self.c_col)[self.resp_col].mean()
        means_ab = self.df.groupby([self.a_col, self.b_col])[self.resp_col].mean()
        means_ac = self.df.groupby([self.a_col, self.c_col])[self.resp_col].mean()
        means_bc = self.df.groupby([self.b_col, self.c_col])[self.resp_col].mean()
        means_abc = self.df.groupby([self.a_col, self.b_col, self.c_col])[self.resp_col].mean()
        
        # Main effect A
        ss_a = self.n_b * self.n_c * self.n_reps * np.sum((means_a - grand_mean) ** 2)
        df_a = self.n_a - 1
        
        # Main effect B
        ss_b = self.n_a * self.n_c * self.n_reps * np.sum((means_b - grand_mean) ** 2)
        df_b = self.n_b - 1
        
        # Main effect C
        ss_c = self.n_a * self.n_b * self.n_reps * np.sum((means_c - grand_mean) ** 2)
        df_c = self.n_c - 1
        
        # Interaction AxB
        ss_ab_total = self.n_c * self.n_reps * np.sum((means_ab - grand_mean) ** 2)
        ss_ab = ss_ab_total - ss_a - ss_b
        df_ab = (self.n_a - 1) * (self.n_b - 1)
        
        # Interaction AxC
        ss_ac_total = self.n_b * self.n_reps * np.sum((means_ac - grand_mean) ** 2)
        ss_ac = ss_ac_total - ss_a - ss_c
        df_ac = (self.n_a - 1) * (self.n_c - 1)
        
        # Interaction BxC
        ss_bc_total = self.n_a * self.n_reps * np.sum((means_bc - grand_mean) ** 2)
        ss_bc = ss_bc_total - ss_b - ss_c
        df_bc = (self.n_b - 1) * (self.n_c - 1)
        
        # Interaction AxBxC
        ss_abc_total = self.n_reps * np.sum((means_abc - grand_mean) ** 2)
        ss_abc = ss_abc_total - ss_a - ss_b - ss_c - ss_ab - ss_ac - ss_bc
        df_abc = (self.n_a - 1) * (self.n_b - 1) * (self.n_c - 1)
        
        # Error (within treatments)
        ss_treatments = ss_a + ss_b + ss_c + ss_ab + ss_ac + ss_bc + ss_abc
        ss_error = ss_total - ss_treatments
        df_error = self.n_a * self.n_b * self.n_c * (self.n_reps - 1)
        
        # Mean squares
        ms_a = ss_a / df_a if df_a > 0 else 0
        ms_b = ss_b / df_b if df_b > 0 else 0
        ms_c = ss_c / df_c if df_c > 0 else 0
        ms_ab = ss_ab / df_ab if df_ab > 0 else 0
        ms_ac = ss_ac / df_ac if df_ac > 0 else 0
        ms_bc = ss_bc / df_bc if df_bc > 0 else 0
        ms_abc = ss_abc / df_abc if df_abc > 0 else 0
        ms_error = ss_error / df_error if df_error > 0 else 0
        
        # F-statistics
        f_a = ms_a / ms_error if ms_error > 0 else 0
        f_b = ms_b / ms_error if ms_error > 0 else 0
        f_c = ms_c / ms_error if ms_error > 0 else 0
        f_ab = ms_ab / ms_error if ms_error > 0 else 0
        f_ac = ms_ac / ms_error if ms_error > 0 else 0
        f_bc = ms_bc / ms_error if ms_error > 0 else 0
        f_abc = ms_abc / ms_error if ms_error > 0 else 0
        
        # P-values
        p_a = 1 - stats.f.cdf(f_a, df_a, df_error) if f_a > 0 else 1
        p_b = 1 - stats.f.cdf(f_b, df_b, df_error) if f_b > 0 else 1
        p_c = 1 - stats.f.cdf(f_c, df_c, df_error) if f_c > 0 else 1
        p_ab = 1 - stats.f.cdf(f_ab, df_ab, df_error) if f_ab > 0 else 1
        p_ac = 1 - stats.f.cdf(f_ac, df_ac, df_error) if f_ac > 0 else 1
        p_bc = 1 - stats.f.cdf(f_bc, df_bc, df_error) if f_bc > 0 else 1
        p_abc = 1 - stats.f.cdf(f_abc, df_abc, df_error) if f_abc > 0 else 1
        
        # Store MSE for later use
        self.mse = ms_error
        self.df_error = df_error
        
        # Build ANOVA table (convert numpy types to Python types for JSON serialization)
        anova_table = {
            "Factor A": {"df": int(df_a), "SS": float(ss_a), "MS": float(ms_a), "F": float(f_a), "P": float(p_a)},
            "Factor B": {"df": int(df_b), "SS": float(ss_b), "MS": float(ms_b), "F": float(f_b), "P": float(p_b)},
            "Factor C": {"df": int(df_c), "SS": float(ss_c), "MS": float(ms_c), "F": float(f_c), "P": float(p_c)},
            "Interaction AxB": {"df": int(df_ab), "SS": float(ss_ab), "MS": float(ms_ab), "F": float(f_ab), "P": float(p_ab)},
            "Interaction AxC": {"df": int(df_ac), "SS": float(ss_ac), "MS": float(ms_ac), "F": float(f_ac), "P": float(p_ac)},
            "Interaction BxC": {"df": int(df_bc), "SS": float(ss_bc), "MS": float(ms_bc), "F": float(f_bc), "P": float(p_bc)},
            "Interaction AxBxC": {"df": int(df_abc), "SS": float(ss_abc), "MS": float(ms_abc), "F": float(f_abc), "P": float(p_abc)},
            "Error": {"df": int(df_error), "SS": float(ss_error), "MS": float(ms_error), "F": None, "P": None},
            "Total": {"df": int(df_total), "SS": float(ss_total), "MS": None, "F": None, "P": None}
        }
        
        return anova_table
    
    def calculate_means_and_comparisons(self, effect, alpha=0.05, method='lsd', control=None, notation='letters'):
        """
        Calculate means and perform multiple comparisons for a specific effect.
        
        Parameters:
        -----------
        effect : str
            One of: 'Factor A', 'Factor B', 'Factor C', 'Interaction AxB', 
                    'Interaction AxC', 'Interaction BxC', 'Interaction AxBxC'
        alpha : float
            Significance level
        method : str
            Comparison method: 'lsd', 'tukey', 'dunnett'
        control : str
            Control group for Dunnett's test (format: "A1 : B1 : C1" for interactions)
        notation : str
            'letters' or 'symbols'
        """
        # Determine grouping columns
        if effect == "Factor A":
            group_cols = [self.a_col]
            n_per_mean = self.n_b * self.n_c * self.n_reps
        elif effect == "Factor B":
            group_cols = [self.b_col]
            n_per_mean = self.n_a * self.n_c * self.n_reps
        elif effect == "Factor C":
            group_cols = [self.c_col]
            n_per_mean = self.n_a * self.n_b * self.n_reps
        elif effect == "Interaction AxB":
            group_cols = [self.a_col, self.b_col]
            n_per_mean = self.n_c * self.n_reps
        elif effect == "Interaction AxC":
            group_cols = [self.a_col, self.c_col]
            n_per_mean = self.n_b * self.n_reps
        elif effect == "Interaction BxC":
            group_cols = [self.b_col, self.c_col]
            n_per_mean = self.n_a * self.n_reps
        elif effect == "Interaction AxBxC":
            group_cols = [self.a_col, self.b_col, self.c_col]
            n_per_mean = self.n_reps
        else:
            raise ValueError(f"Unknown effect: {effect}")
        
        # Calculate means and standard errors
        grouped = self.df.groupby(group_cols)[self.resp_col]
        means = grouped.mean()
        counts = grouped.count()
        
        # Standard error of mean
        se_pooled = np.sqrt(self.mse / n_per_mean)
        
        # Standard error of difference
        sed = np.sqrt(2 * self.mse / n_per_mean)
        
        # Coefficient of variation
        grand_mean = self.df[self.resp_col].mean()
        cv = (np.sqrt(self.mse) / grand_mean) * 100 if grand_mean != 0 else 0
        
        # Critical difference
        if method == 'lsd':
            t_crit = stats.t.ppf(1 - alpha/2, self.df_error)
            cd = t_crit * sed
        elif method == 'tukey':
            from scipy.stats import studentized_range
            q_crit = studentized_range.ppf(1 - alpha, len(means), self.df_error)
            cd = q_crit * se_pooled / np.sqrt(2)
        elif method == 'dunnett':
            # Simplified Dunnett's - using conservative t-value
            t_crit = stats.t.ppf(1 - alpha, self.df_error)
            cd = t_crit * sed
        else:
            cd = None
        
        # Perform pairwise comparisons
        groups = self._assign_groups(means, cd, method, control, notation)
        
        # Build results
        results = []
        for idx, (level, mean_val) in enumerate(means.items()):
            if isinstance(level, tuple):
                level_str = " : ".join(str(x) for x in level)
            else:
                level_str = str(level)
            
            results.append({
                "level": level_str,
                "mean": float(mean_val),
                "se": float(se_pooled),
                "group": groups[idx]
            })
        
        return {
            "means": results,
            "se_pooled": float(se_pooled),
            "sed": float(sed),
            "cv": float(cv),
            "cd": float(cd) if cd else None
        }
    
    def _assign_groups(self, means, cd, method, control, notation):
        """Assign grouping letters or symbols based on mean separation."""
        n = len(means)
        means_sorted = means.sort_values(ascending=False)
        indices = {val: idx for idx, val in enumerate(means_sorted.index)}
        
        if method == 'dunnett' and control:
            # Dunnett's test - compare all to control
            return self._dunnett_groups(means, means_sorted, control, cd, notation)
        else:
            # LSD or Tukey - all pairwise comparisons
            return self._pairwise_groups(means, means_sorted, cd, notation)
    
    def _pairwise_groups(self, means, means_sorted, cd, notation):
        """Assign groups based on pairwise comparisons."""
        n = len(means_sorted)
        groups = [set() for _ in range(n)]
        
        # Compare all pairs
        for i in range(n):
            for j in range(i, n):
                diff = abs(means_sorted.iloc[i] - means_sorted.iloc[j])
                if diff <= cd:
                    # Not significantly different - share a group
                    if not groups[i]:
                        groups[i].add(i)
                    if not groups[j]:
                        groups[j].add(i)
                    groups[i].add(j)
                    groups[j].add(j)
        
        # Assign letters or symbols
        if notation == 'symbols':
            symbols = ['**', '*', 'ns']
            result = ['ns'] * n
            # Highest mean gets **
            result[0] = '**'
            # Check if others are significantly different from highest
            for i in range(1, n):
                diff = abs(means_sorted.iloc[0] - means_sorted.iloc[i])
                if diff > cd:
                    result[i] = 'ns'
                else:
                    result[i] = '*'
        else:
            # Letter notation
            letters = 'abcdefghijklmnopqrstuvwxyz'
            result = [''] * n
            letter_idx = 0
            
            for i in range(n):
                if not result[i]:
                    result[i] = letters[letter_idx]
                    # Assign same letter to non-significant pairs
                    for j in range(i + 1, n):
                        diff = abs(means_sorted.iloc[i] - means_sorted.iloc[j])
                        if diff <= cd:
                            result[j] = letters[letter_idx] if not result[j] else result[j] + letters[letter_idx]
                    letter_idx += 1
        
        # Map back to original order
        final_groups = [''] * n
        for idx, (level, _) in enumerate(means.items()):
            sorted_idx = list(means_sorted.index).index(level)
            final_groups[idx] = result[sorted_idx]
        
        return final_groups
    
    def _dunnett_groups(self, means, means_sorted, control, cd, notation):
        """Assign groups for Dunnett's test (compare all to control)."""
        n = len(means)
        
        # Find control mean
        control_mean = None
        for idx, level in enumerate(means.index):
            level_str = " : ".join(str(x) for x in level) if isinstance(level, tuple) else str(level)
            if level_str == control:
                control_mean = means.iloc[idx]
                break
        
        if control_mean is None:
            # Fallback to pairwise
            return self._pairwise_groups(means, means_sorted, cd, notation)
        
        # Compare each to control
        result = []
        for idx, mean_val in enumerate(means):
            diff = abs(mean_val - control_mean)
            level_str = " : ".join(str(x) for x in means.index[idx]) if isinstance(means.index[idx], tuple) else str(means.index[idx])
            
            if level_str == control:
                result.append('Control' if notation == 'letters' else 'ns')
            elif diff <= cd:
                result.append('a' if notation == 'letters' else 'ns')
            else:
                result.append('b' if notation == 'letters' else '*')
        
        return result
