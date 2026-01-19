from pydantic import BaseModel, Field, validator
from typing import List, Optional

class BaseAnalysisParams(BaseModel):
    alpha: float = Field(0.05, gt=0, lt=1)

class ColumnParams(BaseModel):
    treat_col: Optional[str] = Field(None, max_length=100)
    resp_col: Optional[str] = Field(None, max_length=100)
    rep_col: Optional[str] = Field(None, max_length=100)

class PostHocParams(BaseModel):
    post_hoc: str = Field("lsd", max_length=50)
    mean_order: str = Field("desc", max_length=20)
    notation: str = Field("letters", max_length=50)

class CRDParams(BaseAnalysisParams, PostHocParams):
    treat_col: str = Field(..., max_length=100)
    rep_col: str = Field("", max_length=100)
    control_group: Optional[str] = Field(None, max_length=100)
    comparison_mode: Optional[str] = Field(None, max_length=100)

class FactorialCRDParams(BaseAnalysisParams, PostHocParams):
    treat_cols: str = Field(..., max_length=500)
    rep_col: Optional[str] = Field(None, max_length=100)
    resp_col: str = Field(..., max_length=100)

class ThreeFactorCRDParams(BaseAnalysisParams, PostHocParams):
    a_col: str = Field(..., max_length=100)
    b_col: str = Field(..., max_length=100)
    c_col: str = Field(..., max_length=100)
    rep_col: Optional[str] = Field(None, max_length=100)
    control_col: Optional[str] = Field(None, max_length=100)

class RegressionParams(BaseAnalysisParams):
    y_col: str = Field(..., max_length=100)
    x_cols: str = Field(..., max_length=500)
    model_type: str = Field("linear", max_length=20)
    degree: int = Field(2, ge=1, le=10)

class TTestParams(BaseAnalysisParams):
    category_col: Optional[str] = Field(None, max_length=100)
    value_col: str = Field(..., max_length=100)
    mu_0: float = Field(0.0)
    variance_option: str = Field("bartlett", max_length=20)

class PairedTTestParams(BaseAnalysisParams):
    col1: str = Field(..., max_length=100)
    col2: str = Field(..., max_length=100)
    d0: float = Field(0.0)

class LineTesterParams(BaseAnalysisParams):
    line_col: str = Field(..., max_length=100)
    tester_col: str = Field(..., max_length=100)
    rep_col: str = Field(..., max_length=100)
    trait_col: str = Field(..., max_length=100)

class PCAParams(BaseModel):
    obs_col: str = Field(..., max_length=100)
    var_cols: str = Field(..., max_length=500)

class PathParams(BaseModel):
    dependent_var: str = Field(..., max_length=100)
    independent_vars: str = Field(..., max_length=500)

class CorrelationParams(BaseModel):
    var_cols: str = Field(..., max_length=500)

class GriffingParams(BaseAnalysisParams):
    genotype_col1: str = Field(..., max_length=100)
    genotype_col2: str = Field(..., max_length=100)
    rep_col: str = Field(..., max_length=100)
    trait_col: str = Field(..., max_length=100)
    check_col: Optional[str] = Field(None, max_length=100)

class GeneticParams(BaseAnalysisParams):
    genotype_col: str = Field(..., max_length=100)
    rep_col: str = Field(..., max_length=100)
    trait_cols: str = Field(..., max_length=500)

class StabilityParams(BaseAnalysisParams):
    geno_col: str = Field(..., max_length=100)
    env_col: str = Field(..., max_length=100)
    rep_col: str = Field(..., max_length=100)
    trait_col: str = Field(..., max_length=100)
    model_type: str = Field("fixed", max_length=20)
