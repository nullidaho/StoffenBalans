# ==============================================================================
# REMAS — Complete Integrated Calculation Pipeline 
# ==============================================================================
# Modules:
#   1.1   FarmManureCalculator              Manure Volume & Net/Gross N Excretion
#   1.2   NitrogenPartitioningCalculator    MUN-based UUN / FN split
#   2.1   VEMRequirementCalculator          Energy requirements (kVEM/yr)
#   2.2a  GrasslandProductionCalculator     Vellinga Grassland Yield & N-Uptake Model
#   2.2b  VEMAllocationCalculator           VEM allocation to feed sources
#   2.3   NitrogenIntakeCalculator          DS, N, CP, VRE intake
#   2.4   NitrogenExcretionCalculatorVCRE   VCRE N retention & excretion
#   2.5   NitrogenPartitioningVCRE          VCRE-based UUN / FN split
#   3.1   MineralizationCalculator          Net Mineralization (MUN & VCRE)
#   3.2   CorrectedTANCalculator            Corrected TAN (MUN & VCRE)
#   4.1   EmissionCalculator                Stable / Storage / Grazing emissions
#   4.2   LandApplicationCalculator         N limits, Manure & Fertiliser
#   4.3   ApplicationEmissionCalculator     N residue, application emissions
# ==============================================================================

import numpy as np
import pandas as pd
import sys
import traceback

pd.options.mode.chained_assignment = None  # silence SettingWithCopyWarning

def ensure_numeric(df, col, default=0.0):
    if col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace(',', '.', regex=False).str.strip()
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)
    else:
        df[col] = default
    return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1.1 — Manure Volume & Nitrogen Excretion
# ══════════════════════════════════════════════════════════════════════════════
class FarmManureCalculator:
    CONVERSION_FACTOR_SLURRY = 0.86 
    CONVERSION_FACTOR_SOLID = 0.61 

    def __init__(self, filepath, sheet_name='Main input'):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.df = None

    def load_and_clean(self):
        print(f"--- Loading data from: {self.filepath} ---")
        try:
            df = pd.read_excel(self.filepath, sheet_name=self.sheet_name)
        except Exception as e:
            raise ValueError(f"Failed to load Excel file: {e}")

        df.columns = df.columns.str.strip()

        numeric_targets = [
            'Nr_koe', 'Nr_pink', 'Nr_kalf', 'MilkYield', 'Fat%', 'Pro%', 'MilkUerum',
            'slurry_koe', 'solid_koe', 'slurry_kalf', 'solid_kalf', 'slurry_pink', 'solid_pink',
            'volume_slurry_koe', 'volume_solid_koe', 'volume_slurry_kalf', 'volume_solid_kalf',
            'volume_solid_pink', 'volume_slurry_pink', 'Kg_WholeMilk_Kalf',
            'Kg_KunstMelk_Kalf_0', 'Kg_KunstMelk_Kalf_B', 'Kg_KunstMelk_Kalf_t',
            'slurry%_koe', 'slurry%_kalf', 'slurry%_pink', 'NatureGL%', 'Ha_Grass', 'Ha_Mais', 'Ha_Crop',
            'GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'GD_Unlimited_Kalf', 'GD_Unlimited_Pink',
            'GH_Koe', 'GH_Kalf', 'GH_Pink', 'DS_Freshcut',
            'DS_GrassSilage_0', 'DS_GrassSilage_B', 'DS_GrassSilage_S', 'DS_GrassSilage_t',
            'DS_MaizeSilage_0', 'DS_MaizeSilage_B', 'DS_MaizeSilage_S', 'DS_MaizeSilage_t',
            'DS_OtherSilage1_0', 'DS_OtherSilage1_B', 'DS_OtherSilage1_S', 'DS_OtherSilage1_t',
            'DS_OtherSilage2_0', 'DS_OtherSilage2_B', 'DS_OtherSilage2_S', 'DS_OtherSilage2_t',
            'DS_OtherSilage3_0', 'DS_OtherSilage3_B', 'DS_OtherSilage3_S', 'DS_OtherSilage3_t',
            'DS_Byproducts1_0', 'DS_Byproducts1_B', 'DS_Byproducts1_S', 'DS_Byproducts1_t',
            'DS_Byproducts2_0', 'DS_Byproducts2_B', 'DS_Byproducts2_S', 'DS_Byproducts2_t',
            'DS_Byproducts3_0', 'DS_Byproducts3_B', 'DS_Byproducts3_S', 'DS_Byproducts3_t',
            'Kg_conc1_0', 'Kg_conc1_B', 'Kg_conc1_t', 'Kg_conc2_0', 'Kg_conc2_B', 'Kg_conc2_t',
            'Kg_conc3_0', 'Kg_conc3_B', 'Kg_conc3_t',
            'N_fgrass', 'N_GrassSilage', 'N_MaizeSilage', 'N_OtherSilage1', 'N_OtherSilage2', 'N_OtherSilage3',
            'N_Byproducts1', 'N_Byproducts2', 'N_Byproducts3', 'N_Concentrate1', 'N_Concentrate2', 'N_Concentrate3',
            'VEM_fgrass', 'VEM_GrassSilage', 'VEM_MaizeSilage', 'VEM_OtherSilage1', 'VEM_OtherSilage2', 'VEM_OtherSilage3',
            'VEM_Byproducts1', 'VEM_Byproducts2', 'VEM_Byproducts3', 'VEM_Concentrate1', 'VEM_Concentrate2', 'VEM_Concentrate3',
            'DS_GrassSilage_Harvest_Input', 'DS_MaizeSilage_Harvest_Input',
            'Organic_Matter', 'Cuts_per_year', 'White_Clover'
        ]

        print("--- Cleaning Data (Fixing Comma/Dot issues) ---")
        for col in numeric_targets:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            else:
                if 'slurry%' in col: df[col] = 1.0
                elif col == 'Organic_Matter': df[col] = 5.0
                elif col == 'Cuts_per_year': df[col] = 5.0
                else: df[col] = 0.0

        if 'Grassland_Renovation' not in df.columns:
            df['Grassland_Renovation'] = 'No'

        return df

    def calculate_totals(self):
        self.df = self.load_and_clean()
        df = self.df
        
        slurry_fraction_cow = df['slurry%_koe']
        solid_fraction_cow = 1.0 - df['slurry%_koe']
        slurry_fraction_calf = df['slurry%_kalf']
        solid_fraction_calf = 1.0 - df['slurry%_kalf']
        slurry_fraction_heifer = df['slurry%_pink']
        solid_fraction_heifer = 1.0 - df['slurry%_pink']
        
        df['vol_koe_slurry'] = df['Nr_koe'] * slurry_fraction_cow * df['volume_slurry_koe']
        df['vol_koe_solid'] = df['Nr_koe'] * solid_fraction_cow * df['volume_solid_koe']
        df['vol_kalf_slurry'] = df['Nr_kalf'] * slurry_fraction_calf * df['volume_slurry_kalf']
        df['vol_kalf_solid'] = df['Nr_kalf'] * solid_fraction_calf * df['volume_solid_kalf']
        df['vol_pink_slurry'] = df['Nr_pink'] * slurry_fraction_heifer * df['volume_slurry_pink']
        df['vol_pink_solid'] = df['Nr_pink'] * solid_fraction_heifer * df['volume_solid_pink']
        
        df['Total_Manure_Volume_m3'] = (df['vol_koe_slurry'] + df['vol_koe_solid'] +
                                        df['vol_kalf_slurry'] + df['vol_kalf_solid'] + 
                                        df['vol_pink_slurry'] + df['vol_pink_solid'])
        
        df['net_n_koe_slurry'] = df['Nr_koe'] * slurry_fraction_cow * df['slurry_koe']
        df['net_n_koe_solid'] = df['Nr_koe'] * solid_fraction_cow * df['solid_koe']
        df['net_n_cows'] = df['net_n_koe_slurry'] + df['net_n_koe_solid']
        df['gross_n_cows_mun'] = df['net_n_koe_slurry'] / self.CONVERSION_FACTOR_SLURRY + df['net_n_koe_solid'] / self.CONVERSION_FACTOR_SOLID
        
        df['net_n_kalf_slurry'] = df['Nr_kalf'] * slurry_fraction_calf * df['slurry_kalf']
        df['net_n_kalf_solid'] = df['Nr_kalf'] * solid_fraction_calf * df['solid_kalf']
        df['net_n_calves'] = df['net_n_kalf_slurry'] + df['net_n_kalf_solid']
        df['gross_n_calves_mun'] = df['net_n_kalf_slurry'] / self.CONVERSION_FACTOR_SLURRY + df['net_n_kalf_solid'] / self.CONVERSION_FACTOR_SOLID
        
        df['net_n_pink_slurry'] = df['Nr_pink'] * slurry_fraction_heifer * df['slurry_pink']
        df['net_n_pink_solid'] = df['Nr_pink'] * solid_fraction_heifer * df['solid_pink']
        df['net_n_heifers'] = df['net_n_pink_slurry'] + df['net_n_pink_solid']
        df['gross_n_heifers_mun'] = df['net_n_pink_slurry'] / self.CONVERSION_FACTOR_SLURRY + df['net_n_pink_solid'] / self.CONVERSION_FACTOR_SOLID
        
        df['Total_Net_Nitrogen_kg'] = df['net_n_cows'] + df['net_n_calves'] + df['net_n_heifers']
        df['Total_Nitrogen_Excretion_MUN'] = df['gross_n_cows_mun'] + df['gross_n_calves_mun'] + df['gross_n_heifers_mun']
      
        print(f"  ✓ 1.1 done — Farm1 Gross N: {df['Total_Nitrogen_Excretion_MUN'].iloc[0]:.1f} kg")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1.2 — Nitrogen Partitioning (MUN Method)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningCalculator:
    def __init__(self, df): 
        self.df = df
        
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 1.2: Nitrogen Partitioning (MUN Method)\n{'='*70}")
        df = self.df
        df = ensure_numeric(df, 'MilkUerum', 0.0)
        df['MUN_value'] = df['MilkUerum'] * (28.0 / 60.0)
        df['uun_per_cow_g_day'] = 16.7 + 13.0 + 12.03 * df['MUN_value']
        
        df['uun_total_cows_kg'] = df['uun_per_cow_g_day'] * 365.0 * df['Nr_koe'] / 1000.0
        df['fn_total_cows_kg'] = df['gross_n_cows_mun'] - df['uun_total_cows_kg']
        
        df['cow_uun_ratio'] = np.where(df['gross_n_cows_mun'] > 0, df['uun_total_cows_kg'] / df['gross_n_cows_mun'], 0.0)
        
        df['uun_total_kalf_kg'] = df['gross_n_calves_mun'] * df['cow_uun_ratio']
        df['fn_total_kalf_kg'] = df['gross_n_calves_mun'] - df['uun_total_kalf_kg']
        
        df['uun_total_pink_kg'] = df['gross_n_heifers_mun'] * df['cow_uun_ratio']
        df['fn_total_pink_kg'] = df['gross_n_heifers_mun'] - df['uun_total_pink_kg']
        
        df['Farm_Total_UUN_kg'] = df['uun_total_cows_kg'] + df['uun_total_kalf_kg'] + df['uun_total_pink_kg']
        df['Farm_Total_FN_kg'] = df['fn_total_cows_kg'] + df['fn_total_kalf_kg'] + df['fn_total_pink_kg']
        
        print(f"  ✓ 1.2 done — Farm1 Total UUN: {df['Farm_Total_UUN_kg'].iloc[0]:.1f} kg")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.1 — VEM (Energy) Requirements
# ══════════════════════════════════════════════════════════════════════════════
class VEMRequirementCalculator:
    def __init__(self, df): 
        self.df = df.copy()
        
    def _verify_breed_and_weight(self):
        if 'breed_factor' not in self.df.columns: self.df['breed_factor'] = 1.0  
        if 'avg_weight' not in self.df.columns: self.df['avg_weight'] = 650.0  
        
    def calculate_requirements(self):
        print(f"\n{'='*70}\nMODULE 2.1: Energy (VEM) Requirements\n{'='*70}")
        columns_to_ensure = [
            'MilkYield', 'Fat%', 'Pro%', 'Nr_koe', 'Nr_pink', 'Nr_kalf',
            'GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe',
            'GD_Unlimited_Kalf', 'GD_Unlimited_Pink'
        ]
        for col in columns_to_ensure: 
            self.df = ensure_numeric(self.df, col, 0.0)
            
        self._verify_breed_and_weight()
        
        lactation_days = 326.0
        dry_days = 39.0 
        fat_pct = self.df['Fat%']
        protein_pct = self.df['Pro%']
        breed_factor = self.df['breed_factor']
        
        # Calf Energy Requirements
        vem_kalf_growth = 1323.0 * breed_factor
        vem_kalf_exercise = self.df['GD_Unlimited_Kalf'] * 0.346 * breed_factor
        self.df['vem_req_kalf_per_head_yr'] = (vem_kalf_growth + vem_kalf_exercise) * 1.02
        self.df['Total_VEM_Kalf_Farm'] = self.df['vem_req_kalf_per_head_yr'] * self.df['Nr_kalf']
        
        # Heifer Energy Requirements
        vem_pink_growth = 2259.0 * breed_factor
        vem_pink_exercise = self.df['GD_Unlimited_Pink'] * 0.784 * breed_factor
        vem_pink_processing = 115.9 * breed_factor 
        self.df['vem_req_pink_per_head_yr'] = (vem_pink_growth + vem_pink_exercise + vem_pink_processing) * 1.02
        self.df['Total_VEM_Pink_Farm'] = self.df['vem_req_pink_per_head_yr'] * self.df['Nr_pink']
        
        # Dairy Cow Energy Requirements
        fpcm_yearly = (0.337 + 0.116 * fat_pct + 0.06 * protein_pct) * self.df['MilkYield'] * 365.0 
        fpcm_daily_lactating = fpcm_yearly / lactation_days
        conversion_factor_lactation = 1.0 + (fpcm_daily_lactating - 15.0) * 0.00165
        self.df['vem_cow_milk_yr'] = (442.0 * fpcm_daily_lactating * conversion_factor_lactation / 1000.0) * lactation_days
        
        metabolic_weight = np.power(self.df['avg_weight'], 0.75)
        vem_maint_lactation = 42.4 * metabolic_weight * conversion_factor_lactation * lactation_days / 1000.0
        conversion_factor_dry = 1.0 + (-15.0 * 0.00165)
        vem_maint_dry = 42.4 * metabolic_weight * conversion_factor_dry * dry_days / 1000.0
        self.df['vem_cow_maint_yr'] = vem_maint_lactation + vem_maint_dry
        
        grazing_days_combined = (self.df['GD_Limited_Koe'] * 0.419 + 
                                 self.df['GD_Combi_Koe'] * 0.419 + 
                                 self.df['GD_Unlimited_Koe'] * 0.560)
        self.df['vem_cow_exercise_yr'] = 201.0 + grazing_days_combined * (lactation_days / 365.0) * breed_factor
        self.df['vem_cow_youth_yr'] = 102.0 * breed_factor
        self.df['vem_cow_preg_yr'] = 194.0 * breed_factor
        
        vem_cow_subtotal = (self.df['vem_cow_milk_yr'] + self.df['vem_cow_maint_yr'] + 
                            self.df['vem_cow_exercise_yr'] + self.df['vem_cow_youth_yr'] + self.df['vem_cow_preg_yr'])
                            
        self.df['vem_req_cow_per_head_yr'] = vem_cow_subtotal * 1.02
        self.df['Total_VEM_Cow_Farm'] = self.df['vem_req_cow_per_head_yr'] * self.df['Nr_koe']
        
        self.df['Total_VEM_Requirement_Farm_kVEM'] = (
            self.df['Total_VEM_Cow_Farm'] + self.df['Total_VEM_Kalf_Farm'] + self.df['Total_VEM_Pink_Farm']
        )
        print(f"  ✓ 2.1 done — Farm1 VEM: {self.df['Total_VEM_Requirement_Farm_kVEM'].iloc[0]:.0f}")
        return self.df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.2a — Grassland Production Model (Vellinga & André, 1999)
# ══════════════════════════════════════════════════════════════════════════════
class GrasslandProductionCalculator:
    """
    Calculates empirical grass yields and nitrogen uptake dynamically based on soil, region, 
    and RVO regulatory limits (assuming farmers top-off their N application).
    """
    REGULATORY_TABEL_2_NORMS = {
        'grasland_beweiden': {'klei': 345, 'zand_nwc': 250, 'zand_zuid': 250, 'loss': 250, 'veen': 265},
        'grasland_maaien':   {'klei': 385, 'zand_nwc': 320, 'zand_zuid': 320, 'loss': 320, 'veen': 300}
    }

    def __init__(self, df): 
        self.df = df.copy()

    def _compute_vellinga_model(self, utilization_mode):
        soil_lower = self.df['Soil_Type'].astype(str).str.lower().str.strip()
        region_lower = self.df['Region'].astype(str).str.lower().str.strip()
        
        is_peat = np.where(soil_lower.str.contains('veen') | soil_lower.str.contains('peat'), 1.0, 0.0)
        is_sand = np.where(soil_lower.str.contains('zand') | soil_lower.str.contains('sand'), 1.0, 0.0)
        
        soil_keys = []
        for soil_str, region_str in zip(soil_lower, region_lower):
            if 'klei' in soil_str: soil_keys.append('klei')
            elif 'veen' in soil_str or 'peat' in soil_str: soil_keys.append('veen')
            elif any(term in soil_str for term in ('loss', 'loes', 'löss')): soil_keys.append('loss')
            elif 'zand' in soil_str or 'sand' in soil_str: 
                soil_keys.append('zand_zuid' if ('zuid' in region_str or 'south' in region_str) else 'zand_nwc')
            else: 
                soil_keys.append('zand_nwc')
        
        crop_management_key = 'grasland_beweiden' if utilization_mode == 'Grazing' else 'grasland_maaien'
        nitrogen_applied_top_off = np.array([float(self.REGULATORY_TABEL_2_NORMS[crop_management_key].get(sk, 250.0)) for sk in soil_keys])

        organic_matter_pct = self.df.get('Organic_Matter', pd.Series(5.0, index=self.df.index)).fillna(5.0)
        cuts_per_year = self.df.get('Cuts_per_year', pd.Series(5.0, index=self.df.index)).fillna(5.0)
        white_clover_pct = self.df.get('White_Clover', pd.Series(0.0, index=self.df.index)).fillna(0.0)
        
        renovation_status = self.df.get('Grassland_Renovation', pd.Series('No', index=self.df.index)).astype(str).str.lower().str.strip()
        is_renovated = np.where((renovation_status == 'yes') | (renovation_status == 'ja') | (renovation_status == '1'), 1.0, 0.0)
        
        cuts_deviation = cuts_per_year - 5.0
        is_grazing_flag = 1.0 if utilization_mode == 'Grazing' else 0.0
        
        # Base soil N supply potential
        alpha_0 = 192.45 * (1.0 - np.exp(-(0.984 + np.where(is_sand == 1.0, -0.735, 0.0)) * organic_matter_pct)) * np.exp(0.2691 * is_peat + 0.1335 * is_grazing_flag + np.where(is_sand == 1.0, 0.00364, 0.0) * white_clover_pct)
        # Max N uptake
        alpha_1 = 696.2 * np.exp(-0.004944 * organic_matter_pct + 0.1407 * cuts_deviation - 0.4217 * is_grazing_flag + np.where(is_peat == 1.0, 0.2324, 0.0) * is_grazing_flag + 0.2017 * is_renovated)
        # N Immobilisation
        alpha_2 = 51.42 * np.exp(0.04050 * white_clover_pct + 0.661 * is_renovated)
        
        effective_fert_n = nitrogen_applied_top_off - alpha_2 * (1.0 - np.exp(-nitrogen_applied_top_off / np.where(alpha_2 > 0, alpha_2, 1.0)))
        delta_alpha = alpha_1 - alpha_0
        
        nitrogen_mineralized_pool = np.where(alpha_1 > alpha_0, alpha_0 + delta_alpha * (1.0 - np.exp(-effective_fert_n / np.where(delta_alpha > 0, delta_alpha, 1.0))), alpha_0)
        nitrogen_uptake_final = nitrogen_mineralized_pool * np.exp(-0.0000445 * nitrogen_mineralized_pool)
        
        beta_0_efficiency = 0.018750 * np.exp(0.04143 * cuts_deviation - 0.0926 * is_peat - 0.2108 * is_sand + 0.1875 * is_grazing_flag + 0.0977 * is_renovated)
        beta_1_max_yield = 31418.0 * np.exp(0.0342 * cuts_deviation - 0.0860 * is_peat - 0.2141 * is_sand + 0.2295 * is_grazing_flag + 0.2803 * is_renovated)
        rho_shape = 0.006400 * np.exp(1.218 * cuts_deviation + 3.794 * is_peat + 2.862 * is_sand + 0.0871 * white_clover_pct)
        
        adjusted_n_uptake = np.where((nitrogen_mineralized_pool * np.exp(-0.0001496 * nitrogen_mineralized_pool)) > 0, nitrogen_mineralized_pool * np.exp(-0.0001496 * nitrogen_mineralized_pool), 1e-5)
        dry_matter_yield_final = 1.0 / ((beta_0_efficiency / adjusted_n_uptake) + (1.0 / beta_1_max_yield) * (1.0 - np.exp(-rho_shape * adjusted_n_uptake)))
        
        return nitrogen_uptake_final, dry_matter_yield_final

    def calculate_production(self):
        print(f"\n{'='*70}\nMODULE 2.2a: Grassland Production Model (Vellinga)\n{'='*70}")
        
        n_grazing, yield_grazing = self._compute_vellinga_model('Grazing')
        n_cutting, yield_cutting = self._compute_vellinga_model('Cutting')
        
        self.df['Model_Yield_Grazing_kgDM'] = yield_grazing
        self.df['Model_Yield_Cutting_kgDM'] = yield_cutting
        self.df['Model_N_Uptake_Grazing_kgN'] = n_grazing
        self.df['Model_N_Uptake_Cutting_kgN'] = n_cutting
        
        print(f"  ✓ 2.2a done — Farm1 Grazing Yield: {self.df['Model_Yield_Grazing_kgDM'].iloc[0]:.0f} kg DM/ha")
        return self.df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.2b — VEM Feed Allocation (Reads dynamic yields from 2.2a)
# ══════════════════════════════════════════════════════════════════════════════
class VEMAllocationCalculator:
    VEM_KUNSTMELK = 1500.0
    CONCENTRATE_DRY_MATTER_CONVERSION = 0.876
    LOSS_FRACTION_MILK = 0.02
    LOSS_FRACTION_KUNSTMELK = 0.02
    LOSS_FRACTION_CONCENTRATE = 0.02
    LOSS_FRACTION_SILAGE = 0.05
    LOSS_FRACTION_OTHERS = 0.03 
    
    PCT_GRASS_KALF = 0.75
    PCT_MAIZE_KALF = 0.25
    PCT_GRASS_PINK = 0.90
    PCT_MAIZE_PINK = 0.10
    
    def __init__(self, df): 
        self.df = df.copy()
    
    @staticmethod
    def _lookup_static_vem_parameters(st):
        s = str(st).lower().strip()
        if 'klei' in s:   
            return {'yield_ms':17685, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
        elif 'zand' in s: 
            return {'yield_ms':17773, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
        elif 'veen' in s: 
            return {'yield_ms':16620, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':957.0, 'vem_gs_nat':842.0, 'vem_fg_cult':937.0, 'vem_fg_nat':860.0, 'vem_maize':950.0} 
        else:
            return {'yield_ms':17300, 'yield_nat_gs':6000, 'yield_nat_fg':6000, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
                
    def _milk_vem(self):
        fat_pct, protein_pct = self.df['Fat%'], self.df['Pro%']
        gross_energy = 744.38 + 365.7 * fat_pct + 241.4 * protein_pct
        metabolizable_energy = 584.17 + 376.6 * fat_pct * 0.94 + 171.5 * protein_pct * 0.87
        metabolizability_quotient = np.where(gross_energy > 0, metabolizable_energy / gross_energy * 100.0, 0.0)
        return (0.6 * (1.0 + 0.004 * (metabolizability_quotient - 57.0)) * 0.9752 * metabolizable_energy) / 6.9
        
    def run_allocation(self):
        print(f"\n{'='*70}\nMODULE 2.2b: VEM Allocation\n{'='*70}")
        
        self.df['Kg_KunstMelk_Kalf_Net'] = self.df['Kg_KunstMelk_Kalf_0'] + self.df['Kg_KunstMelk_Kalf_B'] - self.df['Kg_KunstMelk_Kalf_t']
        self.df['Kg_KunstMelk_Kalf'] = np.where(self.df['Kg_WholeMilk_Kalf'] > 0, self.df['Kg_WholeMilk_Kalf'], self.df['Kg_KunstMelk_Kalf_Net']).clip(min=0)
        
        self.df['DS_GS_InvNet'] = self.df['DS_GrassSilage_0'] - self.df['DS_GrassSilage_t']
        self.df['DS_MS_InvNet'] = self.df['DS_MaizeSilage_0'] - self.df['DS_MaizeSilage_t']
        
        self.df['DS_Byproducts_Total'] = sum((self.df[f'DS_Byproducts{i}_0'] + self.df[f'DS_Byproducts{i}_B'] - self.df[f'DS_Byproducts{i}_S'] - self.df[f'DS_Byproducts{i}_t']) for i in [1,2,3]).clip(lower=0)
        self.df['DS_OtherSilage_Total'] = sum((self.df[f'DS_OtherSilage{i}_0'] + self.df[f'DS_OtherSilage{i}_B'] - self.df[f'DS_OtherSilage{i}_S'] - self.df[f'DS_OtherSilage{i}_t']) for i in [1,2,3]).clip(lower=0)
        self.df['DS_CutGrass_Total'] = self.df.get('DS_Freshcut', pd.Series(0, index=self.df.index)) 
        
        soil_parameters = self.df['Soil_Type'].apply(self._lookup_static_vem_parameters).tolist()
        for k in ('yield_ms', 'yield_nat_gs', 'yield_nat_fg', 'vem_gs_cult', 'vem_gs_nat', 'vem_fg_cult', 'vem_fg_nat', 'vem_maize'):
            self.df[k] = [d[k] for d in soil_parameters]
            
        # 🌟 READ YIELDS FROM MODULE 2.2a
        self.df['yield_fg'] = self.df['Model_Yield_Grazing_kgDM']
        self.df['yield_gs'] = self.df['Model_Yield_Cutting_kgDM']
            
        nature_grassland_percentage = self.df['NatureGL%'].clip(0, 100)
        self.df['vem_fg_weighted'] = ((100 - nature_grassland_percentage) * self.df['vem_fg_cult'] + nature_grassland_percentage * self.df['vem_fg_nat']) / 100
        self.df['vem_gs_weighted'] = ((100 - nature_grassland_percentage) * self.df['vem_gs_cult'] + nature_grassland_percentage * self.df['vem_gs_nat']) / 100
        self.df['yield_fg_weighted'] = ((100 - nature_grassland_percentage) * self.df['yield_fg'] + nature_grassland_percentage * self.df['yield_nat_fg']) / 100
        self.df['yield_gs_weighted'] = ((100 - nature_grassland_percentage) * self.df['yield_gs'] + nature_grassland_percentage * self.df['yield_nat_gs']) / 100
        
        self.df['vem_fg_weighted'] = np.where(self.df['VEM_fgrass'] > 0, self.df['VEM_fgrass'], self.df['vem_fg_weighted'])
        self.df['vem_gs_weighted'] = np.where(self.df['VEM_GrassSilage'] > 0, self.df['VEM_GrassSilage'], self.df['vem_gs_weighted'])
        self.df['vem_maize'] = np.where(self.df['VEM_MaizeSilage'] > 0, self.df['VEM_MaizeSilage'], self.df['vem_maize'])
        
        self.df['kVEM_Intake_Byproducts_Total'] = 0.0
        self.df['kVEM_Intake_OtherSilage_Total'] = 0.0
        self.df['kVEM_Intake_Conc_Total'] = 0.0
        
        for i in [1, 2, 3]:
            ds_bp = (self.df[f'DS_Byproducts{i}_0'] + self.df[f'DS_Byproducts{i}_B'] - self.df[f'DS_Byproducts{i}_S'] - self.df[f'DS_Byproducts{i}_t']).clip(lower=0)
            ds_os = (self.df[f'DS_OtherSilage{i}_0'] + self.df[f'DS_OtherSilage{i}_B'] - self.df[f'DS_OtherSilage{i}_S'] - self.df[f'DS_OtherSilage{i}_t']).clip(lower=0)
            kg_cc = (self.df[f'Kg_conc{i}_0'] + self.df[f'Kg_conc{i}_B'] - self.df[f'Kg_conc{i}_t']).clip(lower=0)
            
            self.df['kVEM_Intake_Byproducts_Total'] += ds_bp * (1 - self.LOSS_FRACTION_OTHERS) * self.df[f'VEM_Byproducts{i}'].replace(0, 950.0) / 1000
            self.df['kVEM_Intake_OtherSilage_Total'] += ds_os * (1 - self.LOSS_FRACTION_OTHERS) * self.df[f'VEM_OtherSilage{i}'].replace(0, 950.0) / 1000
            self.df['kVEM_Intake_Conc_Total'] += kg_cc * self.CONCENTRATE_DRY_MATTER_CONVERSION * (1 - self.LOSS_FRACTION_CONCENTRATE) * self.df[f'VEM_Concentrate{i}'].replace(0, 940.0) / 1000
            
        milk_vem_density = self._milk_vem()
        self.df['kVEM_Intake_Milk_Kalf'] = self.df['Kg_WholeMilk_Kalf'] * (1 - self.LOSS_FRACTION_MILK) * milk_vem_density / 1000
        self.df['kVEM_Intake_KunstMelk_Kalf'] = self.df['Kg_KunstMelk_Kalf'] * (1 - self.LOSS_FRACTION_KUNSTMELK) * self.VEM_KUNSTMELK / 1000
        
        ratio_grazing_kalf = (self.df['GD_Unlimited_Kalf'] / 365).clip(0, 1)
        ratio_grazing_pink = (self.df['GD_Unlimited_Pink'] / 365).clip(0, 1)
        
        self.df['kVEM_Intake_Conc_Kalf'] = self.df['Total_VEM_Kalf_Farm'] * (0.10 * ratio_grazing_kalf + 0.25 * (1 - ratio_grazing_kalf))
        self.df['kVEM_Intake_Conc_Pink'] = self.df['Total_VEM_Pink_Farm'] * (0.00 * ratio_grazing_pink + 0.05 * (1 - ratio_grazing_pink))
        
        fresh_grass_factor_kalf = ratio_grazing_kalf * (1323.0 - 101.2) + self.df['GD_Unlimited_Kalf'] * 0.346
        self.df['kVEM_Intake_FreshGrass_Kalf'] = self.df['Nr_kalf'] * fresh_grass_factor_kalf * 0.9 * self.df['breed_factor'] * 1.02 
        
        fresh_grass_factor_pink = ratio_grazing_pink * (2259.0 + 102.9) + self.df['GD_Unlimited_Pink'] * 0.784
        self.df['kVEM_Intake_FreshGrass_Pink'] = self.df['Nr_pink'] * fresh_grass_factor_pink * self.df['breed_factor'] * 1.02
        
        roughage_kalf = (self.df['Total_VEM_Kalf_Farm'] - self.df['kVEM_Intake_Milk_Kalf'] - self.df['kVEM_Intake_KunstMelk_Kalf'] - self.df['kVEM_Intake_Conc_Kalf'] - self.df['kVEM_Intake_FreshGrass_Kalf']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Kalf'] = roughage_kalf * self.PCT_GRASS_KALF
        self.df['kVEM_Intake_MaizeSilage_Kalf'] = roughage_kalf * self.PCT_MAIZE_KALF
        
        roughage_pink = (self.df['Total_VEM_Pink_Farm'] - self.df['kVEM_Intake_Conc_Pink'] - self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Pink'] = roughage_pink * self.PCT_GRASS_PINK
        self.df['kVEM_Intake_MaizeSilage_Pink'] = roughage_pink * self.PCT_MAIZE_PINK
        
        self.df['kVEM_Intake_Conc_Cow'] = (self.df['kVEM_Intake_Conc_Total'] - self.df['kVEM_Intake_Conc_Kalf'] - self.df['kVEM_Intake_Conc_Pink']).clip(lower=0)

        grazing_hours_cow = self.df['GH_Koe']
        grass_intake_rate = np.where(grazing_hours_cow > 2, 2.0 + 0.75 * (grazing_hours_cow - 2), grazing_hours_cow)
        total_grazing_days_cow = self.df['GD_Limited_Koe'] + self.df['GD_Combi_Koe'] + self.df['GD_Unlimited_Koe'] 
        
        fpcm_yearly = (0.337 + 0.116 * self.df['Fat%'] + 0.06 * self.df['Pro%']) * self.df['MilkYield'] * 365
        milk_correction_factor = 1.0 + ((fpcm_yearly - 9500 * self.df['breed_factor']) / 500) * 0.02
        lactating_fraction = (365.0 - 39.0) / 365.0
        
        cow_grazing_vem = total_grazing_days_cow * grass_intake_rate * (self.df['vem_fg_weighted'] / 1000) * lactating_fraction * milk_correction_factor * self.df['breed_factor'] * self.df['Nr_koe']
        kvem_cut_grass_total = self.df['DS_CutGrass_Total'] * (1 - self.LOSS_FRACTION_OTHERS) * (self.df['vem_fg_weighted'] / 1000)
        total_grazing_vem_pool = cow_grazing_vem + self.df['kVEM_Intake_FreshGrass_Pink'] + self.df['kVEM_Intake_FreshGrass_Kalf'] + kvem_cut_grass_total
        
        total_field_required = total_grazing_vem_pool / (self.df['vem_fg_weighted'] / 1000).replace(0, np.nan)
        hectares_fresh_grass_needed = (total_field_required / self.df['yield_fg_weighted'].replace(0, np.nan)).fillna(0)
        hectares_available_for_grass_silage = (self.df['Ha_Grass'] - hectares_fresh_grass_needed).clip(lower=0)
        
        harvest_gs_forfaitair = hectares_available_for_grass_silage * self.df['yield_gs_weighted']
        harvest_ms_forfaitair = self.df['Ha_Mais'] * self.df['yield_ms']
        
        input_source = self.df.get('Input_source_for_silage', pd.Series('Forfaitair', index=self.df.index)).astype(str).str.lower()
        is_eigen = input_source.str.contains('eigen')
        
        harvest_gs_actual = np.where(is_eigen, self.df.get('DS_GrassSilage_Harvest_Input', 0.0), harvest_gs_forfaitair)
        harvest_ms_actual = np.where(is_eigen, self.df.get('DS_MaizeSilage_Harvest_Input', 0.0), harvest_ms_forfaitair)
        
        aanleg_gs = harvest_gs_actual + self.df['DS_GrassSilage_B'] - self.df['DS_GrassSilage_S']
        aanleg_ms = harvest_ms_actual + self.df['DS_MaizeSilage_B'] - self.df['DS_MaizeSilage_S']
        
        ds_gs_avail = (aanleg_gs + self.df['DS_GS_InvNet']).clip(lower=0)
        ds_ms_avail = (aanleg_ms + self.df['DS_MS_InvNet']).clip(lower=0)
        
        vem_grass_silage = ds_gs_avail * (self.df['vem_gs_weighted'] / 1000)
        vem_maize_silage = ds_ms_avail * (self.df['vem_maize'] / 1000)
        
        total_roughage_vem_pool = total_grazing_vem_pool + vem_grass_silage + vem_maize_silage
        has_roughage_supply = total_roughage_vem_pool > 0
        
        ratio_fresh_grass = np.where(has_roughage_supply, total_grazing_vem_pool / total_roughage_vem_pool, 0)
        ratio_grass_silage = np.where(has_roughage_supply, vem_grass_silage / total_roughage_vem_pool, 0)
        ratio_maize_silage = np.where(has_roughage_supply, vem_maize_silage / total_roughage_vem_pool, 0)
        
        remaining_roughage_vem_requirement = (self.df['Total_VEM_Cow_Farm'] + self.df['Total_VEM_Kalf_Farm'] + self.df['Total_VEM_Pink_Farm']
               - self.df['kVEM_Intake_Conc_Total']
               - self.df['kVEM_Intake_Milk_Kalf'] - self.df['kVEM_Intake_KunstMelk_Kalf']
               - self.df['kVEM_Intake_Byproducts_Total'] - self.df['kVEM_Intake_OtherSilage_Total']).clip(lower=0)
        
        self.df['kVEM_Intake_FreshGrass_Cow'] = (remaining_roughage_vem_requirement * ratio_fresh_grass - self.df['kVEM_Intake_FreshGrass_Kalf'] - self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Cow'] = (remaining_roughage_vem_requirement * ratio_grass_silage - self.df['kVEM_Intake_GrassSilage_Kalf'] - self.df['kVEM_Intake_GrassSilage_Pink']).clip(lower=0)
        self.df['kVEM_Intake_MaizeSilage_Cow'] = (remaining_roughage_vem_requirement * ratio_maize_silage  - self.df['kVEM_Intake_MaizeSilage_Kalf'] - self.df['kVEM_Intake_MaizeSilage_Pink']).clip(lower=0)
        
        self.df['kVEM_Intake_Byproducts_Cow'] = self.df['kVEM_Intake_Byproducts_Total']
        self.df['kVEM_Intake_OtherSilage_Cow'] = self.df['kVEM_Intake_OtherSilage_Total']
        
        print(f"  ✓ 2.2b done")
        return self.df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.3 — Nitrogen (N) Intake Calculations
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenIntakeCalculator:
    CONVERSION_PLANT_N_TO_CRUDE_PROTEIN = 6.25
    CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN = 6.38
    REFERENCE_ASH_CONTENT_MAIZE = 40.0
    N_CONTENT_MILK_REPLACER = 32.52
    DRY_MATTER_WHOLE_MILK_FRACTION = 0.228
    DRY_MATTER_MILK_REPLACER_FRACTION = 0.964
    
    def __init__(self, df): 
        self.df = df.copy()
        
    def _calculate_vre_intake(self, dry_matter, crude_protein_content, feed_type):
        safe_crude_protein = np.where(crude_protein_content > 0, crude_protein_content, 1.0)
        
        if feed_type == 'grass_fresh': 
            vre_fraction = (0.931 * crude_protein_content - 43.2) / safe_crude_protein
        elif feed_type == 'grass_silage': 
            vre_fraction = (0.963 * crude_protein_content - 38.3) / safe_crude_protein
        elif feed_type == 'maize_silage': 
            vre_fraction = (0.969 * crude_protein_content + 0.04 * self.REFERENCE_ASH_CONTENT_MAIZE - 40.0) / safe_crude_protein
        elif feed_type == 'concentrate': 
            vre_fraction = (88.7 * (1.0 - np.exp(-0.012 * crude_protein_content))) / 100.0
        elif feed_type == 'dairy': 
            vre_fraction = pd.Series(0.89, index=dry_matter.index) if hasattr(dry_matter, 'index') else 0.89
        else: 
            vre_fraction = 0.0
            
        vre_fraction = np.clip(vre_fraction, 0, 1)
        return dry_matter * (vre_fraction * crude_protein_content) / 1000.0
        
    def run_nutrient_calculation(self):
        print(f"\n{'='*70}\nMODULE 2.3: Nitrogen Intake Calculations\n{'='*70}")
        
        # 🌟 READ N-UPTAKE DIRECTLY FROM MODULE 2.2a Vellinga Model 
        safe_yield_grazing = np.where(self.df['Model_Yield_Grazing_kgDM'] > 0, self.df['Model_Yield_Grazing_kgDM'], 1.0)
        safe_yield_cutting = np.where(self.df['Model_Yield_Cutting_kgDM'] > 0, self.df['Model_Yield_Cutting_kgDM'], 1.0)
        
        self.df['N_cont_fresh_soil'] = (self.df['Model_N_Uptake_Grazing_kgN'] * 1000.0) / safe_yield_grazing
        self.df['N_cont_gs_soil'] = (self.df['Model_N_Uptake_Cutting_kgN'] * 1000.0) / safe_yield_cutting
        
        # Static baselines for non-grass forage 
        self.df['N_cont_ms_soil'] = 12.0
        self.df['N_cont_nat_gs'] = 24.7
        self.df['N_cont_nat_fg'] = 30.2
        
        nature_grassland_percentage = self.df['NatureGL%'].clip(0, 100)
        self.df['N_cont_fresh_weighted'] = ((100 - nature_grassland_percentage) * self.df['N_cont_fresh_soil'] + nature_grassland_percentage * self.df['N_cont_nat_fg']) / 100
        self.df['N_cont_gs_weighted'] = ((100 - nature_grassland_percentage) * self.df['N_cont_gs_soil'] + nature_grassland_percentage * self.df['N_cont_nat_gs']) / 100
        
        # Override with measured laboratory inputs if available
        self.df['N_cont_fresh_weighted'] = np.where(self.df['N_fgrass'] > 0, self.df['N_fgrass'], self.df['N_cont_fresh_weighted'])
        self.df['N_cont_gs_weighted'] = np.where(self.df['N_GrassSilage'] > 0, self.df['N_GrassSilage'], self.df['N_cont_gs_weighted'])
        self.df['N_cont_ms_soil'] = np.where(self.df['N_MaizeSilage'] > 0, self.df['N_MaizeSilage'], self.df['N_cont_ms_soil'])

        vem_fresh_grass_weighted = self.df['vem_fg_weighted'] 
        vem_grass_silage_weighted = self.df['vem_gs_weighted']
        vem_maize_silage = self.df['vem_maize']
        
        n_content_milk_dry_matter = (self.df['Pro%'] * 10.0 / self.DRY_MATTER_WHOLE_MILK_FRACTION) / self.CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN
        
        for animal_category, fresh_col, silage_col, maize_col, conc_col in [
            ('cow', 'kVEM_Intake_FreshGrass_Cow', 'kVEM_Intake_GrassSilage_Cow', 'kVEM_Intake_MaizeSilage_Cow', 'kVEM_Intake_Conc_Cow'),
            ('pink', 'kVEM_Intake_FreshGrass_Pink', 'kVEM_Intake_GrassSilage_Pink', 'kVEM_Intake_MaizeSilage_Pink', 'kVEM_Intake_Conc_Pink'),
            ('kalf', 'kVEM_Intake_FreshGrass_Kalf', 'kVEM_Intake_GrassSilage_Kalf', 'kVEM_Intake_MaizeSilage_Kalf', 'kVEM_Intake_Conc_Kalf')]:
            
            self.df[f'DS_fresh_{animal_category}'] = (self.df[fresh_col] * 1000 / vem_fresh_grass_weighted).fillna(0)
            self.df[f'DS_gs_{animal_category}'] = (self.df[silage_col] * 1000 / vem_grass_silage_weighted).fillna(0)
            self.df[f'DS_ms_{animal_category}'] = (self.df[maize_col] * 1000 / vem_maize_silage.replace(0, np.nan)).fillna(0)
            
            concentrate_vem_average = self.df[['VEM_Concentrate1', 'VEM_Concentrate2', 'VEM_Concentrate3']].replace(0, np.nan).mean(axis=1).fillna(940.0)
            self.df[f'DS_conc_{animal_category}'] = (self.df[conc_col] * 1000 / concentrate_vem_average).fillna(0)
            
        self.df['DS_milk_kalf'] = self.df['Kg_WholeMilk_Kalf'] * self.DRY_MATTER_WHOLE_MILK_FRACTION
        self.df['DS_kunst_kalf'] = self.df['Kg_KunstMelk_Kalf'] * self.DRY_MATTER_MILK_REPLACER_FRACTION
        
        self.df['DS_Total_Byproducts_Cow'] = self.df['DS_Byproducts_Total']
        self.df['DS_Total_OtherSilage_Cow'] = self.df['DS_OtherSilage_Total']
        
        self.df['N_Byproducts_Avg'] = self.df[['N_Byproducts1', 'N_Byproducts2', 'N_Byproducts3']].replace(0, 25.0).mean(axis=1)
        self.df['N_OtherSilage_Avg'] = self.df[['N_OtherSilage1', 'N_OtherSilage2', 'N_OtherSilage3']].replace(0, 20.0).mean(axis=1)
        self.df['N_Concentrate_Avg'] = self.df[['N_Concentrate1', 'N_Concentrate2', 'N_Concentrate3']].replace(0, 27.3).mean(axis=1)
        
        for animal_category in ('cow', 'pink', 'kalf'):
            category_label = animal_category.capitalize()
            feed_matrix = [
                (f'DS_fresh_{animal_category}', 'N_cont_fresh_weighted', 'grass_fresh', False),
                (f'DS_gs_{animal_category}', 'N_cont_gs_weighted', 'grass_silage', False), 
                (f'DS_ms_{animal_category}', 'N_cont_ms_soil', 'maize_silage', False),
                (f'DS_conc_{animal_category}', 'N_Concentrate_Avg', 'concentrate', False)
            ]
                     
            if animal_category == 'kalf':
                feed_matrix.append(('DS_milk_kalf', n_content_milk_dry_matter, 'dairy', True))
                feed_matrix.append(('DS_kunst_kalf', self.N_CONTENT_MILK_REPLACER, 'dairy', True))
                
            if animal_category == 'cow':
                feed_matrix.append(('DS_Total_Byproducts_Cow', 'N_Byproducts_Avg', 'concentrate', False)) 
                feed_matrix.append(('DS_Total_OtherSilage_Cow', 'N_OtherSilage_Avg', 'grass_silage', False))
                
            total_n = pd.Series(0.0, index=self.df.index)
            total_cp = total_n.copy()
            total_vre = total_n.copy()
            
            for dry_matter_col, nitrogen_src, feed_type, is_dairy in feed_matrix:
                dry_matter = self.df[dry_matter_col]
                nitrogen_content = self.df[nitrogen_src] if isinstance(nitrogen_src, str) else nitrogen_src
                protein_factor = self.CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN if is_dairy else self.CONVERSION_PLANT_N_TO_CRUDE_PROTEIN
                
                n_intake_kg = dry_matter * nitrogen_content / 1000
                cp_intake_kg = n_intake_kg * protein_factor
                cp_content_concentration = nitrogen_content * protein_factor
                
                vre_intake_kg = self._calculate_vre_intake(dry_matter, cp_content_concentration, feed_type)
                
                total_n += n_intake_kg
                total_cp += cp_intake_kg
                total_vre += vre_intake_kg
                
            self.df[f'Total_N_Intake_{category_label}'] = total_n
            self.df[f'Total_CP_Intake_{category_label}'] = total_cp
            self.df[f'Total_VRE_Intake_{category_label}'] = total_vre
            
        for category_label in ('Cow', 'Pink', 'Kalf'):
            self.df[f'VCRE_Factor_{category_label}'] = (self.df[f'Total_VRE_Intake_{category_label}'] / self.df[f'Total_CP_Intake_{category_label}'].replace(0, np.nan)).fillna(0)
            
        total_vre_farm = sum(self.df[f'Total_VRE_Intake_{category_label}'] for category_label in ('Cow', 'Pink', 'Kalf'))
        total_cp_farm = sum(self.df[f'Total_CP_Intake_{category_label}'] for category_label in ('Cow', 'Pink', 'Kalf'))
        
        self.df['VCRE_Factor_Farm'] = (total_vre_farm / total_cp_farm.replace(0, np.nan)).fillna(0)
        self.df['Total_N_Intake_Farm'] = sum(self.df[f'Total_N_Intake_{category_label}'] for category_label in ('Cow', 'Pink', 'Kalf'))
        
        print(f"  ✓ 2.3 done — Farm1 N intake: {self.df['Total_N_Intake_Farm'].iloc[0]:.1f} kg")
        return self.df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.4 — Nitrogen (N) Excretion & Retention (VCRE Method)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenExcretionCalculatorVCRE:
    N_CONCENTRATION_CALF_BIRTH = 29.4
    N_CONCENTRATION_HEIFER_GROWTH = 24.1
    N_CONCENTRATION_REPLACEMENT_CALVING = 23.1
    N_CONCENTRATION_MATURE_COW = 22.5
    
    RATIO_WEIGHT_BIRTH_TO_MATURE = 44.0 / 650.0
    RATIO_WEIGHT_HEIFER_TO_MATURE = 320.0 / 650.0
    RATIO_WEIGHT_CALVING_TO_MATURE = 540.0 / 650.0
    
    CALVING_RATE_MULTIPLIER = 0.70
    REPLACEMENT_RATE_FRACTION = 0.27
    CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN = 6.38
    
    def __init__(self, df): 
        self.df = df.copy()
        
    def calculate_excretion(self):
        print(f"\n{'='*70}\nMODULE 2.4: Nitrogen Excretion (VCRE Method)\n{'='*70}")
        df = self.df
        cow_weight = df['avg_weight']
        
        weight_birth = cow_weight * self.RATIO_WEIGHT_BIRTH_TO_MATURE
        weight_heifer_1year = cow_weight * self.RATIO_WEIGHT_HEIFER_TO_MATURE
        weight_calving = cow_weight * self.RATIO_WEIGHT_CALVING_TO_MATURE
        
        n_birth_kg = weight_birth * self.N_CONCENTRATION_CALF_BIRTH / 1000
        n_heifer_1year_kg = weight_heifer_1year * self.N_CONCENTRATION_HEIFER_GROWTH / 1000
        n_calving_kg = weight_calving * self.N_CONCENTRATION_REPLACEMENT_CALVING / 1000
        n_mature_cow_kg = cow_weight * self.N_CONCENTRATION_MATURE_COW / 1000
        
        year_multiplier = np.where(df['MilkYield'] < 100, 365, 1)
        total_milk_yield = df['MilkYield'] * year_multiplier * df['Nr_koe']
        
        n_retention_milk = total_milk_yield * df['Pro%'] * 10 / self.CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN / 1000
        n_retention_fetus = n_birth_kg * self.CALVING_RATE_MULTIPLIER * df['Nr_koe']
        
        n_replacement_inflow = self.REPLACEMENT_RATE_FRACTION * n_calving_kg * df['Nr_koe']
        n_cull_outflow = self.REPLACEMENT_RATE_FRACTION * n_mature_cow_kg * df['Nr_koe']
        df['N_Retention_Cow_VCRE'] = n_retention_milk + n_retention_fetus + (n_cull_outflow - n_replacement_inflow)
        
        growth_target_n = n_heifer_1year_kg - n_birth_kg
        growth_constant = 0.36 * df['breed_factor']
        temp_term_growth = growth_target_n * (0.376 / 0.407)
        temp_term_maintenance = (growth_constant / 2 * 24) * (0.031 / 0.407)
        
        calf_n_retention_ratio = np.divide(temp_term_growth + temp_term_maintenance, growth_target_n, out=np.ones_like(growth_target_n.values, dtype=float), where=growth_target_n.values != 0)
        df['N_Retention_Kalf_VCRE'] = growth_target_n * df['Nr_kalf'] * calf_n_retention_ratio
        
        growth_heifer_n = (n_calving_kg - n_he