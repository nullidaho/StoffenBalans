# ==============================================================================
# REMAS — Complete Integrated Calculation Pipeline 
# ==============================================================================
# Modules:
#   1.1  FarmManureCalculator              Manure Volume & Net/Gross N Excretion
#   1.2  NitrogenPartitioningCalculator    MUN-based UUN / FN split
#   2.1  VEMRequirementCalculator          Energy requirements (kVEM/yr)
#   2.2  VEMAllocationCalculator           VEM allocation to feed sources
#   2.3  NitrogenIntakeCalculator          DS, N, CP, VRE intake
#   2.4  NitrogenExcretionCalculatorVCRE   VCRE N retention & excretion
#   2.5  NitrogenPartitioningVCRE          VCRE-based UUN / FN split
#   3.1    MineralizationCalculator          Net Mineralization (MUN & VCRE)
#   3.2  CorrectedTANCalculator            Corrected TAN (MUN & VCRE)
#   4.1    EmissionCalculator                Stable / Storage / Grazing emissions
#   4.2  LandApplicationCalculator         N limits, Manure & Fertiliser
#.  4.3  ApplicationEmissionCalculator      N residue, application emisisons
# ==============================================================================
#
# INPUT:  InputREMAS.xlsx  (sheet 'Main input')
# OUTPUT: Output_REMAS_Complete.xlsx
#
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
# MODULE 1.1 — Manure Volume & N Excretion
# ══════════════════════════════════════════════════════════════════════════════
class FarmManureCalculator:
    CF_slurry = 0.86 #Net to gross excretion
    CF_solid = 0.61 #Net to gross excretion
    
    def __init__(self, input_filepath, sheet_name='Main input'):
        self.input_filepath = input_filepath
        self.sheet_name = sheet_name
        self.df = None

    def __init__(self, filepath, sheet_name='Main input'):
        self.filepath = filepath
        self.sheet_name = sheet_name

    def load_and_clean(self):
        print(f"--- Loading data from: {self.filepath} ---")
        try:
            df = pd.read_excel(self.filepath, sheet_name=self.sheet_name)
        except Exception as e:
            raise ValueError(f"Failed to load Excel file: {e}")

        df.columns = df.columns.str.strip()

        # Added all new inventory and nutrient override columns to numeric_targets
        numeric_targets = [
            'Nr_koe', 'Nr_pink', 'Nr_kalf', 'MilkYield', 'Fat%', 'Pro%', 'MilkUerum',
            'slurry_koe', 'solid_koe', 'slurry_kalf', 'solid_kalf', 'slurry_pink', 'solid_pink',
            'volume_slurry_koe', 'volume_solid_koe', 'volume_slurry_kalf', 'volume_solid_kalf',
            'volume_slurry_pink', 'volume_solid_pink', 'Kg_WholeMilk_Kalf',
            'Kg_KunstMelk_Kalf_0', 'Kg_KunstMelk_Kalf_B', 'Kg_KunstMelk_Kalf_t',
            'slurry%_koe', 'slurry%_kalf', 'slurry%_pink', 'NatureGL%', 'Ha_Grass', 'Ha_Mais', 
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
            'DS_GrassSilage_Harvest_Input', 'DS_MaizeSilage_Harvest_Input'
        ]

        print("--- Cleaning Data (Fixing Comma/Dot issues) ---")
        for col in numeric_targets:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            else:
                if 'slurry%' in col: df[col] = 1.0
                else: df[col] = 0.0

        return df

    def calculate_totals(self):
        if self.df is None: raise ValueError("Call load_and_clean() first")
        df = self.df
        sk, fk = df['slurry%_koe'], 1.0 - df['slurry%_koe']
        sc, fc = df['slurry%_kalf'], 1.0 - df['slurry%_kalf']
        sp, fp = df['slurry%_pink'], 1.0 - df['slurry%_pink']
        df['vol_koe_slurry'] = df['Nr_koe']*sk*df['volume_slurry_koe']
        df['vol_koe_solid'] = df['Nr_koe']*fk*df['volume_solid_koe']
        df['vol_kalf_slurry'] = df['Nr_kalf']*sc*df['volume_slurry_kalf']
        df['vol_kalf_solid'] = df['Nr_kalf']*fc*df['volume_solid_kalf']
        df['vol_pink_slurry'] = df['Nr_pink']*sp*df['volume_slurry_pink']
        df['vol_pink_solid'] = df['Nr_pink']*fp*df['volume_solid_pink']
        df['Total_Manure_Volume_m3'] = (df['vol_koe_slurry']+df['vol_koe_solid']+
            df['vol_kalf_slurry']+df['vol_kalf_solid']+df['vol_pink_slurry']+df['vol_pink_solid'])
        df['net_n_koe_slurry'] = df['Nr_koe']*sk*df['slurry_koe']
        df['net_n_koe_solid'] = df['Nr_koe']*fk*df['solid_koe']
        df['net_n_cows'] = df['net_n_koe_slurry']+df['net_n_koe_solid']
        df['gross_n_cows_mun'] = df['net_n_koe_slurry']/self.CF_slurry+df['net_n_koe_solid']/self.CF_solid
        
        df['net_n_kalf_slurry'] = df['Nr_kalf']*sc*df['slurry_kalf']
        df['net_n_kalf_solid'] = df['Nr_kalf']*fc*df['solid_kalf']
        df['net_n_calves'] = df['net_n_kalf_slurry']+df['net_n_kalf_solid']
        df['gross_n_calves_mun'] = df['net_n_kalf_slurry']/self.CF_slurry+df['net_n_kalf_solid']/self.CF_solid
        
        df['net_n_pink_slurry'] = df['Nr_pink']*sp*df['slurry_pink']
        df['net_n_pink_solid'] = df['Nr_pink']*fp*df['solid_pink']
        df['net_n_heifers'] = df['net_n_pink_slurry']+df['net_n_pink_solid']
        df['gross_n_heifers_mun'] = df['net_n_pink_slurry']/self.CF_slurry+df['net_n_pink_solid']/self.CF_solid
        
        df['Total_Net_Nitrogen_kg'] = df['net_n_cows']+df['net_n_calves']+df['net_n_heifers']
        df['Total_Nitrogen_Excretion_MUN'] = df['gross_n_cows_mun']+df['gross_n_calves_mun']+df['gross_n_heifers_mun']
      
        print(f"  ✓ 1.1 done — Farm1 Gross N: {df['Total_Nitrogen_Excretion_MUN'].iloc[0]:.1f} kg")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1.2 — N Partitioning (MUN)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningCalculator:
    
    def __init__(self, df): self.df = df
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 1.2: N Partitioning (MUN)\n{'='*70}")
        df = self.df
        df = ensure_numeric(df, 'MilkUerum', 0.0)
        df['MUN_value'] = df['MilkUerum']*(28.0/60.0)
        df['uun_per_cow_g_day'] = 16.7+13.0+12.03*df['MUN_value']
        df['uun_total_cows_kg'] = df['uun_per_cow_g_day']*365.0*df['Nr_koe']/1000.0
        df['fn_total_cows_kg'] = df['gross_n_cows_mun']-df['uun_total_cows_kg']
        df['cow_uun_ratio'] = np.where(df['gross_n_cows_mun']>0, df['uun_total_cows_kg']/df['gross_n_cows_mun'], 0.0)
        df['uun_total_kalf_kg'] = df['gross_n_calves_mun']*df['cow_uun_ratio']
        df['fn_total_kalf_kg'] = df['gross_n_calves_mun']-df['uun_total_kalf_kg']
        df['uun_total_pink_kg'] = df['gross_n_heifers_mun']*df['cow_uun_ratio']
        df['fn_total_pink_kg'] = df['gross_n_heifers_mun']-df['uun_total_pink_kg']
        df['Farm_Total_UUN_kg'] = df['uun_total_cows_kg']+df['uun_total_kalf_kg']+df['uun_total_pink_kg']
        df['Farm_Total_FN_kg'] = df['fn_total_cows_kg']+df['fn_total_kalf_kg']+df['fn_total_pink_kg']
        print(f"  ✓ 1.2 done — Farm1 UUN: {df['Farm_Total_UUN_kg'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.1 — VEM Requirements
# ══════════════════════════════════════════════════════════════════════════════

class VEMRequirementCalculator:
    
    # 1. ADDED: __init__ method to accept the dataframe
    def __init__(self, df):
        self.df = df.copy()
        
    # 2. ADDED: _breed method to prevent the 'self._breed()' call from failing
    def _breed(self):
        if 'breed_factor' not in self.df.columns:
            self.df['breed_factor'] = 1.0  # Default to 1.0 (Holstein) if not provided
        if 'avg_weight' not in self.df.columns:
            self.df['avg_weight'] = 650.0  # Default cow weight if not provided

    def calculate_requirements(self):
        print(f"\n{'='*70}\nMODULE 2.1: VEM Requirements\n{'='*70}")
        
        columns_to_ensure = ['MilkYield','Fat%','Pro%','Nr_koe','Nr_pink','Nr_kalf',
                             'GD_Limited_Koe','GD_Combi_Koe','GD_Unlimited_Koe',
                             'GD_Unlimited_Kalf','GD_Unlimited_Pink']
        for col in columns_to_ensure:
            self.df = ensure_numeric(self.df, col, 0.0)
            
        self._breed()
        
        # Define constants for lactation and dry periods (consistent with legacy logic)
        lactation_days = 326.0
        dry_days = 39.0 
        
        # Extract frequently used variables
        fat_pct = self.df['Fat%']
        protein_pct = self.df['Pro%']
        breed_factor = self.df['breed_factor']
        
        # ---------------------------------------------------------
        # Kalf (Calves < 1 year)
        # ---------------------------------------------------------
        vem_kalf_growth = 1323.0 * breed_factor
        vem_kalf_exercise = self.df['GD_Unlimited_Kalf'] * 0.346 * breed_factor
        
        # 1.02 accounts for feed processing/losses (voederverliezen/opslag)
        self.df['vem_req_kalf_per_head_yr'] = (vem_kalf_growth + vem_kalf_exercise) * 1.02
        self.df['Total_VEM_Kalf_Farm'] = self.df['vem_req_kalf_per_head_yr'] * self.df['Nr_kalf']
        
        # ---------------------------------------------------------
        # Pink (Heifers > 1 year)
        # ---------------------------------------------------------
        vem_pink_growth = 2259.0 * breed_factor
        vem_pink_exercise = self.df['GD_Unlimited_Pink'] * 0.784 * breed_factor
        vem_pink_processing = 115.9 * breed_factor # Processing/Pregnancy maintenance
        
        self.df['vem_req_pink_per_head_yr'] = (vem_pink_growth + vem_pink_exercise + vem_pink_processing) * 1.02
        self.df['Total_VEM_Pink_Farm'] = self.df['vem_req_pink_per_head_yr'] * self.df['Nr_pink']
        
        # ---------------------------------------------------------
        # Cow (Dairy Cows)
        # ---------------------------------------------------------
        # 1. Milk Production Requirements
        # FPCM: Fat and Protein Corrected Milk
        fpcm_yearly = (0.337 + 0.116 * fat_pct + 0.06 * protein_pct) * self.df['MilkYield'] * 365.0 
        fpcm_daily_lactating = fpcm_yearly / lactation_days
        
        # Conversion factor for lactating cows based on daily yield
        conversion_factor_lactation = 1.0 + (fpcm_daily_lactating - 15.0) * 0.00165
        
        self.df['vem_cow_milk_yr'] = (442.0 * fpcm_daily_lactating * conversion_factor_lactation / 1000.0) * lactation_days
        
        # 2. Maintenance Requirements (Metabolic Weight)
        metabolic_weight = np.power(self.df['avg_weight'], 0.75)
        
        # Maintenance during lactation
        vem_maint_lactation = 42.4 * metabolic_weight * conversion_factor_lactation * lactation_days / 1000.0
        
        # Maintenance during dry period (FPCM = 0)
        conversion_factor_dry = 1.0 + (-15.0 * 0.00165)
        vem_maint_dry = 42.4 * metabolic_weight * conversion_factor_dry * dry_days / 1000.0
        
        self.df['vem_cow_maint_yr'] = vem_maint_lactation + vem_maint_dry
        
        # 3. Exercise Requirements (Grazing)
        # Calculate combined grazing days factor based on grazing system
        grazing_days_combined = (self.df['GD_Limited_Koe'] * 0.419 + 
                                 self.df['GD_Combi_Koe'] * 0.419 + 
                                 self.df['GD_Unlimited_Koe'] * 0.560)
                                 
        self.df['vem_cow_exercise_yr'] = 201.0 + grazing_days_combined * (lactation_days / 365.0) * breed_factor
        
        # 4. Growth/Youth and Pregnancy Requirements
        self.df['vem_cow_youth_yr'] = 102.0 * breed_factor
        self.df['vem_cow_preg_yr'] = 194.0 * breed_factor
        
        vem_cow_subtotal = (self.df['vem_cow_milk_yr']+self.df['vem_cow_maint_yr']+
              self.df['vem_cow_exercise_yr']+self.df['vem_cow_youth_yr']+self.df['vem_cow_preg_yr'])
        
        self.df['vem_req_cow_per_head_yr'] = vem_cow_subtotal*1.02
        
        self.df['Total_VEM_Cow_Farm'] = self.df['vem_req_cow_per_head_yr']*self.df['Nr_koe']
        self.df['Total_VEM_Requirement_Farm_kVEM'] = (
            self.df['Total_VEM_Cow_Farm']+self.df['Total_VEM_Kalf_Farm']+self.df['Total_VEM_Pink_Farm'])
        print(f"  ✓ 2.1 done — Farm1 VEM: {self.df['Total_VEM_Requirement_Farm_kVEM'].iloc[0]:.0f}")
        return self.df
    
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.2 — VEM Allocation (Restored BEX Ratio Logic + Dynamic Inventory)
# ══════════════════════════════════════════════════════════════════════════════
class VEMAllocationCalculator:
    VEM_KuntsMelk = 1500.0
    Concentrate_DS_Conversion = 0.876
    Loss_Milk = 0.02
    Loss_Kuntsmelk = 0.02
    Loss_Concentrate = 0.02
    Loss_Silage = 0.05
    Loss_Others = 0.03 
    Pct_Grass_Kalf = 0.75
    Pct_Maize_Kalf = 0.25
    Pct_Grass_Pink = 0.90
    Pct_Maize_Pink = 0.10
    
    def __init__(self, df): 
        self.df = df.copy()
        
        self.PGK = self.Pct_Grass_Kalf
        self.PMK = self.Pct_Maize_Kalf
        self.PGP = self.Pct_Grass_Pink
        self.PMP = self.Pct_Maize_Pink
    
    @staticmethod
    def _soil_p(st):
        s = str(st).lower().strip()
        if 'klei' in s:   
            return {'yield_gs':10005, 'yield_fg':8893, 'yield_ms':17685, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
        elif 'zand' in s: 
            return {'yield_gs':9360, 'yield_fg':8320, 'yield_ms':17773, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
        elif 'veen' in s: 
            return {'yield_gs':9751, 'yield_fg':8668, 'yield_ms':16620, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':957.0, 'vem_gs_nat':842.0, 'vem_fg_cult':937.0, 'vem_fg_nat':860.0, 'vem_maize':950.0} 
        else:
            return {'yield_gs':9700, 'yield_fg':9700, 'yield_ms':17300, 'yield_nat_gs':6000, 'yield_nat_fg':6000, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
                
    def _milk_vem(self):
        fat_pct, protein_pct = self.df['Fat%'], self.df['Pro%']
        GE = 744.38 + 365.7 * fat_pct + 241.4 * protein_pct
        ME = 584.17 + 376.6 * fat_pct * 0.94 + 171.5 * protein_pct * 0.87
        Q = np.where(GE > 0, ME / GE * 100.0, 0.0)
        return (0.6 * (1.0 + 0.004 * (Q - 57.0)) * 0.9752 * ME) / 6.9
        
    def run_allocation(self):
        print(f"\n{'='*70}\nMODULE 2.2: VEM Allocation (Dynamic Input Method)\n{'='*70}")
        
        # --- 1. Calculate Aggregated Inventories & Nets ---
        
        # Kunstmelk
        self.df['Kg_KunstMelk_Kalf_Net'] = self.df['Kg_KunstMelk_Kalf_0'] + self.df['Kg_KunstMelk_Kalf_B'] - self.df['Kg_KunstMelk_Kalf_t']
        self.df['Kg_KunstMelk_Kalf'] = np.where(self.df['Kg_WholeMilk_Kalf'] > 0, self.df['Kg_WholeMilk_Kalf'], self.df['Kg_KunstMelk_Kalf_Net']).clip(min=0)
        
        # Silage Inventory Net (Opening - Closing)
        self.df['DS_GS_InvNet'] = self.df['DS_GrassSilage_0'] - self.df['DS_GrassSilage_t']
        self.df['DS_MS_InvNet'] = self.df['DS_MaizeSilage_0'] - self.df['DS_MaizeSilage_t']
        
        # Sum up the 3 columns for Conc, Byprod, OtherSilage
        self.df['DS_Byproducts_Total'] = sum((self.df[f'DS_Byproducts{i}_0'] + self.df[f'DS_Byproducts{i}_B'] - self.df[f'DS_Byproducts{i}_S'] - self.df[f'DS_Byproducts{i}_t']) for i in [1,2,3]).clip(lower=0)
        self.df['DS_OtherSilage_Total'] = sum((self.df[f'DS_OtherSilage{i}_0'] + self.df[f'DS_OtherSilage{i}_B'] - self.df[f'DS_OtherSilage{i}_S'] - self.df[f'DS_OtherSilage{i}_t']) for i in [1,2,3]).clip(lower=0)
        
        self.df['DS_CutGrass_Total'] = self.df.get('DS_Freshcut', pd.Series(0, index=self.df.index)) 
        
        # --- 2. Soil logic and weighted VEM ---
        soil_parameters = self.df['Soil_Type'].apply(self._soil_p).tolist()
        
        for k in ('yield_gs','yield_fg','yield_ms','yield_nat_gs','yield_nat_fg',
                  'vem_gs_cult','vem_gs_nat','vem_fg_cult','vem_fg_nat','vem_maize'):
            self.df[k] = [d[k] for d in soil_parameters]
            
        pn = self.df['NatureGL%'].clip(0, 100)
        self.df['vem_fg_weighted'] = ((100 - pn) * self.df['vem_fg_cult'] + pn * self.df['vem_fg_nat']) / 100
        self.df['vem_gs_weighted'] = ((100 - pn) * self.df['vem_gs_cult'] + pn * self.df['vem_gs_nat']) / 100
        self.df['yield_fg_weighted'] = ((100 - pn) * self.df['yield_fg'] + pn * self.df['yield_nat_fg']) / 100
        self.df['yield_gs_weighted'] = ((100 - pn) * self.df['yield_gs'] + pn * self.df['yield_nat_gs']) / 100
        
        # Overrides if user provided measured VEM
        self.df['vem_fg_weighted'] = np.where(self.df['VEM_fgrass'] > 0, self.df['VEM_fgrass'], self.df['vem_fg_weighted'])
        self.df['vem_gs_weighted'] = np.where(self.df['VEM_GrassSilage'] > 0, self.df['VEM_GrassSilage'], self.df['vem_gs_weighted'])
        self.df['vem_maize'] = np.where(self.df['VEM_MaizeSilage'] > 0, self.df['VEM_MaizeSilage'], self.df['vem_maize'])
        
        # --- 3. Calculate Total VEM Intake for the fixed/known feeds ---
        self.df['kVEM_Intake_Byproducts_Total'] = 0.0
        self.df['kVEM_Intake_OtherSilage_Total'] = 0.0
        self.df['kVEM_Intake_Conc_Total'] = 0.0
        
        for i in [1, 2, 3]:
            vem_byproducts = self.df[f'VEM_Byproducts{i}'].replace(0, 950.0)
            vem_othersilages = self.df[f'VEM_OtherSilage{i}'].replace(0, 950.0)
            vem_concentrate = self.df[f'VEM_Concentrate{i}'].replace(0, 940.0)
            
            ds_bp = (self.df[f'DS_Byproducts{i}_0'] + self.df[f'DS_Byproducts{i}_B'] - self.df[f'DS_Byproducts{i}_S'] - self.df[f'DS_Byproducts{i}_t']).clip(lower=0)
            ds_os = (self.df[f'DS_OtherSilage{i}_0'] + self.df[f'DS_OtherSilage{i}_B'] - self.df[f'DS_OtherSilage{i}_S'] - self.df[f'DS_OtherSilage{i}_t']).clip(lower=0)
            kg_cc = (self.df[f'Kg_conc{i}_0'] + self.df[f'Kg_conc{i}_B'] - self.df[f'Kg_conc{i}_t']).clip(lower=0)
            
            self.df['kVEM_Intake_Byproducts_Total'] += ds_bp * (1 - self.Loss_Others) * vem_byproducts / 1000
            self.df['kVEM_Intake_OtherSilage_Total'] += ds_os * (1 - self.Loss_Others) * vem_othersilages / 1000
            self.df['kVEM_Intake_Conc_Total'] += kg_cc * self.Concentrate_DS_Conversion * (1 - self.Loss_Concentrate) * vem_concentrate / 1000
            
        # --- 4. Young Stock Allocation ---
        vmd = self._milk_vem()
        self.df['kVEM_Intake_Milk_Kalf'] = self.df['Kg_WholeMilk_Kalf'] * (1 - self.Loss_Milk) * vmd / 1000
        self.df['kVEM_Intake_KunstMelk_Kalf'] = self.df['Kg_KunstMelk_Kalf'] * (1 - self.Loss_Kuntsmelk) * self.VEM_KuntsMelk / 1000
        
        ratio_grazing_kalf = (self.df['GD_Unlimited_Kalf'] / 365).clip(0, 1)
        ratio_grazing_pink = (self.df['GD_Unlimited_Pink'] / 365).clip(0, 1)
        
        # Concentrate rules
        self.df['kVEM_Intake_Conc_Kalf'] = self.df['Total_VEM_Kalf_Farm'] * (0.10 * ratio_grazing_kalf + 0.25 * (1 - ratio_grazing_kalf))
        self.df['kVEM_Intake_Conc_Pink'] = self.df['Total_VEM_Pink_Farm'] * (0.00 * ratio_grazing_pink + 0.05 * (1 - ratio_grazing_pink))
        
        # Fresh grass factors
        fresh_grass_factor_kalf = ratio_grazing_kalf * (1323.0 - 101.2) + self.df['GD_Unlimited_Kalf'] * 0.346
        self.df['kVEM_Intake_FreshGrass_Kalf'] = self.df['Nr_kalf'] * fresh_grass_factor_kalf * 0.9 * self.df['breed_factor'] * 1.02 
        
        fresh_grass_factor_pink = ratio_grazing_pink * (2259.0 + 102.9) + self.df['GD_Unlimited_Pink'] * 0.784
        self.df['kVEM_Intake_FreshGrass_Pink'] = self.df['Nr_pink'] * fresh_grass_factor_pink * self.df['breed_factor'] * 1.02
        
        # Roughage gaps
        roughage_kalf = (self.df['Total_VEM_Kalf_Farm'] - self.df['kVEM_Intake_Milk_Kalf'] - self.df['kVEM_Intake_KunstMelk_Kalf'] - self.df['kVEM_Intake_Conc_Kalf'] - self.df['kVEM_Intake_FreshGrass_Kalf']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Kalf'] = roughage_kalf * self.PGK
        self.df['kVEM_Intake_MaizeSilage_Kalf'] = roughage_kalf * self.PMK
        
        roughage_pink = (self.df['Total_VEM_Pink_Farm'] - self.df['kVEM_Intake_Conc_Pink'] - self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Pink'] = roughage_pink * self.PGP
        self.df['kVEM_Intake_MaizeSilage_Pink'] = roughage_pink * self.PMP
        
        self.df['kVEM_Intake_Conc_Cow'] = (self.df['kVEM_Intake_Conc_Total'] - self.df['kVEM_Intake_Conc_Kalf'] - self.df['kVEM_Intake_Conc_Pink']).clip(lower=0)

        # --- 5. The BEX Ratio Squeeze Pool for Roughage ---
        hrs = self.df['GH_Koe']
        grass_intake_rate = np.where(hrs > 2, 2.0 + 0.75 * (hrs - 2), hrs)
        total_grazing_days = self.df['GD_Limited_Koe'] + self.df['GD_Combi_Koe'] + self.df['GD_Unlimited_Koe'] 
        
        fpcm_yearly = (0.337 + 0.116 * self.df['Fat%'] + 0.06 * self.df['Pro%']) * self.df['MilkYield'] * 365
        milk_correction_factor = 1.0 + ((fpcm_yearly - 9500 * self.df['breed_factor']) / 500) * 0.02
        lactating_fraction = (365 - 39) / 365
        
        # Total Fresh Grass Pool = Cow physical grazing + Young stock grazing + Cut Grass
        cow_grazing_vem = total_grazing_days * grass_intake_rate * (self.df['vem_fg_weighted'] / 1000) * lactating_fraction * milk_correction_factor * self.df['breed_factor'] * self.df['Nr_koe']
        kVEM_CutGrass = self.df['DS_CutGrass_Total'] * (1 - self.Loss_Others) * (self.df['vem_fg_weighted'] / 1000)
        total_grazing_vem = cow_grazing_vem + self.df['kVEM_Intake_FreshGrass_Pink'] + self.df['kVEM_Intake_FreshGrass_Kalf'] + kVEM_CutGrass
        
        # Land used for Fresh Grass
        total_field_required = total_grazing_vem / (self.df['vem_fg_weighted'] / 1000).replace(0, np.nan)
        hectares_fresh_grass_needed = (total_field_required / self.df['yield_fg_weighted'].replace(0, np.nan)).fillna(0)
        hectares_available_for_grass_silage = (self.df['Ha_Grass'] - hectares_fresh_grass_needed).clip(lower=0)
        
        # ─── Dynamic Silage Harvest & Aanleg ───
        # Forfaitair Output
        harvest_gs_forfaitair = hectares_available_for_grass_silage * self.df['yield_gs_weighted']
        harvest_ms_forfaitair = self.df['Ha_Mais'] * self.df['yield_ms']
        
        input_source = self.df.get('Input_source_for_silage', pd.Series('Forfaitair', index=self.df.index)).astype(str).str.lower()
        is_eigen = input_source.str.contains('eigen')
        
        # Actual Harvested Volume
        harvest_gs_actual = np.where(is_eigen, self.df.get('DS_GrassSilage_Harvest_Input', 0.0), harvest_gs_forfaitair)
        harvest_ms_actual = np.where(is_eigen, self.df.get('DS_MaizeSilage_Harvest_Input', 0.0), harvest_ms_forfaitair)
        
        # Aanleg (Harvest + Bought - Sold)
        aanleg_gs = harvest_gs_actual + self.df['DS_GrassSilage_B'] - self.df['DS_GrassSilage_S']
        aanleg_ms = harvest_ms_actual + self.df['DS_MaizeSilage_B'] - self.df['DS_MaizeSilage_S']
        
        # Total Silage Availability = Aanleg + (Opening - Closing Inventory)
        ds_gs_avail = (aanleg_gs + self.df['DS_GS_InvNet']).clip(lower=0)
        ds_ms_avail = (aanleg_ms + self.df['DS_MS_InvNet']).clip(lower=0)
        
        vem_grass_silage = ds_gs_avail * (self.df['vem_gs_weighted'] / 1000)
        vem_maize_silage = ds_ms_avail * (self.df['vem_maize'] / 1000)
        
        # Calculate Ratios
        total_roughage_vem_pool = total_grazing_vem + vem_grass_silage + vem_maize_silage
        has_roughage_supply = total_roughage_vem_pool > 0
        ratio_fresh_grass = np.where(has_roughage_supply, total_grazing_vem / total_roughage_vem_pool, 0)
        ratio_grass_silage = np.where(has_roughage_supply, vem_grass_silage / total_roughage_vem_pool, 0)
        ratio_maize_silage = np.where(has_roughage_supply, vem_maize_silage / total_roughage_vem_pool, 0)
        
        # --- 6. Calculate the remaining VEM gap for the farm ---
        remaining_roughage_vem_requirement = (self.df['Total_VEM_Cow_Farm'] + self.df['Total_VEM_Kalf_Farm'] + self.df['Total_VEM_Pink_Farm']
               - self.df['kVEM_Intake_Conc_Total']
               - self.df['kVEM_Intake_Milk_Kalf'] - self.df['kVEM_Intake_KunstMelk_Kalf']
               - self.df['kVEM_Intake_Byproducts_Total'] - self.df['kVEM_Intake_OtherSilage_Total']).clip(lower=0)
        
        # --- 7. Apply the Ratio Squeeze and assign residual to Cows ---
        self.df['kVEM_Intake_FreshGrass_Cow'] = (remaining_roughage_vem_requirement * ratio_fresh_grass - self.df['kVEM_Intake_FreshGrass_Kalf'] - self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Cow'] = (remaining_roughage_vem_requirement * ratio_grass_silage - self.df['kVEM_Intake_GrassSilage_Kalf'] - self.df['kVEM_Intake_GrassSilage_Pink']).clip(lower=0)
        self.df['kVEM_Intake_MaizeSilage_Cow'] = (remaining_roughage_vem_requirement * ratio_maize_silage  - self.df['kVEM_Intake_MaizeSilage_Kalf'] - self.df['kVEM_Intake_MaizeSilage_Pink']).clip(lower=0)
        
        # Record Byproducts & Other Silages explicitly to cows
        self.df['kVEM_Intake_Byproducts_Cow'] = self.df['kVEM_Intake_Byproducts_Total']
        self.df['kVEM_Intake_OtherSilage_Cow'] = self.df['kVEM_Intake_OtherSilage_Total']
        
        print(f"  ✓ 2.2 done")
        return self.df
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.3 — Nitrogen (N) Intake (Restored to Simple Dry Matter Conversion)
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
    
    def _get_soil_nitrogen_content(self, soil_type):
        """Returns standard regional reference nitrogen content parameters per soil type."""
        soil_str = str(soil_type).lower().strip()
        if 'klei' in soil_str:   
            return {
                'n_content_grass_silage': 28.2, 
                'n_content_fresh_grass': 31.6, 
                'n_content_maize_silage': 12.0, 
                'n_content_natural_grass_silage': 24.7, 
                'n_content_natural_fresh_grass': 30.2
            } 
        elif 'zand' in soil_str: 
            return {
                'n_content_grass_silage': 28.6, 
                'n_content_fresh_grass': 32.0, 
                'n_content_maize_silage': 12.0, 
                'n_content_natural_grass_silage': 25.1, 
                'n_content_natural_fresh_grass': 30.2
            }
        elif 'veen' in soil_str: 
            return {
                'n_content_grass_silage': 27.8, 
                'n_content_fresh_grass': 31.1, 
                'n_content_maize_silage': 12.0, 
                'n_content_natural_grass_silage': 24.5, 
                'n_content_natural_fresh_grass': 30.2
            }
        else:             
            return {
                'n_content_grass_silage': 28.2, 
                'n_content_fresh_grass': 31.6, 
                'n_content_maize_silage': 12.0, 
                'n_content_natural_grass_silage': 24.7, 
                'n_content_natural_fresh_grass': 30.2
            }
        
    def _calculate_vre_intake(self, dry_matter, crude_protein_content, feed_type):
        """Calculates VRE (True Digestible Protein) intake based on dry matter and crude protein."""
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
        soil_nitrogen_params = self.df['Soil_Type'].apply(self._get_soil_nitrogen_content).tolist()
        
        self.df['N_cont_fresh_soil'] = [d['n_content_fresh_grass'] for d in soil_nitrogen_params]
        self.df['N_cont_gs_soil'] = [d['n_content_grass_silage'] for d in soil_nitrogen_params]
        self.df['N_cont_ms_soil'] = [d['n_content_maize_silage'] for d in soil_nitrogen_params]
        self.df['N_cont_nat_gs'] = [d['n_content_natural_grass_silage'] for d in soil_nitrogen_params]
        self.df['N_cont_nat_fg'] = [d['n_content_natural_fresh_grass'] for d in soil_nitrogen_params]
        
        nature_grassland_percentage = self.df['NatureGL%'].clip(0, 100)
        self.df['N_cont_fresh_weighted'] = ((100 - nature_grassland_percentage) * self.df['N_cont_fresh_soil'] + nature_grassland_percentage * self.df['N_cont_nat_fg']) / 100
        self.df['N_cont_gs_weighted'] = ((100 - nature_grassland_percentage) * self.df['N_cont_gs_soil'] + nature_grassland_percentage * self.df['N_cont_nat_gs']) / 100
        
        # Apply user nitrogen content overrides if available
        self.df['N_cont_fresh_weighted'] = np.where(self.df['N_fgrass'] > 0, self.df['N_fgrass'], self.df['N_cont_fresh_weighted'])
        self.df['N_cont_gs_weighted'] = np.where(self.df['N_GrassSilage'] > 0, self.df['N_GrassSilage'], self.df['N_cont_gs_weighted'])
        self.df['N_cont_ms_soil'] = np.where(self.df['N_MaizeSilage'] > 0, self.df['N_MaizeSilage'], self.df['N_cont_ms_soil'])

        vem_fresh_grass_weighted = self.df['vem_fg_weighted'] 
        vem_grass_silage_weighted = self.df['vem_gs_weighted']
        vem_maize_silage = self.df['vem_maize']
        
        n_content_milk_dry_matter = (self.df['Pro%'] * 10.0 / self.DRY_MATTER_WHOLE_MILK_FRACTION) / self.CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN
        
        # 1. Convert Allocated VEM back to Dry Matter (DS) Units
        for animal_category, fresh_col, silage_col, maize_col, conc_col in [
            ('cow', 'kVEM_Intake_FreshGrass_Cow', 'kVEM_Intake_GrassSilage_Cow', 'kVEM_Intake_MaizeSilage_Cow', 'kVEM_Intake_Conc_Cow'),
            ('pink', 'kVEM_Intake_FreshGrass_Pink', 'kVEM_Intake_GrassSilage_Pink', 'kVEM_Intake_MaizeSilage_Pink', 'kVEM_Intake_Conc_Pink'),
            ('kalf', 'kVEM_Intake_FreshGrass_Kalf', 'kVEM_Intake_GrassSilage_Kalf', 'kVEM_Intake_MaizeSilage_Kalf', 'kVEM_Intake_Conc_Kalf')]:
            
            self.df[f'DS_fresh_{animal_category}'] = (self.df[fresh_col] * 1000 / vem_fresh_grass_weighted).fillna(0)
            self.df[f'DS_gs_{animal_category}'] = (self.df[silage_col] * 1000 / vem_grass_silage_weighted).fillna(0)
            self.df[f'DS_ms_{animal_category}'] = (self.df[maize_col] * 1000 / vem_maize_silage.replace(0, np.nan)).fillna(0)
            
            # Weighted average of the 3 concentrates for accurate reverse DS conversion
            concentrate_vem_average = self.df[['VEM_Concentrate1', 'VEM_Concentrate2', 'VEM_Concentrate3']].replace(0, np.nan).mean(axis=1).fillna(940.0)
            self.df[f'DS_conc_{animal_category}'] = (self.df[conc_col] * 1000 / concentrate_vem_average).fillna(0)
            
        self.df['DS_milk_kalf'] = self.df['Kg_WholeMilk_Kalf'] * self.DRY_MATTER_WHOLE_MILK_FRACTION
        self.df['DS_kunst_kalf'] = self.df['Kg_KunstMelk_Kalf'] * self.DRY_MATTER_MILK_REPLACER_FRACTION
        
        # 2. Add Byproducts & Other Silage back to Cow's DS Pool
        self.df['DS_Total_Byproducts_Cow'] = self.df['DS_Byproducts_Total']
        self.df['DS_Total_OtherSilage_Cow'] = self.df['DS_OtherSilage_Total']
        
        # Aggregate mean values for specialized nutrient content streams
        self.df['N_Byproducts_Avg'] = self.df[['N_Byproducts1', 'N_Byproducts2', 'N_Byproducts3']].replace(0, 25.0).mean(axis=1)
        self.df['N_OtherSilage_Avg'] = self.df[['N_OtherSilage1', 'N_OtherSilage2', 'N_OtherSilage3']].replace(0, 20.0).mean(axis=1)
        self.df['N_Concentrate_Avg'] = self.df[['N_Concentrate1', 'N_Concentrate2', 'N_Concentrate3']].replace(0, 27.3).mean(axis=1)
        
        # 3. Calculate Final Nutrient Aggregates (N, CP, VRE)
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
        
        # Scaling animal weights dynamically based on reference mature weight curves
        weight_birth = cow_weight * self.RATIO_WEIGHT_BIRTH_TO_MATURE
        weight_heifer_1year = cow_weight * self.RATIO_WEIGHT_HEIFER_TO_MATURE
        weight_calving = cow_weight * self.RATIO_WEIGHT_CALVING_TO_MATURE
        
        # Convert total body composition targets into absolute Nitrogen quantities
        n_birth_kg = weight_birth * self.N_CONCENTRATION_CALF_BIRTH / 1000
        n_heifer_1year_kg = weight_heifer_1year * self.N_CONCENTRATION_HEIFER_GROWTH / 1000
        n_calving_kg = weight_calving * self.N_CONCENTRATION_REPLACEMENT_CALVING / 1000
        n_mature_cow_kg = cow_weight * self.N_CONCENTRATION_MATURE_COW / 1000
        
        year_multiplier = np.where(df['MilkYield'] < 100, 365, 1)
        total_milk_yield = df['MilkYield'] * year_multiplier * df['Nr_koe']
        
        # Biological N Retention outputs (Milk production and fetal tissue growth)
        n_retention_milk = total_milk_yield * df['Pro%'] * 10 / self.CONVERSION_DAIRY_N_TO_CRUDE_PROTEIN / 1000
        n_retention_fetus = n_birth_kg * self.CALVING_RATE_MULTIPLIER * df['Nr_koe']
        
        # N flow balances resulting from herd rotation (cull outflows vs replacement inflows)
        n_replacement_inflow = self.REPLACEMENT_RATE_FRACTION * n_calving_kg * df['Nr_koe']
        n_cull_outflow = self.REPLACEMENT_RATE_FRACTION * n_mature_cow_kg * df['Nr_koe']
        df['N_Retention_Cow_VCRE'] = n_retention_milk + n_retention_fetus + (n_cull_outflow - n_replacement_inflow)
        
        # Dynamic growth modeling for young stock calves
        growth_target_n = n_heifer_1year_kg - n_birth_kg
        growth_constant = 0.36 * df['breed_factor']
        temp_term_growth = growth_target_n * (0.376 / 0.407)
        temp_term_maintenance = (growth_constant / 2 * 24) * (0.031 / 0.407)
        
        calf_n_retention_ratio = np.divide(temp_term_growth + temp_term_maintenance, growth_target_n, out=np.ones_like(growth_target_n.values, dtype=float), where=growth_target_n.values != 0)
        df['N_Retention_Kalf_VCRE'] = growth_target_n * df['Nr_kalf'] * calf_n_retention_ratio
        
        # Dynamic growth modeling for replacement heifers
        growth_heifer_n = (n_calving_kg - n_heifer_1year_kg) * (12 / 14)
        df['N_Retention_Pink_VCRE'] = (n_birth_kg + growth_heifer_n) * df['Nr_pink']
        
        # System aggregation across the entire farm inventory
        df['Total_N_Retention_Farm_VCRE'] = df['N_Retention_Cow_VCRE'] + df['N_Retention_Kalf_VCRE'] + df['N_Retention_Pink_VCRE']
        df['Total_N_Excretion_Cow_VCRE'] = df['Total_N_Intake_Cow'] - df['N_Retention_Cow_VCRE']
        df['Total_N_Excretion_Kalf_VCRE'] = df['Total_N_Intake_Kalf'] - df['N_Retention_Kalf_VCRE']
        df['Total_N_Excretion_Pink_VCRE'] = df['Total_N_Intake_Pink'] - df['N_Retention_Pink_VCRE']
        df['Total_N_Excretion_Farm_VCRE'] = df['Total_N_Excretion_Cow_VCRE'] + df['Total_N_Excretion_Kalf_VCRE'] + df['Total_N_Excretion_Pink_VCRE']
        
        print(f"  ✓ 2.4 done — Farm1 total excretion: {df['Total_N_Excretion_Farm_VCRE'].iloc[0]:.1f} kg")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.5 — Nitrogen Partitioning into Urinary (UUN) & Fecal (FN) (VCRE)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningVCRE:
    DAIRY_COW_CONVERSION_EFFICIENCY_FACTOR = 0.91  
    
    def __init__(self, df): 
        self.df = df.copy()
        
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 2.5: Nitrogen Partitioning (VCRE Method)\n{'='*70}")
        df = self.df
        for category_code, category_label in [('cow', 'Cow'), ('pink', 'Pink'), ('kalf', 'Kalf')]:
            # Scale total crude digestible protein intake using core efficiency factors
            digestible_n_intake = df[f'Total_N_Intake_{category_label}'] * df[f'VCRE_Factor_{category_label}'] * self.DAIRY_COW_CONVERSION_EFFICIENCY_FACTOR
            
            # Urinary Urea Nitrogen (UUN) represents the excess digestible pool over tissue retention
            urinary_urea_n = digestible_n_intake - df[f'N_Retention_{category_label}_VCRE']
            
            # Fecal Nitrogen (FN) constitutes the remaining excreted mass balance
            fecal_n = df[f'Total_N_Excretion_{category_label}_VCRE'] - urinary_urea_n
            
            df[f'UUN_{category_label}_VCRE'] = urinary_urea_n
            df[f'FN_{category_label}_VCRE'] = fecal_n
            
        df['Total_UUN_Farm_VCRE'] = df['UUN_Cow_VCRE'] + df['UUN_Kalf_VCRE'] + df['UUN_Pink_VCRE']
        df['Total_FN_Farm_VCRE'] = df['FN_Cow_VCRE'] + df['FN_Kalf_VCRE'] + df['FN_Pink_VCRE']
        
        print(f"  ✓ 2.5 done — Farm1 VCRE UUN Pool: {df['Total_UUN_Farm_VCRE'].iloc[0]:.1f} kg")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 3.1 — Manure Net Mineralization (MUN & VCRE Systems)
# ══════════════════════════════════════════════════════════════════════════════
class MineralizationCalculator:
    MINERALIZATION_RATE_SLURRY_FN = 0.10   # Mineralization release coefficient of slurry fecal organic bound N
    MINERALIZATION_RATE_SOLID_UUN = -0.25   # Volatilization tie-down / loss parameter for solid manure systems
    
    def __init__(self, df): 
        self.df = df.copy()
        
    def calculate_mineralization(self):
        print(f"\n{'='*70}\nMODULE 3.1: Net Manure Mineralization Calculations\n{'='*70}")
        df = self.df
        
        # --- A. Mineralization tracking for traditional MUN-based matrix ---
        for animal_name, fecal_n_col, urinary_n_col, slurry_fraction_col, suffix in [
            ('Cow', 'fn_total_cows_kg', 'uun_total_cows_kg', 'slurry%_koe', 'Cow'),
            ('Kalf', 'fn_total_kalf_kg', 'uun_total_kalf_kg', 'slurry%_kalf', 'Kalf'),
            ('Pink', 'fn_total_pink_kg', 'uun_total_pink_kg', 'slurry%_pink', 'Pink')]:
            
            min_from_fecal_n = df[fecal_n_col] * df[slurry_fraction_col] * self.MINERALIZATION_RATE_SLURRY_FN
            min_from_urinary_n = df[urinary_n_col] * (1 - df[slurry_fraction_col]) * self.MINERALIZATION_RATE_SOLID_UUN
            df[f'Net_Min_{suffix}_MUN'] = min_from_fecal_n + min_from_urinary_n
            
        df['Total_Net_Mineralization_MUN'] = df['Net_Min_Cow_MUN'] + df['Net_Min_Kalf_MUN'] + df['Net_Min_Pink_MUN']
        
        # --- B. Mineralization tracking for updated VCRE-based matrix ---
        for animal_name, fecal_n_col, urinary_n_col, slurry_fraction_col, suffix in [
            ('Cow', 'FN_Cow_VCRE', 'UUN_Cow_VCRE', 'slurry%_koe', 'Cow'),
            ('Kalf', 'FN_Kalf_VCRE', 'UUN_Kalf_VCRE', 'slurry%_kalf', 'Kalf'),
            ('Pink', 'FN_Pink_VCRE', 'UUN_Pink_VCRE', 'slurry%_pink', 'Pink')]:
            
            min_from_fecal_n = df[fecal_n_col] * df[slurry_fraction_col] * self.MINERALIZATION_RATE_SLURRY_FN
            min_from_urinary_n = df[urinary_n_col] * (1 - df[slurry_fraction_col]) * self.MINERALIZATION_RATE_SOLID_UUN
            df[f'Net_Min_{suffix}_VCRE'] = min_from_fecal_n + min_from_urinary_n
            
        df['Total_Net_Mineralization_VCRE'] = df['Net_Min_Cow_VCRE'] + df['Net_Min_Kalf_VCRE'] + df['Net_Min_Pink_VCRE']
        df['Mineralization_Diff'] = df['Total_Net_Mineralization_VCRE'] - df['Total_Net_Mineralization_MUN']
        
        print(f"  ✓ 3.1 done — Mineralization MUN: {df['Total_Net_Mineralization_MUN'].iloc[0]:.2f}, VCRE: {df['Total_Net_Mineralization_VCRE'].iloc[0]:.2f}")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 3.2 — Corrected Total Ammoniacal Nitrogen (TAN) Calculations
# ══════════════════════════════════════════════════════════════════════════════
class CorrectedTANCalculator:
    def __init__(self, df): 
        self.df = df.copy()
        
    def calculate_corrected_tan(self):
        print(f"\n{'='*70}\nMODULE 3.2: Corrected Total Ammoniacal Nitrogen (TAN)\n{'='*70}")
        df = self.df
        
        # Loop through both diagnostic methodologies to dynamically evaluate corrected target pools
        for method_name, uun_columns, mineralization_columns in [
            ('MUN', ['uun_total_cows_kg', 'uun_total_kalf_kg', 'uun_total_pink_kg'], ['Net_Min_Cow_MUN', 'Net_Min_Kalf_MUN', 'Net_Min_Pink_MUN']),
            ('VCRE', ['UUN_Cow_VCRE', 'UUN_Kalf_VCRE', 'UUN_Pink_VCRE'], ['Net_Min_Cow_VCRE', 'Net_Min_Kalf_VCRE', 'Net_Min_Pink_VCRE'])]:
            
            for animal_name, uun_col, mineralization_col in zip(['Cow', 'Kalf', 'Pink'], uun_columns, mineralization_columns):
                # Corrected TAN = Soluble Urinary N (UUN) + Net Mineralization releases from the organic pool
                df[f'Corrected_TAN_{animal_name}_{method_name}'] = df[uun_col] + df[mineralization_col]
                
            df[f'Total_Corrected_TAN_{method_name}'] = (
                df[f'Corrected_TAN_Cow_{method_name}'] + 
                df[f'Corrected_TAN_Kalf_{method_name}'] + 
                df[f'Corrected_TAN_Pink_{method_name}']
            )
            
        print(f"  ✓ 3.2 done — Total TAN Pool MUN: {df['Total_Corrected_TAN_MUN'].iloc[0]:.1f} kg, VCRE: {df['Total_Corrected_TAN_VCRE'].iloc[0]:.1f} kg")
        return df
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.1 — Ammonia Emissions (Stable / Storage / Grazing)
# ══════════════════════════════════════════════════════════════════════════════
class EmissionCalculator:
    EMISSION_FACTOR_STABLE = 0.143
    EMISSION_FACTOR_GRAZING = 0.04
    EMISSION_FACTOR_STORAGE_SOLID_MANURE = 0.01
    EMISSION_FACTOR_STORAGE_SLURRY_OUTDOOR = 0.024
    EMISSION_FACTOR_STORAGE_SOLID_OUTDOOR = 0.035
    FRACTION_MANURE_TO_SOLID_STORAGE = 0.20
    MOLECULAR_WEIGHT_RATIO_N_TO_NH3 = 14.0 / 17.0

    def __init__(self, df): 
        self.df = df.copy()
        
    def calculate_emissions(self):
        print(f"\n{'='*70}\nMODULE 4.1: Ammonia Emissions (Stable, Storage, Grazing)\n{'='*70}")
        df = self.df
        
        numeric_columns_to_verify = [
            'GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'GH_Koe',
            'GD_Unlimited_Pink', 'GH_Pink', 'GD_Unlimited_Kalf', 'GH_Kalf',
            'slurry%_koe', 'slurry%_kalf', 'slurry%_pink'
        ]
        for column in numeric_columns_to_verify:
            df = ensure_numeric(df, column, 0.0)
            
        df['total_grazing_days_cow'] = df['GD_Limited_Koe'] + df['GD_Combi_Koe'] + df['GD_Unlimited_Koe']
        
        # Animal categories configuration dictionary mapping to target inventory matrices
        animal_categories_configuration = [
            {
                'animal_name': 'Cow', 
                'grazing_days_col': 'total_grazing_days_cow', 
                'grazing_hours_col': 'GH_Koe', 
                'slurry_fraction_col': 'slurry%_koe',
                'gross_n_mun_col': 'gross_n_cows_mun', 
                'uun_mun_col': 'uun_total_cows_kg', 
                'corrected_tan_mun_col': 'Corrected_TAN_Cow_MUN',
                'gross_n_vcre_col': 'Total_N_Excretion_Cow_VCRE', 
                'uun_vcre_col': 'UUN_Cow_VCRE', 
                'corrected_tan_vcre_col': 'Corrected_TAN_Cow_VCRE'
            },
            {
                'animal_name': 'Pink', 
                'grazing_days_col': 'GD_Unlimited_Pink', 
                'grazing_hours_col': 'GH_Pink', 
                'slurry_fraction_col': 'slurry%_pink',
                'gross_n_mun_col': 'gross_n_heifers_mun', 
                'uun_mun_col': 'uun_total_pink_kg', 
                'corrected_tan_mun_col': 'Corrected_TAN_Pink_MUN',
                'gross_n_vcre_col': 'Total_N_Excretion_Pink_VCRE', 
                'uun_vcre_col': 'UUN_Pink_VCRE', 
                'corrected_tan_vcre_col': 'Corrected_TAN_Pink_VCRE'
            },
            {
                'animal_name': 'Kalf', 
                'grazing_days_col': 'GD_Unlimited_Kalf', 
                'grazing_hours_col': 'GH_Kalf', 
                'slurry_fraction_col': 'slurry%_kalf',
                'gross_n_mun_col': 'gross_n_calves_mun', 
                'uun_mun_col': 'uun_total_kalf_kg', 
                'corrected_tan_mun_col': 'Corrected_TAN_Kalf_MUN',
                'gross_n_vcre_col': 'Total_N_Excretion_Kalf_VCRE', 
                'uun_vcre_col': 'UUN_Kalf_VCRE', 
                'corrected_tan_vcre_col': 'Corrected_TAN_Kalf_VCRE'
            }
        ]
        
        for category in animal_categories_configuration:
            name = category['animal_name']
            grazing_days = df[category['grazing_days_col']].fillna(0)
            grazing_hours = df[category['grazing_hours_col']].fillna(0)
            slurry_fraction = df[category['slurry_fraction_col']].fillna(1.0)
            
            # Grazing reduction factor on stable emissions based on NEMA model assumptions
            grazing_reduction_factor = (0.0261 * grazing_hours * (grazing_days / 365.0)).clip(0, 1)
            
            # Fraction of temporal period animals stay indoors vs outdoor pastures
            indoor_emission_period_fraction = (1.0 - (grazing_days * grazing_hours) / (365.0 * 24.0)).clip(0, 1)
            
            for methodology in ['mun', 'vcre']:
                method_suffix = 'MUN' if methodology == 'mun' else 'VCRE'
                gross_n_col = category['gross_n_mun_col'] if methodology == 'mun' else category['gross_n_vcre_col']
                uun_col = category['uun_mun_col'] if methodology == 'mun' else category['uun_vcre_col']
                tan_col = category['corrected_tan_mun_col'] if methodology == 'mun' else category['corrected_tan_vcre_col']
                
                gross_n_excretion = df[gross_n_col].fillna(0)
                urinary_urea_n = df[uun_col].fillna(0)
                total_ammoniacal_nitrogen = df[tan_col].fillna(0)
                
                # 1. House/Stable Ammonia Emissions calculation step
                stable_emission = (total_ammoniacal_nitrogen * self.EMISSION_FACTOR_STABLE * (1.0 - grazing_reduction_factor)) / self.MOLECULAR_WEIGHT_RATIO_N_TO_NH3
                df[f'Emission_Stable_{name}_{method_suffix}'] = stable_emission
                
                # 2. Outdoor Storage Emissions calculation step
                nitrogen_excreted_indoors = gross_n_excretion * indoor_emission_period_fraction
                outdoor_storage_loss = nitrogen_excreted_indoors * (slurry_fraction * self.EMISSION_FACTOR_STORAGE_SLURRY_OUTDOOR + (1.0 - slurry_fraction) * self.EMISSION_FACTOR_STORAGE_SOLID_OUTDOOR)
                
                remaining_tan_for_solid_manure_storage = ((nitrogen_excreted_indoors - outdoor_storage_loss) / self.MOLECULAR_WEIGHT_RATIO_N_TO_NH3 - stable_emission).clip(lower=0)
                storage_emission_solid = remaining_tan_for_solid_manure_storage * self.FRACTION_MANURE_TO_SOLID_STORAGE * self.EMISSION_FACTOR_STORAGE_SOLID_MANURE
                df[f'Emission_Storage_{name}_{method_suffix}'] = storage_emission_solid
                
                # 3. Grazing/Pasture Emissions calculation step
                urinary_urea_n_excreted_outdoors = urinary_urea_n * (1.0 - indoor_emission_period_fraction)
                grazing_emission = (urinary_urea_n_excreted_outdoors * self.EMISSION_FACTOR_GRAZING) / self.MOLECULAR_WEIGHT_RATIO_N_TO_NH3
                df[f'Emission_Grazing_{name}_{method_suffix}'] = grazing_emission
                
        for method_suffix in ['MUN', 'VCRE']:
            for emission_type in ['Stable', 'Storage', 'Grazing']:
                df[f'Total_Emission_{emission_type}_{method_suffix}'] = sum(df[f'Emission_{emission_type}_{n}_{method_suffix}'] for n in ['Cow', 'Pink', 'Kalf'])
                
            df[f'Total_Emission_All_{method_suffix}'] = (
                df[f'Total_Emission_Stable_{method_suffix}'] + 
                df[f'Total_Emission_Storage_{method_suffix}'] + 
                df[f'Total_Emission_Grazing_{method_suffix}']
            )
            
        print(f"  ✓ 4.1 done — Total MUN: {df['Total_Emission_All_MUN'].iloc[0]:.2f}, VCRE: {df['Total_Emission_All_VCRE'].iloc[0]:.2f}")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.2 — Land Application Limits (VCRE-based, Dutch RVO Regulatory Rules)
# ══════════════════════════════════════════════════════════════════════════════
class LandApplicationCalculator:
    TABEL_2_NITROGEN_USE_NORMS = {
        'grasland_beweiden':        {'klei': 345, 'zand_nwc': 250, 'zand_zuid': 250, 'loss': 250, 'veen': 265},
        'grasland_maaien':          {'klei': 385, 'zand_nwc': 320, 'zand_zuid': 320, 'loss': 320, 'veen': 300},
        'mais_derogatie':           {'klei': 160, 'zand_nwc': 140, 'zand_zuid': 112, 'loss': 112, 'veen': 150},
        'mais_geen_derogatie':      {'klei': 185, 'zand_nwc': 140, 'zand_zuid': 112, 'loss': 112, 'veen': 150},
        'consumptieaardappel_hoog': {'klei': 275, 'zand_nwc': 260, 'zand_zuid': 208, 'loss': 204, 'veen': 270},
        'consumptieaardappel_overig': {'klei': 250, 'zand_nwc': 235, 'zand_zuid': 188, 'loss': 184, 'veen': 245},
        'wintertarwe':              {'klei': 245, 'zand_nwc': 160, 'zand_zuid': 160, 'loss': 190, 'veen': 160},
        'zomertarwe':               {'klei': 150, 'zand_nwc': 140, 'zand_zuid': 140, 'loss': 140, 'veen': 140},
        'wintergerst':              {'klei': 140, 'zand_nwc': 140, 'zand_zuid': 140, 'loss': 140, 'veen': 140},
        'suikerbieten':             {'klei': 150, 'zand_nwc': 145, 'zand_zuid': 116, 'loss': 116, 'veen': 145},
        'zetmeelaardappelen':       {'klei': 240, 'zand_nwc': 230, 'zand_zuid': 184, 'loss': 184, 'veen': 230},
        'overig':                   {'klei': 200, 'zand_nwc': 185, 'zand_zuid': 148, 'loss': 148, 'veen': 190},
    }

    WORKING_COEFFICIENT_SLURRY_GRAZING = 0.45
    WORKING_COEFFICIENT_SLURRY_NO_GRAZING = 0.60
    WORKING_COEFFICIENT_SOLID_GRAZING = 0.45
    WORKING_COEFFICIENT_SOLID_NO_GRAZING = 0.60

    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self._ensure_columns()

    def _ensure_columns(self):
        """Validates that input datasets conform to structural requirements."""
        if 'HA_Crop' in self.df.columns and 'Ha_Crop' not in self.df.columns:
            self.df.rename(columns={'HA_Crop': 'Ha_Crop'}, inplace=True)

        if 'CropType' not in self.df.columns:
            self.df['CropType'] = 'overig'

        columns_to_verify_numeric = ['Ha_Grass', 'Ha_Mais', 'Ha_Crop', 'NatureGL%']
        for col in columns_to_verify_numeric:
            self.df = ensure_numeric(self.df, col, 0.0)

        conditional_columns = ['GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'slurry%_koe', 'Derogation', 'NV_Area']
        for col in conditional_columns:
            if col in ['Derogation', 'NV_Area']:
                if col not in self.df.columns:
                    self.df[col] = 'No'
            else:
                self.df = ensure_numeric(self.df, col, 0.0)

        if self.df['slurry%_koe'].max() > 1.0:
            self.df['slurry%_koe'] = (self.df['slurry%_koe'] / 100.0).clip(0.0, 1.0)
        else:
            self.df['slurry%_koe'] = self.df['slurry%_koe'].clip(0.0, 1.0)

        if 'Soil_Type' not in self.df.columns:
            self.df['Soil_Type'] = 'zand'
        if 'Region' not in self.df.columns:
            self.df['Region'] = 'Others'

    @staticmethod
    def _is_yes_string(value) -> bool:
        return str(value).strip().lower() in ('yes', 'ja', 'true', '1', 'y')

    @staticmethod
    def _get_soil_key(soil_type, region):
        soil_str = str(soil_type).lower()
        region_str = str(region).lower()
        if 'klei' in soil_str:
            return 'klei'
        if 'veen' in soil_str:
            return 'veen'
        if any(term in soil_str for term in ('loss', 'loes', 'löss')):
            return 'loss'
        if 'zand' in soil_str:
            return 'zand_zuid' if ('zuid' in region_str or 'south' in region_str) else 'zand_nwc'
        return 'zand_nwc'

    def _get_crop_nitrogen_norm(self, crop_key, soil_key) -> float:
        cleaned_crop_key = str(crop_key).strip().lower().replace('ï', 'i')
        entry = self.TABEL_2_NITROGEN_USE_NORMS.get(cleaned_crop_key, self.TABEL_2_NITROGEN_USE_NORMS['overig'])
        return float(entry.get(soil_key, 185.0))

    def _get_grassland_management_key(self, row) -> str:
        total_grazing_days = float(row.get('GD_Limited_Koe', 0) or 0) + float(row.get('GD_Combi_Koe', 0) or 0) + float(row.get('GD_Unlimited_Koe', 0) or 0)
        return 'grasland_beweiden' if total_grazing_days > 0 else 'grasland_maaien'

    def _get_maize_derogation_key(self, row) -> str:
        is_derogation = self._is_yes_string(row.get('Derogation', 'No'))
        return 'mais_derogatie' if is_derogation else 'mais_geen_derogatie'

    def _get_manure_limit_per_hectare(self, row) -> float:
        is_derogation = self._is_yes_string(row.get('Derogation', 'No'))
        is_nitrate_vulnerable_zone = self._is_yes_string(row.get('NV_Area', 'No'))
        if is_derogation:
            return 190.0 if is_nitrate_vulnerable_zone else 200.0
        return 170.0

    def _check_if_grazing_practiced(self, row) -> bool:
        total_grazing_days = float(row.get('GD_Limited_Koe', 0) or 0) + float(row.get('GD_Combi_Koe', 0) or 0) + float(row.get('GD_Unlimited_Koe', 0) or 0)
        return total_grazing_days > 0

    def _calculate_working_coefficient(self, row) -> float:
        slurry_fraction = float(row.get('slurry%_koe', 1.0) or 1.0)
        solid_fraction = 1.0 - slurry_fraction
        if self._check_if_grazing_practiced(row):
            return slurry_fraction * self.WORKING_COEFFICIENT_SLURRY_GRAZING + solid_fraction * self.WORKING_COEFFICIENT_SOLID_GRAZING
        return slurry_fraction * self.WORKING_COEFFICIENT_SLURRY_NO_GRAZING + solid_fraction * self.WORKING_COEFFICIENT_SOLID_NO_GRAZING

    def calculate_land_application(self) -> pd.DataFrame:
        print(f"\n{'='*70}\nMODULE 4.2: Land Application Limit Calculations (VCRE-based)\n{'='*70}")
        df = self.df

        df['Soil_Key_5_1'] = [self._get_soil_key(s, r) for s, r in zip(df['Soil_Type'], df['Region'])]
        df['Ha_Total_For_ManureLimit'] = (df['Ha_Grass'] + df['Ha_Mais'] + df['Ha_Crop']).clip(lower=0)

        df['Manure_Limit_kgN_per_ha'] = df.apply(self._get_manure_limit_per_hectare, axis=1)
        df['Manure_Limit_Total_kgN'] = df['Manure_Limit_kgN_per_ha'] * df['Ha_Total_For_ManureLimit']

        grassland_type_key = df.apply(self._get_grassland_management_key, axis=1)
        maize_type_key = df.apply(self._get_maize_derogation_key, axis=1)

        df['Norm_Grass_kgN_per_ha'] = [self._get_crop_nitrogen_norm(k, sk) for k, sk in zip(grassland_type_key, df['Soil_Key_5_1'])]
        df['Norm_Maize_kgN_per_ha'] = [self._get_crop_nitrogen_norm(k, sk) for k, sk in zip(maize_type_key, df['Soil_Key_5_1'])]
        df['Norm_Crop_kgN_per_ha'] = [self._get_crop_nitrogen_norm(k, sk) for k, sk in zip(df['CropType'], df['Soil_Key_5_1'])]

        df['UsageSpace_Grass_kgN'] = df['Norm_Grass_kgN_per_ha'] * df['Ha_Grass']
        df['UsageSpace_Maize_kgN'] = df['Norm_Maize_kgN_per_ha'] * df['Ha_Mais']
        df['UsageSpace_Crop_kgN'] = df['Norm_Crop_kgN_per_ha'] * df['Ha_Crop']
        df['UsageSpace_Total_kgN'] = df['UsageSpace_Grass_kgN'] + df['UsageSpace_Maize_kgN'] + df['UsageSpace_Crop_kgN']

        for col in ['Total_N_Excretion_Farm_VCRE', 'Total_Emission_All_VCRE']:
            if col not in df.columns:
                df[col] = 0.0
            df = ensure_numeric(df, col, 0.0)

        df['ManureN_Available_After_Emissions_kgN'] = (df['Total_N_Excretion_Farm_VCRE'] - df['Total_Emission_All_VCRE']).clip(lower=0)
        df['ManureN_Applied_kgN'] = np.minimum(df['ManureN_Available_After_Emissions_kgN'], df['Manure_Limit_Total_kgN'])

        df['Working_Coefficient'] = df.apply(self._calculate_working_coefficient, axis=1)
        df['ManureN_Effective_kgN'] = df['ManureN_Applied_kgN'] * df['Working_Coefficient']
        df['FertiliserSpace_kgN'] = (df['UsageSpace_Total_kgN'] - df['ManureN_Effective_kgN']).clip(lower=0)

        df['BindingConstraint'] = np.where(
            df['ManureN_Available_After_Emissions_kgN'] >= df['Manure_Limit_Total_kgN'],
            'ManureLimit',
            'Availability'
        )

        print(f"  ✓ 4.2 done — Farm1 UsageSpace: {df['UsageSpace_Total_kgN'].iloc[0]:.1f} kg N; Applied manure: {df['ManureN_Applied_kgN'].iloc[0]:.1f} kg N ({df['BindingConstraint'].iloc[0]})")
        return df


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.3 — Land Application Ammonia Emissions (Manure & Synthetic Fertiliser)
# ══════════════════════════════════════════════════════════════════════════════
class ApplicationEmissionCalculator:
    EMISSION_FACTOR_SYNTHETIC_FERTILISER = {
        '100%ammonium': 11.3,
        '100%nitraat': 0.0,
        'combinatievanammoniumetnitraat': 2.5, 
        'combinatievanammoniumennitraat': 2.5, 
        'ureum,gekorreld,zonderurease-remmer': 14.3,
        'ureum,gekorreld,meturease-remmer': 5.9,
        'voeibaarureumzonderurease-remmerofzuur': 7.5, 
        'vloeibaarureumzonderurease-remmerofzuur': 7.5,
        'vloeibaarureummeturease-remmerofzuuer': 3.1,  
        'vloeibaarureummeturease-remmerofzuur': 3.1,
        'vloeibaarureumtoegviainjectie': 1.5,          
        'vloeibaarureumtoegediendviainjectie': 1.5
    }

    EMISSION_FACTOR_MANURE_GRASSLAND = {
        'bovengronds': 68.0,
        'sleepvoet': 26.4,
        'sleepvoetverdund': 17.0,
        'sleufkouterverdund': 17.0,
        'sleufkouter': 21.7,
        'zodebemester': 17.0
    }

    EMISSION_FACTOR_MANURE_CROPLAND = {
        'bovengronds': 69.0,
        'ineenwerkgangonderwerken': 22.0,
        'sleepvoet': 36.0,
        'diepeinjectie': 2.0,
        'ondiepeinjectie': 24.0
    }
    
    MOLECULAR_WEIGHT_RATIO_N_TO_NH3 = 14.0 / 17.0
    MOLECULAR_WEIGHT_RATIO_NH3_TO_N = 17.0 / 14.0

    def __init__(self, df):
        self.df = df.copy()

    def _normalize_and_clean_string(self, string_value):
        return str(string_value).lower().replace(' ', '').strip()

    def run_application_emission(self):
        print(f"\n{'='*70}\nMODULE 4.3: Land Application Ammonia Emissions\n{'='*70}")
        
        fertiliser_forms_series = self.df.get('Fertiliser_Form', pd.Series('100%ammonium', index=self.df.index))
        cleaned_fertiliser_forms = fertiliser_forms_series.apply(self._normalize_and_clean_string)
        self.df['EF_Fertiliser_%'] = cleaned_fertiliser_forms.map(self.EMISSION_FACTOR_SYNTHETIC_FERTILISER).fillna(0.0)

        manure_application_technology_grassland = self.df.get('AM_Grassland', pd.Series('zodebemester', index=self.df.index))
        self.df['EF_Manure_Grass_%'] = manure_application_technology_grassland.apply(self._normalize_and_clean_string).map(self.EMISSION_FACTOR_MANURE_GRASSLAND).fillna(17.0)
        
        manure_application_technology_cropland = self.df.get('AM_Cropland', pd.Series('ondiepeinjectie', index=self.df.index))
        self.df['EF_Manure_Arable_%'] = manure_application_technology_cropland.apply(self._normalize_and_clean_string).map(self.EMISSION_FACTOR_MANURE_CROPLAND).fillna(24.0)

        corrected_tan_mun = self.df['Total_Corrected_TAN_MUN']
        indoor_emissions_total_mun = self.df['Total_Emission_All_MUN']
        
        if 'Total_Nitrogen_Excretion_kg' in self.df.columns:
            nitrogen_excretion_total_mun = self.df['Total_Nitrogen_Excretion_kg']
        else:
            nitrogen_excretion_total_mun = self.df.get('Total_Nitrogen_Excretion_MUN', 0.0) 
        
        corrected_tan_vcre = self.df['Total_Corrected_TAN_VCRE']
        indoor_emissions_total_vcre = self.df['Total_Emission_All_VCRE']
        nitrogen_excretion_total_vcre = self.df['Total_N_Excretion_Farm_VCRE'] 
        
        applied_manure_nitrogen = self.df['ManureN_Applied_kgN']
        applied_fertiliser_nitrogen = self.df['FertiliserSpace_kgN']
        
        net_tan_pool_after_indoor_losses_mun = (corrected_tan_mun - (indoor_emissions_total_mun * self.MOLECULAR_WEIGHT_RATIO_N_TO_NH3)).clip(lower=0)
        net_tan_pool_after_indoor_losses_vcre = (corrected_tan_vcre - (indoor_emissions_total_vcre * self.MOLECULAR_WEIGHT_RATIO_N_TO_NH3)).clip(lower=0)
        
        average_net_tan_pool = (net_tan_pool_after_indoor_losses_mun + net_tan_pool_after_indoor_losses_vcre) / 2.0
        self.df['Avg_Net_TAN_Excreted'] = average_net_tan_pool
        
        average_nitrogen_excretion = (nitrogen_excretion_total_mun + nitrogen_excretion_total_vcre) / 2.0
        safe_average_nitrogen_excretion = np.where(average_nitrogen_excretion > 0, average_nitrogen_excretion, 1.0)
        
        average_tan_percentage = average_net_tan_pool / safe_average_nitrogen_excretion
        self.df['Avg_TAN_Pct'] = average_tan_percentage.clip(0, 1)
        
        hectares_grassland = self.df.get('Ha_Grass', 0.0)
        hectares_maize = self.df.get('Ha_Mais', 0.0)
        hectares_other_crops = self.df.get('Ha_Crop', 0.0)
        
        total_farm_hectares = hectares_grassland + hectares_maize + hectares_other_crops
        safe_total_farm_hectares = np.where(total_farm_hectares > 0, total_farm_hectares, 1.0) 
        
        fraction_area_grassland = hectares_grassland / safe_total_farm_hectares
        fraction_area_cropland = (hectares_maize + hectares_other_crops) / safe_total_farm_hectares
        
        applied_manure_nitrogen_grassland = applied_manure_nitrogen * fraction_area_grassland
        applied_manure_nitrogen_cropland = applied_manure_nitrogen * fraction_area_cropland
        
        tan_applied_manure_grassland = applied_manure_nitrogen_grassland * self.df['Avg_TAN_Pct']
        tan_applied_manure_cropland = applied_manure_nitrogen_cropland * self.df['Avg_TAN_Pct']
        
        self.df['TAN_Applied_Manure_Grass'] = tan_applied_manure_grassland
        self.df['TAN_Applied_Manure_Arable'] = tan_applied_manure_cropland
        self.df['TAN_Applied_Manure_Total'] = tan_applied_manure_grassland + tan_applied_manure_cropland
        
        # Multiply by molecular scale factor 17/14 to convert from N back to raw NH3 gas equivalents
        application_emission_manure_grassland = (tan_applied_manure_grassland * self.df['EF_Manure_Grass_%'] / 100.0) * self.MOLECULAR_WEIGHT_RATIO_NH3_TO_N
        application_emission_manure_cropland = (tan_applied_manure_cropland * self.df['EF_Manure_Arable_%'] / 100.0) * self.MOLECULAR_WEIGHT_RATIO_NH3_TO_N
        
        total_application_emission_manure = application_emission_manure_grassland + application_emission_manure_cropland
        
        self.df['Emission_ManureApp_Grass'] = application_emission_manure_grassland
        self.df['Emission_ManureApp_Arable'] = application_emission_manure_cropland
        self.df['Emission_ManureApp_Total'] = total_application_emission_manure
        
        total_application_emission_fertiliser = (applied_fertiliser_nitrogen * self.df['EF_Fertiliser_%'] / 100.0) * self.MOLECULAR_WEIGHT_RATIO_NH3_TO_N
        self.df['Emission_FertiliserApp'] = total_application_emission_fertiliser

        self.df['Total_Farm_NH3_Emission_MUN'] = indoor_emissions_total_mun + total_application_emission_manure + total_application_emission_fertiliser
        self.df['Total_Farm_NH3_Emission_VCRE'] = indoor_emissions_total_vcre + total_application_emission_manure + total_application_emission_fertiliser
            
        # Dutch standard livestock unit conversion (Grootvee-eenheid / GVE) values
        self.df['Total_GVE_Farm'] = (self.df['Nr_koe'] * 1.0) + (self.df['Nr_pink'] * 0.50) + (self.df['Nr_kalf'] * 0.25)
        
        safe_total_gve_farm = np.where(self.df['Total_GVE_Farm'] > 0, self.df['Total_GVE_Farm'], 1.0)
        
        self.df['NH3_Emission_per_GVE_MUN'] = self.df['Total_Farm_NH3_Emission_MUN'] / safe_total_gve_farm
        self.df['NH3_Emission_per_GVE_VCRE'] = self.df['Total_Farm_NH3_Emission_VCRE'] / safe_total_gve_farm

        self.df['Stable_Emission_per_GVE_MUN'] = self.df['Total_Emission_Stable_MUN'] / safe_total_gve_farm
        self.df['Stable_Emission_per_GVE_VCRE'] = self.df['Total_Emission_Stable_VCRE'] / safe_total_gve_farm
        self.df['Fertiliser_Emission_per_GVE'] = self.df['Emission_FertiliserApp'] / safe_total_gve_farm

        print(f"  ✓ 4.3 Application Emissions Calculated")
        print(f"      Farm GVE: {self.df['Total_GVE_Farm'].iloc[0]:.1f}, Total NH3/GVE (VCRE): {self.df['NH3_Emission_per_GVE_VCRE'].iloc[0]:.1f} kg")
        return self.df   
     
def run_pipeline(INPUT_PATH='InputREMAS.xlsx', OUTPUT_PATH='Output_REMAS_Complete.xlsx', SHEET='Main input'):
    # ── 1.1 Load + manure totals
    m11 = FarmManureCalculator(INPUT_PATH, SHEET)
    df = m11.load_and_clean()
    m11.df = df
    df = m11.calculate_totals()

    # ── 1.2 MUN partitioning (kept for comparison / dependencies)
    df = NitrogenPartitioningCalculator(df).calculate_partitioning()

    # ── 2.1 VEM requirements
    df = VEMRequirementCalculator(df).calculate_requirements()

    # ── 2.2 VEM allocation
    df = VEMAllocationCalculator(df).run_allocation()

    # ── 2.3 N intake
    df = NitrogenIntakeCalculator(df).run_nutrient_calculation()

    # ── 2.4 VCRE excretion
    df = NitrogenExcretionCalculatorVCRE(df).calculate_excretion()

    # ── 2.5 VCRE partitioning
    df = NitrogenPartitioningVCRE(df).calculate_partitioning()

    # ── 3.1 mineralization
    df = MineralizationCalculator(df).calculate_mineralization()

    # ── 3.2 corrected TAN
    df = CorrectedTANCalculator(df).calculate_corrected_tan()

    # ── 4.1 emissions during stable, grazing, storage
    df = EmissionCalculator(df).calculate_emissions()

    # ── 4.2 land application (VCRE)
    df = LandApplicationCalculator(df).calculate_land_application()
    
    # ── 4.3 Application Emission
    df = ApplicationEmissionCalculator(df).run_application_emission()

    # Write output
    with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as xl:
        df.to_excel(xl, sheet_name='Results', index=False)

    print(f"\nWROTE: {OUTPUT_PATH}")
    return df


if __name__ == '__main__':
    INPUT = '/Users/shuaij/Desktop/2905 DMS data - eigen kuilwaarden.xlsx'
    OUTPUT = '/Users/shuaij/Desktop/Output_DMS_Complete_eigen.xlsx'
    run_pipeline(INPUT, OUTPUT)