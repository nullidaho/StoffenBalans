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
    CF = 0.86 #Net to gross excretion
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

        # 1. Clean Column Headers
        df.columns = df.columns.str.strip()

        # 2. Force Numeric Conversion (Crucial for Mac/Excel locale issues)
        # Lists all columns that MUST be numbers
        numeric_targets = [
            'Nr_koe', 'Nr_pink', 'Nr_kalf', 'MilkYield', 'Fat%', 'Pro%', 'MilkUerum',
            'slurry_koe', 'solid_koe', 'slurry_kalf', 'solid_kalf', 'slurry_pink', 'solid_pink',
            'volume_slurry_koe', 'volume_solid_koe', 'volume_slurry_kalf', 'volume_solid_kalf',
            'volume_slurry_pink', 'volume_solid_pink', 'Vol_WholeMilk_Kalf', 'Vol_KunstMelk_Kalf',
            'slurry%_koe', 'slurry%_kalf', 'slurry%_pink', 'NatureGL%', 'Ha_Grass', 'Ha_Mais', 
            'Kg_conc', 'VEM_Concentrate', 'N_Concentrate',
            'GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'GD_Unlimited_Kalf', 'GD_Unlimited_Pink',
            'GH_Koe', 'GH_Kalf', 'GH_Pink'
        ]

        print("--- Cleaning Data (Fixing Comma/Dot issues) ---")
        for col in numeric_targets:
            if col in df.columns:
                # Convert to string -> replace comma with dot -> convert to float
                df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            else:
                # Set reasonable defaults for missing columns to prevent crashes
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
        df['net_n_kalf_slurry'] = df['Nr_kalf']*sc*df['slurry_kalf']
        df['net_n_kalf_solid'] = df['Nr_kalf']*fc*df['solid_kalf']
        df['net_n_calves'] = df['net_n_kalf_slurry']+df['net_n_kalf_solid']
        df['net_n_pink_slurry'] = df['Nr_pink']*sp*df['slurry_pink']
        df['net_n_pink_solid'] = df['Nr_pink']*fp*df['solid_pink']
        df['net_n_heifers'] = df['net_n_pink_slurry']+df['net_n_pink_solid']
        df['Total_Net_Nitrogen_kg'] = df['net_n_cows']+df['net_n_calves']+df['net_n_heifers']
        df['Total_Nitrogen_Excretion_MUN'] = df['Total_Net_Nitrogen_kg']/self.CF
        df['gross_n_cows_mun'] = df['net_n_cows']/self.CF
        df['gross_n_calves_mun'] = df['net_n_calves']/self.CF
        df['gross_n_heifers_mun'] = df['net_n_heifers']/self.CF
        print(f"  ✓ 1.1 done — Farm1 Gross N: {df['Total_Nitrogen_Excretion_MUN'].iloc[0]:.1f} kg")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1.2 — N Partitioning (MUN)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningCalculator:
    CF = 0.86
    def __init__(self, df): self.df = df
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 1.2: N Partitioning (MUN)\n{'='*70}")
        df = self.df
        df['gross_n_cows'] = df['net_n_cows']/self.CF
        df['gross_n_calves'] = df['net_n_calves']/self.CF
        df['gross_n_heifers'] = df['net_n_heifers']/self.CF
        df = ensure_numeric(df, 'MilkUerum', 0.0)
        df['MUN_value'] = df['MilkUerum']*(28.0/60.0)
        df['uun_per_cow_g_day'] = 16.7+13.0+12.03*df['MUN_value']
        df['uun_total_cows_kg'] = df['uun_per_cow_g_day']*365.0*df['Nr_koe']/1000.0
        df['fn_total_cows_kg'] = df['gross_n_cows']-df['uun_total_cows_kg']
        df['cow_uun_ratio'] = np.where(df['gross_n_cows']>0, df['uun_total_cows_kg']/df['gross_n_cows'], 0.0)
        df['uun_total_kalf_kg'] = df['gross_n_calves']*df['cow_uun_ratio']
        df['fn_total_kalf_kg'] = df['gross_n_calves']-df['uun_total_kalf_kg']
        df['uun_total_pink_kg'] = df['gross_n_heifers']*df['cow_uun_ratio']
        df['fn_total_pink_kg'] = df['gross_n_heifers']-df['uun_total_pink_kg']
        df['Farm_Total_UUN_kg'] = df['uun_total_cows_kg']+df['uun_total_kalf_kg']+df['uun_total_pink_kg']
        df['Farm_Total_FN_kg'] = df['fn_total_cows_kg']+df['fn_total_kalf_kg']+df['fn_total_pink_kg']
        print(f"  ✓ 1.2 done — Farm1 UUN: {df['Farm_Total_UUN_kg'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.1 — VEM Requirements
# ══════════════════════════════════════════════════════════════════════════════
class VEMRequirementCalculator:
    def __init__(self, df): self.df = df.copy()
    def _breed(self):
        self.df['avg_weight']=650.0; self.df['weight_factor']=1.0; self.df['breed_factor']=1.0
        if 'Breed' in self.df.columns:
            b = self.df['Breed'].astype(str).str.lower().str.strip()
            mj = b.str.contains('jersey', na=False)
            self.df.loc[mj,['avg_weight','weight_factor','breed_factor']] = [400.0, 400/650, 0.695]
            mc = b.str.contains('cross', na=False)
            self.df.loc[mc,['avg_weight','weight_factor','breed_factor']] = [525.0, 525/650, 0.852]
    def calculate_requirements(self):
        print(f"\n{'='*70}\nMODULE 2.1: VEM Requirements\n{'='*70}")
        for c in ['MilkYield','Fat%','Pro%','Nr_koe','Nr_pink','Nr_kalf',
                   'GD_Limited_Koe','GD_Combi_Koe','GD_Unlimited_Koe',
                   'GD_Unlimited_Kalf','GD_Unlimited_Pink']:
            self.df = ensure_numeric(self.df, c, 0.0)
        self._breed()
        L,D = 326.0, 39.0
        f,p,bf = self.df['Fat%'], self.df['Pro%'], self.df['breed_factor']
        # Kalf
        vkg = 1323.0*bf; vke = self.df['GD_Unlimited_Kalf']*0.346*bf
        self.df['vem_req_kalf_per_head_yr'] = (vkg+vke)*1.02
        self.df['Total_VEM_Kalf_Farm'] = self.df['vem_req_kalf_per_head_yr']*self.df['Nr_kalf']
        # Pink
        vpg = 2259.0*bf; vpe = self.df['GD_Unlimited_Pink']*0.784*bf; vpp = 115.9*bf
        self.df['vem_req_pink_per_head_yr'] = (vpg+vpe+vpp)*1.02
        self.df['Total_VEM_Pink_Farm'] = self.df['vem_req_pink_per_head_yr']*self.df['Nr_pink']
        # Cow
        fpcm_yr = (0.337+0.116*f+0.06*p)*self.df['MilkYield']*365.0
        fpcm_dl = fpcm_yr/L; cf = 1.0+(fpcm_dl-15.0)*0.00165
        self.df['vem_cow_milk_yr'] = (442.0*fpcm_dl*cf/1000.0)*L
        mw = np.power(self.df['avg_weight'], 0.75)
        vml = 42.4*mw*cf*L/1000.0; cd = 1.0+(-15.0*0.00165); vmd = 42.4*mw*cd*D/1000.0
        self.df['vem_cow_maint_yr'] = vml+vmd
        gdc = self.df['GD_Limited_Koe']*0.419+self.df['GD_Combi_Koe']*0.419+self.df['GD_Unlimited_Koe']*0.560
        self.df['vem_cow_exercise_yr'] = 201.0+gdc*(L/365.0)*bf
        self.df['vem_cow_youth_yr'] = 102.0*bf
        self.df['vem_cow_preg_yr'] = 194.0*bf
        vs = (self.df['vem_cow_milk_yr']+self.df['vem_cow_maint_yr']+
              self.df['vem_cow_exercise_yr']+self.df['vem_cow_youth_yr']+self.df['vem_cow_preg_yr'])
        self.df['vem_req_cow_per_head_yr'] = vs*1.02
        self.df['Total_VEM_Cow_Farm'] = self.df['vem_req_cow_per_head_yr']*self.df['Nr_koe']
        self.df['Total_VEM_Requirement_Farm_kVEM'] = (
            self.df['Total_VEM_Cow_Farm']+self.df['Total_VEM_Kalf_Farm']+self.df['Total_VEM_Pink_Farm'])
        print(f"  ✓ 2.1 done — Farm1 VEM: {self.df['Total_VEM_Requirement_Farm_kVEM'].iloc[0]:.0f}")
        return self.df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.2 — VEM Allocation (Corrected Fresh Grass Logic & English Comments)
# ══════════════════════════════════════════════════════════════════════════════
class VEMAllocationCalculator:
    VEM_KM=1500.0; VEM_CF=940.0; DS_DC=0.876
    LM=0.02; LK=0.02; LC=0.02; LR=0.05
    PGK=0.75; PMK=0.25; PGP=0.90; PMP=0.10
    
    def __init__(self, df): 
        self.df = df.copy()
    
    @staticmethod
    def _soil_p(st):
        s = str(st).lower().strip()
        if 'klei' in s:   
            yg, yf, ym, yn_gs, yn_fg = 10005, 8893, 17685, 6000, 5330
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 960.0, 940.0, 842.0, 860.0, 950.0
        elif 'zand' in s: 
            yg, yf, ym, yn_gs, yn_fg = 9360, 8320, 17773, 6000, 5333
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 960.0, 940.0, 842.0, 860.0, 950.0
        elif 'veen' in s: 
            yg, yf, ym, yn_gs, yn_fg = 9751, 8668, 16620, 6000, 5333
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 957.0, 937.0, 842.0, 860.0, 950.0 
        else: 
            yg, yf, ym, yn_gs, yn_fg = 9700, 9700, 17300, 6000, 6000
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 960.0, 940.0, 842.0, 860.0, 950.0
        
        return {'yield_gs':yg, 'yield_fg':yf, 'yield_ms':ym, 
                'yield_nat_gs':yn_gs, 'yield_nat_fg':yn_fg, 
                'vem_gs_cult':vc_gs, 'vem_gs_nat':vn_gs,
                'vem_fg_cult':vc_fg, 'vem_fg_nat':vn_fg,
                'vem_maize':vm}
                
    def _milk_vem(self):
        f,p = self.df['Fat%'], self.df['Pro%']
        GE = 744.38+365.7*f+241.4*p; ME = 584.17+376.6*f*0.94+171.5*p*0.87
        Q = np.where(GE>0, ME/GE*100.0, 0.0)
        return (0.6*(1.0+0.004*(Q-57.0))*0.9752*ME)/6.9
        
    def run_allocation(self):
        print(f"\n{'='*70}\nMODULE 2.2: VEM Allocation\n{'='*70}")
        
        # Default initialization to prevent missing column crashes.
        # Added Kg_freshcut and DM_density_fresh_grass (Default: 0.176 based on literature)
        defs = {'Vol_WholeMilk_Kalf':0, 'Vol_KunstMelk_Kalf':0, 'Kg_conc':0,
                'GH_Koe':6, 'GH_Kalf':6, 'GH_Pink':6, 'Ha_Grass':0, 'NatureGL%':0, 'Ha_Mais':0,
                'GD_Unlimited_Kalf':0, 'GD_Unlimited_Pink':0,
                'GD_Limited_Koe':0, 'GD_Combi_Koe':0, 'GD_Unlimited_Koe':0,
                'DS_OtherSilage':0, 'DS_Byproducts':0,
                'Kg_freshcut': 0, 'DM_density_fresh_grass': 0.176} 
        
        for c,d in defs.items(): 
            if c not in self.df.columns: self.df[c] = d
            self.df[c] = pd.to_numeric(self.df[c], errors='coerce').fillna(d)
        
        sp = self.df['Soil_Type'].apply(self._soil_p).tolist()
        
        for k in ('yield_gs','yield_fg','yield_ms','yield_nat_gs','yield_nat_fg',
                  'vem_gs_cult','vem_gs_nat','vem_fg_cult','vem_fg_nat','vem_maize'):
            self.df[k] = [d[k] for d in sp]
            
        pn = self.df['NatureGL%'].clip(0,100)
        
        # 1. Fresh Grass VEM (Weighted by Natural Grassland %)
        self.df['vem_fg_weighted'] = ((100-pn)*self.df['vem_fg_cult'] + pn*self.df['vem_fg_nat'])/100
        
        # 2. SMART VEM FALLBACKS
        vem_gs_soil = ((100-pn)*self.df['vem_gs_cult'] + pn*self.df['vem_gs_nat'])/100
        self.df['vem_gs_weighted'] = vem_gs_soil 
        
        def get_vem(col_name, default_series):
            val = pd.to_numeric(self.df.get(col_name, 0), errors='coerce').fillna(0)
            return np.where(val > 0, val, default_series)

        self.df['VEM_GS_Actual'] = get_vem('VEM_GrassSilage', vem_gs_soil)
        self.df['VEM_MS_Actual'] = get_vem('VEM_MaizeSilage', self.df['vem_maize'])
        self.df['VEM_OS_Actual'] = get_vem('VEM_OtherSilage', 950.0)
        self.df['VEM_BP_Actual'] = get_vem('VEM_Byproducts', 950.0) 
        self.df['VEM_Conc_Actual'] = get_vem('VEM_Concentrate', self.VEM_CF)
        
        # Yields (Weighted)
        self.df['yield_fg_weighted'] = ((100-pn)*self.df['yield_fg'] + pn*self.df['yield_nat_fg'])/100
        self.df['yield_gs_weighted'] = ((100-pn)*self.df['yield_gs'] + pn*self.df['yield_nat_gs'])/100
        
        # ----------------------------------------------------------------------
        # Young stock liquid + grazing + conc
        # ----------------------------------------------------------------------
        vmd = self._milk_vem()
        self.df['kVEM_Intake_Milk_Kalf'] = self.df['Vol_WholeMilk_Kalf']*(1-self.LM)*vmd/1000
        self.df['kVEM_Intake_KunstMelk_Kalf'] = self.df['Vol_KunstMelk_Kalf']*(1-self.LK)*self.VEM_KM/1000
        rgk = (self.df['GD_Unlimited_Kalf']/365).clip(0,1)
        rgp = (self.df['GD_Unlimited_Pink']/365).clip(0,1)
        self.df['kVEM_Intake_Conc_Kalf'] = self.df['Total_VEM_Kalf_Farm']*(0.10*rgk+0.25*(1-rgk))
        self.df['kVEM_Intake_Conc_Pink'] = self.df['Total_VEM_Pink_Farm']*(0.00*rgp+0.05*(1-rgp))
        tk = rgk*(1323.0-101.2)+self.df['GD_Unlimited_Kalf']*0.346
        self.df['kVEM_Intake_FreshGrass_Kalf'] = self.df['Nr_kalf']*tk*0.9*self.df['breed_factor']*1.02
        tp = rgp*(2259.0+102.9)+self.df['GD_Unlimited_Pink']*0.784
        self.df['kVEM_Intake_FreshGrass_Pink'] = self.df['Nr_pink']*tp*self.df['breed_factor']*1.02
        
        # Young stock Roughage Balance
        rk = (self.df['Total_VEM_Kalf_Farm']-self.df['kVEM_Intake_Milk_Kalf']-self.df['kVEM_Intake_KunstMelk_Kalf']-self.df['kVEM_Intake_Conc_Kalf']-self.df['kVEM_Intake_FreshGrass_Kalf']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Kalf'] = rk*self.PGK
        self.df['kVEM_Intake_MaizeSilage_Kalf'] = rk*self.PMK
        rp = (self.df['Total_VEM_Pink_Farm']-self.df['kVEM_Intake_Conc_Pink']-self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Pink'] = rp*self.PGP
        self.df['kVEM_Intake_MaizeSilage_Pink'] = rp*self.PMP
        
        # ----------------------------------------------------------------------
        # Cow Supply (Corrected to assign Fresh Grass strictly by BEX physical formula)
        # ----------------------------------------------------------------------
        hrs = self.df['GH_Koe']
        ir = np.where(hrs>2, 2.0+0.75*(hrs-2), hrs)
        tgd = self.df['GD_Limited_Koe']+self.df['GD_Combi_Koe']+self.df['GD_Unlimited_Koe']
        
        # FPCM and correction factors
        fy = (0.337+0.116*self.df['Fat%']+0.06*self.df['Pro%'])*self.df['MilkYield']*365
        mf = 1.0+((fy-9500*self.df['breed_factor'])/500)*0.02
        dc = (365-39)/365
        
        # Direct calculation of physical grazing intake
        cgv = tgd*ir*(self.df['vem_fg_weighted']/1000)*dc*mf*self.df['breed_factor']*self.df['Nr_koe']
        
        # Assign directly - bypassing proportional logic
        self.df['kVEM_Intake_FreshGrass_Cow'] = cgv
        
        # Calculate energy from Byproducts and Other Silage 
        ds_os = self.df['DS_OtherSilage']
        ds_bp = self.df['DS_Byproducts']
        kvem_os_cow = ds_os * self.df['VEM_OS_Actual'] / 1000.0
        kvem_bp_cow = ds_bp * self.df['VEM_BP_Actual'] / 1000.0
        
        # Save them for fallback referencing
        self.df['kVEM_Intake_OtherSilage_Cow'] = kvem_os_cow
        self.df['kVEM_Intake_Byproducts_Cow'] = kvem_bp_cow
        
        # Cow Concentrate Allocation (Residual after young stock)
        vc = self.df['VEM_Conc_Actual'] 
        dsc = self.df['Kg_conc'] * self.DS_DC
        mkc = self.df['kVEM_Intake_Conc_Kalf']*1000/vc/(1-self.LC)
        mpc = self.df['kVEM_Intake_Conc_Pink']*1000/vc/(1-self.LC)
        mcc = (dsc-mkc-mpc).clip(lower=0)
        
        self.df['kVEM_Intake_Conc_Cow'] = mcc*(1-self.LC)*vc/1000
        
        print(f"  ✓ 2.2 done")
        return self.df
    
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.3 — N Intake (Hybrid: Records + Sluitpost + Cut Fresh Grass)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenIntakeCalculator:
    FP=6.25; FD=6.38; RAS=40.0; NK=32.52; DSW=0.228; DSK=0.964
    DS_DC=0.876
    VC_DEF=940.0; VB_DEF=950.0; VO_DEF=950.0
    NC_DEF=27.3;  NB_DEF=25.0;  NO_DEF=25.0

    def __init__(self, df): self.df = df.copy()
    
    def _safe_get(self, col):
        if col in self.df.columns:
            return pd.to_numeric(self.df[col], errors='coerce').fillna(0.0)
        return pd.Series(0.0, index=self.df.index)
    
    def _soil_N(self, st):
        s = str(st).lower().strip()
        if 'klei' in s:   return {'ng':28.2,'nf':31.6,'nm':12.0,'nn_gs':24.7,'nn_fg':30.2} 
        elif 'zand' in s: return {'ng':28.6,'nf':32.0,'nm':12.0,'nn_gs':25.1,'nn_fg':30.2}
        elif 'veen' in s: return {'ng':27.8,'nf':31.1,'nm':12.0,'nn_gs':24.5,'nn_fg':30.2}
        else:             return {'ng':28.2,'nf':31.6,'nm':12.0,'nn_gs':24.7,'nn_fg':30.2}
        
    def _vre(self, ds, cpg, ft):
        cp = cpg; cps = np.where(cp>0, cp, 1.0)
        if ft=='grass_fresh': v = (0.931*cp-43.2)/cps
        elif ft=='grass_silage': v = (0.963*cp-38.3)/cps
        elif ft=='maize_silage': v = (0.969*cp+0.04*self.RAS-40.0)/cps
        elif ft in ('concentrate', 'byproduct', 'other_silage'): 
            v = (88.7*(1.0-np.exp(-0.012*cp)))/100.0
        elif ft=='dairy':
            v = pd.Series(0.89, index=ds.index) if hasattr(ds, 'index') else 0.89
        else: v = 0.0
        v = np.clip(v, 0, 1)
        return ds*(v*cp)/1000.0
        
    def run_nutrient_calculation(self):
        print(f"\n{'='*70}\nMODULE 2.3: N Intake (Sluitpost Balancing & Cut Fresh Grass)\n{'='*70}")
        sn = self.df['Soil_Type'].apply(self._soil_N).tolist()
        
        n_nf = np.array([d['nf'] for d in sn]); n_ng = np.array([d['ng'] for d in sn])
        n_nm = np.array([d['nm'] for d in sn])
        n_nn_fg = np.array([d['nn_fg'] for d in sn]); n_nn_gs = np.array([d['nn_gs'] for d in sn])
        
        pn = self.df['NatureGL%'].clip(0,100)
        nf_w = ((100-pn)*n_nf + pn*n_nn_fg)/100
        ng_w = ((100-pn)*n_ng + pn*n_nn_gs)/100

        def _get_v(col, fb):
            v = self._safe_get(col)
            return pd.Series(np.where(v>0, v, fb), index=self.df.index)
            
        def _get_n(r_col, n_col, fb):
            r = self._safe_get(r_col)
            n = self._safe_get(n_col)
            return pd.Series(np.where(r>0, r/self.FP, np.where(n>0, n, fb)), index=self.df.index)

        v_fg = _get_v('VEM_fgrass', self.df['vem_fg_weighted']) 
        v_gs = _get_v('VEM_GrassSilage', self.df['vem_gs_weighted'])
        v_ms = _get_v('VEM_MaizeSilage', self.df['vem_maize'])
        v_cc = _get_v('VEM_Concentrate', self.VC_DEF)
        v_bp = _get_v('VEM_Byproducts', self.VB_DEF)
        v_os = _get_v('VEM_OtherSilage', self.VO_DEF)
        
        n_fg = _get_n('RE_fgrass', 'N_fgrass', nf_w)
        n_gs = _get_n('RE_GrassSilage', 'N_GrassSilage', ng_w)
        n_ms = _get_n('RE_MaizeSilage', 'N_MaizeSilage', n_nm)
        n_cc = _get_n('RE_Concentrate', 'N_Concentrate', self.NC_DEF)
        n_bp = _get_n('RE_Byproducts', 'N_Byproducts', self.NB_DEF)
        n_os = _get_n('RE_OtherSilage', 'N_OtherSilage', self.NO_DEF)
        n_milk_ds = (self.df['Pro%']*10.0/self.DSW)/self.FD
        
        # ----------------------------------------------------------------------------------
        # 1. Physical Grazing Fresh Grass (Strictly based on Grazing Hours)
        # ----------------------------------------------------------------------------------
        self.df['DS_fresh_cow'] = (self.df['kVEM_Intake_FreshGrass_Cow'] * 1000 / v_fg).fillna(0)
        self.df['DS_fresh_pink'] = (self.df['kVEM_Intake_FreshGrass_Pink'] * 1000 / v_fg).fillna(0)
        self.df['DS_fresh_kalf'] = (self.df['kVEM_Intake_FreshGrass_Kalf'] * 1000 / v_fg).fillna(0)
        
        self.df['kVEM_Intake_FreshGrass_Cow'] = self.df['DS_fresh_cow'] * v_fg / 1000.0
        self.df['kVEM_Intake_FreshGrass_Pink'] = self.df['DS_fresh_pink'] * v_fg / 1000.0
        self.df['kVEM_Intake_FreshGrass_Kalf'] = self.df['DS_fresh_kalf'] * v_fg / 1000.0

        # ----------------------------------------------------------------------------------
        # 2. NEW: Cut Fresh Grass fed in barn (Zomerstalvoedering)
        # ----------------------------------------------------------------------------------
        kg_cut_grass = self._safe_get('Kg_freshcut')
        
        # Use provided DM density, fallback to 0.176 (literature mean: 176 g/kg) if empty
        dm_density_fg = self._safe_get('DM_density_fresh_grass')
        dm_density_fg = pd.Series(np.where(dm_density_fg > 0, dm_density_fg, 0.176), index=self.df.index)
        
        # Total Farm Dry Matter (DS) of cut grass
        fds_cut_grass = kg_cut_grass * dm_density_fg
        
        # Energy provided by the cut grass
        kvem_cut_grass_tot = fds_cut_grass * v_fg / 1000.0

        # ----------------------------------------------------------------------------------
        # 3. Liquids (Calves)
        # ----------------------------------------------------------------------------------
        self.df['DS_milk_kalf'] = self.df['Vol_WholeMilk_Kalf']*self.DSW
        self.df['DS_kunst_kalf'] = self.df['Vol_KunstMelk_Kalf']*self.DSK
        kvem_milk_tot = self.df['kVEM_Intake_Milk_Kalf'] + self.df['kVEM_Intake_KunstMelk_Kalf']
        
        # ----------------------------------------------------------------------------------
        # 4. OS and BP (100% assigned to Cows)
        # ----------------------------------------------------------------------------------
        fds_os = self._safe_get('DS_OtherSilage')
        fds_bp = self._safe_get('DS_Byproducts')
        self.df['DS_other_cow'] = fds_os
        self.df['DS_byprod_cow'] = fds_bp
        kvem_os_bp_cow = (fds_os * v_os / 1000.0) + (fds_bp * v_bp / 1000.0)

        # ----------------------------------------------------------------------------------
        # 5. Concentrate and Maize (Input vs Fallback)
        # ----------------------------------------------------------------------------------
        fds_cc = self._safe_get('DS_Concentrate')
        kg_cc = self._safe_get('Kg_conc')
        fds_cc = np.where(fds_cc > 0, fds_cc, kg_cc * self.DS_DC)
        
        fds_ms = self._safe_get('DS_MaizeSilage')
        fallback_ms = self.df['Ha_Mais'] * self.df['yield_ms'] # Maize Fallback follows Land Yield
        fds_ms = pd.Series(np.where(fds_ms > 0, fds_ms, fallback_ms), index=self.df.index)

        # ----------------------------------------------------------------------------------
        # 6. GRASS SILAGE FALLBACK: THE SLUITPOST (Ultimate Balancing Feed)
        # ----------------------------------------------------------------------------------
        total_vem_req = self.df['Total_VEM_Cow_Farm'] + self.df['Total_VEM_Pink_Farm'] + self.df['Total_VEM_Kalf_Farm']
        kvem_fg_tot = self.df['kVEM_Intake_FreshGrass_Cow'] + self.df['kVEM_Intake_FreshGrass_Pink'] + self.df['kVEM_Intake_FreshGrass_Kalf']
        kvem_cc_tot = fds_cc * v_cc / 1000.0
        kvem_ms_tot = fds_ms * v_ms / 1000.0
        
        # Energy Deficit = Total Req - (Grazing + Milk + BP/OS + Conc + Maize + Cut Grass)
        # Notice we are now deducting kvem_cut_grass_tot to correctly balance the energy!
        kvem_gs_needed = (total_vem_req - kvem_fg_tot - kvem_milk_tot - kvem_os_bp_cow - kvem_cc_tot - kvem_ms_tot - kvem_cut_grass_tot).clip(lower=0)
        fallback_gs = kvem_gs_needed * 1000.0 / v_gs.replace(0, np.nan)
        
        fds_gs = self._safe_get('DS_GrassSilage')
        fds_gs = pd.Series(np.where(fds_gs > 0, fds_gs, fallback_gs.fillna(0)), index=self.df.index)

        # ----------------------------------------------------------------------------------
        # 7. Barn Energy Deficit Ratios (Determines how shared barn feeds are split)
        # ----------------------------------------------------------------------------------
        # Cut grass is treated as a shared barn feed. Therefore, we calculate the deficit
        # BEFORE distributing Cut Grass, Maize, Concentrate, and Grass Silage.
        rcow = (self.df['Total_VEM_Cow_Farm'] - self.df['kVEM_Intake_FreshGrass_Cow'] - kvem_os_bp_cow).clip(lower=0)
        rpnk = (self.df['Total_VEM_Pink_Farm'] - self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        rklf = (self.df['Total_VEM_Kalf_Farm'] - self.df['kVEM_Intake_FreshGrass_Kalf'] - self.df['kVEM_Intake_Milk_Kalf'] - self.df['kVEM_Intake_KunstMelk_Kalf']).clip(lower=0)
        treq = rcow+rpnk+rklf
        
        pcow = np.divide(rcow, treq, out=np.zeros_like(rcow), where=treq!=0)
        ppnk = np.divide(rpnk, treq, out=np.zeros_like(rpnk), where=treq!=0)
        pklf = np.divide(rklf, treq, out=np.zeros_like(rklf), where=treq!=0)

        # ----------------------------------------------------------------------------------
        # 8. Apply Proportions to Shared Barn Feeds
        # ----------------------------------------------------------------------------------
        for sx, r in [('cow', pcow), ('pink', ppnk), ('kalf', pklf)]:
            # Distribute shared feeds according to relative stomach capacity
            self.df[f'DS_cut_grass_{sx}'] = fds_cut_grass * r
            self.df[f'DS_gs_{sx}'] = fds_gs * r
            self.df[f'DS_ms_{sx}'] = fds_ms * r
            self.df[f'DS_conc_{sx}'] = fds_cc * r
            
            # Combine Grazing Fresh Grass and Cut Fresh Grass into a total
            self.df[f'DS_fresh_total_{sx}'] = self.df[f'DS_fresh_{sx}'] + self.df[f'DS_cut_grass_{sx}']

            sfx_cap = sx.capitalize()
            self.df[f'kVEM_Intake_GrassSilage_{sfx_cap}'] = self.df[f'DS_gs_{sx}'] * v_gs / 1000.0
            self.df[f'kVEM_Intake_MaizeSilage_{sfx_cap}'] = self.df[f'DS_ms_{sx}'] * v_ms / 1000.0
            self.df[f'kVEM_Intake_Conc_{sfx_cap}'] = self.df[f'DS_conc_{sx}'] * v_cc / 1000.0
            self.df[f'kVEM_Intake_CutGrass_{sfx_cap}'] = self.df[f'DS_cut_grass_{sx}'] * v_fg / 1000.0
          
        # ----------------------------------------------------------------------------------
        # 9. Calculate Nutrients (N, CP, VRE)
        # ----------------------------------------------------------------------------------
        for sfx in ('cow','pink','kalf'):
            lbl = sfx.capitalize()
            # Note: We now use DS_fresh_total_{sfx} which includes both grazed and cut grass
            feeds = [(f'DS_fresh_total_{sfx}', n_fg, 'grass_fresh', False),
                     (f'DS_gs_{sfx}',          n_gs, 'grass_silage', False), 
                     (f'DS_ms_{sfx}',          n_ms, 'maize_silage', False),
                     (f'DS_conc_{sfx}',        n_cc, 'concentrate', False)]
            
            if sfx == 'cow':
                feeds.append(('DS_other_cow', n_os, 'other_silage', False))
                feeds.append(('DS_byprod_cow', n_bp, 'byproduct', False))
                
            if sfx == 'kalf':
                feeds.append(('DS_milk_kalf', n_milk_ds, 'dairy', True))
                feeds.append(('DS_kunst_kalf', self.NK, 'dairy', True))
                
            tn=pd.Series(0.0,index=self.df.index); tcp=tn.copy(); tvre=tn.copy()
            for dc, nc, ft, dairy in feeds:
                ds = self.df[dc]
                nf = self.FD if dairy else self.FP
                nk = ds*nc/1000; cpk = nk*nf; cpc = nc*nf
                vk = self._vre(ds, cpc, ft)
                tn+=nk; tcp+=cpk; tvre+=vk
                
            self.df[f'Total_N_Intake_{lbl}'] = tn
            self.df[f'Total_CP_Intake_{lbl}'] = tcp
            self.df[f'Total_VRE_Intake_{lbl}'] = tvre
            
        for l in ('Cow','Pink','Kalf'):
            self.df[f'VCRE_Factor_{l}'] = (self.df[f'Total_VRE_Intake_{l}']/self.df[f'Total_CP_Intake_{l}'].replace(0,np.nan)).fillna(0)
            
        tvf = sum(self.df[f'Total_VRE_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        tcf = sum(self.df[f'Total_CP_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        self.df['VCRE_Factor_Farm'] = (tvf/tcf.replace(0,np.nan)).fillna(0)
        self.df['Total_N_Intake_Farm'] = sum(self.df[f'Total_N_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        
        print(f"  ✓ 2.3 done — Farm1 N intake: {self.df['Total_N_Intake_Farm'].iloc[0]:.1f} kg")
        return self.df
          
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.4 — N Excretion (VCRE)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenExcretionCalculatorVCRE:
    NK=29.4; NP=24.1; NV=23.1; NC=22.5
    RB=44/650; R1=320/650; RC=540/650
    BR=0.70; RR=0.27; FD=6.38
    def __init__(self, df): self.df = df.copy()
    def calculate_excretion(self):
        print(f"\n{'='*70}\nMODULE 2.4: N Excretion (VCRE)\n{'='*70}")
        df = self.df; wc = df['avg_weight']
        wb=wc*self.RB; wp=wc*self.R1; wv=wc*self.RC
        nb=wb*self.NK/1000; npk=wp*self.NP/1000; nv=wv*self.NV/1000; nc=wc*self.NC/1000
        ym = np.where(df['MilkYield']<100, 365, 1)
        tm = df['MilkYield']*ym*df['Nr_koe']
        rm = tm*df['Pro%']*10/self.FD/1000
        rf = nb*self.BR*df['Nr_koe']
        nih = self.RR*nv*df['Nr_koe']; noc = self.RR*nc*df['Nr_koe']
        df['N_Retention_Cow_VCRE'] = rm+rf+(noc-nih)
        gt = npk-nb; g1 = 0.36*df['breed_factor']
        t1 = gt*(0.376/0.407); t2 = (g1/2*24)*(0.031/0.407)
        ncr = np.divide(t1+t2, gt, out=np.ones_like(gt.values,dtype=float), where=gt.values!=0)
        df['N_Retention_Kalf_VCRE'] = gt*df['Nr_kalf']*ncr
        gha = (nv-npk)*(12/14)
        df['N_Retention_Pink_VCRE'] = (nb+gha)*df['Nr_pink']
        df['Total_N_Retention_Farm_VCRE'] = df['N_Retention_Cow_VCRE']+df['N_Retention_Kalf_VCRE']+df['N_Retention_Pink_VCRE']
        df['Total_N_Excretion_Cow_VCRE'] = df['Total_N_Intake_Cow']-df['N_Retention_Cow_VCRE']
        df['Total_N_Excretion_Kalf_VCRE'] = df['Total_N_Intake_Kalf']-df['N_Retention_Kalf_VCRE']
        df['Total_N_Excretion_Pink_VCRE'] = df['Total_N_Intake_Pink']-df['N_Retention_Pink_VCRE']
        df['Total_N_Excretion_Farm_VCRE'] = df['Total_N_Excretion_Cow_VCRE']+df['Total_N_Excretion_Kalf_VCRE']+df['Total_N_Excretion_Pink_VCRE']
        print(f"  ✓ 2.4 done — Farm1 excr: {df['Total_N_Excretion_Farm_VCRE'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.5 — N Partitioning (VCRE)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningVCRE:
    DCF = 0.91  #converting to dairy cow
    def __init__(self, df): self.df = df.copy()
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 2.5: N Partitioning (VCRE)\n{'='*70}")
        df = self.df
        for co,lb in [('cow','Cow'),('pink','Pink'),('kalf','Kalf')]:
            dn = df[f'Total_N_Intake_{lb}']*df[f'VCRE_Factor_{lb}']*self.DCF
            uun = dn - df[f'N_Retention_{lb}_VCRE']
            fn = df[f'Total_N_Excretion_{lb}_VCRE'] - uun
            df[f'UUN_{lb}_VCRE'] = uun; df[f'FN_{lb}_VCRE'] = fn
        df['Total_UUN_Farm_VCRE'] = df['UUN_Cow_VCRE']+df['UUN_Kalf_VCRE']+df['UUN_Pink_VCRE']
        df['Total_FN_Farm_VCRE'] = df['FN_Cow_VCRE']+df['FN_Kalf_VCRE']+df['FN_Pink_VCRE']
        print(f"  ✓ 2.5 done — Farm1 VCRE UUN: {df['Total_UUN_Farm_VCRE'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 3.1 — Mineralization
# ══════════════════════════════════════════════════════════════════════════════
class MineralizationCalculator:
    RS=0.10; RV=-0.25
    def __init__(self, df): self.df = df.copy()
    def calculate_mineralization(self):
        print(f"\n{'='*70}\nMODULE 3.1: Mineralization\n{'='*70}")
        df = self.df
        for an,fnc,uunc,sl,sfx in [
            ('Cow','fn_total_cows_kg','uun_total_cows_kg','slurry%_koe','Cow'),
            ('Kalf','fn_total_kalf_kg','uun_total_kalf_kg','slurry%_kalf','Kalf'),
            ('Pink','fn_total_pink_kg','uun_total_pink_kg','slurry%_pink','Pink')]:
            ns = df[fnc]*df[sl]*self.RS; nv = df[uunc]*(1-df[sl])*self.RV
            df[f'Net_Min_{sfx}_MUN'] = ns+nv
        df['Total_Net_Mineralization_MUN'] = df['Net_Min_Cow_MUN']+df['Net_Min_Kalf_MUN']+df['Net_Min_Pink_MUN']
        for an,fnc,uunc,sl,sfx in [
            ('Cow','FN_Cow_VCRE','UUN_Cow_VCRE','slurry%_koe','Cow'),
            ('Kalf','FN_Kalf_VCRE','UUN_Kalf_VCRE','slurry%_kalf','Kalf'),
            ('Pink','FN_Pink_VCRE','UUN_Pink_VCRE','slurry%_pink','Pink')]:
            ns = df[fnc]*df[sl]*self.RS; nv = df[uunc]*(1-df[sl])*self.RV
            df[f'Net_Min_{sfx}_VCRE'] = ns+nv
        df['Total_Net_Mineralization_VCRE'] = df['Net_Min_Cow_VCRE']+df['Net_Min_Kalf_VCRE']+df['Net_Min_Pink_VCRE']
        df['Mineralization_Diff'] = df['Total_Net_Mineralization_VCRE']-df['Total_Net_Mineralization_MUN']
        print(f"  ✓ 3.1 done — MUN: {df['Total_Net_Mineralization_MUN'].iloc[0]:.2f}, VCRE: {df['Total_Net_Mineralization_VCRE'].iloc[0]:.2f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 3.2 — Corrected TAN
# ══════════════════════════════════════════════════════════════════════════════
class CorrectedTANCalculator:
    def __init__(self, df): self.df = df.copy()
    def calculate_corrected_tan(self):
        print(f"\n{'='*70}\nMODULE 3.2: Corrected TAN\n{'='*70}")
        df = self.df
        for m,us,ms in [
            ('MUN',['uun_total_cows_kg','uun_total_kalf_kg','uun_total_pink_kg'],
                   ['Net_Min_Cow_MUN','Net_Min_Kalf_MUN','Net_Min_Pink_MUN']),
            ('VCRE',['UUN_Cow_VCRE','UUN_Kalf_VCRE','UUN_Pink_VCRE'],
                    ['Net_Min_Cow_VCRE','Net_Min_Kalf_VCRE','Net_Min_Pink_VCRE'])]:
            for an,uc,mc in zip(['Cow','Kalf','Pink'],us,ms):
                df[f'Corrected_TAN_{an}_{m}'] = df[uc]+df[mc]
            df[f'Total_Corrected_TAN_{m}'] = df[f'Corrected_TAN_Cow_{m}']+df[f'Corrected_TAN_Kalf_{m}']+df[f'Corrected_TAN_Pink_{m}']
        print(f"  ✓ 3.2 done — TAN MUN: {df['Total_Corrected_TAN_MUN'].iloc[0]:.1f}, VCRE: {df['Total_Corrected_TAN_VCRE'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.1 — Emissions (Stable / Storage / Grazing)
# ══════════════════════════════════════════════════════════════════════════════
class EmissionCalculator:
    EFS=0.143; EFG=0.04; EFST=0.01; EFSO=0.024; EFVO=0.035; MSO=0.20; FN2N=14/17
    def __init__(self, df): self.df = df.copy()
    def calculate_emissions(self):
        print(f"\n{'='*70}\nMODULE 4.1: Emissions\n{'='*70}")
        df = self.df
        for c in ['GD_Limited_Koe','GD_Combi_Koe','GD_Unlimited_Koe','GH_Koe',
                   'GD_Unlimited_Pink','GH_Pink','GD_Unlimited_Kalf','GH_Kalf',
                   'slurry%_koe','slurry%_kalf','slurry%_pink']:
            df = ensure_numeric(df, c, 0.0)
        df['total_gd_cow'] = df['GD_Limited_Koe']+df['GD_Combi_Koe']+df['GD_Unlimited_Koe']
        cats = [
            {'n':'Cow','gd':'total_gd_cow','gh':'GH_Koe','sl':'slurry%_koe',
             'gm':'gross_n_cows_mun','um':'uun_total_cows_kg','tm':'Corrected_TAN_Cow_MUN',
             'gv':'Total_N_Excretion_Cow_VCRE','uv':'UUN_Cow_VCRE','tv':'Corrected_TAN_Cow_VCRE'},
            {'n':'Pink','gd':'GD_Unlimited_Pink','gh':'GH_Pink','sl':'slurry%_pink',
             'gm':'gross_n_heifers_mun','um':'uun_total_pink_kg','tm':'Corrected_TAN_Pink_MUN',
             'gv':'Total_N_Excretion_Pink_VCRE','uv':'UUN_Pink_VCRE','tv':'Corrected_TAN_Pink_VCRE'},
            {'n':'Kalf','gd':'GD_Unlimited_Kalf','gh':'GH_Kalf','sl':'slurry%_kalf',
             'gm':'gross_n_calves_mun','um':'uun_total_kalf_kg','tm':'Corrected_TAN_Kalf_MUN',
             'gv':'Total_N_Excretion_Kalf_VCRE','uv':'UUN_Kalf_VCRE','tv':'Corrected_TAN_Kalf_VCRE'}]
        for ct in cats:
            nm = ct['n']; gd = df[ct['gd']].fillna(0); gh = df[ct['gh']].fillna(0)
            sp = df[ct['sl']].fillna(1)
            gr = (0.0261*gh*(gd/365)).clip(0,1)
            epi = (1-(gd*gh)/(365*24)).clip(0,1)
            for method in ['mun','vcre']:
                M = 'MUN' if method=='mun' else 'VCRE'
                gk = 'gm' if method=='mun' else 'gv'
                uk = 'um' if method=='mun' else 'uv'
                tk = 'tm' if method=='mun' else 'tv'
                gn = df[ct[gk]].fillna(0); uun = df[ct[uk]].fillna(0); tan = df[ct[tk]].fillna(0)
                se = (tan*self.EFS*(1-gr))/self.FN2N
                df[f'Emission_Stable_{nm}_{M}'] = se
                ni = gn*epi
                no = ni*(sp*self.EFSO+(1-sp)*self.EFVO)
                nts = ((ni-no)/self.FN2N-se).clip(lower=0)
                ste = nts*self.MSO*self.EFST
                df[f'Emission_Storage_{nm}_{M}'] = ste
                uo = uun*(1-epi)
                ge = uo*self.EFG/self.FN2N
                df[f'Emission_Grazing_{nm}_{M}'] = ge
        for M in ['MUN','VCRE']:
            for tp in ['Stable','Storage','Grazing']:
                df[f'Total_Emission_{tp}_{M}'] = sum(df[f'Emission_{tp}_{n}_{M}'] for n in ['Cow','Pink','Kalf'])
            df[f'Total_Emission_All_{M}'] = df[f'Total_Emission_Stable_{M}']+df[f'Total_Emission_Storage_{M}']+df[f'Total_Emission_Grazing_{M}']
        print(f"  ✓ 4.1 done — Total MUN: {df['Total_Emission_All_MUN'].iloc[0]:.2f}, VCRE: {df['Total_Emission_All_VCRE'].iloc[0]:.2f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.2 — Land Application (VCRE-based, RVO 2025 rules)
# ══════════════════════════════════════════════════════════════════════════════
#
# Key logic per your instruction:
#   "remaining TAN from total manure excretion is usually MORE than the
#    usage space → so N applied = usage space (the limit is binding)"
#
# Steps per farm:
#   1. Determine manure N limit (170/190/200 kg N/ha, based on derogation)
#   2. Determine total N usage norms per land parcel (Tabel 2)
#   3. Manure N available = N excretion (VCRE) after emission losses
#   4. Manure N applied = min(available, limit × ha)   [limit is binding]
#   5. Effective manure N = applied × working coefficient (Tabel 9)
#   6. Fertiliser N space = total N norm − effective manure N
# ══════════════════════════════════════════════════════════════════════════════

class LandApplicationCalculator:
    """
    Computes:
      - Total N usage space (kg N) based on land areas and crop norms
      - Manure N application limit (kg N) based on derogation/NV area
      - Available manure N (VCRE excretion minus NH3-N losses)
      - Applied manure N = min(available, manure_limit)
      - Effective manure N for N-space = applied * working_coefficient
      - Fertiliser N space = usage_space - effective_manure_N

    Expected columns (from earlier modules / input):
      Soil_Type, Region, Derogation, NV_Area, CropType, Ha_Crop, Ha_Grass, Ha_Mais,
      GD_* and slurry%_koe (for working coefficient), and emissions columns from Module 5.
    """

    # ── Tabel 2: N use norms (kg N/ha/yr) ─────────────────────────────────────
    TABEL_2 = {
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

    # ── "Tabel 9" working coefficients (apparent N effectiveness) ─────────────
    WC_SLURRY_GRAZING = 0.45
    WC_SLURRY_NOGRAZING = 0.60
    WC_SOLID_GRAZING = 0.45
    WC_SOLID_NOGRAZING = 0.60

    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self._ensure_columns()

    def _ensure_columns(self):
        """Make sure required columns exist; coerce numerics."""
        # Normalize HA_Crop capitalization if needed
        if 'HA_Crop' in self.df.columns and 'Ha_Crop' not in self.df.columns:
            self.df.rename(columns={'HA_Crop': 'Ha_Crop'}, inplace=True)

        # Defaults
        if 'CropType' not in self.df.columns:
            self.df['CropType'] = 'overig'

        for c in ['Ha_Grass', 'Ha_Mais', 'Ha_Crop', 'NatureGL%']:
            self.df = ensure_numeric(self.df, c, 0.0)

        for c in ['GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'slurry%_koe', 'Derogation', 'NV_Area']:
            if c in ['Derogation', 'NV_Area']:
                if c not in self.df.columns:
                    self.df[c] = 'No'
            else:
                self.df = ensure_numeric(self.df, c, 0.0)

        # slurry% fraction handling
        if self.df['slurry%_koe'].max() > 1.0:
            self.df['slurry%_koe'] = (self.df['slurry%_koe'] / 100.0).clip(0.0, 1.0)
        else:
            self.df['slurry%_koe'] = self.df['slurry%_koe'].clip(0.0, 1.0)

        if 'Soil_Type' not in self.df.columns:
            self.df['Soil_Type'] = 'zand'
        if 'Region' not in self.df.columns:
            self.df['Region'] = 'Others'

    @staticmethod
    def _yes(x) -> bool:
        return str(x).strip().lower() in ('yes', 'ja', 'true', '1', 'y')

    @staticmethod
    def _soil_key(soil_type, region):
        s = str(soil_type).lower()
        r = str(region).lower()
        if 'klei' in s:
            return 'klei'
        if 'veen' in s:
            return 'veen'
        if any(t in s for t in ('loss', 'loes', 'löss')):
            return 'loss'
        if 'zand' in s:
            return 'zand_zuid' if ('zuid' in r or 'south' in r) else 'zand_nwc'
        return 'zand_nwc'

    def _crop_norm(self, crop_key, soil_key) -> float:
        ck = str(crop_key).strip().lower().replace('ï', 'i')
        entry = self.TABEL_2.get(ck, self.TABEL_2['overig'])
        return float(entry.get(soil_key, 185.0))

    def _grass_key(self, row) -> str:
        gd = float(row.get('GD_Limited_Koe', 0) or 0) + float(row.get('GD_Combi_Koe', 0) or 0) + float(row.get('GD_Unlimited_Koe', 0) or 0)
        return 'grasland_beweiden' if gd > 0 else 'grasland_maaien'

    def _maize_key(self, row) -> str:
        derog = self._yes(row.get('Derogation', 'No'))
        return 'mais_derogatie' if derog else 'mais_geen_derogatie'

    def _manure_limit_per_ha(self, row) -> float:
        derog = self._yes(row.get('Derogation', 'No'))
        nv = self._yes(row.get('NV_Area', 'No'))
        if derog:
            return 190.0 if nv else 200.0
        return 170.0

    def _is_grazing(self, row) -> bool:
        gd = float(row.get('GD_Limited_Koe', 0) or 0) + float(row.get('GD_Combi_Koe', 0) or 0) + float(row.get('GD_Unlimited_Koe', 0) or 0)
        return gd > 0

    def _working_coefficient(self, row) -> float:
        slurry_frac = float(row.get('slurry%_koe', 1.0) or 1.0)
        solid_frac = 1.0 - slurry_frac
        if self._is_grazing(row):
            return slurry_frac * self.WC_SLURRY_GRAZING + solid_frac * self.WC_SOLID_GRAZING
        return slurry_frac * self.WC_SLURRY_NOGRAZING + solid_frac * self.WC_SOLID_NOGRAZING

    def calculate_land_application(self) -> pd.DataFrame:
        print(f"\n{'='*70}\nMODULE 4.2: Land Application (VCRE-based)\n{'='*70}")
        df = self.df

        # Soil key & per-ha norms
        df['Soil_Key_5_1'] = [self._soil_key(s, r) for s, r in zip(df['Soil_Type'], df['Region'])]

        # Total land area for manure placement limit:
        # - Use grass + maize + crop. (If you need a different definition, adjust here.)
        df['Ha_Total_For_ManureLimit'] = (df['Ha_Grass'] + df['Ha_Mais'] + df['Ha_Crop']).clip(lower=0)

        # Manure placement limit (kg N/ha) and total (kg N/farm)
        df['Manure_Limit_kgN_per_ha'] = df.apply(self._manure_limit_per_ha, axis=1)
        df['Manure_Limit_Total_kgN'] = df['Manure_Limit_kgN_per_ha'] * df['Ha_Total_For_ManureLimit']

        # N usage norms ("gebruiksruimte") per land type
        # Grass norm depends on grazing
        grass_key = df.apply(self._grass_key, axis=1)
        maize_key = df.apply(self._maize_key, axis=1)

        df['Norm_Grass_kgN_per_ha'] = [self._crop_norm(k, sk) for k, sk in zip(grass_key, df['Soil_Key_5_1'])]
        df['Norm_Maize_kgN_per_ha'] = [self._crop_norm(k, sk) for k, sk in zip(maize_key, df['Soil_Key_5_1'])]
        df['Norm_Crop_kgN_per_ha'] = [self._crop_norm(k, sk) for k, sk in zip(df['CropType'], df['Soil_Key_5_1'])]

        df['UsageSpace_Grass_kgN'] = df['Norm_Grass_kgN_per_ha'] * df['Ha_Grass']
        df['UsageSpace_Maize_kgN'] = df['Norm_Maize_kgN_per_ha'] * df['Ha_Mais']
        df['UsageSpace_Crop_kgN'] = df['Norm_Crop_kgN_per_ha'] * df['Ha_Crop']
        df['UsageSpace_Total_kgN'] = df['UsageSpace_Grass_kgN'] + df['UsageSpace_Maize_kgN'] + df['UsageSpace_Crop_kgN']

        # Available manure N (VCRE) after NH3 emissions.
        # Use total VCRE excretion and subtract NH3-N losses (Stable+Storage+Grazing) in N units.
        for c in ['Total_N_Excretion_Farm_VCRE', 'Total_Emission_All_VCRE']:
            if c not in df.columns:
                df[c] = 0.0
            df = ensure_numeric(df, c, 0.0)

        df['ManureN_Available_After_Emissions_kgN'] = (df['Total_N_Excretion_Farm_VCRE'] - df['Total_Emission_All_VCRE']).clip(lower=0)

        # Applied manure N is capped by manure placement limit
        df['ManureN_Applied_kgN'] = np.minimum(df['ManureN_Available_After_Emissions_kgN'], df['Manure_Limit_Total_kgN'])

        # Working coefficient (effective fraction counting towards usage space)
        df['Working_Coefficient'] = df.apply(self._working_coefficient, axis=1)

        # Effective manure N for usage space
        df['ManureN_Effective_kgN'] = df['ManureN_Applied_kgN'] * df['Working_Coefficient']

        # Remaining fertiliser space
        df['FertiliserSpace_kgN'] = (df['UsageSpace_Total_kgN'] - df['ManureN_Effective_kgN']).clip(lower=0)

        # Optional: show which constraint binds
        df['BindingConstraint'] = np.where(
            df['ManureN_Available_After_Emissions_kgN'] >= df['Manure_Limit_Total_kgN'],
            'ManureLimit',
            'Availability'
        )

        print(f"  ✓ 4.2 done — Farm1 UsageSpace: {df['UsageSpace_Total_kgN'].iloc[0]:.1f} kg N; "
              f"Applied manure: {df['ManureN_Applied_kgN'].iloc[0]:.1f} kg N "
              f"({df['BindingConstraint'].iloc[0]})")
        return df

# ==============================================================================
# MODULE 4.3: Application Emission Calculation (Manure & Fertiliser)
# ==============================================================================

class ApplicationEmissionCalculator:
    """
    Calculates NH3 emissions from the application of Manure and Synthetic Fertilizers.
    Integrates results from stable/storage/grazing (calculated previously via EF & NEMA methods).
    """
    
    # --- EF Tables (from the provided screenshots) ---
    # We remove all spaces and lowercase everything to make matching immune to Excel typos
    
    # 1. Synthetic Fertiliser EFs (% of N applied)
    EF_FERTILISER = {
        '100%ammonium': 11.3,
        '100%nitraat': 0.0,
        'combinatievanammoniumetnitraat': 2.5, # Typo in excel matched here
        'combinatievanammoniumennitraat': 2.5, # Correct spelling fallback
        'ureum,gekorreld,zonderurease-remmer': 14.3,
        'ureum,gekorreld,meturease-remmer': 5.9,
        'voeibaarureumzonderurease-remmerofzuur': 7.5, # Typo 'voeibaar'
        'vloeibaarureumzonderurease-remmerofzuur': 7.5,
        'vloeibaarureummeturease-remmerofzuuer': 3.1,  # Typo 'zuuer'
        'vloeibaarureummeturease-remmerofzuur': 3.1,
        'vloeibaarureumtoegviainjectie': 1.5,          # Typo 'toeg'
        'vloeibaarureumtoegediendviainjectie': 1.5
    }

    # 2. Manure Application EFs on Grassland (% of TAN applied)
    EF_MANURE_GRASS = {
        'bovengronds': 68.0,
        'sleepvoet': 26.4,
        'sleepvoetverdund': 17.0,
        'sleufkouterverdund': 17.0,
        'sleufkouter': 21.7,
        'zodebemester': 17.0
    }

    # 3. Manure Application EFs on Cropland (Bouwland) (% of TAN applied)
    EF_MANURE_ARABLE = {
        'bovengronds': 69.0,
        'ineenwerkgangonderwerken': 22.0,
        'sleepvoet': 36.0,
        'diepeinjectie': 2.0,
        'ondiepeinjectie': 24.0
    }

    def __init__(self, df):
        self.df = df.copy()

    def _clean_string(self, s):
        """Helper to lowercase and strip ALL spaces for robust dictionary matching"""
        return str(s).lower().replace(' ', '').strip()

    def run_application_emission(self):
        print(f"\n{'='*70}\nMODULE 4.3: Application Emissions\n{'='*70}")
        
        # ---------------------------------------------------------
        # 1. Lookup Emission Factors (EF)
        # ---------------------------------------------------------
        
        # Synthetic Fertilizer EF
        fert_forms = self.df.get('Fertiliser_Form', pd.Series('100%ammonium', index=self.df.index))
        cleaned_fert_forms = fert_forms.apply(self._clean_string)
        self.df['EF_Fertiliser_%'] = cleaned_fert_forms.map(self.EF_FERTILISER).fillna(0.0)

        # Manure Application EF (Grassland)
        tech_grass = self.df.get('AM_Grassland', pd.Series('zodebemester', index=self.df.index))
        self.df['EF_Manure_Grass_%'] = tech_grass.apply(self._clean_string).map(self.EF_MANURE_GRASS).fillna(17.0)
        
        # Manure Application EF (Arable/Cropland)
        tech_arable = self.df.get('AM_Cropland', pd.Series('ondiepeinjectie', index=self.df.index))
        self.df['EF_Manure_Arable_%'] = tech_arable.apply(self._clean_string).map(self.EF_MANURE_ARABLE).fillna(24.0)

        # ---------------------------------------------------------
        # 2. Calculate TAN Applied & Manure Emissions (Average Method)
        # ---------------------------------------------------------
        
        # --- A. Retrieve correct column data from previous modules ---
        tan_correct_mun = self.df['Total_Corrected_TAN_MUN']
        prev_emission_mun = self.df['Total_Emission_All_MUN']
        
        # Handle potential column name variations from Module 1.1 safely
        if 'Total_Nitrogen_Excretion_kg' in self.df.columns:
            n_excr_mun = self.df['Total_Nitrogen_Excretion_kg']
        else:
            n_excr_mun = self.df.get('Total_Nitrogen_Excretion_MUN', 0.0) 
        
        tan_correct_vcre = self.df['Total_Corrected_TAN_VCRE']
        prev_emission_vcre = self.df['Total_Emission_All_VCRE']
        n_excr_vcre = self.df['Total_N_Excretion_Farm_VCRE'] 
        
        # Nitrogen input for the application stage (from Module 4.2)
        n_applied_manure = self.df['ManureN_Applied_kgN']
        n_applied_fert = self.df['FertiliserSpace_kgN']
        
        # --- B. Calculate Net TAN (deduct previous NH3 emissions, multiply by 14/17 to convert to N) ---
        net_tan_mun = (tan_correct_mun - (prev_emission_mun * (14.0 / 17.0))).clip(lower=0)
        net_tan_vcre = (tan_correct_vcre - (prev_emission_vcre * (14.0 / 17.0))).clip(lower=0)
        
        # --- C. Calculate averages (Numerator & Denominator) ---
        avg_net_tan = (net_tan_mun + net_tan_vcre) / 2.0
        self.df['Avg_Net_TAN_Excreted'] = avg_net_tan
        
        avg_n_excr = (n_excr_mun + n_excr_vcre) / 2.0
        avg_n_excr_safe = np.where(avg_n_excr > 0, avg_n_excr, 1.0) # Prevent division by zero
        
        # --- D. Calculate unified TAN % ---
        tan_pct = avg_net_tan / avg_n_excr_safe
        self.df['Avg_TAN_Pct'] = tan_pct.clip(0, 1) # Cap between 0 and 100%
        
        # ---------------------------------------------------------
        # 3. Calculate Manure Application Emissions (Land-Specific)
        # ---------------------------------------------------------
        
        # Get land areas (assuming arable land = Maize + Other Crops)
        ha_grass = self.df.get('Ha_Grass', 0.0)
        ha_mais = self.df.get('Ha_Mais', 0.0)
        ha_crop = self.df.get('Ha_Crop', 0.0)
        
        total_ha = ha_grass + ha_mais + ha_crop
        total_ha_safe = np.where(total_ha > 0, total_ha, 1.0) 
        
        # Calculate the fraction of manure going to each land type (proportional to area)
        frac_grass = ha_grass / total_ha_safe
        frac_arable = (ha_mais + ha_crop) / total_ha_safe
        
        # Distribute Applied Manure N per land type
        applied_n_grass = n_applied_manure * frac_grass
        applied_n_arable = n_applied_manure * frac_arable
        
        # Calculate TAN Applied per land type
        tan_applied_grass = applied_n_grass * self.df['Avg_TAN_Pct']
        tan_applied_arable = applied_n_arable * self.df['Avg_TAN_Pct']
        
        self.df['TAN_Applied_Manure_Grass'] = tan_applied_grass
        self.df['TAN_Applied_Manure_Arable'] = tan_applied_arable
        self.df['TAN_Applied_Manure_Total'] = tan_applied_grass + tan_applied_arable
        
        # Calculate emissions using land-specific EFs (convert NH3-N back to NH3 gas via 17/14)
        emission_grass = (tan_applied_grass * self.df['EF_Manure_Grass_%'] / 100.0) * (17.0 / 14.0)
        emission_arable = (tan_applied_arable * self.df['EF_Manure_Arable_%'] / 100.0) * (17.0 / 14.0)
        
        app_emission_manure = emission_grass + emission_arable
        
        self.df['Emission_ManureApp_Grass'] = emission_grass
        self.df['Emission_ManureApp_Arable'] = emission_arable
        self.df['Emission_ManureApp_Total'] = app_emission_manure
        
        # ---------------------------------------------------------
        # 4. Calculate Fertiliser Emission
        # ---------------------------------------------------------

        app_emission_fert = (n_applied_fert * self.df['EF_Fertiliser_%'] / 100.0) * (17.0 / 14.0)
        self.df['Emission_FertiliserApp'] = app_emission_fert

        # ---------------------------------------------------------
        # 5. Final Total Emissions (Previous + App Emissions)
        # ---------------------------------------------------------
        # Calculate final total emissions separately for both MUN and VCRE systems
        self.df['Total_Farm_NH3_Emission_MUN'] = prev_emission_mun + app_emission_manure + app_emission_fert
        self.df['Total_Farm_NH3_Emission_VCRE'] = prev_emission_vcre + app_emission_manure + app_emission_fert
            
        print(f"  ✓ 4.3 Application Emissions Calculated")
        print(f"      Avg TAN%: {self.df['Avg_TAN_Pct'].iloc[0]*100:.1f}%, App Manure Emission: {app_emission_manure.iloc[0]:.1f}")
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
    # If you want hard paths, edit here:
    INPUT = '/Users/shuaij/Desktop/0104 DMS data.xlsx'
    OUTPUT = '/Users/shuaij/Desktop/Output_DMS_Complete_0331v3.xlsx'
    run_pipeline(INPUT, OUTPUT)
    
    
