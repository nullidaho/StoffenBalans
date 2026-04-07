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
import matplotlib.pyplot as plt

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
            
        self.df = ensure_numeric(self.df, 'VEM_Milk_Mult', 1.0) # NEW: Multiplier
        
        self._breed()
        L,D = 326.0, 39.0
        f,p,bf = self.df['Fat%'], self.df['Pro%'], self.df['breed_factor']
        
        # Kalf & Pink
        vkg = 1323.0*bf; vke = self.df['GD_Unlimited_Kalf']*0.346*bf
        self.df['vem_req_kalf_per_head_yr'] = (vkg+vke)*1.02
        self.df['Total_VEM_Kalf_Farm'] = self.df['vem_req_kalf_per_head_yr']*self.df['Nr_kalf']
        
        vpg = 2259.0*bf; vpe = self.df['GD_Unlimited_Pink']*0.784*bf; vpp = 115.9*bf
        self.df['vem_req_pink_per_head_yr'] = (vpg+vpe+vpp)*1.02
        self.df['Total_VEM_Pink_Farm'] = self.df['vem_req_pink_per_head_yr']*self.df['Nr_pink']
        
        # Cow
        fpcm_yr = (0.337+0.116*f+0.06*p)*self.df['MilkYield']*365.0
        fpcm_dl = fpcm_yr/L; cf = 1.0+(fpcm_dl-15.0)*0.00165
        
        # --- NEW: Apply VEM_Milk_Mult to the milk production requirement ---
        self.df['vem_cow_milk_yr'] = ((442.0*fpcm_dl*cf/1000.0)*L) * self.df['VEM_Milk_Mult']
        
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
# MODULE 2.2 — VEM Allocation (updated 17.Mar)
# ══════════════════════════════════════════════════════════════════════════════
class VEMAllocationCalculator:
    VEM_KM=1500.0; VEM_CF=940.0; DS_DC=0.876
    LM=0.02; LK=0.02; LC=0.02; LR=0.05
    PGK=0.75; PMK=0.25; PGP=0.90; PMP=0.10
    def __init__(self, df): self.df = df.copy()
    
    @staticmethod
    def _soil_p(st):
        s = str(st).lower().strip()
        
        # We assign Yields AND VEM values based on soil type
        # yg: yield grass silage (cult), yf: yield fresh grass (cult), ym: maize
        # yn_gs: yield nature grass silage, yn_fg: yield nature fresh grass
        # vc_gs: VEM cult grass silage, vc_fg: VEM cult fresh grass
        # vn_gs: VEM nature grass silage, vn_fg: VEM nature fresh grass
        
        if 'klei' in s:   
            yg, yf, ym, yn_gs, yn_fg = 10005, 8893, 17685, 6000, 5330
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 960.0, 940.0, 842.0, 860.0, 950.0
            
        elif 'zand' in s: 
            yg, yf, ym, yn_gs, yn_fg = 9360, 8320, 17773, 6000, 5333
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 960.0, 940.0, 842.0, 860.0, 950.0
            
        elif 'veen' in s: 
            yg, yf, ym, yn_gs, yn_fg = 9751, 8668, 16620, 6000, 5333
            vc_gs, vc_fg, vn_gs, vn_fg, vm = 957.0, 937.0, 842.0, 860.0, 950.0 
            
        else: # Default values if soil is unknown or 'loss'
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
        
        defs = {'Vol_WholeMilk_Kalf':0,'Vol_KunstMelk_Kalf':0,'Kg_conc':0,
                'GH_Koe':6,'GH_Kalf':6,'GH_Pink':6,'Ha_Grass':0,'NatureGL%':0,'Ha_Mais':0,
                'VEM_Concentrate':self.VEM_CF,'GD_Unlimited_Kalf':0,'GD_Unlimited_Pink':0,
                'GD_Limited_Koe':0,'GD_Combi_Koe':0,'GD_Unlimited_Koe':0}
        
        for c,d in defs.items(): self.df = ensure_numeric(self.df, c, d)
        
        sp = self.df['Soil_Type'].apply(self._soil_p).tolist()
        
        for k in ('yield_gs','yield_fg','yield_ms','yield_nat_gs','yield_nat_fg',
                  'vem_gs_cult','vem_gs_nat','vem_fg_cult','vem_fg_nat','vem_maize'):
            self.df[k] = [d[k] for d in sp]
            
        # --- Apply Yield Multipliers ---
        self.df = ensure_numeric(self.df, 'Yield_FG_Mult', 1.0)
        self.df = ensure_numeric(self.df, 'Yield_GS_Mult', 1.0)
        self.df = ensure_numeric(self.df, 'Yield_Maize_Mult', 1.0)
        
        self.df['yield_fg'] *= self.df['Yield_FG_Mult']
        self.df['yield_nat_fg'] *= self.df['Yield_FG_Mult']
        self.df['yield_gs'] *= self.df['Yield_GS_Mult']
        self.df['yield_nat_gs'] *= self.df['Yield_GS_Mult']
        self.df['yield_ms'] *= self.df['Yield_Maize_Mult']
     
            
        pn = self.df['NatureGL%'].clip(0,100)
        # Calculate weighted VEM for both Fresh Grass and Grass Silage
        self.df['vem_fg_weighted'] = ((100-pn)*self.df['vem_fg_cult'] + pn*self.df['vem_fg_nat'])/100
        self.df['vem_gs_weighted'] = ((100-pn)*self.df['vem_gs_cult'] + pn*self.df['vem_gs_nat'])/100
        # Calculate weighted dry matter yields combining cultivated and nature grassland
        self.df['yield_fg_weighted'] = ((100-pn)*self.df['yield_fg'] + pn*self.df['yield_nat_fg'])/100
        self.df['yield_gs_weighted'] = ((100-pn)*self.df['yield_gs'] + pn*self.df['yield_nat_gs'])/100
    
        
        # Young stock
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
        rk = (self.df['Total_VEM_Kalf_Farm']-self.df['kVEM_Intake_Milk_Kalf']-self.df['kVEM_Intake_KunstMelk_Kalf']-self.df['kVEM_Intake_Conc_Kalf']-self.df['kVEM_Intake_FreshGrass_Kalf']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Kalf'] = rk*self.PGK
        self.df['kVEM_Intake_MaizeSilage_Kalf'] = rk*self.PMK
        rp = (self.df['Total_VEM_Pink_Farm']-self.df['kVEM_Intake_Conc_Pink']-self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Pink'] = rp*self.PGP
        self.df['kVEM_Intake_MaizeSilage_Pink'] = rp*self.PMP
        
        # Supply
        hrs = self.df['GH_Koe']; ir = np.where(hrs>2, 2.0+0.75*(hrs-2), hrs)
        tgd = self.df['GD_Limited_Koe']+self.df['GD_Combi_Koe']+self.df['GD_Unlimited_Koe']
        fy = (0.337+0.116*self.df['Fat%']+0.06*self.df['Pro%'])*self.df['MilkYield']*365
        mf = 1.0+((fy-9500*self.df['breed_factor'])/500)*0.02; dc = (365-39)/365
        
        # Use vem_fg_weighted for fresh grazing intake
        cgv = tgd*ir*(self.df['vem_fg_weighted']/1000)*dc*mf*self.df['breed_factor']*self.df['Nr_koe']
        tgv = cgv+self.df['kVEM_Intake_FreshGrass_Pink']+self.df['kVEM_Intake_FreshGrass_Kalf']
        
        # Use vem_fg_weighted for translating VEM to Dry Matter of fresh grass
        tfm = tgv/(self.df['vem_fg_weighted']/1000).replace(0,np.nan)
        hfn = (tfm/self.df['yield_fg_weighted'].replace(0,np.nan)).fillna(0)
        hag = (self.df['Ha_Grass']-hfn).clip(lower=0)
        
        # Use vem_gs_weighted for calculating available Grass Silage VEM
        vgs = hag*self.df['yield_gs_weighted']*(self.df['vem_gs_weighted']/1000)
        vms = self.df['Ha_Mais']*self.df['yield_ms']*(self.df['vem_maize']/1000)
        
        thg = tgv+vgs+vms; mk = thg>0
        rfg = np.where(mk, tgv/thg, 0); rgs = np.where(mk, vgs/thg, 0); rms = np.where(mk, vms/thg, 0)
        
        # Cow
        vc = self.df['VEM_Concentrate'].replace(0, self.VEM_CF) 
        dsc = self.df['Kg_conc'] * self.DS_DC
        mkc = self.df['kVEM_Intake_Conc_Kalf']*1000/vc/(1-self.LC)
        mpc = self.df['kVEM_Intake_Conc_Pink']*1000/vc/(1-self.LC)
        mcc = (dsc-mkc-mpc).clip(lower=0)
        self.df['kVEM_Intake_Conc_Cow'] = mcc*(1-self.LC)*vc/1000
        hgr = (self.df['Total_VEM_Cow_Farm']+self.df['Total_VEM_Kalf_Farm']+self.df['Total_VEM_Pink_Farm']
               -self.df['kVEM_Intake_Conc_Cow']-self.df['kVEM_Intake_Conc_Kalf']-self.df['kVEM_Intake_Conc_Pink']
               -self.df['kVEM_Intake_Milk_Kalf']-self.df['kVEM_Intake_KunstMelk_Kalf']).clip(lower=0)
        self.df['kVEM_Intake_FreshGrass_Cow'] = (hgr*rfg-self.df['kVEM_Intake_FreshGrass_Kalf']-self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Cow'] = (hgr*rgs-self.df['kVEM_Intake_GrassSilage_Kalf']-self.df['kVEM_Intake_GrassSilage_Pink']).clip(lower=0)
        self.df['kVEM_Intake_MaizeSilage_Cow'] = (hgr*rms-self.df['kVEM_Intake_MaizeSilage_Kalf']-self.df['kVEM_Intake_MaizeSilage_Pink']).clip(lower=0)
        
        print(f"  ✓ 2.2 done")
        return self.df
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.3 — N Intake (DS → N → CP → VRE)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenIntakeCalculator:
    FP=6.25; FD=6.38; RAS=40.0; NK=32.52; DSW=0.228; DSK=0.964
    def __init__(self, df): self.df = df.copy()
    
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
        elif ft=='concentrate': v = (88.7*(1.0-np.exp(-0.012*cp)))/100.0
        elif ft=='dairy':
            v = pd.Series(0.89, index=ds.index) if hasattr(ds, 'index') else 0.89
        else: v = 0.0
        v = np.clip(v, 0, 1)
        return ds*(v*cp)/1000.0
        
    def run_nutrient_calculation(self):
        print(f"\n{'='*70}\nMODULE 2.3: N Intake\n{'='*70}")
        sn = self.df['Soil_Type'].apply(self._soil_N).tolist()
        
        self.df['N_cont_fresh_soil'] = [d['nf'] for d in sn]
        self.df['N_cont_gs_soil'] = [d['ng'] for d in sn]
        self.df['N_cont_ms_soil'] = [d['nm'] for d in sn]
        self.df['N_cont_nat_gs'] = [d['nn_gs'] for d in sn]
        self.df['N_cont_nat_fg'] = [d['nn_fg'] for d in sn]
        
        # --- Apply Nitrogen Content Multipliers for Roughage ---
        self.df = ensure_numeric(self.df, 'N_Grass_Mult', 1.0)
        self.df = ensure_numeric(self.df, 'N_Maize_Mult', 1.0)
        
        self.df['N_cont_fresh_soil'] *= self.df['N_Grass_Mult']
        self.df['N_cont_gs_soil'] *= self.df['N_Grass_Mult']
        self.df['N_cont_nat_gs'] *= self.df['N_Grass_Mult']
        self.df['N_cont_nat_fg'] *= self.df['N_Grass_Mult']
        self.df['N_cont_ms_soil'] *= self.df['N_Maize_Mult']
        # -----------------------------------------------
        
        pn = self.df['NatureGL%'].clip(0,100)
        self.df['N_cont_fresh_weighted'] = ((100-pn)*self.df['N_cont_fresh_soil'] + pn*self.df['N_cont_nat_fg'])/100
        self.df['N_cont_gs_weighted'] = ((100-pn)*self.df['N_cont_gs_soil'] + pn*self.df['N_cont_nat_gs'])/100
        
        # Prepare base N_Concentrate value (default 27.3 if 0 or missing)
        self.df = ensure_numeric(self.df, 'N_Concentrate', 27.3)
        self.df['N_Concentrate'] = self.df['N_Concentrate'].replace(0, 27.3)
        
        # --- NEW: Apply Nitrogen Content Multiplier for Concentrate ---
        self.df = ensure_numeric(self.df, 'N_Conc_Mult', 1.0)
        self.df['N_Concentrate'] *= self.df['N_Conc_Mult']
        # ------------------------------------------------------------
        
        vfgw = self.df['vem_fg_weighted'] 
        vgsw = self.df['vem_gs_weighted']
        vmc = self.df['vem_maize']
        vc = self.df['VEM_Concentrate'].replace(0, 940.0)
        
        n_milk_ds = (self.df['Pro%']*10.0/self.DSW)/self.FD
        
        for sfx,fg,gs,ms,cc in [
            ('cow','kVEM_Intake_FreshGrass_Cow','kVEM_Intake_GrassSilage_Cow','kVEM_Intake_MaizeSilage_Cow','kVEM_Intake_Conc_Cow'),
            ('pink','kVEM_Intake_FreshGrass_Pink','kVEM_Intake_GrassSilage_Pink','kVEM_Intake_MaizeSilage_Pink','kVEM_Intake_Conc_Pink'),
            ('kalf','kVEM_Intake_FreshGrass_Kalf','kVEM_Intake_GrassSilage_Kalf','kVEM_Intake_MaizeSilage_Kalf','kVEM_Intake_Conc_Kalf')]:
            
            self.df[f'DS_fresh_{sfx}'] = (self.df[fg]*1000/vfgw).fillna(0)
            self.df[f'DS_gs_{sfx}'] = (self.df[gs]*1000/vgsw).fillna(0)
            self.df[f'DS_ms_{sfx}'] = (self.df[ms]*1000/vmc.replace(0,np.nan)).fillna(0)
            self.df[f'DS_conc_{sfx}'] = (self.df[cc]*1000/vc.replace(0,np.nan)).fillna(0)
            
        self.df['DS_milk_kalf'] = self.df['Vol_WholeMilk_Kalf']*self.DSW
        self.df['DS_kunst_kalf'] = self.df['Vol_KunstMelk_Kalf']*self.DSK
        
        for sfx in ('cow','pink','kalf'):
            lbl = sfx.capitalize()
            feeds = [(f'DS_fresh_{sfx}','N_cont_fresh_weighted','grass_fresh',False),
                     (f'DS_gs_{sfx}','N_cont_gs_weighted','grass_silage',False), 
                     (f'DS_ms_{sfx}','N_cont_ms_soil','maize_silage',False),
                     (f'DS_conc_{sfx}','N_Concentrate','concentrate',False)]
            if sfx=='kalf':
                feeds.append(('DS_milk_kalf', n_milk_ds, 'dairy', True))
                feeds.append(('DS_kunst_kalf', self.NK, 'dairy', True))
            tn=pd.Series(0.0,index=self.df.index); tcp=tn.copy(); tvre=tn.copy()
            for dc,nr,ft,dairy in feeds:
                ds = self.df[dc]; nc = self.df[nr] if isinstance(nr,str) else nr
                nf = self.FD if dairy else self.FP
                nk = ds*nc/1000; cpk = nk*nf; cpc = nc*nf
                vk = self._vre(ds, cpc, ft)
                tn+=nk; tcp+=cpk; tvre+=vk
            self.df[f'Total_N_Intake_{lbl}'] = tn
            self.df[f'Total_CP_Intake_{lbl}'] = tcp
            self.df[f'Total_VRE_Intake_{lbl}'] = tvre
            
        # --- Apply VCRE Coefficient Multiplier ---
        self.df = ensure_numeric(self.df, 'VCRE_Coeff_Mult', 1.0)
        mult_vcre = self.df['VCRE_Coeff_Mult']
        
        for l in ('Cow','Pink','Kalf'):
            self.df[f'VCRE_Factor_{l}'] = (self.df[f'Total_VRE_Intake_{l}']/self.df[f'Total_CP_Intake_{l}'].replace(0,np.nan)).fillna(0) * mult_vcre
            
        tvf = sum(self.df[f'Total_VRE_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        tcf = sum(self.df[f'Total_CP_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        self.df['VCRE_Factor_Farm'] = (tvf/tcf.replace(0,np.nan)).fillna(0)
        self.df['Total_N_Intake_Farm'] = sum(self.df[f'Total_N_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        print(f"  ✓ 2.3 done — Farm1 N intake: {self.df['Total_N_Intake_Farm'].iloc[0]:.1f}")
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
        
        df = ensure_numeric(df, 'Ret_Meat_Mult', 1.0) # NEW: Multiplier
        mult_meat = df['Ret_Meat_Mult']
        
        wb=wc*self.RB; wp=wc*self.R1; wv=wc*self.RC
        nb=wb*self.NK/1000; npk=wp*self.NP/1000; nv=wv*self.NV/1000; nc=wc*self.NC/1000
        ym = np.where(df['MilkYield']<100, 365, 1)
        tm = df['MilkYield']*ym*df['Nr_koe']
        
        # Milk retention stays standard
        rm = tm*df['Pro%']*10/self.FD/1000
        
        # Meat/Fetus retention
        rf = nb*self.BR*df['Nr_koe']
        nih = self.RR*nv*df['Nr_koe']; noc = self.RR*nc*df['Nr_koe']
        
        # Apply meat multiplier to fetal growth (rf) and weight change (noc-nih)
        df['N_Retention_Cow_VCRE'] = rm + ((rf + (noc-nih)) * mult_meat)
        
        gt = npk-nb; g1 = 0.36*df['breed_factor']
        t1 = gt*(0.376/0.407); t2 = (g1/2*24)*(0.031/0.407)
        ncr = np.divide(t1+t2, gt, out=np.ones_like(gt.values,dtype=float), where=gt.values!=0)
        
        # Youngstock retention is entirely growth (meat)
        df['N_Retention_Kalf_VCRE'] = (gt*df['Nr_kalf']*ncr) * mult_meat
        
        gha = (nv-npk)*(12/14)
        df['N_Retention_Pink_VCRE'] = ((nb+gha)*df['Nr_pink']) * mult_meat
        
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
        tech_grass = self.df.get('AM_Grass', pd.Series('zodebemester', index=self.df.index))
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
# ==============================================================================
# 1. PIPELINE WRAPPER 
# ==============================================================================
def run_pipeline_from_df(df_input):
    """Runs the REMAS pipeline directly from a DataFrame."""
    numeric_cols = ['Nr_koe', 'Nr_pink', 'Nr_kalf', 'MilkYield', 'Fat%', 'Pro%', 'MilkUerum',
                    'slurry_koe', 'solid_koe', 'slurry_kalf', 'solid_kalf', 'slurry_pink', 'solid_pink',
                    'volume_slurry_koe', 'volume_solid_koe', 'volume_slurry_kalf', 'volume_solid_kalf',
                    'volume_slurry_pink', 'volume_solid_pink', 'Vol_WholeMilk_Kalf', 'Vol_KunstMelk_Kalf',
                    'slurry%_koe', 'slurry%_kalf', 'slurry%_pink', 'NatureGL%', 'Ha_Grass', 'Ha_Mais', 
                    'Kg_conc', 'VEM_Concentrate', 'N_Concentrate', 'GD_Limited_Koe', 'GD_Combi_Koe', 
                    'GD_Unlimited_Koe', 'GD_Unlimited_Kalf', 'GD_Unlimited_Pink', 'GH_Koe', 'GH_Kalf', 'GH_Pink']
    
    for col in numeric_cols:
        if col in df_input.columns:
            df_input[col] = pd.to_numeric(df_input[col], errors='coerce').fillna(0.0)

    m11 = FarmManureCalculator(filepath=None) 
    m11.df = df_input.copy()
    
    df = m11.calculate_totals()
    df = NitrogenPartitioningCalculator(df).calculate_partitioning()
    df = VEMRequirementCalculator(df).calculate_requirements()
    df = VEMAllocationCalculator(df).run_allocation()
    df = NitrogenIntakeCalculator(df).run_nutrient_calculation()
    df = NitrogenExcretionCalculatorVCRE(df).calculate_excretion()
    df = NitrogenPartitioningVCRE(df).calculate_partitioning()
    df = MineralizationCalculator(df).calculate_mineralization()
    df = CorrectedTANCalculator(df).calculate_corrected_tan()
    df = EmissionCalculator(df).calculate_emissions()
    df = LandApplicationCalculator(df).calculate_land_application()
    df = ApplicationEmissionCalculator(df).run_application_emission()
    
    return df

# ==============================================================================
# 2. HELPER FUNCTIONS (Excel Baseline & Excretion Lookup)
# ==============================================================================
def get_baseline_from_excel(filepath, sheet_name='Main input'):
    """Extracts Farm 1 directly from your input Excel file."""
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()
    return df.iloc[[0]].copy() # Extract only the first row as the baseline

def estimate_excretion(yield_daily, urea):
    """
    Approximates the RVO table lookup for slurry and solid excretion based on your screenshots.
    (Yield in the table is annual, so we multiply daily yield * 365).
    """
    annual_yield = yield_daily * 365.0
    base_slurry = 79.0
    base_solid = 56.0
    
    slurry = base_slurry + 1.5 * (urea - 14.0) + 0.008 * (annual_yield - 5500)
    solid = base_solid + 1.0 * (urea - 14.0) + 0.006 * (annual_yield - 5500)
    
    return np.clip(slurry, 77.5, 163.5), np.clip(solid, 55.0, 116.0)

# ==============================================================================
# 3. SENSITIVITY ANALYSIS EXECUTION (Split Stages + Housing Tornado)
# ==============================================================================
def run_sensitivity_analysis(filepath):
    import matplotlib.pyplot as plt
    import numpy as np
    import pandas as pd
    
    baseline_df = get_baseline_from_excel(filepath)
    
    base_nr_koe = baseline_df['Nr_koe'].iloc[0]
    base_kg_conc = baseline_df['Kg_conc'].iloc[0]
    base_daily_conc = base_kg_conc / (base_nr_koe * 365.0) if base_nr_koe > 0 else 0
    base_gd_young = baseline_df['GD_Unlimited_Pink'].iloc[0]
    base_slurry_young = baseline_df['slurry%_pink'].iloc[0]

    # --- 1. Run Base Model to get Normalization Standard ---
    base_yield = baseline_df['MilkYield'].iloc[0]
    base_urea = baseline_df['MilkUerum'].iloc[0]
    
    # Slurry is calculated using RAW milk yield (physical volume matters for water intake/urine)
    baseline_df['slurry_koe'], baseline_df['solid_koe'] = estimate_excretion(base_yield, base_urea)
    
    res_base = run_pipeline_from_df(baseline_df)
    
    base_housing_nh3 = res_base['Total_Emission_All_VCRE'].iloc[0] 
    base_app_nh3 = res_base['Emission_ManureApp_Total'].iloc[0] + res_base['Emission_FertiliserApp'].iloc[0]
    base_total_nh3 = base_housing_nh3 + base_app_nh3
    
    # --- NEW: Baseline FPCM Calculation for Normalization ---
    base_fat = baseline_df['Fat%'].iloc[0]
    base_pro = baseline_df['Pro%'].iloc[0]
    base_fpcm_yield = base_yield * (0.337 + 0.116 * base_fat + 0.06 * base_pro)
    base_milk_total_fpcm = base_fpcm_yield * base_nr_koe * 365.0
    
    # Calculate base normalized emissions (g NH3 / kg FPCM)
    base_norm_total = (base_total_nh3 * 1000.0) / base_milk_total_fpcm if base_milk_total_fpcm > 0 else 0
    base_norm_housing = (base_housing_nh3 * 1000.0) / base_milk_total_fpcm if base_milk_total_fpcm > 0 else 0 
    
    print(f"\n--- BASELINE RESULTS ---")
    print(f"Total FPCM Yield:   {base_milk_total_fpcm:.0f} kg (Corrected for Fat & Pro)")
    print(f"Housing NH3:        {base_housing_nh3:.2f} kg")
    print(f"Application NH3:    {base_app_nh3:.2f} kg")
    print(f"Housing Normalized: {base_norm_housing:.2f} g NH3 / kg FPCM\n")

    categorical_grid = {
        'Region': ['Southern', 'Others'],
        'NV_Area': ['Yes', 'No'],
        'Derogation': ['Yes', 'No'],
        'Breed': ['Jersey', 'cross', 'others'],
        'Soil_Type': ['Zand', 'Klei', 'Veen', 'loss'],
        'AM_Grassland': ['bovengronds', 'sleepvoet', 'sleepvoet verdund', 'sleufkouter verdund', 'sleufkouter', 'zodebemester'],
        'AM_Cropland' : ['bovengronds', 'sleepvoet', 'diepe injectie', 'ondiepe injectie', 'in eed werkgangonderwerken'],
        'Fertiliser_Form': ['100%ammonium', '100%nitraat', 'combinatievanammoniumennitraat', 'Ureum, gekorreld,zonderurease-remmer', 'Ureum, gekorreld, meturease-remmer', 'Vloeibaarureumzonderurease-remmerofzuur', 'Vloeibaarureummeturease-remmerofzuur', 'Vloeibaarureumtoegviainjectie']
    }

    continuous_grid = {
        'Ha_Grass': np.linspace(0.0, 400.0, 5).tolist(),
        'Ha_Mais': np.linspace(0.0, 200.0, 5).tolist(), 
        'NatureGL%': np.linspace(0.0, 100.0, 5).tolist(),
        'GD_Unlimited_Koe': np.linspace(0.0, 300.0, 5).tolist(),
        'GH_Koe': np.linspace(0.0, 24.0, 5).tolist(),
        'Nr_koe': np.linspace(10.0, 500.0, 5).tolist(),
        'Nr_pink': np.linspace(0.0, 100.0, 5).tolist(), 
        'Nr_kalf': np.linspace(0.0, 100.0, 5).tolist(),
        'MilkUerum': np.linspace(11.0, 30.0, 5).tolist(),
        'MilkYield': np.linspace(15.0, 30.0, 5).tolist(),
        'Pro%': np.linspace(4.0, 5.0, 5).tolist(),
        'Fat%': np.linspace(3.0, 4.0, 5).tolist(),
        'slurry%_koe': np.linspace(0.0, 1.0, 5).tolist(),
        
        # Multipliers
        'Yield_FG_Mult': np.linspace(0.9, 1.1, 5).tolist(),
        'Yield_GS_Mult': np.linspace(0.6, 1.4, 5).tolist(),
        'Yield_Maize_Mult': np.linspace(0.83, 1.17, 5).tolist(),
        'N_Grass_Mult': np.linspace(0.76, 1.24, 5).tolist(),
        'N_Maize_Mult': np.linspace(0.82, 1.18, 5).tolist(),
        'N_Conc_Mult': np.linspace(0.7, 1.3, 5).tolist(),      # NEW: Concentrate N Multiplier
        'VEM_Milk_Mult': np.linspace(0.95, 1.05, 5).tolist(),
        'Ret_Meat_Mult': np.linspace(0.9, 1.1, 5).tolist(),
        'VCRE_Coeff_Mult': np.linspace(0.9, 1.1, 5).tolist(),
        
        # Pseudo-parameters
        'Daily_Conc_Per_Cow': np.linspace(3.0, 15.0, 5).tolist(), 
        'GD_Youngstock': np.linspace(0.0, 300.0, 5).tolist(),   
        'slurry%_youngstock': np.linspace(0.0, 1.0, 5).tolist() 
    }

    sensitivity_grid = {**categorical_grid, **continuous_grid}
    results = []

    print("--- RUNNING SENSITIVITY SCRIPT (SPLIT STAGES WITH FPCM) ---")
    for param, values in sensitivity_grid.items():
        for val in values:
            test_df = baseline_df.copy() 
            
            current_daily_conc = base_daily_conc
            current_gd_young = base_gd_young
            current_slurry_young = base_slurry_young
            
            if param == 'Daily_Conc_Per_Cow': current_daily_conc = val
            elif param == 'GD_Youngstock': current_gd_young = val
            elif param == 'slurry%_youngstock': current_slurry_young = val
            else: test_df[param] = val
                
            test_df['Kg_conc'] = test_df['Nr_koe'].iloc[0] * current_daily_conc * 365.0
            test_df['GD_Unlimited_Pink'] = current_gd_young
            test_df['GD_Unlimited_Kalf'] = current_gd_young
            test_df['slurry%_pink'] = current_slurry_young
            test_df['slurry%_kalf'] = current_slurry_young
            
            # Use raw milk yield to estimate physical slurry volume
            current_yield = test_df['MilkYield'].iloc[0]
            current_urea = test_df['MilkUerum'].iloc[0]
            new_slurry, new_solid = estimate_excretion(current_yield, current_urea)
            test_df['slurry_koe'] = new_slurry
            test_df['solid_koe'] = new_solid
            
            try:
                res = run_pipeline_from_df(test_df)
                
                # Extract Split Emissions
                housing_nh3 = res['Total_Emission_All_VCRE'].iloc[0] 
                app_nh3 = res['Emission_ManureApp_Total'].iloc[0] + res['Emission_FertiliserApp'].iloc[0]
                total_nh3 = housing_nh3 + app_nh3
                
                # --- NEW: FPCM calculation for Normalization step ---
                res_fat = test_df['Fat%'].iloc[0]
                res_pro = test_df['Pro%'].iloc[0]
                res_yield = res['MilkYield'].iloc[0]
                res_nr_koe = res['Nr_koe'].iloc[0]
                
                current_fpcm_yield = res_yield * (0.337 + 0.116 * res_fat + 0.06 * res_pro)
                total_fpcm = current_fpcm_yield * res_nr_koe * 365.0
                
                if total_fpcm > 0:
                    norm_housing = (housing_nh3 * 1000.0) / total_fpcm
                    norm_app = (app_nh3 * 1000.0) / total_fpcm
                    norm_total = (total_nh3 * 1000.0) / total_fpcm
                else:
                    norm_housing, norm_app, norm_total = 0, 0, 0
                
                # Calculate the Percentage Difference specifically for HOUSING
                diff_total_pct = ((norm_total - base_norm_total) / base_norm_total) * 100 if base_norm_total else 0
                diff_housing_pct = ((norm_housing - base_norm_housing) / base_norm_housing) * 100 if base_norm_housing else 0
                
                results.append({
                    'Parameter': param,
                    'Test_Value': val,
                    'Norm_Housing_NH3': norm_housing,
                    'Norm_App_NH3': norm_app,
                    'Norm_Total_NH3': norm_total,
                    'Norm_Total_Diff_%': diff_total_pct,
                    'Norm_Housing_Diff_%': diff_housing_pct 
                })
            except Exception as e:
                print(f"Error testing {param} = {val}: {e}")

    results_df = pd.DataFrame(results)
    results_df.to_csv("Sensitivity_Results_Split_22.csv", index=False)
    print("\n✓ Sensitivity Analysis Complete. CSV Saved.")
    
    # Call BOTH chart plotting functions (ensure they are defined in your script)
    plot_tornado(results_df)
    plot_line_charts_split(results_df, list(continuous_grid.keys()))
    
    return results_df

# ==============================================================================
# 4. TORNADO CHART GENERATION (Targeting Housing Emission Only)
# ==============================================================================

def plot_tornado(results_df):
    import matplotlib.pyplot as plt
    temp_df = results_df.copy()
    temp_df['Test_Value'] = temp_df['Test_Value'].astype(str)
    
    # --- CHANGED: Plot based on the Maximum Absolute impact on the HOUSING normalized emission ---
    impacts = temp_df.groupby('Parameter')['Norm_Housing_Diff_%'].apply(lambda x: x.abs().max()).sort_values(ascending=True)
    impacts = impacts[impacts > 0] 

    plt.figure(figsize=(10, 8))
    bars = plt.barh(impacts.index, impacts.values, color='steelblue', edgecolor='black')
    
    # --- CHANGED: Updated Labels to reflect Housing ---
    plt.xlabel('Maximum Absolute Change in Normalized Housing NH3 (%)', fontsize=12)
    plt.title('Sensitivity Analysis: Tornado Chart (Normalized Housing NH3)', fontsize=14, fontweight='bold')
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    
    for bar in bars:
        plt.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2, 
                 f'{bar.get_width():.1f}%', va='center', fontsize=10)

    plt.tight_layout()
    plt.savefig('Tornado_Chart_Housing_22.png', dpi=300)
    plt.close() 
    print("✓ Tornado Chart saved as 'Tornado_Chart_Housing_22.png'")
# ==============================================================================
# 5. LINE CHARTS FOR CONTINUOUS VARIABLES (SPLIT STAGES)
# ==============================================================================
def plot_line_charts_split(results_df, continuous_params):
    import math
    num_plots = len(continuous_params)
    cols = 4
    rows = math.ceil(num_plots / cols)
    
    fig, axes = plt.subplots(rows, cols, figsize=(16, 3.5 * rows))
    axes = axes.flatten() 
    
    for i, param in enumerate(continuous_params):
        param_data = results_df[results_df['Parameter'] == param].copy()
        param_data['Test_Value'] = pd.to_numeric(param_data['Test_Value'])
        param_data = param_data.sort_values(by='Test_Value')
        
        ax = axes[i]
        
        # Plot all three lines
        ax.plot(param_data['Test_Value'], param_data['Norm_Total_NH3'], marker='o', color='darkred', linewidth=2, label='Total')
        ax.plot(param_data['Test_Value'], param_data['Norm_Housing_NH3'], marker='s', color='steelblue', linewidth=1.5, linestyle='--', label='Housing/Storage')
        ax.plot(param_data['Test_Value'], param_data['Norm_App_NH3'], marker='^', color='forestgreen', linewidth=1.5, linestyle=':', label='Application')
        
        ax.set_title(param, fontsize=10, fontweight='bold')
        ax.set_ylabel('g NH3 / kg milk' if i % cols == 0 else '') 
        ax.grid(True, linestyle=':', alpha=0.6)
        
        # Add a legend only to the very first plot to keep things clean
        if i == 0:
            ax.legend(loc='best', fontsize=8)
    
    for j in range(num_plots, len(axes)):
        fig.delaxes(axes[j])
        
    plt.tight_layout()
    plt.savefig('Continuous_Trends_Split_22.png', dpi=300)
    plt.close()
    print("✓ Line Charts grid saved as 'Continuous_Trends_Split_22.png'")

if __name__ == '__main__':
    
    INPUT_FILE = '/Users/shuaij/Desktop/1803 Sensitivity Data copy farm 22.xlsx'
    # Run the sensitivity analysis
    df_results = run_sensitivity_analysis(INPUT_FILE)