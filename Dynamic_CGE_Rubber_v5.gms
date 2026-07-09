*
* RubberCGE_Dynamic.gms
* ├── SECTION 0: File I/O
* ├── SECTION 1: Sets
* ├── SECTION 2: Parameters (Static + Dynamic)
* ├── SECTION 3: Calibration (จาก SAM)
* ├── SECTION 4: Elasticities & Shares
* ├── SECTION 5: Variables & Equations
* ├── SECTION 6: Model Definition
* ├── SECTION 7: Base Solve (ปี 0)
* └── SECTION 8: Dynamic Loop (ปี 1→T)
$TITLE Rubber CGE Thailand — Recursive Dynamic (2024-2034)
*==============================================================================
* SECTION 0: FILE I/O
*==============================================================================
$setglobal XL    "SAM2024B.xlsx"
$setglobal SHEET "sam"
$setglobal GDX   "sam2024B.gdx"

$if not exist "%XL%" $abort "Excel file not found: %XL%"
$call gdxxrw i="%XL%" o="%GDX%" trace=3 par=SAM_data rng=%SHEET%!A1 rdim=1 cdim=1
$ifE errorlevel<>0 $abort "gdxxrw failed."
$if not exist "%GDX%" $abort "GDX not created."

*==============================================================================
* SECTION 1: SETS
*==============================================================================
Set acc /
  com1*com29, act1*act29
  TAX_IND, TAX_M, TAX_DIR
  LAB, CAP, GOV, ROW, SAV
  HH_Q1*HH_Q5
/;
Alias(acc,ii,jj);

Set COM(acc) / com1*com29 /;
Set ACT(acc) / act1*act29 /;
Set HH(acc)  / HH_Q1*HH_Q5 /;

* 1:1 commodity-activity mapping
Set MAP(COM,ACT);
MAP(COM,ACT)$(ord(COM)=ord(ACT)) = yes;

* CET sectors
Set CETCOM(COM) / com10,com11,com12,com14,com15,com16,
                  com26,com27,com28,com29 /;

* Rubber commodity set (used in EQINCOME rubber channel)
Set RUBCOM(COM) / com26, com27, com28, com29 /;

* Rubber activity set (used for partial capital adjustment bounds in scenarios)
Set RUBACT(ACT) / act26, act27, act28, act29 /;

* Input cost commodity set (for S7 input cost shock)
Set COST_COM(COM) / com6, com13, com25 /;

* ★ Time dimension (Recursive Dynamic)
Set T    "Simulation years"   / 2024*2034 /;
Set TFIRST(T)                 / 2024 /;
Set TLAST(T)                  / 2034 /;
Alias(T,TT);

*==============================================================================
* SECTION 2: ALL PARAMETERS — Static + Dynamic
*==============================================================================
Parameter SAM_data(acc,acc) "Raw SAM from Excel";
$gdxin "%GDX%"
$load SAM_data
$gdxin

*--- 2A: Static base parameters (calibrated from SAM)
Parameter
  io(COM,ACT)        "Intermediate input coefficients"
  LZ(ACT)            "Base labor demand"
  KZ(ACT)            "Base capital demand"
  XDZ(ACT)           "Base activity output"
  CZ(COM,HH)         "Base household consumption"
  YZ(HH)             "Base household income"
  SHZ(HH)            "Base household saving"
  CGZ(COM)           "Base government demand"
  IZ(COM)            "Base investment demand"
  MZ(COM)            "Base imports"
  EZ(COM)            "Base exports"
  tc(COM)            "Commodity tax rate"
  ty(HH)             "Direct income tax rate"
  tm(COM)            "Import tariff rate"
  mps(HH)            "Marginal propensity to save"
  h_lab_sh(HH)       "HH labor income share"
  h_cap_sh(HH)       "HH capital income share"
  betaC(COM,HH)      "Budget shares"
  alphaI(COM)        "Investment commodity shares"
  alphaCG(COM)       "Government commodity shares"
  alphaKG            "Govt capital share"
  alphaLG            "Govt labor share"
  PWMZ(COM)          "World import price"
  PWEZ(COM)          "World export price"
  PLZ, PKZ           "Base wage, rental rate"
  PCINDEXZ           "Base CPI"
  ERZ                "Base exchange rate"
  LSZ                "Labor supply"
  KSZ                "Capital supply"
  UNEMPZ             "Base unemployment"
  SFZ                "Foreign saving"
  fac_lab_sh(HH)     "HH share of total labor income (sum=1)"
  fac_cap_sh(HH)     "HH share of total capital income (sum=1)"
  useARM(COM)        "Armington switch"
  rub_inc_sh(HH)      "rubber income share by quintile"
;

*--- 2B: Elasticity & Share parameters
Parameter
  sigmaF(ACT)    "CES VA elasticity"
  gammaF(ACT)    "CES capital share"
  aF(ACT)        "CES scale parameter"
  tk(ACT)        "Capital tax on activities"
  tl(ACT)        "Labor tax on activities"
  sigmaT(COM)    "CET elasticity"
  gammaT(COM)    "CET export share"
  aT(COM)        "CET scale"
  useCET(COM)    "CET switch"
  sigmaA(COM)    "Armington elasticity"
  gammaA(COM)    "Armington import share"
  aA(COM)        "Armington scale"
;

Scalar
  phillips      "Wage curve slope"          / 0.04 /
  trep          "Unemployment benefit rate" / 0.0 /
  TRO           "Other transfers index"     / 0.0 /
  unemp_floor   "Min unemployment floor"    / 0.001 /
  CPI_denom     "CPI denominator"
  unemp_base_rate "Base unemployment rate"
  kappa         "Max annual capital reallocation rate per sector" / 0.20 /
* S12 targeted transfer scalars
  transfer_per_hh   "15000 THB converted to SAM units"   / 0 /
  transfer_switch   "0=off 1=on transfer in EQINCOME+EQTAXREV" / 0 /
  SAM_unit          "SAM unit: 1e3=thousands THB"  / 1e3 /
  subsidy_equiv     "Effective subsidy rate for S10" / 0.15 /
;

*--- 2C: ★ Dynamic parameters (calibrated from Thailand data)
Parameter
  deprate(ACT)     "Capital depreciation rate by sector"
  popgrow_t(T)     "Labor force growth rate by year (time-varying)"
  tfpgrow(ACT)     "Annual TFP growth rate by sector"
;

*--------------------------------------------------------------
* DEPRATE — อัตราค่าเสื่อมราคาทุนรายสาขา
*--------------------------------------------------------------
deprate(ACT) = 0.05;
deprate('act1')=0.035; deprate('act2')=0.045; deprate('act3')=0.030;
deprate('act4')=0.050; deprate('act5')=0.075;
deprate('act6')=0.055; deprate('act7')=0.050; deprate('act8')=0.060;
deprate('act9')=0.055; deprate('act10')=0.065; deprate('act11')=0.060;
deprate('act12')=0.065; deprate('act13')=0.055; deprate('act14')=0.065;
deprate('act15')=0.070; deprate('act16')=0.080; deprate('act17')=0.060;
deprate('act18')=0.035; deprate('act19')=0.040;
deprate('act20')=0.045; deprate('act21')=0.045; deprate('act22')=0.055;
deprate('act23')=0.040; deprate('act24')=0.025; deprate('act25')=0.045;
deprate('act26')=0.040; deprate('act27')=0.055;
deprate('act28')=0.080; deprate('act29')=0.065;

*--------------------------------------------------------------
* POPGROW_T — Labor force growth rate (time-varying)
*--------------------------------------------------------------
popgrow_t('2024')=0.000;
popgrow_t('2025')=-0.002; popgrow_t('2026')=-0.004; popgrow_t('2027')=-0.005;
popgrow_t('2028')=-0.007; popgrow_t('2029')=-0.008; popgrow_t('2030')=-0.009;
popgrow_t('2031')=-0.010; popgrow_t('2032')=-0.011;
popgrow_t('2033')=-0.012; popgrow_t('2034')=-0.013;

*--------------------------------------------------------------
* TFPGROW — TFP growth rate by sector
*--------------------------------------------------------------
tfpgrow(ACT)=0.008;
tfpgrow('act1')=0.012; tfpgrow('act2')=0.010; tfpgrow('act3')=0.005;
tfpgrow('act4')=0.003; tfpgrow('act5')=0.002;
tfpgrow('act6')=0.012; tfpgrow('act7')=0.006; tfpgrow('act8')=0.007;
tfpgrow('act9')=0.004; tfpgrow('act10')=0.010; tfpgrow('act11')=0.007;
tfpgrow('act12')=0.009; tfpgrow('act13')=0.006; tfpgrow('act14')=0.010;
tfpgrow('act15')=0.008; tfpgrow('act16')=0.018; tfpgrow('act17')=0.008;
tfpgrow('act18')=0.010; tfpgrow('act19')=0.005;
tfpgrow('act20')=0.010; tfpgrow('act21')=0.008; tfpgrow('act22')=0.012;
tfpgrow('act23')=0.015; tfpgrow('act24')=0.003; tfpgrow('act25')=0.005;
tfpgrow('act26')=0.005; tfpgrow('act27')=0.008;
tfpgrow('act28')=0.018; tfpgrow('act29')=0.012;

Display deprate, popgrow_t, tfpgrow;

*--- 2D: ★ Dynamic tracking parameters
Parameter
  KZ_t(ACT,T)       "Capital stock by sector and year"
  LSZ_t(T)          "Labor supply by year"
  aF_t(ACT,T)       "TFP level by sector and year"
  Y_path(HH,T)      "HH income path"
  XD_path(ACT,T)    "Sector output path"
  E_path(COM,T)     "Export path"
  M_path(COM,T)     "Import path"
  PL_path(T)        "Wage path"
  PK_path(T)        "Rental rate path"
  UNEMP_path(T)     "Unemployment path"
  TAXR_path(T)      "Tax revenue path"
  S_path(T)         "Total saving path"
  PCINDEX_path(T)   "CPI path"
  CH_path(COM,HH,T) "Consumption path"
  INV_path(COM,T)   "Investment path"
  L_path(ACT,T)     "Labor demand path"
  gY(HH,T)          "HH income growth rate"
  gXD(ACT,T)        "Sector output growth rate"
  gPL(T)            "Wage growth rate"
  gUNEMP(T)         "Unemployment change"
  kapshare(COM,ACT) "Investment allocation matrix"
  K_prev(ACT)       "Capital stock in previous period (for partial adjustment bounds)"
  K_path(ACT,T)     "Baseline capital stock path by sector (used as scenario bound anchor)"
* --- Scenario S3-S12 result storage ---
  Y_path_s3(HH,T)    "S3: price -30% temporary 3yr"
  XD_path_s3(ACT,T)
  UNEMP_path_s3(T)
  dY_s3(HH,T)
  Y_path_s3b(HH,T)   "S3b: price -20% temporary 3yr"
  XD_path_s3b(ACT,T)
  UNEMP_path_s3b(T)
  dY_s3b(HH,T)
  Y_path_s5(HH,T)    "S5: TFP rubber -10%"
  XD_path_s5(ACT,T)
  UNEMP_path_s5(T)
  dY_s5(HH,T)
  Y_path_s6(HH,T)    "S6: TFP rubber +10%"
  XD_path_s6(ACT,T)
  UNEMP_path_s6(T)
  dY_s6(HH,T)
  Y_path_s7(HH,T)    "S7: Input cost +20%"
  XD_path_s7(ACT,T)
  UNEMP_path_s7(T)
  dY_s7(HH,T)
  Y_path_s8(HH,T)    "S8: Export demand -15%"
  XD_path_s8(ACT,T)
  UNEMP_path_s8(T)
  dY_s8(HH,T)
  Y_path_s9(HH,T)    "S9: Real ER +10% depreciation"
  XD_path_s9(ACT,T)
  UNEMP_path_s9(T)
  dY_s9(HH,T)
  Y_path_s10(HH,T)   "S10: S1 + Output subsidy"
  XD_path_s10(ACT,T)
  UNEMP_path_s10(T)
  dY_s10(HH,T)
  Y_path_s11(HH,T)   "S11: S1 + TFP support phased"
  XD_path_s11(ACT,T)
  UNEMP_path_s11(T)
  dY_s11(HH,T)
  Y_path_s12(HH,T)   "S12: S1 + Cash transfer Q1-Q2"
  XD_path_s12(ACT,T)
  UNEMP_path_s12(T)
  dY_s12(HH,T)
* --- Policy instrument parameters ---
  subsidy_rate(ACT)      "Output subsidy rate on rubber sectors"
  tfp_support(ACT,T)     "Additional TFP increment from policy (phased)"
  io_base(COM,ACT)       "Original io coefficients (backup before S7 shock)"
  aF_base(ACT)           "Original TFP before scenario"
  EZ_base(COM)           "Original export demand (backup before S8 shock)"
  tfpgrow_s5(ACT)        "TFP growth rate for S5 scenario"
  tfpgrow_s6(ACT)        "TFP growth rate for S6 scenario"
* --- S12 targeted transfer ---
  rub_share(HH)          "Rubber farmer share per quintile (RAOT registry)"
  n_rub_hh(HH)           "Number of eligible rubber smallholders ≤15 rai"
  GOV_TRANSFER_RUB(HH)   "Total transfer per quintile (SAM units)"
  total_transfer_cost    "Total transfer cost S12 (SAM units)"
;

* [S12 init] ตั้งค่าเริ่มต้น = 0 ก่อน base solve
* เพื่อให้ EQINCOME และ EQTAXREV compile ได้โดยไม่ error
GOV_TRANSFER_RUB(HH) = 0;
total_transfer_cost   = 0;
transfer_switch       = 0;

*--- 2E: Base solution storage
Parameter
  YZ_base(HH), E_base2(COM), XD_base3(ACT)
  PL_base2, PK_base2, UNEMP_base2, TAXR_base2
;

*==============================================================================
* SECTION 3: CALIBRATION FROM SAM
*==============================================================================
*--- 3A: Factor endowments
Parameter sumLabHH, sumCapHH, total_xd(ACT);
sumLabHH = sum(HH, SAM_data(HH,"LAB"));
sumCapHH = sum(HH, SAM_data(HH,"CAP"));
sumLabHH$(sumLabHH=0) = 1;
sumCapHH$(sumCapHH=0) = 1;

total_xd(ACT) = sum(COM, SAM_data(COM,ACT))
              + SAM_data("LAB",ACT) + SAM_data("CAP",ACT);
total_xd(ACT)$(total_xd(ACT)=0) = 1;

io(COM,ACT) = SAM_data(COM,ACT) / total_xd(ACT);

LSZ = sum(ACT, SAM_data("LAB",ACT));
KSZ = sum(ACT, SAM_data("CAP",ACT));
LSZ$(LSZ=0) = 1;
KSZ$(KSZ=0) = 1;
UNEMPZ = 0.01 * LSZ;

unemp_base_rate = UNEMPZ / max(1, LSZ);
Display unemp_base_rate;

*--- 3B: Output, factor demands
XDZ(ACT) = total_xd(ACT);
LZ(ACT)  = SAM_data("LAB",ACT);
KZ(ACT)  = SAM_data("CAP",ACT);
LZ(ACT)$(LZ(ACT)=0) = 1e-6;
KZ(ACT)$(KZ(ACT)=0) = 1e-6;

*--- 3C: Household income & tax
YZ(HH) = SAM_data(HH,"LAB") + SAM_data(HH,"CAP")
        + SAM_data(HH,"GOV") + SAM_data(HH,"ROW");
YZ(HH)$(YZ(HH)=0) = 1;
ty(HH) = SAM_data("TAX_DIR",HH) / YZ(HH);
ty(HH)$(ty(HH)<0) = 0;

*--- 3D: Consumption
CZ(COM,HH) = SAM_data(COM,HH);
CZ(COM,HH)$(CZ(COM,HH)=0) = 1e-6;

*--- 3E: Trade
MZ(COM) = SAM_data("ROW",COM);
EZ(COM) = SAM_data(COM,"ROW");
MZ(COM)$(MZ(COM)=0) = 1e-6;
EZ(COM)$(EZ(COM)=0) = 1e-6;

Display MZ;

*--- 3F: Government & investment
CGZ(COM) = SAM_data(COM,"GOV");
CGZ(COM)$(CGZ(COM)=0) = 1e-6;
IZ(COM)  = SAM_data("SAV",COM);
IZ(COM)$(IZ(COM)=0) = 1e-6;
SFZ = sum(COM, SAM_data("ROW",COM)) - sum(COM, SAM_data(COM,"ROW"));

*--- 3G: Tax rates
Parameter taxind_act(ACT), tc_act(ACT);
taxind_act(ACT) = SAM_data("TAX_IND",ACT);
tc_act(ACT)$(total_xd(ACT)>0) = taxind_act(ACT) / total_xd(ACT);
tc(COM) = sum(ACT$MAP(COM,ACT), tc_act(ACT));
tc(COM)$(tc(COM)<0)    = 0;
tc(COM)$(tc(COM)>0.99) = 0.99;
tm(COM) = SAM_data("TAX_M",COM) / max(1, MZ(COM));
tm(COM)$(tm(COM)<0)   = 0;
tm(COM)$(tm(COM)>2.0) = 2.0;

PLZ=1; PKZ=1; PCINDEXZ=1; ERZ=1;
PWMZ(COM) = 1 / max(1e-4, (1+tm(COM)));
PWEZ(COM) = 1;

*--- 3H: Shares
Parameter sumCZ(HH), sumCGZ, IZ_total;
sumCZ(HH) = sum(COM, CZ(COM,HH));
sumCZ(HH)$(sumCZ(HH)=0) = 1;
betaC(COM,HH) = CZ(COM,HH) / sumCZ(HH);
betaC(COM,HH)$(SAM_data(COM,HH)=0) = 0;

mps(HH) = 1 - sumCZ(HH) / max(1, (1-ty(HH))*YZ(HH));
mps(HH)$(mps(HH)<=0.001) = 0.001;
mps(HH)$(mps(HH)>=0.990) = 0.990;

CPI_denom = sum((COM,HH), (1+tc(COM))*CZ(COM,HH));
CPI_denom$(CPI_denom=0) = 1;

sumCGZ = sum(COM, CGZ(COM));
sumCGZ$(sumCGZ=0) = 1;
alphaCG(COM) = CGZ(COM) / sumCGZ;

IZ_total = sum(COM, IZ(COM));
IZ_total$(IZ_total=0) = 1;
alphaI(COM) = IZ(COM) / IZ_total;
alphaKG = 0;
alphaLG = 0;

*--- 3I: Investment allocation matrix
Parameter tot_k;
tot_k = sum(ACT, KZ(ACT));
tot_k$(tot_k=0) = 1;
kapshare(COM,ACT)$MAP(COM,ACT) = KZ(ACT) / tot_k;

*--- 3J: Factor income shares — calibrated from YZ + NSO SES
h_lab_sh('HH_Q1')=0.85; h_lab_sh('HH_Q2')=0.78; h_lab_sh('HH_Q3')=0.68;
h_lab_sh('HH_Q4')=0.55; h_lab_sh('HH_Q5')=0.35;
h_cap_sh(HH) = 1 - h_lab_sh(HH);

Parameter lab_inc(HH), cap_inc(HH), tot_lab2, tot_cap2;
lab_inc(HH) = h_lab_sh(HH) * YZ(HH);
cap_inc(HH) = (1 - h_lab_sh(HH)) * YZ(HH);
tot_lab2 = sum(HH, lab_inc(HH));
tot_cap2 = sum(HH, cap_inc(HH));
tot_lab2$(tot_lab2=0) = 1;
tot_cap2$(tot_cap2=0) = 1;
fac_lab_sh(HH) = lab_inc(HH) / tot_lab2;
fac_cap_sh(HH) = cap_inc(HH) / tot_cap2;

* mps — empirical override (NSO SES 2021/2023)
mps('HH_Q1')=0.02; mps('HH_Q2')=0.05; mps('HH_Q3')=0.10;
mps('HH_Q4')=0.18; mps('HH_Q5')=0.32;

Parameter chk_lab, chk_cap;
chk_lab = sum(HH, fac_lab_sh(HH));
chk_cap = sum(HH, fac_cap_sh(HH));
Display fac_lab_sh, fac_cap_sh, chk_lab, chk_cap, mps;

*==============================================================================
* SECTION 3K: Rubber income shares + NESDC LFS override
*==============================================================================
* rub_inc_sh(HH) = สัดส่วนรายได้แรงงานของ HH ที่มาจากภาคยางพารา
* ที่มา: ทะเบียน กยท. 1.1M ราย / HH ไทย 22M ≈ 5% ของ HH ทั้งหมด
*        NSO SES 2023: Q1 มีชาวสวนยาง ~18% → สัดส่วนรายได้ Q1 จากยาง ≈ 30%
rub_inc_sh('HH_Q1') = 0.30;
rub_inc_sh('HH_Q2') = 0.20;
rub_inc_sh('HH_Q3') = 0.08;
rub_inc_sh('HH_Q4') = 0.02;
rub_inc_sh('HH_Q5') = 0.005;

* ★ NESDC LFS 2023 override — SAM ให้ LAB share เท่ากันทุก quintile (37.3%)
*    ซึ่งผิดจากความเป็นจริง ปรับให้ตรงกับโครงสร้างรายได้จาก NSO SES 2021/2023
fac_lab_sh('HH_Q1') = 0.085;
fac_lab_sh('HH_Q2') = 0.142;
fac_lab_sh('HH_Q3') = 0.196;
fac_lab_sh('HH_Q4') = 0.240;
fac_lab_sh('HH_Q5') = 0.337;
* sum = 1.000 ✓  (ตาม NESDC LFS 2023 labor income distribution)

* Re-derive fac_cap_sh จาก NSO SES 2021 (capital/property income by quintile)
fac_cap_sh('HH_Q1') = 0.025;
fac_cap_sh('HH_Q2') = 0.055;
fac_cap_sh('HH_Q3') = 0.120;
fac_cap_sh('HH_Q4') = 0.280;
fac_cap_sh('HH_Q5') = 0.520;
* sum = 1.000 ✓  (ตาม NSO SES 2021 capital income distribution)

Parameter chk_lab2, chk_cap2;
chk_lab2 = sum(HH, fac_lab_sh(HH));
chk_cap2 = sum(HH, fac_cap_sh(HH));
Display "=== 3K: rub_inc_sh + NESDC override ===";
Display rub_inc_sh, fac_lab_sh, fac_cap_sh, chk_lab2, chk_cap2;
* ★ chk_lab2 = 1.000 ✓   chk_cap2 = 1.000 ✓
*   Display fac_lab_sh, fac_cap_sh, chk_lab, chk_cap, mps;
* ════════════════════════════════════════════════════════════════════

*==============================================================================
* ★ DIAGNOSTIC 1 — SAM unit, rubber sector size, income linkage
*==============================================================================
Parameter
  diag_rub_xd_share      "Rubber sector share of total output (%)"
  diag_rub_lab_share     "Rubber labor share of total labor (%)"
  diag_sam_unit_gdp      "..."
  diag_lsz_labor         "..."
  diag_yz_per_hh(HH)     "implied income per HH (YZ ÷ 4.4M)"
  diag_rubber_hh_link(HH) "..."
;
Set RUBACT_D(ACT) / act26, act27, act28, act29 /;
diag_rub_xd_share  = sum(ACT$RUBACT_D(ACT), XDZ(ACT)) / sum(ACT, XDZ(ACT)) * 100;
diag_rub_lab_share = sum(ACT$RUBACT_D(ACT), LZ(ACT))  / sum(ACT, LZ(ACT)) * 100;
diag_sam_unit_gdp  = sum(HH, YZ(HH)) / 17e6;
diag_lsz_labor     = LSZ / 1e6;
diag_yz_per_hh(HH) = YZ(HH) / 4.4e6;
diag_rubber_hh_link(HH) = fac_lab_sh(HH) * diag_rub_lab_share / 100;

Display "=== DIAG 1: SAM unit & rubber linkage ===";
Display diag_rub_xd_share, diag_rub_lab_share;
* ★ คาดหวัง: rubber ≈ 8-12% ของ XDZ | ≈ 10-15% ของ LZ
* ★ ถ้า < 3% → crowding-in ชนะ price shock → dY บวกผิดปกติ
Display diag_sam_unit_gdp;
* ★ ≈ 1.0 = SAM หน่วยล้านบาท ✓  ← ใช้ตรวจ SAM_unit สำหรับ S12 transfer
Display diag_lsz_labor;
* ★ ≈ 40 = แรงงานล้านคน  ถ้า ≈ 5400 = income units → UNEMP ไม่ใช่จำนวนคน
Display diag_yz_per_hh;
* ★ Q1 ควรได้ ~0.0002 ถ้าหน่วยล้านบาท (= 200,000 บาท/ปี/HH)
Display fac_lab_sh, fac_cap_sh;
* ★ fac_lab_sh(Q1) = 0.062 → Q1 ได้ 6.2% ของ labor income รวม
* ★ ควร ≥ 0.08 (NESDC LFS 2023 พบ Q1 ≈ 8-10%)
Display diag_rubber_hh_link;
* ★ diag_rubber_hh_link(Q1) = fac_lab_sh(Q1) × rub_lab_share%
* ★ ถ้า < 0.003 → rubber price shock แทบไม่ถึง Q1 → ต้องเพิ่ม direct channel
Display "=== END DIAG 1 ===";


*==============================================================================
* SECTION 4: ELASTICITIES & SHARES (CES, CET, ARMINGTON)
*==============================================================================
*--- 4A: Armington (sigmaA)
sigmaA('com1')=2.80; sigmaA('com2')=2.44; sigmaA('com3')=2.80;
sigmaA('com4')=2.80; sigmaA('com5')=5.60; sigmaA('com6')=2.60;
sigmaA('com7')=2.60; sigmaA('com8')=5.56; sigmaA('com9')=4.00;
sigmaA('com10')=5.30; sigmaA('com11')=5.20; sigmaA('com12')=4.40;
sigmaA('com13')=4.00; sigmaA('com14')=5.60; sigmaA('com15')=4.40;
sigmaA('com16')=5.40; sigmaA('com17')=4.40; sigmaA('com18')=2.80;
sigmaA('com19')=1.90; sigmaA('com20')=1.90; sigmaA('com21')=1.90;
sigmaA('com22')=1.90; sigmaA('com23')=1.90; sigmaA('com24')=1.90;
sigmaA('com25')=1.90;
sigmaA('com26')=3.50; sigmaA('com27')=4.40;
sigmaA('com28')=4.40; sigmaA('com29')=4.40;

*--- 4B: CET
useCET(COM)=0;
useCET('com10')=1; useCET('com11')=1; useCET('com12')=1;
useCET('com14')=1; useCET('com15')=1; useCET('com16')=1;
useCET('com26')=1; useCET('com27')=1; useCET('com28')=1; useCET('com29')=1;

sigmaT(COM)=2.0;
sigmaT('com10')=2.5; sigmaT('com11')=2.5; sigmaT('com12')=2.5;
sigmaT('com14')=3.0; sigmaT('com15')=3.0; sigmaT('com16')=3.0;
*sigmaT('com26')=2.0; sigmaT('com27')=3.0;
sigmaT('com26')=2.0; sigmaT('com27')=1.5;
*sigmaT('com28')=3.0; sigmaT('com29')=2.5;
sigmaT('com28')=1.5; sigmaT('com29')=2.0;
gammaT(COM)=0.3; aT(COM)=1.0;

*--- 4C: CES production
sigmaF('act1')=0.30; sigmaF('act2')=0.30; sigmaF('act3')=0.25;
sigmaF('act4')=0.25; sigmaF('act5')=0.20;
sigmaF('act6')=1.12; sigmaF('act7')=1.12; sigmaF('act8')=1.26;
sigmaF('act9')=1.10; sigmaF('act10')=1.20; sigmaF('act11')=1.20;
sigmaF('act12')=1.20; sigmaF('act13')=1.20; sigmaF('act14')=1.20;
sigmaF('act15')=1.20; sigmaF('act16')=1.26; sigmaF('act17')=1.20;
sigmaF('act18')=1.26; sigmaF('act19')=1.40; sigmaF('act20')=1.20;
sigmaF('act21')=1.20; sigmaF('act22')=1.20; sigmaF('act23')=1.20;
sigmaF('act24')=1.20; sigmaF('act25')=0.99;
sigmaF('act26')=0.50; sigmaF('act27')=0.80;
sigmaF('act28')=1.10; sigmaF('act29')=1.10;
sigmaF(ACT)$(sigmaF(ACT)=0) = 0.50;
sigmaF(ACT)$(sigmaF(ACT)=1) = 0.999;
sigmaA(COM)$(sigmaA(COM)=1) = 0.999;
tk(ACT)=0; tl(ACT)=0;

*--- 4D: Calibrate gammaF, aF
Parameter KL_ratio(ACT), cost_idx(ACT);
KL_ratio(ACT)$(LZ(ACT)>0) = KZ(ACT)/LZ(ACT);
gammaF(ACT)$(KL_ratio(ACT)>0) =
    KL_ratio(ACT)**(1/sigmaF(ACT)) /
    (1 + KL_ratio(ACT)**(1/sigmaF(ACT)));
gammaF(ACT)$(gammaF(ACT)<=0.01)=0.01;
gammaF(ACT)$(gammaF(ACT)>=0.99)=0.99;
cost_idx(ACT) = gammaF(ACT)**sigmaF(ACT) + (1-gammaF(ACT))**sigmaF(ACT);
aF(ACT)$(KZ(ACT)>0) =
    (XDZ(ACT)/KZ(ACT)) * gammaF(ACT)**sigmaF(ACT) *
    cost_idx(ACT)**(sigmaF(ACT)/(1-sigmaF(ACT)));
aF(ACT)$(aF(ACT)<=1e-10)=1e-10;

*--- 4E: Calibrate gammaA, aA
Parameter XDD_base(COM), X_base(COM), MXratio(COM), cost_idxA(COM);
XDD_base(COM) = sum(ACT$MAP(COM,ACT), XDZ(ACT)) - EZ(COM);
EZ(COM)$(XDD_base(COM)<=0) = sum(ACT$MAP(COM,ACT), XDZ(ACT)) - 1e-6;
XDD_base(COM) = max(1e-6, sum(ACT$MAP(COM,ACT), XDZ(ACT)) - EZ(COM));
EZ(COM)$(EZ(COM)>=0.999*sum(ACT$MAP(COM,ACT), XDZ(ACT))) =
    0.950*sum(ACT$MAP(COM,ACT), XDZ(ACT));
XDD_base(COM) = max(0, sum(ACT$MAP(COM,ACT), XDZ(ACT)) - EZ(COM));
X_base(COM)   = MZ(COM) + XDD_base(COM);
MXratio(COM)$(XDD_base(COM)>0) = MZ(COM)/XDD_base(COM);

gammaA(COM)$(MXratio(COM)>0) =
    MXratio(COM)**(1/sigmaA(COM)) /
    (1 + MXratio(COM)**(1/sigmaA(COM)));
gammaA(COM)$(gammaA(COM)<=0.01)=0.01;
gammaA(COM)$(gammaA(COM)>=0.99)=0.99;
cost_idxA(COM) = gammaA(COM)**sigmaA(COM) + (1-gammaA(COM))**sigmaA(COM);
aA(COM)$(MZ(COM)>0) =
    (X_base(COM)/MZ(COM)) * gammaA(COM)**sigmaA(COM) *
    cost_idxA(COM)**(sigmaA(COM)/(1-sigmaA(COM)));
aA(COM)$(aA(COM)<=1e-10)=1e-10;
useARM(COM) = 1$(MZ(COM) > 1e-4);

*--- 4F: CET calibration
Parameter EX_ratio(COM);
EX_ratio(CETCOM)$(XDD_base(CETCOM)>0) = EZ(CETCOM)/XDD_base(CETCOM);
gammaT(CETCOM)$(EX_ratio(CETCOM)>0) =
    EX_ratio(CETCOM)**(1/sigmaT(CETCOM)) /
    (1 + EX_ratio(CETCOM)**(1/sigmaT(CETCOM)));
gammaT(CETCOM)$(gammaT(CETCOM)<=0.01)=0.01;
gammaT(CETCOM)$(gammaT(CETCOM)>=0.99)=0.99;


*==============================================================================
* SECTION 5: VARIABLES & EQUATIONS
*==============================================================================
Variables
  PK, PL, ER, PCINDEX
  P(COM), PM(COM), PE(COM), PDD(COM), PD(ACT)
  XD(ACT), K(ACT), L(ACT)
   M(COM), E(COM), XDD(COM),  X(COM)
  CH(COM,HH), Y(HH), SH(HH), CBUD(HH)
  INV(COM), CG(COM), KG, LG,
  TAXR, TRF, S, SF, UNEMP
;

Positive Variables
  P, PM, PE, PDD, PD,
  XD, K, L, XDD,
  CG, KG, LG,
  TAXR, TRF, S
;

* Free Variables — ไม่มี LO bound จาก declaration
Free Variables
  PK, PL, ER, PCINDEX,
  SF, UNEMP,
  E(COM),
  M(COM),
  X(COM),
  CH(COM,HH),
  Y(HH),
  SH(HH),
  CBUD(HH),
  INV(COM)
;

Equations
  EQC(COM,HH), EQSH(HH), EQINCOME(HH), EQCBUD(HH)
  EQK(ACT), EQL(ACT), EQPROFIT(ACT)
  EQS, EQI(COM), EQCG(COM), EQKG, EQLG
  EQTAXREV, EQTRANSFER
  EQEXPORT(COM), EQXDD(COM), EQPROFITT(COM)
  EQIMPORT(COM), EQARMD(COM), EQPROFITA(COM)
  EQMARKETL, EQMARKETK, EQMARKETC(COM)
  EQIMPRICE(COM), EQEXPRICE(COM), EQPCINDEX
  EQPHILLIPS
;

*--- Household
EQC(COM,HH)..
    (1+tc(COM))*P(COM)*CH(COM,HH) =E= betaC(COM,HH)*CBUD(HH);

EQSH(HH)..
    SH(HH) =E= mps(HH)*(1-ty(HH))*Y(HH);

EQINCOME(HH)..
    Y(HH) =E=
*   --- Non-rubber labor income ---
        fac_lab_sh(HH) * (1 - rub_inc_sh(HH))
            * PL * (LSZ - UNEMP)
*   --- Rubber-specific labor income (scales with rubber export price index) ---
      + fac_lab_sh(HH) * rub_inc_sh(HH)
            * PL * (LSZ - UNEMP)
            * (  sum(COM$RUBCOM(COM), PE(COM)*E(COM))
               / max(1e-10, sum(COM$RUBCOM(COM), PWEZ(COM)*ERZ*EZ(COM)))  )
*   --- Capital income ---
      + fac_cap_sh(HH) * PK * KSZ
*   --- S12: Targeted government transfer (switch=1 only during S12) ---
      + transfer_switch * GOV_TRANSFER_RUB(HH);

EQCBUD(HH)..
    CBUD(HH) =E= (1-ty(HH))*Y(HH) - SH(HH);

*--- Production (CES Value Added)
EQK(ACT)..
    K(ACT) =E= (XD(ACT)/aF(ACT)) *
    (gammaF(ACT)/((1+tk(ACT))*PK))**sigmaF(ACT) *
    (gammaF(ACT)**sigmaF(ACT)*((1+tk(ACT))*PK)**(1-sigmaF(ACT))
    +(1-gammaF(ACT))**sigmaF(ACT)*((1+tl(ACT))*PL)**(1-sigmaF(ACT))
    )**(sigmaF(ACT)/(1-sigmaF(ACT)));

EQL(ACT)..
    L(ACT) =E= (XD(ACT)/aF(ACT)) *
    ((1-gammaF(ACT))/((1+tl(ACT))*PL))**sigmaF(ACT) *
    (gammaF(ACT)**sigmaF(ACT)*((1+tk(ACT))*PK)**(1-sigmaF(ACT))
    +(1-gammaF(ACT))**sigmaF(ACT)*((1+tl(ACT))*PL)**(1-sigmaF(ACT))
    )**(sigmaF(ACT)/(1-sigmaF(ACT)));

EQPROFIT(ACT)..
    PD(ACT)*XD(ACT) =E= (1+tk(ACT))*PK*K(ACT)
                      + (1+tl(ACT))*PL*L(ACT)
                      + sum(COM, io(COM,ACT)*XD(ACT)*PDD(COM));

*--- Saving & Investment
EQS..
    S =E= sum(HH, SH(HH)) + (TAXR-TRF) + SF*ER;

EQI(COM)..
    P(COM)*INV(COM) =E= alphaI(COM)*S;

*--- Government
EQCG(COM)..  P(COM)*CG(COM) =E= alphaCG(COM)*(TAXR-TRF);
EQKG..       PK*KG =E= alphaKG*(TAXR-TRF);
EQLG..       PL*LG =E= alphaLG*(TAXR-TRF);

*--- Tax & Transfers
EQTAXREV..
    TAXR =E= sum(HH, ty(HH)*Y(HH))
           + sum(COM, tc(COM)*P(COM)*sum(HH, CH(COM,HH)))
           + sum(ACT, tk(ACT)*PK*K(ACT) + tl(ACT)*PL*L(ACT))
           + sum(COM, tm(COM)*ER*PWMZ(COM)*M(COM))
           - transfer_switch * sum(HH, GOV_TRANSFER_RUB(HH));
*          ↑ S12: fiscal cost of targeted transfer (switch=0 in all other scenarios)

EQTRANSFER..
    TRF =E= trep*PL*UNEMP + TRO*PCINDEX;

*--- Prices
EQIMPRICE(COM).. PM(COM) =E= (1+tm(COM))*ER*PWMZ(COM);
EQEXPRICE(COM).. PE(COM) =E= PWEZ(COM)*ER;

*--- CET Export supply
EQEXPORT(COM)..
    E(COM) =E= EZ(COM)*(PE(COM)/PDD(COM))**(sigmaT(COM)*useCET(COM));

EQXDD(COM)..
    XDD(COM) =E= sum(ACT$MAP(COM,ACT), XD(ACT)) - E(COM);

EQPROFITT(COM)..
    PDD(COM) =E= sum(ACT$MAP(COM,ACT), PD(ACT));

*--- Armington Import demand
EQIMPORT(COM)..
    M(COM) =E= useARM(COM) *
    (
      (X(COM)/aA(COM)) *
      (gammaA(COM)/PM(COM))**sigmaA(COM) *
      (gammaA(COM)**sigmaA(COM)*PM(COM)**(1-sigmaA(COM))
      +(1-gammaA(COM))**sigmaA(COM)*PDD(COM)**(1-sigmaA(COM))
      )**(sigmaA(COM)/(1-sigmaA(COM)))
    )
    + (1-useARM(COM)) * MZ(COM);

EQARMD(COM)..
    XDD(COM) =E= useARM(COM) *
    (
      (X(COM)/aA(COM)) *
      ((1-gammaA(COM))/PDD(COM))**sigmaA(COM) *
      (gammaA(COM)**sigmaA(COM)*PM(COM)**(1-sigmaA(COM))
      +(1-gammaA(COM))**sigmaA(COM)*PDD(COM)**(1-sigmaA(COM))
      )**(sigmaA(COM)/(1-sigmaA(COM)))
    )
    + (1-useARM(COM)) * XDD_base(COM);

EQPROFITA(COM)..
    P(COM)*X(COM) =E= PM(COM)*M(COM) + PDD(COM)*XDD(COM);

*--- CPI
EQPCINDEX..
    PCINDEX =E= sum((COM,HH), (1+tc(COM))*P(COM)*CZ(COM,HH)) / CPI_denom;

*--- Markets
EQMARKETL..
    sum(ACT, L(ACT)) + LG =E= LSZ - UNEMP;

EQMARKETK..
    sum(ACT, K(ACT)) + KG =E= KSZ;

EQMARKETC(COM)..
    sum(HH, CH(COM,HH)) + INV(COM) + CG(COM)
    + sum(ACT, io(COM,ACT)*XD(ACT)) =E= X(COM);

*--- Phillips Curve
EQPHILLIPS..
    PL / PCINDEX =E= (PLZ/PCINDEXZ) *
    (1 + phillips * (UNEMPZ/max(1e-10,UNEMP) - 1));

*==============================================================================
* SECTION 6: MODEL DEFINITION
*==============================================================================
Model RubberDynamic / all /;

Model RubberScenario / RubberDynamic - EQPHILLIPS /;

* Default bounds หลัง model declaration
* (Free Variables ด้านบนไม่มี LO อยู่แล้ว — ตั้งได้อิสระ)
M.LO(COM)     = 0;
INV.LO(COM)   = 0;
CH.LO(COM,HH) = 0;
Y.LO(HH)      = 0;
SH.LO(HH)     = 0;
CBUD.LO(HH)   = 0;
E.LO(COM)     = 0;
UNEMP.L       = max(1, UNEMPZ);
PK.LO     = 1e-4;
PL.LO     = 1e-4;
ER.LO     = 1e-4;
PCINDEX.LO = 1e-4;


*==============================================================================
* SECTION 7: BASE YEAR INITIALIZATION & SOLVE (ปี 2024)
*==============================================================================
*--- 7A: Prices
PK.L=1; PL.L=1; ER.L=1; PCINDEX.L=1;
P.L(COM)=1; PM.L(COM)=1; PE.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1;
PM.L(COM) = (1+tm(COM))*ERZ*PWMZ(COM);

*--- 7B: Quantities
XD.L(ACT) = XDZ(ACT);
K.L(ACT)  = KZ(ACT);
L.L(ACT)  = LZ(ACT);

*--- 7C: Household
Y.L(HH)    = YZ(HH);
SH.L(HH)   = mps(HH)*(1-ty(HH))*YZ(HH);
CBUD.L(HH) = (1-ty(HH))*YZ(HH) - SH.L(HH);

*--- 7D: Trade
E.L(COM)   = EZ(COM);
XDD.L(COM) = XDD_base(COM);
M.L(COM)   = MZ(COM);
X.L(COM)   = X_base(COM);
UNEMP.L    = UNEMPZ;

*--- 7E: Fiscal
TAXR.L = sum(HH, ty(HH)*YZ(HH))
       + sum(COM, tc(COM)*sum(HH, CZ(COM,HH)))
       + sum(COM, tm(COM)*ERZ*PWMZ(COM)*MZ(COM));
TAXR.L$(TAXR.L<=0) = 1e-6;
TRF.L = 0;
S.L   = sum(HH, SH.L(HH)) + TAXR.L + SFZ;
S.L$(S.L<=0) = 1e-6;
INV.L(COM) = alphaI(COM)*S.L;
CG.L(COM)  = alphaCG(COM)*TAXR.L;
CG.L(COM)$(CG.L(COM)<=0) = 1e-6;
KG.L=0; KG.LO=0;
LG.L=0; LG.LO=0;
SF.L = SFZ;

*--- 7F: Zero-consumption fix
Parameter strict_zero(COM,HH);
strict_zero(COM,HH) = 1$(SAM_data(COM,HH) = 0);
betaC(COM,HH)$strict_zero(COM,HH) = 0;
CH.L(COM,HH)$(not strict_zero(COM,HH)) = max(1e-8, CZ(COM,HH));
CH.L(COM,HH)$strict_zero(COM,HH)       = 0;
* ไม่ตั้ง CH.LO ตรงนี้ — ใช้ค่าจาก Section 6 (=0)

*--- 7G: Bounds
ER.FX = ERZ;
SF.FX = SFZ;
PK.LO=1e-4;      PK.UP=1e4;
PL.LO=1e-4;      PL.UP=1e4;
PCINDEX.LO=1e-4;
P.LO(COM)=1e-4;  PDD.LO(COM)=1e-4;  PD.LO(ACT)=1e-4;
XD.LO(ACT)=1e-6; X.LO(COM)=0;
XDD.LO(COM)=-INF;
E.L(COM)=EZ(COM);
TAXR.LO=1e-6;    S.LO=1e-6;
UNEMP.L  = LSZ * 0.01;
UNEMP.LO = 0;
UNEMP.UP = LSZ;

Display UNEMP.L, UNEMP.LO, UNEMP.UP, LSZ, UNEMPZ;
Display X.LO;

*--- 7H: Solve
Solve RubberDynamic using CNS;

*--- 7I: Store base results
YZ_base(HH)        = Y.L(HH);
E_base2(COM)       = E.L(COM);
XD_base3(ACT)      = XD.L(ACT);
PL_base2           = PL.L;
PK_base2           = PK.L;
UNEMP_base2        = UNEMP.L;
TAXR_base2         = TAXR.L;
L_path(ACT,'2024') = L.L(ACT);

Display "=== BASE YEAR SOLVED ===";
Display PL.L, PK.L, UNEMP.L, TAXR.L;

* ════════════════════════════════════════════════════════════════════
* BLOCK 2 — แทรกหลังบรรทัดนี้ใน SECTION 7I:
*   Display PL.L, PK.L, UNEMP.L, TAXR.L;
* ════════════════════════════════════════════════════════════════════

*==============================================================================
* ★ DIAGNOSTIC 2 — Post-base-solve structure
*==============================================================================
Parameter
  diag2_rub_rev_total    "Rubber revenue (PD×XD, rubber sectors)"
  diag2_gdp_sum          "GDP-at-producer-prices (sum PD×XD)"
  diag2_rub_gdp_pct      "Rubber % of GDP"
  diag2_pk_impact(HH)    "dY/Y(%) if PK rises 7.4% (= S1a result)"
  diag2_unemp_rate       "Baseline unemployment rate (%)"
  diag2_eqincome_res(HH) "..."
  rub_rev_base           "Base rubber export revenue (denominator for rubber income ratio)"
;
Set RUBACT_D2(ACT) / act26, act27, act28, act29 /;
diag2_rub_rev_total = sum(ACT$RUBACT_D2(ACT), PD.L(ACT)*XD.L(ACT));
diag2_gdp_sum       = sum(ACT, PD.L(ACT)*XD.L(ACT));
diag2_rub_gdp_pct   = diag2_rub_rev_total / max(1,diag2_gdp_sum) * 100;
diag2_pk_impact(HH) = fac_cap_sh(HH) * 0.074 * PK.L * KSZ / max(1,Y.L(HH)) * 100;
diag2_unemp_rate    = UNEMP.L / max(1,LSZ) * 100;

* ★ rubber price ratio at base = 1.0 (PE=PWEZ*ER=1, EZ=base exports)
rub_rev_base = sum(COM$RUBCOM(COM), PWEZ(COM)*ERZ*EZ(COM));

diag2_eqincome_res(HH) = Y.L(HH)
    - fac_lab_sh(HH)*(1-rub_inc_sh(HH))*PL.L*(LSZ-UNEMP.L)
    - fac_lab_sh(HH)*rub_inc_sh(HH)*PL.L*(LSZ-UNEMP.L)
        * (sum(COM$RUBCOM(COM), PE.L(COM)*E.L(COM)) / max(1e-10, rub_rev_base))
    - fac_cap_sh(HH)*PK.L*KSZ;

Display "=== DIAG 2: Post-Solve Structure ===";
Display diag2_rub_rev_total, diag2_gdp_sum, diag2_rub_gdp_pct;
* ★ คาดหวัง rubber ≈ 8-12%
* ★ ถ้า < 5% = ROOT CAUSE ของ crowding-in ที่ทำให้ dY บวกทั้งหมด
Display diag2_pk_impact;
* ★ ถ้า PK ขึ้น 7.4% → HH income เปลี่ยน diag2_pk_impact(HH) %
* ★ ถ้าทุก quintile ≈ กัน → model ถูก driven โดย capital เป็นหลัก → ผิดพลาด
Display diag2_unemp_rate;
* ★ ควร 1-5%   ถ้า ≈ 0 หรือ < 0 → over-employment
Display diag2_eqincome_res;
* ★ ควร = 0.000  ถ้าไม่ใช่ → calibration ผิดหรือ equation ไม่ consistent
Display PK.L, PL.L, KSZ, LSZ;
Display "=== END DIAG 2 ===";

*==============================================================================
* SECTION 8: RECURSIVE DYNAMIC LOOP
*==============================================================================
*--- Price path for EV (declare before baseline loop)
Parameter P_path_base2(COM,T) "Baseline commodity price path";

*--- 8A: Initialize dynamic stocks
KZ_t(ACT,'2024') = KZ(ACT);
LSZ_t('2024')    = LSZ;
aF_t(ACT,'2024') = aF(ACT);
aF_base(ACT)     = aF(ACT);

* [FIX-K] Initialize K_prev = base year capital stock
K_prev(ACT) = KZ(ACT);

* [FIX-K] Store baseline K path for T=2024 (used as scenario bound anchor)
K_path(ACT,'2024') = KZ(ACT);

Y_path(HH,'2024')      = Y.L(HH);
XD_path(ACT,'2024')    = XD.L(ACT);
E_path(COM,'2024')     = E.L(COM);
M_path(COM,'2024')     = M.L(COM);
PL_path('2024')        = PL.L;
PK_path('2024')        = PK.L;
UNEMP_path('2024')     = UNEMP.L;
TAXR_path('2024')      = TAXR.L;
S_path('2024')         = S.L;
PCINDEX_path('2024')   = PCINDEX.L;
CH_path(COM,HH,'2024') = CH.L(COM,HH);
INV_path(COM,'2024')   = INV.L(COM);

*--- 8B: Main Dynamic Loop
Loop(T$(not TFIRST(T)),

  Display T;

* STEP 1: Capital accumulation
  KZ_t(ACT,T) = KZ_t(ACT,T-1)*(1-deprate(ACT))
              + sum(COM$MAP(COM,ACT), kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0) = 1e-6;

* STEP 2: Labor supply
  LSZ_t(T) = LSZ_t(T-1)*(1+popgrow_t(T));

* STEP 3: TFP
  aF_t(ACT,T) = aF_t(ACT,T-1)*(1+tfpgrow(ACT));

* STEP 4: Feed stocks into parameters
  KSZ     = sum(ACT, KZ_t(ACT,T));
  LSZ     = LSZ_t(T);
  aF(ACT) = aF_t(ACT,T);

* STEP 5: Warm-start
  K.L(ACT)     = max(1e-6, KZ_t(ACT,T));
  L.L(ACT)     = max(1e-6, L_path(ACT,T-1));
  XD.L(ACT)    = max(1e-6, XD_path(ACT,T-1));
  Y.L(HH)      = Y_path(HH,T-1);
  CH.L(COM,HH) = CH_path(COM,HH,T-1);
  INV.L(COM)   = INV_path(COM,T-1);
  M.L(COM)     = max(1e-6, M_path(COM,T-1));
  UNEMP.L      = max(UNEMP_path(T-1), 1e-4);

* STEP 6: Relax bounds ก่อน solve
  M.LO(COM)=-INF;    INV.LO(COM)=-INF;   CH.LO(COM,HH)=-INF;
  Y.LO(HH)=-INF;     CBUD.LO(HH)=-INF;   SH.LO(HH)=-INF;
  XDD.LO(COM)=-INF;  X.LO(COM)=0;        E.LO(COM)=-INF;
  XD.LO(ACT)=1e-6;   L.LO(ACT)=0;        K.LO(ACT)=0;
  PDD.LO(COM)=1e-6;  PD.LO(ACT)=1e-6;    P.LO(COM)=1e-6;
  PM.LO(COM)=1e-6;   PE.LO(COM)=1e-6;
  PK.LO=0;           PL.LO=0;            PCINDEX.LO=0;
  PK.UP=+INF;        PL.UP=+INF;
  CG.LO(COM)=0;      TAXR.LO=-INF;       S.LO=-INF;    TRF.LO=-INF;
  UNEMP.LO=-INF;     UNEMP.UP=LSZ;
* [FIX-K] Baseline K is FREE — no bounds here, let equations determine K naturally

  RubberDynamic.tolinfeas = 1e-6;
  RubberDynamic.iterlim   = 10000;

  Solve RubberDynamic using CNS;

  Display RubberDynamic.modelstat, RubberDynamic.solvestat;
  if(RubberDynamic.modelstat <> 16,
    Display "INFEASIBLE at T=", T;
    Display XD.L, PL.L, PK.L, UNEMP.L, Y.L;
    abort "...";
  );

* STEP 6b: Restore bounds หลัง solve
  M.LO(COM)=0;       INV.LO(COM)=0;      CH.LO(COM,HH)=0;
  Y.LO(HH)=0;        CBUD.LO(HH)=0;      SH.LO(HH)=0;
  X.LO(COM)=0;       E.LO(COM)=0;
  XD.LO(ACT)=1e-6;   L.LO(ACT)=0;        K.LO(ACT)=1e-6;
  K.UP(ACT)=+INF;
  PDD.LO(COM)=1e-4;  PD.LO(ACT)=1e-4;    P.LO(COM)=1e-4;
  PM.LO(COM)=0;      PE.LO(COM)=0;
  PK.LO=1e-4;        PL.LO=1e-4;         PCINDEX.LO=1e-4;
  PK.UP=1e4;         PL.UP=1e4;
  CG.LO(COM)=0;      TAXR.LO=1e-6;       S.LO=1e-6;
  UNEMP.LO=0;        UNEMP.UP=LSZ;

* STEP 7: Store results
  Y_path(HH,T)      = Y.L(HH);
  XD_path(ACT,T)    = XD.L(ACT);
  E_path(COM,T)     = E.L(COM);
  M_path(COM,T)     = M.L(COM);
  PL_path(T)        = PL.L;
  PK_path(T)        = PK.L;
  UNEMP_path(T)     = max(0, UNEMP.L);
  TAXR_path(T)      = TAXR.L;
  S_path(T)         = S.L;
  PCINDEX_path(T)   = PCINDEX.L;
  CH_path(COM,HH,T) = CH.L(COM,HH);
  INV_path(COM,T)   = INV.L(COM);
  L_path(ACT,T)     = L.L(ACT);
  K_path(ACT,T)     = K.L(ACT);

* 1. ใน baseline dynamic loop (Section 8B) — เพิ่มหลัง STEP 7:
P_path_base2(COM,T) = P.L(COM);

);
* End Dynamic Loop

*--- Save LSZ0 หลัง baseline loop (declare ครั้งเดียว)
Scalar LSZ0 "Base labor supply year 2024";
LSZ0 = LSZ_t('2024');

* ════════════════════════════════════════════════════════════════════
* BLOCK 3 — แทรกหลังบรรทัดนี้ (หลัง baseline dynamic loop):
*   LSZ0 = LSZ_t('2024');
* ════════════════════════════════════════════════════════════════════

*==============================================================================
* ★ DIAGNOSTIC 3 — Baseline dynamics validity
*==============================================================================
Parameter
  diag3_unemp_rate(T)  "Baseline unemployment rate % by year"
  diag3_rub_idx(T)     "Rubber output index (2024=100)"
  diag3_rub_base       "Rubber base output denominator"
;
Set RUBACT_D3(ACT) / act26, act27, act28, act29 /;
diag3_rub_base = sum(ACT$RUBACT_D3(ACT), XD_path(ACT,'2024'));
diag3_rub_idx(T)$(diag3_rub_base>0)
    = sum(ACT$RUBACT_D3(ACT), XD_path(ACT,T)) / diag3_rub_base * 100;
diag3_unemp_rate(T)$(LSZ_t(T)>0)
    = UNEMP_path(T) / LSZ_t(T) * 100;

Display "=== DIAG 3: Baseline Dynamics ===";
Display diag3_unemp_rate;
* ★ ควรอยู่ที่ 1-5% ตลอด 2024-2034
* ★ ถ้าลดเหลือ 0 หรือ missing ที่ 2034 → UNEMP baseline loop มีปัญหา
Display diag3_rub_idx;
* ★ ควรขยายตัวตาม TFP ~0.5-1.2% ต่อปี (2034 ≈ 105-115)
Display "=== END DIAG 3 ===";


*
* ════════════════════════════════════════════════════════════════════
* FIX CODE — แทรก/แก้ใน SECTION 8 (ก่อน scenario loops)
* ════════════════════════════════════════════════════════════════════

*==============================================================================
* ★ FIX 1 (PRIORITY): RELAX_BOUNDS macro — เพิ่ม UNEMP.LO = 0
*    แก้ $macro RELAX_BOUNDS ที่มีอยู่แล้ว โดยเพิ่มบรรทัดสุดท้าย:
*    UNEMP.LO=0;  UNEMP.UP=LSZ;
*
*    และในทุก scenario loop ที่มี:
*      UNEMP.LO = -INF; UNEMP.UP = +INF;
*    เปลี่ยนเป็น:
*      UNEMP.LO = 0;    UNEMP.UP = LSZ;
*
*    ผลที่คาดหวัง: dY จะเปลี่ยนเป็นลบสำหรับ S1a/S2a
*    (เพราะ employment ไม่สามารถเกิน LSZ ได้อีกแล้ว)
*==============================================================================

*==============================================================================
* SCENARIO 1: RUBBER PRICE SHOCK
*==============================================================================
Parameter
  Y_path_base(HH,T)      "Baseline HH income"
  XD_path_base(ACT,T)    "Baseline sector output"
  UNEMP_path_base(T)     "Baseline unemployment"
  INV_path_base(COM,T)   "Baseline investment"
  CH_path_base(COM,HH,T) "Baseline consumption"
  Y_path_s1a(HH,T)       "S1a: price -20%"
  XD_path_s1a(ACT,T)
  UNEMP_path_s1a(T)
  Y_path_s1b(HH,T)       "S1b: price +20%"
  XD_path_s1b(ACT,T)
  UNEMP_path_s1b(T)
  dY_s1a(HH,T)           "% dev from baseline S1a"
  dY_s1b(HH,T)           "% dev from baseline S1b"
  dXD_s1a(ACT,T)
  dXD_s1b(ACT,T)
  PL_s1a(T)              "Wage in S1a"
  PK_s1a(T)              "Rental rate in S1a"
  dPL_s1a(T)             "% dev of wage S1a vs baseline"
  dPK_s1a(T)             "% dev of rental rate S1a vs baseline"
  PL_path_base(T)        "Baseline wage path for scenario anchor"
  Y_path_s2a(HH,T)       "S2a: price -30%"
  XD_path_s2a(ACT,T)
  UNEMP_path_s2a(T)
  Y_path_s2b(HH,T)       "S2b: price +30%"
  XD_path_s2b(ACT,T)
  UNEMP_path_s2b(T)
  dY_s2a(HH,T)           "% dev from baseline S2a"
  dY_s2b(HH,T)           "% dev from baseline S2b"
  dXD_s2a(ACT,T)
  dXD_s2b(ACT,T)
  PL_s2a(T)              "Wage in S2a"
  PK_s2a(T)              "Rental rate in S2a"
  dPL_s2a(T)             "% dev of wage S2a vs baseline"
  dPK_s2a(T)             "% dev of rental rate S2a vs baseline"
  PL_s2b(T)              "Wage in S2b"
  PK_s2b(T)              "Rental rate in S2b"
  dPL_s2b(T)             "% dev of wage S2b vs baseline"
  dPK_s2b(T)             "% dev of rental rate S2b vs baseline" 
;

* เก็บ baseline
Y_path_base(HH,T)      = Y_path(HH,T);
XD_path_base(ACT,T)    = XD_path(ACT,T);
UNEMP_path_base(T)     = UNEMP_path(T);
INV_path_base(COM,T)   = INV_path(COM,T);
CH_path_base(COM,HH,T) = CH_path(COM,HH,T);

* ★ Save baseline PL path สำหรับ anchor ใน scenario
PL_path_base(T) = PL_path(T);

* RELAX_BOUNDS — ไม่มี PL bounds เพราะ PL.FX จัดการแยก
$macro RELAX_BOUNDS \
  M.LO(COM)=-INF;    INV.LO(COM)=-INF;   CH.LO(COM,HH)=-INF; \
  Y.LO(HH)=-INF;     CBUD.LO(HH)=-INF;   SH.LO(HH)=-INF; \
  XDD.LO(COM)=-INF;  X.LO(COM)=0;        E.LO(COM)=-INF; \
  XD.LO(ACT)=1e-6;   L.LO(ACT)=0; \
  PDD.LO(COM)=1e-6;  PD.LO(ACT)=1e-6;    P.LO(COM)=1e-6; \
  PM.LO(COM)=1e-6;   PE.LO(COM)=1e-6; \
  PK.LO=1e-6;        PCINDEX.LO=1e-6; \
  PK.UP=+INF; \
  CG.LO(COM)=0;      TAXR.LO=-INF;       S.LO=-INF;    TRF.LO=-INF; \
*  UNEMP.LO=-INF;     UNEMP.UP=LSZ;
  UNEMP.LO=0.005*LSZ; UNEMP.UP=LSZ;

* RESTORE_BOUNDS — คืน PL bounds หลัง solve (unfix PL)
$macro RESTORE_BOUNDS \
  M.LO(COM)=0;       INV.LO(COM)=0;      CH.LO(COM,HH)=0; \
  Y.LO(HH)=0;        CBUD.LO(HH)=0;      SH.LO(HH)=0; \
  X.LO(COM)=0;       E.LO(COM)=0; \
  XD.LO(ACT)=1e-6;   L.LO(ACT)=0; \
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF; \
  PDD.LO(COM)=1e-4;  PD.LO(ACT)=1e-4;    P.LO(COM)=1e-4; \
  PK.LO=1e-4;        PL.LO=1e-4;         PCINDEX.LO=1e-4; \
  PK.UP=1e4;         PL.UP=1e4; \
  TAXR.LO=1e-6;      S.LO=1e-6; \
*  UNEMP.LO=0;        UNEMP.UP=LSZ;
  UNEMP.LO=0.005*LSZ; UNEMP.UP=LSZ;

$macro RESET_STOCKS \
  KZ_t(ACT,T)=0; LSZ_t(T)=0; aF_t(ACT,T)=0; \
  KZ_t(ACT,'2024')=KZ(ACT); \
  LSZ_t('2024')=LSZ0; \
  aF_t(ACT,'2024')=aF_base(ACT); \
  KSZ=sum(ACT,KZ_t(ACT,'2024')); LSZ=LSZ_t('2024'); aF(ACT)=aF_base(ACT);

RubberScenario.scaleopt  = 1;
RubberScenario.tolinfeas = 1e-6;
RubberScenario.iterlim   = 10000;
option nlp = conopt;
$onecho > conopt.opt
* [FIX 6] Removed unknown option "Lmfac" (not valid in CONOPT 4.37)
* Add valid CONOPT options here if needed, e.g.:
* MaxFuEval = 100000
$offecho
RubberScenario.optfile = 1;

*=== SCENARIO 1A: Rubber price -20% ===
PWEZ(RUBCOM) = 0.80;
RESET_STOCKS
* K_prev(ACT) = KZ(ACT);  * [FIX-K] reset K_prev to base capital at scenario start

K.L(ACT)     = KZ(ACT);
L.L(ACT)     = max(1e-6, L_path(ACT,'2024'));
XD.L(ACT)    = max(1e-6, XD_path_base(ACT,'2024'));
Y.L(HH)      = Y_path_base(HH,'2024');
CH.L(COM,HH) = CH_path_base(COM,HH,'2024');
INV.L(COM)   = INV_path_base(COM,'2024');
M.L(COM)     = max(1e-6, MZ(COM));

* T=2024: Fix PL = baseline → UNEMP determined by EQMARKETL
* Squareness: 574 eq - EQPHILLIPS, Fixed: ER SF PL → 577-3=574 ✓
UNEMP.LO = -INF; UNEMP.UP = LSZ;
UNEMP.L  = UNEMP_path_base('2024');
PL.FX    = PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
Solve RubberScenario using CNS;
RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

PL_s1a('2024') = PL.L;
PK_s1a('2024') = PK.L;
Y_path_s1a(HH,'2024')   = Y.L(HH);
XD_path_s1a(ACT,'2024') = XD.L(ACT);
UNEMP_path_s1a('2024')  = max(0, UNEMP.L);
INV_path(COM,'2024')     = INV.L(COM);
CH_path(COM,HH,'2024')   = CH.L(COM,HH);

Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T) = KZ_t(ACT,T-1)*(1-deprate(ACT))
              + sum(COM$MAP(COM,ACT), kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0) = 1e-6;
  LSZ_t(T)    = LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T) = aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);

  K.L(ACT)     = max(1e-6, KZ_t(ACT,T));
  L.L(ACT)     = max(1e-6, L_path(ACT,T-1));
  XD.L(ACT)    = max(1e-6, XD_path_s1a(ACT,T-1));
  Y.L(HH)      = Y_path_s1a(HH,T-1);
  CH.L(COM,HH) = CH_path(COM,HH,T-1);
  INV.L(COM)   = INV_path(COM,T-1);
  M.L(COM)     = max(1e-6, MZ(COM));

  UNEMP.LO = -INF; UNEMP.UP = LSZ;
  UNEMP.L  = UNEMP_path_base(T);
  PL.FX    = PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

  PL_s1a(T)  = PL.L;
  PK_s1a(T)  = PK.L;
  Y_path_s1a(HH,T)   = Y.L(HH);
  XD_path_s1a(ACT,T) = XD.L(ACT);
  UNEMP_path_s1a(T)  = max(0, UNEMP.L);
  INV_path(COM,T)     = INV.L(COM);
  CH_path(COM,HH,T)   = CH.L(COM,HH);
);

*=== SCENARIO 1B: Rubber price +20% ===
PWEZ(RUBCOM) = 1.20;
RESET_STOCKS
* K_prev(ACT) = KZ(ACT);  * [FIX-K] reset K_prev to base capital at scenario start

INV_path(COM,T)   = INV_path_base(COM,T);
CH_path(COM,HH,T) = CH_path_base(COM,HH,T);

K.L(ACT)     = KZ(ACT);
L.L(ACT)     = max(1e-6, L_path(ACT,'2024'));
XD.L(ACT)    = max(1e-6, XD_path_base(ACT,'2024'));
Y.L(HH)      = Y_path_base(HH,'2024');
CH.L(COM,HH) = CH_path_base(COM,HH,'2024');
INV.L(COM)   = INV_path_base(COM,'2024');
M.L(COM)     = max(1e-6, MZ(COM));

UNEMP.LO = -INF; UNEMP.UP = LSZ;
UNEMP.L  = UNEMP_path_base('2024');
PL.FX    = PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
Solve RubberScenario using CNS;
RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

Y_path_s1b(HH,'2024')   = Y.L(HH);
XD_path_s1b(ACT,'2024') = XD.L(ACT);
UNEMP_path_s1b('2024')  = max(0, UNEMP.L);
INV_path(COM,'2024')     = INV.L(COM);
CH_path(COM,HH,'2024')   = CH.L(COM,HH);

Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T) = KZ_t(ACT,T-1)*(1-deprate(ACT))
              + sum(COM$MAP(COM,ACT), kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0) = 1e-6;
  LSZ_t(T)    = LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T) = aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);

  K.L(ACT)     = max(1e-6, KZ_t(ACT,T));
  L.L(ACT)     = max(1e-6, L_path(ACT,T-1));
  XD.L(ACT)    = max(1e-6, XD_path_s1b(ACT,T-1));
  Y.L(HH)      = Y_path_s1b(HH,T-1);
  CH.L(COM,HH) = CH_path(COM,HH,T-1);
  INV.L(COM)   = INV_path(COM,T-1);
  M.L(COM)     = max(1e-6, MZ(COM));

  UNEMP.LO = -INF; UNEMP.UP = LSZ;
  UNEMP.L  = UNEMP_path_base(T);
  PL.FX    = PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

  Y_path_s1b(HH,T)   = Y.L(HH);
  XD_path_s1b(ACT,T) = XD.L(ACT);
  UNEMP_path_s1b(T)  = max(0, UNEMP.L);
  INV_path(COM,T)     = INV.L(COM);
  CH_path(COM,HH,T)   = CH.L(COM,HH);
);

PWEZ(RUBCOM) = 1.0;

*==============================================================================
* % DEVIATION FROM BASELINE
*==============================================================================
dY_s1a(HH,T)$(Y_path_base(HH,T)<>0) =
    (Y_path_s1a(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
dY_s1b(HH,T)$(Y_path_base(HH,T)<>0) =
    (Y_path_s1b(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
dXD_s1a(ACT,T)$(XD_path_base(ACT,T)<>0) =
    (XD_path_s1a(ACT,T)-XD_path_base(ACT,T))/abs(XD_path_base(ACT,T))*100;
dXD_s1b(ACT,T)$(XD_path_base(ACT,T)<>0) =
    (XD_path_s1b(ACT,T)-XD_path_base(ACT,T))/abs(XD_path_base(ACT,T))*100;

dPL_s1a(T)$(PL_path(T)>0) = (PL_s1a(T)-PL_path(T))/PL_path(T)*100;
dPK_s1a(T)$(PK_path(T)>0) = (PK_s1a(T)-PK_path(T))/PK_path(T)*100;

Display "=== S1A: Rubber price -20% ===";
Display dY_s1a, dXD_s1a, UNEMP_path_s1a, dPL_s1a, dPK_s1a;
Display "=== S1B: Rubber price +20% ===";
Display dY_s1b, dXD_s1b, UNEMP_path_s1b;

Execute_Unload "scenario1_results.gdx",
  Y_path_base, Y_path_s1a, Y_path_s1b,
  XD_path_base, XD_path_s1a, XD_path_s1b,
  UNEMP_path_base, UNEMP_path_s1a, UNEMP_path_s1b,
  dY_s1a, dY_s1b, dXD_s1a, dXD_s1b;

Display dY_s1a, dY_s1b;
Display dXD_s1a, dXD_s1b;
Display UNEMP_path_s1a, UNEMP_path_s1b;

*==============================================================================
* % DEVIATION FROM BASELINE
*==============================================================================
dY_s1a(HH,T)$(Y_path_base(HH,T)<>0) =
    (Y_path_s1a(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
dY_s1b(HH,T)$(Y_path_base(HH,T)<>0) =
    (Y_path_s1b(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
dXD_s1a(ACT,T)$(XD_path_base(ACT,T)<>0) =
    (XD_path_s1a(ACT,T)-XD_path_base(ACT,T))/abs(XD_path_base(ACT,T))*100;
dXD_s1b(ACT,T)$(XD_path_base(ACT,T)<>0) =
    (XD_path_s1b(ACT,T)-XD_path_base(ACT,T))/abs(XD_path_base(ACT,T))*100;
    
dPL_s1a(T)$(PL_path(T)>0) = (PL_s1a(T)-PL_path(T))/PL_path(T)*100;
dPK_s1a(T)$(PK_path(T)>0) = (PK_s1a(T)-PK_path(T))/PK_path(T)*100;

Display dPL_s1a, dPK_s1a;
Display UNEMP_path_base;

Display "=== S1A: Rubber price -20% ===";
Display dY_s1a, dXD_s1a, UNEMP_path_s1a;
Display "=== S1B: Rubber price +20% ===";
Display dY_s1b, dXD_s1b, UNEMP_path_s1b;

Execute_Unload "scenario1_results.gdx",
  Y_path_base, Y_path_s1a, Y_path_s1b,
  XD_path_base, XD_path_s1a, XD_path_s1b,
  UNEMP_path_base, UNEMP_path_s1a, UNEMP_path_s1b,
  dY_s1a, dY_s1b, dXD_s1a, dXD_s1b;

Display dY_s1a, dY_s1b;
Display dXD_s1a, dXD_s1b, dPL_s1a, PK_s1a;
Display UNEMP_path_s1a, UNEMP_path_s1b;

*=== SCENARIO 2A: Rubber price -30% ===
PWEZ(RUBCOM) = 0.70;
RESET_STOCKS
* K_prev(ACT) = KZ(ACT);  * [FIX-K] reset K_prev to base capital at scenario start

K.L(ACT)     = KZ(ACT);
L.L(ACT)     = max(1e-6, L_path(ACT,'2024'));
XD.L(ACT)    = max(1e-6, XD_path_base(ACT,'2024'));
Y.L(HH)      = Y_path_base(HH,'2024');
CH.L(COM,HH) = CH_path_base(COM,HH,'2024');
INV.L(COM)   = INV_path_base(COM,'2024');
M.L(COM)     = max(1e-6, MZ(COM));

UNEMP.LO = -INF; UNEMP.UP = LSZ;
UNEMP.L  = UNEMP_path_base('2024');
PL.FX    = PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
Solve RubberScenario using CNS;
RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

PL_s2a('2024') = PL.L;
PK_s2a('2024') = PK.L;
Y_path_s2a(HH,'2024')   = Y.L(HH);
XD_path_s2a(ACT,'2024') = XD.L(ACT);
UNEMP_path_s2a('2024')  = max(0, UNEMP.L);
INV_path(COM,'2024')     = INV.L(COM);
CH_path(COM,HH,'2024')   = CH.L(COM,HH);

Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T) = KZ_t(ACT,T-1)*(1-deprate(ACT))
              + sum(COM$MAP(COM,ACT), kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0) = 1e-6;
  LSZ_t(T)    = LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T) = aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);

  K.L(ACT)     = max(1e-6, KZ_t(ACT,T));
  L.L(ACT)     = max(1e-6, L_path(ACT,T-1));
  XD.L(ACT)    = max(1e-6, XD_path_s2a(ACT,T-1));
  Y.L(HH)      = Y_path_s2a(HH,T-1);
  CH.L(COM,HH) = CH_path(COM,HH,T-1);
  INV.L(COM)   = INV_path(COM,T-1);
  M.L(COM)     = max(1e-6, MZ(COM));

  UNEMP.LO = -INF; UNEMP.UP = LSZ;
  UNEMP.L  = UNEMP_path_base(T);
  PL.FX    = PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

  PL_s2a(T)  = PL.L;
  PK_s2a(T)  = PK.L;
  Y_path_s2a(HH,T)   = Y.L(HH);
  XD_path_s2a(ACT,T) = XD.L(ACT);
  UNEMP_path_s2a(T)  = max(0, UNEMP.L);
  INV_path(COM,T)     = INV.L(COM);
  CH_path(COM,HH,T)   = CH.L(COM,HH);
);

*=== SCENARIO 2B: Rubber price +30% ===
PWEZ(RUBCOM) = 1.30;
RESET_STOCKS
* K_prev(ACT) = KZ(ACT);  * [FIX-K] reset K_prev to base capital at scenario start

INV_path(COM,T)   = INV_path_base(COM,T);
CH_path(COM,HH,T) = CH_path_base(COM,HH,T);

K.L(ACT)     = KZ(ACT);
L.L(ACT)     = max(1e-6, L_path(ACT,'2024'));
XD.L(ACT)    = max(1e-6, XD_path_base(ACT,'2024'));
Y.L(HH)      = Y_path_base(HH,'2024');
CH.L(COM,HH) = CH_path_base(COM,HH,'2024');
INV.L(COM)   = INV_path_base(COM,'2024');
M.L(COM)     = max(1e-6, MZ(COM));

UNEMP.LO = -INF; UNEMP.UP = LSZ;
UNEMP.L  = UNEMP_path_base('2024');
PL.FX    = PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
Solve RubberScenario using CNS;
RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

PL_s2b('2024') = PL.L;
PK_s2b('2024') = PK.L;
Y_path_s2b(HH,'2024')   = Y.L(HH);
XD_path_s2b(ACT,'2024') = XD.L(ACT);
UNEMP_path_s2b('2024')  = max(0, UNEMP.L);
INV_path(COM,'2024')     = INV.L(COM);
CH_path(COM,HH,'2024')   = CH.L(COM,HH);

Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T) = KZ_t(ACT,T-1)*(1-deprate(ACT))
              + sum(COM$MAP(COM,ACT), kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0) = 1e-6;
  LSZ_t(T)    = LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T) = aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);

  K.L(ACT)     = max(1e-6, KZ_t(ACT,T));
  L.L(ACT)     = max(1e-6, L_path(ACT,T-1));
  XD.L(ACT)    = max(1e-6, XD_path_s2b(ACT,T-1));
  Y.L(HH)      = Y_path_s2b(HH,T-1);
  CH.L(COM,HH) = CH_path(COM,HH,T-1);
  INV.L(COM)   = INV_path(COM,T-1);
  M.L(COM)     = max(1e-6, MZ(COM));

  UNEMP.LO = -INF; UNEMP.UP = LSZ;
  UNEMP.L  = UNEMP_path_base(T);
  PL.FX    = PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  K.LO(ACT)=1e-6;    K.UP(ACT)=+INF;
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
* K_prev(ACT) = max(1e-6, K.L(ACT));              * [FIX-K] update K_prev for next period

  PL_s2b(T)  = PL.L;
  PK_s2b(T)  = PK.L;
  Y_path_s2b(HH,T)   = Y.L(HH);
  XD_path_s2b(ACT,T) = XD.L(ACT);
  UNEMP_path_s2b(T)  = max(0, UNEMP.L);
  INV_path(COM,T)     = INV.L(COM);
  CH_path(COM,HH,T)   = CH.L(COM,HH);
);

PWEZ(RUBCOM) = 1.0;

*==============================================================================
* % DEVIATION FROM BASELINE
*==============================================================================
dY_s2a(HH,T)$(Y_path_base(HH,T)<>0) =
    (Y_path_s2a(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
dY_s2b(HH,T)$(Y_path_base(HH,T)<>0) =
    (Y_path_s2b(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
dXD_s2a(ACT,T)$(XD_path_base(ACT,T)<>0) =
    (XD_path_s2a(ACT,T)-XD_path_base(ACT,T))/abs(XD_path_base(ACT,T))*100;
dXD_s2b(ACT,T)$(XD_path_base(ACT,T)<>0) =
    (XD_path_s2b(ACT,T)-XD_path_base(ACT,T))/abs(XD_path_base(ACT,T))*100;

dPL_s2a(T)$(PL_path(T)>0) = (PL_s2a(T)-PL_path(T))/PL_path(T)*100;
dPK_s2a(T)$(PK_path(T)>0) = (PK_s2a(T)-PK_path(T))/PK_path(T)*100;
dPL_s2b(T)$(PL_path(T)>0) = (PL_s2b(T)-PL_path(T))/PL_path(T)*100;
dPK_s2b(T)$(PK_path(T)>0) = (PK_s2b(T)-PK_path(T))/PK_path(T)*100;

Display "=== S2A: Rubber price -30% ===";
Display dY_s2a, dXD_s2a, UNEMP_path_s2a, dPL_s2a, dPK_s2a;
Display "=== S2B: Rubber price +30% ===";
Display dY_s2b, dXD_s2b, UNEMP_path_s2b, dPL_s2b, dPK_s2b;

Execute_Unload "scenario2_results.gdx",
  Y_path_base, Y_path_s2a, Y_path_s2b,
  XD_path_base, XD_path_s2a, XD_path_s2b,
  UNEMP_path_base, UNEMP_path_s2a, UNEMP_path_s2b,
  dY_s2a, dY_s2b, dXD_s2a, dXD_s2b,
  dPL_s2a, dPK_s2a, dPL_s2b, dPK_s2b,
  PK_s2a, PK_s2b;

Display dY_s2a, dY_s2b;
Display dXD_s2a, dXD_s2b;
Display dPL_s2a, dPK_s2a, dPL_s2b, dPK_s2b;
Display UNEMP_path_s2a, UNEMP_path_s2b;

*==============================================================================
* SECTION 8C: ADDITIONAL SCENARIOS S3-S12
*==============================================================================
*--- Price paths for EV computation (declare before policy scenario loops)
Parameter
  P_path_s10(COM,T)  "Commodity price path under S10"
  P_path_s11(COM,T)  "Commodity price path under S11"
  P_path_s12(COM,T)  "Commodity price path under S12"
;
*------------------------------------------------------------------------------
* STEP B: Policy parameter initialization (B1-B4)
*------------------------------------------------------------------------------

* B1: Output subsidy (S10)
subsidy_rate(ACT) = 0;
subsidy_rate('act26') = 0.15;
subsidy_rate('act27') = 0.10;

* B2: TFP support phased (S11)
tfp_support(RUBACT,T) = 0;
tfp_support('act26','2027') = 0.005; tfp_support('act26','2028') = 0.005;
tfp_support('act26','2029') = 0.005; tfp_support('act26','2030') = 0.010;
tfp_support('act26','2031') = 0.010; tfp_support('act26','2032') = 0.010;
tfp_support('act26','2033') = 0.010; tfp_support('act26','2034') = 0.010;
tfp_support('act27','2027') = 0.005; tfp_support('act27','2028') = 0.008;
tfp_support('act27','2029') = 0.010;
tfp_support(RUBACT,T)$(ord(T)>=7) = 0.010;

* B3: S12 targeted transfer — calibrate จากทะเบียน กยท.
rub_share('HH_Q1') = 0.1502; rub_share('HH_Q2') = 0.0626;
rub_share('HH_Q3') = 0.0300; rub_share('HH_Q4') = 0.0063;
rub_share('HH_Q5') = 0.0013;

n_rub_hh('HH_Q1') = 519617; n_rub_hh('HH_Q2') = 215916;
n_rub_hh('HH_Q3') = 0;      n_rub_hh('HH_Q4') = 0;
n_rub_hh('HH_Q5') = 0;

transfer_per_hh      = 15000 / SAM_unit;
GOV_TRANSFER_RUB(HH) = n_rub_hh(HH) * transfer_per_hh;
total_transfer_cost  = sum(HH, GOV_TRANSFER_RUB(HH));
Display GOV_TRANSFER_RUB, total_transfer_cost;
* คาดหวัง (SAM=ล้านบาท): Q1=7,794 | Q2=3,239 | รวม≈11,033 ล้านบาท

* B4: TFP growth rates for S5/S6
tfpgrow_s5(ACT) = tfpgrow(ACT);
tfpgrow_s5(RUBACT) = tfpgrow(RUBACT) - 0.010;
tfpgrow_s6(ACT) = tfpgrow(ACT);
tfpgrow_s6(RUBACT) = tfpgrow(RUBACT) + 0.010;

* Backup base parameters
io_base(COM,ACT) = io(COM,ACT);
EZ_base(COM)     = EZ(COM);

*==============================================================================
* BLOCK I: PRICE SCENARIOS
*==============================================================================

*--- SCENARIO S3: Rubber price -30% TEMPORARY (3 years 2024-2026) ---
RESET_STOCKS
PWEZ(RUBCOM) = 0.70;
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024');
PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s3(HH,'2024')=Y.L(HH); XD_path_s3(ACT,'2024')=XD.L(ACT);
UNEMP_path_s3('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  if(ord(T)<=3, PWEZ(RUBCOM)=0.70; else PWEZ(RUBCOM)=1.00; );
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s3(ACT,T-1));
  Y.L(HH)=Y_path_s3(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s3(HH,T)=Y_path_base(HH,T); XD_path_s3(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s3(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s3(HH,T)=Y.L(HH); XD_path_s3(ACT,T)=XD.L(ACT);
    UNEMP_path_s3(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
PWEZ(RUBCOM)=1.0;
dY_s3(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s3(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S3: Rubber price -30% TEMPORARY (3yr) ===";
Display dY_s3, UNEMP_path_s3;


*--- SCENARIO S3b: Rubber price -20% TEMPORARY (3 years 2024-2026) ---
RESET_STOCKS
PWEZ(RUBCOM)=0.80;
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s3b(HH,'2024')=Y.L(HH); XD_path_s3b(ACT,'2024')=XD.L(ACT);
UNEMP_path_s3b('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  if(ord(T)<=3, PWEZ(RUBCOM)=0.80; else PWEZ(RUBCOM)=1.00; );
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s3b(ACT,T-1));
  Y.L(HH)=Y_path_s3b(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s3b(HH,T)=Y_path_base(HH,T); XD_path_s3b(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s3b(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s3b(HH,T)=Y.L(HH); XD_path_s3b(ACT,T)=XD.L(ACT);
    UNEMP_path_s3b(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
PWEZ(RUBCOM)=1.0;
dY_s3b(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s3b(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S3b: Rubber price -20% TEMPORARY (3yr) ===";
Display dY_s3b, UNEMP_path_s3b;


*==============================================================================
* BLOCK II: SUPPLY / COST SCENARIOS
*==============================================================================

*--- SCENARIO S5: TFP rubber -10% (drought, disease) ---
RESET_STOCKS
aF_t(RUBACT,'2024')=aF(RUBACT)*0.90;
KSZ=sum(ACT,KZ_t(ACT,'2024')); LSZ=LSZ_t('2024'); aF(ACT)=aF_t(ACT,'2024');
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s5(HH,'2024')=Y.L(HH); XD_path_s5(ACT,'2024')=XD.L(ACT);
UNEMP_path_s5('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)$(not RUBACT(ACT))=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  aF_t(ACT,T)$RUBACT(ACT)=aF_t(ACT,T-1)*(1+tfpgrow_s5(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s5(ACT,T-1));
  Y.L(HH)=Y_path_s5(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s5(HH,T)=Y_path_base(HH,T); XD_path_s5(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s5(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s5(HH,T)=Y.L(HH); XD_path_s5(ACT,T)=XD.L(ACT);
    UNEMP_path_s5(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
dY_s5(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s5(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S5: TFP rubber -10% ===";
Display dY_s5, XD_path_s5;


*--- SCENARIO S6: TFP rubber +10% (smart farming, technology) ---
RESET_STOCKS
aF_t(RUBACT,'2024')=aF(RUBACT)*1.10;
KSZ=sum(ACT,KZ_t(ACT,'2024')); LSZ=LSZ_t('2024'); aF(ACT)=aF_t(ACT,'2024');
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s6(HH,'2024')=Y.L(HH); XD_path_s6(ACT,'2024')=XD.L(ACT);
UNEMP_path_s6('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)$(not RUBACT(ACT))=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  aF_t(ACT,T)$RUBACT(ACT)=aF_t(ACT,T-1)*(1+tfpgrow_s6(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s6(ACT,T-1));
  Y.L(HH)=Y_path_s6(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s6(HH,T)=Y_path_base(HH,T); XD_path_s6(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s6(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s6(HH,T)=Y.L(HH); XD_path_s6(ACT,T)=XD.L(ACT);
    UNEMP_path_s6(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
dY_s6(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s6(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S6: TFP rubber +10% ===";
Display dY_s6, XD_path_s6;


*--- SCENARIO S7: Input cost +20% (oil, fertilizer, transport) ---
io(COST_COM,RUBACT) = io_base(COST_COM,RUBACT) * 1.20;
K.L(ACT) = KZ(ACT);
RESET_STOCKS
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s7(HH,'2024')=Y.L(HH); XD_path_s7(ACT,'2024')=XD.L(ACT);
UNEMP_path_s7('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s7(ACT,T-1));
  Y.L(HH)=Y_path_s7(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s7(HH,T)=Y_path_base(HH,T); XD_path_s7(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s7(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s7(HH,T)=Y.L(HH); XD_path_s7(ACT,T)=XD.L(ACT);
    UNEMP_path_s7(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
io(COM,ACT) = io_base(COM,ACT);
dY_s7(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s7(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S7: Input cost +20% ===";
Display dY_s7, XD_path_s7;


*==============================================================================
* BLOCK III: TRADE SCENARIOS
*==============================================================================

*--- SCENARIO S8: Export demand -15% ---
EZ(RUBCOM) = EZ_base(RUBCOM) * 0.85;
RESET_STOCKS
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s8(HH,'2024')=Y.L(HH); XD_path_s8(ACT,'2024')=XD.L(ACT);
UNEMP_path_s8('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s8(ACT,T-1));
  Y.L(HH)=Y_path_s8(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s8(HH,T)=Y_path_base(HH,T); XD_path_s8(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s8(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s8(HH,T)=Y.L(HH); XD_path_s8(ACT,T)=XD.L(ACT);
    UNEMP_path_s8(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
EZ(COM) = EZ_base(COM);
dY_s8(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s8(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S8: Export demand -15% ===";
Display dY_s8, XD_path_s8;


*--- SCENARIO S9: Real ER +10% (baht depreciation) ---
RESET_STOCKS
ER.FX = 1.10 * ERZ;
SF.FX = SFZ/1.10;
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
PM.L(COM)=(1+tm(COM))*1.10*PWMZ(COM);
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s9(HH,'2024')=Y.L(HH); XD_path_s9(ACT,'2024')=XD.L(ACT);
UNEMP_path_s9('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s9(ACT,T-1));
  Y.L(HH)=Y_path_s9(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM)); PM.L(COM)=(1+tm(COM))*1.10*PWMZ(COM);
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s9(HH,T)=Y_path_base(HH,T); XD_path_s9(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s9(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s9(HH,T)=Y.L(HH); XD_path_s9(ACT,T)=XD.L(ACT);
    UNEMP_path_s9(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
  );
);
ER.FX = ERZ;
dY_s9(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s9(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S9: Real ER +10% (depreciation) ===";
Display dY_s9, XD_path_s9;


*==============================================================================
* BLOCK IV: POLICY SCENARIOS
*==============================================================================

*--- SCENARIO S10: S1 (price -20%) + Output Subsidy 15% ---
PWEZ(RUBCOM) = 0.80 * (1 + subsidy_equiv);
RESET_STOCKS
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s10(HH,'2024')=Y.L(HH); XD_path_s10(ACT,'2024')=XD.L(ACT);
UNEMP_path_s10('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);
* และสำหรับ T=2024 (ก่อน loop):
P_path_s10(COM,'2024') = P.L(COM);

INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s10(ACT,T-1));
  Y.L(HH)=Y_path_s10(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s10(HH,T)=Y_path_base(HH,T); XD_path_s10(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s10(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s10(HH,T)=Y.L(HH); XD_path_s10(ACT,T)=XD.L(ACT);
    UNEMP_path_s10(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
* 2. ใน S10 loop — เพิ่มใน if/else block ฝั่ง solved:
P_path_s10(COM,T) = P.L(COM); 

  );
);
PWEZ(RUBCOM)=1.0;
dY_s10(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s10(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S10: S1 + Output Subsidy 15% ===";
Display dY_s10, XD_path_s10;


*--- SCENARIO S11: S1 (price -20%) + TFP Support phased ---
PWEZ(RUBCOM)=0.80;
RESET_STOCKS
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024')); Y.L(HH)=Y_path_base(HH,'2024');
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s11(HH,'2024')=Y.L(HH); XD_path_s11(ACT,'2024')=XD.L(ACT);
UNEMP_path_s11('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);
P_path_s11(COM,'2024') = P.L(COM);
INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)$(not RUBACT(ACT))=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  aF_t(ACT,T)$RUBACT(ACT)=aF_t(ACT,T-1)*(1+tfpgrow(ACT)+tfp_support(ACT,T));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s11(ACT,T-1));
  Y.L(HH)=Y_path_s11(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s11(HH,T)=Y_path_base(HH,T); XD_path_s11(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s11(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s11(HH,T)=Y.L(HH); XD_path_s11(ACT,T)=XD.L(ACT);
    UNEMP_path_s11(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
    P_path_s11(COM,T) = P.L(COM);
  );
);
PWEZ(RUBCOM)=1.0;
dY_s11(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s11(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "=== S11: S1 + TFP Support (phased) ===";
Display dY_s11, XD_path_s11;


*--- SCENARIO S12: S1 (price -20%) + Targeted Transfer 15,000 บาท/ราย ---
transfer_switch=1; PWEZ(RUBCOM)=0.80;
RESET_STOCKS
K.L(ACT)=KZ(ACT); L.L(ACT)=max(1e-6,L_path(ACT,'2024'));
XD.L(ACT)=max(1e-6,XD_path_base(ACT,'2024'));
Y.L(HH)=Y_path_base(HH,'2024')+GOV_TRANSFER_RUB(HH);
CH.L(COM,HH)=CH_path_base(COM,HH,'2024'); INV.L(COM)=INV_path_base(COM,'2024');
M.L(COM)=max(1e-6,MZ(COM));
UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base('2024'); PL.FX=PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
RELAX_BOUNDS
Solve RubberScenario using CNS;
RESTORE_BOUNDS
Y_path_s12(HH,'2024')=Y.L(HH); XD_path_s12(ACT,'2024')=XD.L(ACT);
UNEMP_path_s12('2024')=max(0,UNEMP.L);
INV_path(COM,'2024')=INV.L(COM); CH_path(COM,HH,'2024')=CH.L(COM,HH);
P_path_s12(COM,'2024') = P.L(COM);
INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
Loop(T$(not TFIRST(T)),
  KZ_t(ACT,T)=KZ_t(ACT,T-1)*(1-deprate(ACT))+sum(COM$MAP(COM,ACT),kapshare(COM,ACT)*INV_path(COM,T-1));
  KZ_t(ACT,T)$(KZ_t(ACT,T)<=0)=1e-6;
  LSZ_t(T)=LSZ_t(T-1)*(1+popgrow_t(T));
  aF_t(ACT,T)=aF_t(ACT,T-1)*(1+tfpgrow(ACT));
  KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);
  K.L(ACT)=max(1e-6,KZ_t(ACT,T)); L.L(ACT)=max(1e-6,L_path(ACT,T-1));
  XD.L(ACT)=max(max(1e-6,0.01*XD_path_base(ACT,'2024')),XD_path_s12(ACT,T-1));
  Y.L(HH)=Y_path_s12(HH,T-1);
  CH.L(COM,HH)=CH_path(COM,HH,T-1); INV.L(COM)=INV_path(COM,T-1);
  M.L(COM)=max(1e-6,MZ(COM));
  UNEMP.LO=-INF; UNEMP.UP=LSZ; UNEMP.L=UNEMP_path_base(T); PL.FX=PL_path_base(T);
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1; PE.L(COM)=PWEZ(COM)*ER.L; PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  Solve RubberScenario using CNS;
  RESTORE_BOUNDS
  if(RubberScenario.modelstat <> 16,
    Y_path_s12(HH,T)=Y_path_base(HH,T); XD_path_s12(ACT,T)=XD_path_base(ACT,T);
    UNEMP_path_s12(T)=UNEMP_path_base(T);
    INV_path(COM,T)=INV_path_base(COM,T); CH_path(COM,HH,T)=CH_path_base(COM,HH,T);
  else
    Y_path_s12(HH,T)=Y.L(HH); XD_path_s12(ACT,T)=XD.L(ACT);
    UNEMP_path_s12(T)=max(0,UNEMP.L);
    INV_path(COM,T)=INV.L(COM); CH_path(COM,HH,T)=CH.L(COM,HH);
    P_path_s12(COM,T) = P.L(COM);

  );
);
transfer_switch=0; PWEZ(RUBCOM)=1.0;
dY_s12(HH,T)$(Y_path_base(HH,T)<>0)=(Y_path_s12(HH,T)-Y_path_base(HH,T))/abs(Y_path_base(HH,T))*100;
Display "...";
Display total_transfer_cost, GOV_TRANSFER_RUB, dY_s12, UNEMP_path_s12;


*--- EXPORT ALL SCENARIO RESULTS ---
Execute_Unload "scenario_all_results.gdx",
  Y_path_base,
  Y_path_s3,  XD_path_s3,  UNEMP_path_s3,  dY_s3,
  Y_path_s3b, XD_path_s3b, UNEMP_path_s3b, dY_s3b,
  Y_path_s5,  XD_path_s5,  UNEMP_path_s5,  dY_s5,
  Y_path_s6,  XD_path_s6,  UNEMP_path_s6,  dY_s6,
  Y_path_s7,  XD_path_s7,  UNEMP_path_s7,  dY_s7,
  Y_path_s8,  XD_path_s8,  UNEMP_path_s8,  dY_s8,
  Y_path_s9,  XD_path_s9,  UNEMP_path_s9,  dY_s9,
  Y_path_s10, XD_path_s10, UNEMP_path_s10, dY_s10,
  Y_path_s11, XD_path_s11, UNEMP_path_s11, dY_s11,
  Y_path_s12, XD_path_s12, UNEMP_path_s12, dY_s12;

Display "=== ALL SCENARIOS S3-S12 COMPLETE ===";

*==============================================================================
* SECTION 9: POST-LOOP ANALYSIS
*==============================================================================

*--- 9A: Compute annual growth rates
gY(HH,T)$(Y_path(HH,T-1)>0) =
    (Y_path(HH,T) - Y_path(HH,T-1)) / Y_path(HH,T-1) * 100;

gXD(ACT,T)$(XD_path(ACT,T-1)>0) =
    (XD_path(ACT,T) - XD_path(ACT,T-1)) / XD_path(ACT,T-1) * 100;

gPL(T)$(PL_path(T-1)>0) =
    (PL_path(T) - PL_path(T-1)) / PL_path(T-1) * 100;

gUNEMP(T) = UNEMP_path(T) - UNEMP_path(T-1);

*--- 9B: Display key results
Display "=== DYNAMIC RESULTS ===";
Display Y_path, XD_path, PL_path, UNEMP_path;
Display gY, gXD, gPL, LSZ_t, UNEMP_path_base;
Display PK_path, KZ_t, YZ;
Display SAM_data;

*--- 9C: Export to GDX for Excel/visualization
Execute_Unload "dynamic_results.gdx",
  Y_path, XD_path, E_path, M_path,
  PL_path, PK_path, UNEMP_path, TAXR_path,
  S_path, PCINDEX_path, INV_path,
  KZ_t, LSZ_t, aF_t,
  gY, gXD, gPL, gUNEMP;
  
*==============================================================================
* SECTION 10: SENSITIVITY ANALYSIS
* Varies sigmaA, sigmaT, sigmaF for rubber sectors (com/act 26-29)
* across a 3x3x3 grid and reports Q1 welfare loss under S1a (2024, 2034)
*==============================================================================

*--- 10A: Sensitivity grid sets and index parameters
Set
  SA  "Armington elasticity levels"  / sa1, sa2, sa3 /
  ST  "CET elasticity levels"        / st1, st2, st3 /
  SF2 "CES factor elasticity levels" / sf1, sf2, sf3 /
;

Parameter
  sa_val(SA)   "sigmaA values for rubber"  / sa1 2.5, sa2 3.5, sa3 4.5 /
  st_val(ST)   "sigmaT values for rubber"  / st1 1.0, st2 1.5, st3 2.0 /
  sf_val(SF2)  "sigmaF values for rubber"  / sf1 0.3, sf2 0.5, sf3 0.7 /
;

*--- 10B: Results storage
Parameter
  sens_dY_s1a_2024(SA,ST,SF2,HH)   "Q1 income dev S1a 2024 by elasticity combo"
  sens_dY_s1a_2034(SA,ST,SF2,HH)   "Q1 income dev S1a 2034 by elasticity combo"
  sens_ratio(SA,ST,SF2)             "Q1/Q5 welfare ratio under S1a 2024"
  sens_gdp_2024(SA,ST,SF2)          "GDP deviation S1a 2024"
  sens_unemp_2024(SA,ST,SF2)        "Unemployment deviation S1a 2024"
;

*--- 10C: Backup baseline elasticities for rubber sectors
Parameter
  sigmaA_base(COM)   "Baseline Armington elasticity"
  sigmaT_base(COM)   "Baseline CET elasticity"
  sigmaF_base(ACT)   "Baseline CES factor elasticity"
  gammaA_base(COM)   "Baseline Armington share"
  aA_base(COM)       "Baseline Armington scale"
  gammaT_base(COM)   "Baseline CET share"
  gammaF_base(ACT)   "Baseline CES capital share"
  aF_base2(ACT)      "Baseline CES scale"
;

sigmaA_base(COM)  = sigmaA(COM);
sigmaT_base(COM)  = sigmaT(COM);
sigmaF_base(ACT)  = sigmaF(ACT);
gammaA_base(COM)  = gammaA(COM);
aA_base(COM)      = aA(COM);
gammaT_base(COM)  = gammaT(COM);
gammaF_base(ACT)  = gammaF(ACT);
aF_base2(ACT)     = aF_base(ACT);

*--- 10D: GDP path parameter (for sensitivity macro GDP dev)
Parameter
  gdp_base_2024   "Baseline GDP proxy (sum PD*XD) at 2024"
  gdp_sens        "Sensitivity GDP proxy under S1a 2024"
  Y_path_sens(HH,T)    "HH income path for current sensitivity run"
  XD_path_sens(ACT,T)  "Output path for current sensitivity run"
  UNEMP_path_sens(T)   "Unemployment path for current sensitivity run"
;

gdp_base_2024 = sum(ACT, XD_path_base(ACT,'2024'));

*==============================================================================
* 10E: RECALIBRATION SUBROUTINE — rubber-sector elasticities only
*      Called after changing sigmaA/sigmaT/sigmaF for RUBCOM/RUBACT
*==============================================================================

* Macro: recalibrate gammaF and aF for rubber activities after sigmaF change
$macro RECAL_CES_RUB \
  KL_ratio(ACT)$(LZ(ACT)>0 and RUBACT(ACT)) = KZ(ACT)/LZ(ACT); \
  gammaF(ACT)$(KL_ratio(ACT)>0 and RUBACT(ACT)) = \
      KL_ratio(ACT)**(1/sigmaF(ACT)) / \
      (1 + KL_ratio(ACT)**(1/sigmaF(ACT))); \
  gammaF(ACT)$(gammaF(ACT)<=0.01 and RUBACT(ACT))=0.01; \
  gammaF(ACT)$(gammaF(ACT)>=0.99 and RUBACT(ACT))=0.99; \
  cost_idx(ACT)$RUBACT(ACT) = gammaF(ACT)**sigmaF(ACT) \
      + (1-gammaF(ACT))**sigmaF(ACT); \
  aF(ACT)$(KZ(ACT)>0 and RUBACT(ACT)) = \
      (XDZ(ACT)/KZ(ACT)) * gammaF(ACT)**sigmaF(ACT) * \
      cost_idx(ACT)**(sigmaF(ACT)/(1-sigmaF(ACT))); \
  aF(ACT)$(aF(ACT)<=1e-10 and RUBACT(ACT))=1e-10; \
  aF_base(ACT)$RUBACT(ACT) = aF(ACT);

* Macro: recalibrate gammaA and aA for rubber commodities after sigmaA change
$macro RECAL_ARM_RUB \
  MXratio(COM)$(XDD_base(COM)>0 and RUBCOM(COM)) = MZ(COM)/XDD_base(COM); \
  gammaA(COM)$(MXratio(COM)>0 and RUBCOM(COM)) = \
      MXratio(COM)**(1/sigmaA(COM)) / \
      (1 + MXratio(COM)**(1/sigmaA(COM))); \
  gammaA(COM)$(gammaA(COM)<=0.01 and RUBCOM(COM))=0.01; \
  gammaA(COM)$(gammaA(COM)>=0.99 and RUBCOM(COM))=0.99; \
  cost_idxA(COM)$RUBCOM(COM) = gammaA(COM)**sigmaA(COM) \
      + (1-gammaA(COM))**sigmaA(COM); \
  aA(COM)$(MZ(COM)>0 and RUBCOM(COM)) = \
      (X_base(COM)/MZ(COM)) * gammaA(COM)**sigmaA(COM) * \
      cost_idxA(COM)**(sigmaA(COM)/(1-sigmaA(COM))); \
  aA(COM)$(aA(COM)<=1e-10 and RUBCOM(COM))=1e-10;

* Macro: recalibrate gammaT for rubber CET sectors after sigmaT change
$macro RECAL_CET_RUB \
  EX_ratio(COM)$(XDD_base(COM)>0 and RUBCOM(COM)) = EZ(COM)/XDD_base(COM); \
  gammaT(COM)$(EX_ratio(COM)>0 and RUBCOM(COM)) = \
      EX_ratio(COM)**(1/sigmaT(COM)) / \
      (1 + EX_ratio(COM)**(1/sigmaT(COM))); \
  gammaT(COM)$(gammaT(COM)<=0.01 and RUBCOM(COM))=0.01; \
  gammaT(COM)$(gammaT(COM)>=0.99 and RUBCOM(COM))=0.99;

* Macro: restore all rubber elasticities to baseline
$macro RESTORE_ELAST_RUB \
  sigmaA(COM)$RUBCOM(COM)  = sigmaA_base(COM); \
  sigmaT(COM)$RUBCOM(COM)  = sigmaT_base(COM); \
  sigmaF(ACT)$RUBACT(ACT)  = sigmaF_base(ACT); \
  gammaA(COM)$RUBCOM(COM)  = gammaA_base(COM); \
  aA(COM)$RUBCOM(COM)      = aA_base(COM); \
  gammaT(COM)$RUBCOM(COM)  = gammaT_base(COM); \
  gammaF(ACT)$RUBACT(ACT)  = gammaF_base(ACT); \
  aF(ACT)$RUBACT(ACT)      = aF_base2(ACT); \
  aF_base(ACT)$RUBACT(ACT) = aF_base2(ACT);

*==============================================================================
* 10F: SENSITIVITY LOOP — 3 x 3 x 3 = 27 combinations
*==============================================================================
Display "=== STARTING SENSITIVITY ANALYSIS (27 combinations) ===";

Loop(SA,
Loop(ST,
Loop(SF2,

  Display SA, ST, SF2;

*--- Step 1: Set rubber elasticities to grid values
  sigmaA(RUBCOM)  = sa_val(SA);
  sigmaT(RUBCOM)  = st_val(ST);
  sigmaF(RUBACT)  = sf_val(SF2);

*--- Step 2: Recalibrate shares and scales for rubber sectors
  RECAL_CES_RUB
  RECAL_ARM_RUB
  RECAL_CET_RUB

*--- Step 3: Reset dynamic stocks to baseline year 2024
  RESET_STOCKS

*--- Step 4: Run S1a (price -20%) — year 2024
  PWEZ(RUBCOM) = 0.80;

  K.L(ACT)     = KZ(ACT);
  L.L(ACT)     = max(1e-6, L_path(ACT,'2024'));
  XD.L(ACT)    = max(1e-6, XD_path_base(ACT,'2024'));
  Y.L(HH)      = Y_path_base(HH,'2024');
  CH.L(COM,HH) = CH_path_base(COM,HH,'2024');
  INV.L(COM)   = INV_path_base(COM,'2024');
  M.L(COM)     = max(1e-6, MZ(COM));

  UNEMP.LO = -INF;  UNEMP.UP = LSZ;
  UNEMP.L  = UNEMP_path_base('2024');
  PL.FX    = PL_path_base('2024');
  P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1;
  PE.L(COM)=PWEZ(COM)*ER.L;
  PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
  RELAX_BOUNDS
  K.LO(ACT)=1e-6; K.UP(ACT)=+INF;

  Solve RubberScenario using CNS;

  RESTORE_BOUNDS

  Y_path_sens(HH,'2024')    = Y.L(HH);
  XD_path_sens(ACT,'2024')  = XD.L(ACT);
  UNEMP_path_sens('2024')   = max(0, UNEMP.L);
  INV_path(COM,'2024')      = INV.L(COM);
  CH_path(COM,HH,'2024')    = CH.L(COM,HH);

*--- Step 5: Run S1a dynamic loop through 2034
  INV_path(COM,T)   = INV_path_base(COM,T);
  CH_path(COM,HH,T) = CH_path_base(COM,HH,T);

  Loop(T$(not TFIRST(T)),
    KZ_t(ACT,T) = KZ_t(ACT,T-1)*(1-deprate(ACT))
                + sum(COM$MAP(COM,ACT), kapshare(COM,ACT)*INV_path(COM,T-1));
    KZ_t(ACT,T)$(KZ_t(ACT,T)<=0) = 1e-6;
    LSZ_t(T) = LSZ_t(T-1)*(1+popgrow_t(T));
    aF_t(ACT,T) = aF_t(ACT,T-1)*(1+tfpgrow(ACT));
    KSZ=sum(ACT,KZ_t(ACT,T)); LSZ=LSZ_t(T); aF(ACT)=aF_t(ACT,T);

    K.L(ACT)     = max(1e-6, KZ_t(ACT,T));
    L.L(ACT)     = max(1e-6, L_path(ACT,T-1));
    XD.L(ACT)    = max(max(1e-6, 0.01*XD_path_base(ACT,'2024')),
                       XD_path_sens(ACT,T-1));
    Y.L(HH)      = Y_path_sens(HH,T-1);
    CH.L(COM,HH) = CH_path(COM,HH,T-1);
    INV.L(COM)   = INV_path(COM,T-1);
    M.L(COM)     = max(1e-6, MZ(COM));

    UNEMP.LO = -INF; UNEMP.UP = LSZ;
    UNEMP.L  = UNEMP_path_base(T);
    PL.FX    = PL_path_base(T);
    P.L(COM)=1; PDD.L(COM)=1; PD.L(ACT)=1;
    PE.L(COM)=PWEZ(COM)*ER.L;
    PM.L(COM)=max(1e-6,(1+tm(COM))*ER.L*PWMZ(COM));
    RELAX_BOUNDS
    K.LO(ACT)=1e-6; K.UP(ACT)=+INF;

    Solve RubberScenario using CNS;
    RESTORE_BOUNDS

    if(RubberScenario.modelstat <> 16,
      Y_path_sens(HH,T)   = Y_path_base(HH,T);
      XD_path_sens(ACT,T) = XD_path_base(ACT,T);
      UNEMP_path_sens(T)  = UNEMP_path_base(T);
      INV_path(COM,T)     = INV_path_base(COM,T);
      CH_path(COM,HH,T)   = CH_path_base(COM,HH,T);
    else
      Y_path_sens(HH,T)   = Y.L(HH);
      XD_path_sens(ACT,T) = XD.L(ACT);
      UNEMP_path_sens(T)  = max(0, UNEMP.L);
      INV_path(COM,T)     = INV.L(COM);
      CH_path(COM,HH,T)   = CH.L(COM,HH);
    );
  );

  PWEZ(RUBCOM) = 1.0;

*--- Step 6: Compute and store deviations
  sens_dY_s1a_2024(SA,ST,SF2,HH)$(Y_path_base(HH,'2024')<>0) =
      (Y_path_sens(HH,'2024') - Y_path_base(HH,'2024'))
      / abs(Y_path_base(HH,'2024')) * 100;

  sens_dY_s1a_2034(SA,ST,SF2,HH)$(Y_path_base(HH,'2034')<>0) =
      (Y_path_sens(HH,'2034') - Y_path_base(HH,'2034'))
      / abs(Y_path_base(HH,'2034')) * 100;

  sens_ratio(SA,ST,SF2)$(sens_dY_s1a_2024(SA,ST,SF2,'HH_Q5')<>0) =
      sens_dY_s1a_2024(SA,ST,SF2,'HH_Q1')
      / abs(sens_dY_s1a_2024(SA,ST,SF2,'HH_Q5'));

  gdp_sens = sum(ACT, XD_path_sens(ACT,'2024'));
  sens_gdp_2024(SA,ST,SF2)$(gdp_base_2024<>0) =
      (gdp_sens - gdp_base_2024) / abs(gdp_base_2024) * 100;

  sens_unemp_2024(SA,ST,SF2) = UNEMP_path_sens('2024') - UNEMP_path_base('2024');

*--- Step 7: Restore baseline elasticities before next iteration
  RESTORE_ELAST_RUB

);
* end SF2 loop
);
* end ST loop
);
* end SA loop

Display "=== SENSITIVITY ANALYSIS COMPLETE ===";

*==============================================================================
* 10G: DISPLAY AND EXPORT SENSITIVITY RESULTS
*==============================================================================

*--- Q1 welfare loss by combo (2024)
Display "--- Q1 income deviation (%) under S1a 2024 by (sigmaA, sigmaT, sigmaF) ---";
Display sens_dY_s1a_2024;

*--- Q1/Q5 ratio by combo
Display "--- Q1/Q5 welfare ratio under S1a 2024 ---";
Display sens_ratio;

*--- GDP deviation by combo
Display "--- GDP deviation (%) under S1a 2024 ---";
Display sens_gdp_2024;

*--- Diagonal: baseline combination (sa2, st2, sf2) should match main S1a result
Display "--- Baseline combination check (sa2=3.5, st2=1.5, sf2=0.5) ---";
Parameter sens_check_Q1, sens_check_ratio;
sens_check_Q1    = sens_dY_s1a_2024('sa2','st2','sf2','HH_Q1');
sens_check_ratio = sens_ratio('sa2','st2','sf2');
Display sens_check_Q1, sens_check_ratio;
* Expected: sens_check_Q1 ≈ -6.62  |  sens_check_ratio ≈ 6.2

*--- Range across all 27 combinations
Parameter
  sens_Q1_min   "Minimum Q1 welfare loss across all combos"
  sens_Q1_max   "Maximum Q1 welfare loss across all combos"
  sens_Q1_range "Range (max - min)"
  sens_ratio_min
  sens_ratio_max
;
sens_Q1_min   = smin((SA,ST,SF2), sens_dY_s1a_2024(SA,ST,SF2,'HH_Q1'));
sens_Q1_max   = smax((SA,ST,SF2), sens_dY_s1a_2024(SA,ST,SF2,'HH_Q1'));
sens_Q1_range = sens_Q1_max - sens_Q1_min;
sens_ratio_min = smin((SA,ST,SF2), sens_ratio(SA,ST,SF2));
sens_ratio_max = smax((SA,ST,SF2), sens_ratio(SA,ST,SF2));

Display "--- Range summary ---";
Display sens_Q1_min, sens_Q1_max, sens_Q1_range;
Display sens_ratio_min, sens_ratio_max;
* Expected (from paper §4.7): range ≈ 1.6pp  |  ratio range 5.9–6.5

*--- Marginal contribution of each elasticity dimension
* (Fix other two at baseline, vary one across levels)
Parameter
  sens_margA(SA)   "Q1 loss varying sigmaA only (sigmaT=1.5, sigmaF=0.5)"
  sens_margT(ST)   "Q1 loss varying sigmaT only (sigmaA=3.5, sigmaF=0.5)"
  sens_margF(SF2)  "Q1 loss varying sigmaF only (sigmaA=3.5, sigmaT=1.5)"
;
sens_margA(SA)  = sens_dY_s1a_2024(SA,'st2','sf2','HH_Q1');
sens_margT(ST)  = sens_dY_s1a_2024('sa2',ST,'sf2','HH_Q1');
sens_margF(SF2) = sens_dY_s1a_2024('sa2','st2',SF2,'HH_Q1');

Display "--- Marginal elasticity contributions ---";
Display sens_margA, sens_margT, sens_margF;
* Expected: sigmaT contributes most (~0.9pp), sigmaA and sigmaF < 0.4pp each

*--- Export to GDX
Execute_Unload "sensitivity_results.gdx",
  sens_dY_s1a_2024, sens_dY_s1a_2034,
  sens_ratio, sens_gdp_2024, sens_unemp_2024,
  sens_margA, sens_margT, sens_margF,
  sens_Q1_min, sens_Q1_max, sens_Q1_range,
  sa_val, st_val, sf_val;

Display "=== SENSITIVITY RESULTS EXPORTED TO sensitivity_results.gdx ===";

*==============================================================================
* SECTION 11: EQUIVALENT VARIATION (EV) — Policy Scenarios S10–S12
* Reports EV alongside %ΔY for Table 4 cross-instrument comparison
*==============================================================================

*--- 11A: Parameters for EV computation
Parameter
  P_base(COM,T)       "Baseline commodity prices by year"
  P_s10(COM,T)        "Commodity prices under S10"
  P_s11(COM,T)        "Commodity prices under S11"
  P_s12(COM,T)        "Commodity prices under S12"
  CBUD_base(HH,T)     "Baseline consumption budget by quintile"
  CBUD_s10(HH,T)      "Consumption budget under S10"
  CBUD_s11(HH,T)      "Consumption budget under S11"
  CBUD_s12(HH,T)      "Consumption budget under S12"
  price_idx(HH,T)     "Cobb-Douglas price index: prod_i (P0/P1)^beta"
  EV_s10(HH,T)        "EV under S10 (SAM units = billion THB)"
  EV_s11(HH,T)        "EV under S11 (SAM units)"
  EV_s12(HH,T)        "EV under S12 (SAM units)"
  EV_s10_pct(HH,T)    "EV under S10 as % of baseline consumption budget"
  EV_s11_pct(HH,T)    "EV under S11 as % of baseline consumption budget"
  EV_s12_pct(HH,T)    "EV under S12 as % of baseline consumption budget"
  CBUD_base_ann(HH)   "Baseline annual consumption budget (2024)"
;

*--- 11B: Store baseline commodity prices and consumption budgets
*    NOTE: P_path must be stored during scenario runs below.
*    Re-run S10–S12 saving P.L(COM,T) at each period.
*    Here we first re-extract from the CH_path already stored:
*    CBUD(h,t) = sum_i (1+tc(i))*P(i,t)*CH(i,h,t) / (1-ty(h))
*    Simpler: store P.L during the policy scenario loops (add below).

* Baseline consumption budget: CBUD = (1-ty)*Y - SH
CBUD_base(HH,T) = (1 - ty(HH)) * Y_path_base(HH,T)
                - mps(HH) * (1 - ty(HH)) * Y_path_base(HH,T);
* = (1 - mps(HH)) * (1 - ty(HH)) * Y_path_base(HH,T)
CBUD_base(HH,T) = (1 - mps(HH)) * (1 - ty(HH)) * Y_path_base(HH,T);

CBUD_base_ann(HH) = CBUD_base(HH,'2024');
Display CBUD_base_ann;

*--- 11C: Price path storage parameters (add to policy scenario loops)
Parameter
  P_path_s10(COM,T)   "P.L stored during S10 loop"
  P_path_s11(COM,T)   "P.L stored during S11 loop"
  P_path_s12(COM,T)   "P.L stored during S12 loop"
  P_path_base2(COM,T) "Baseline P.L stored during baseline loop"
;

*--- 11D: Re-run S10 saving P.L — insert after S10 loop
*    (or patch the existing S10 loop to add one line per period)

* PATCH FOR EXISTING S10 LOOP — add after each Solve:
*   P_path_s10(COM,T) = P.L(COM);
* and after T=2024 solve:
*   P_path_s10(COM,'2024') = P.L(COM);

* Similarly for S11, S12, and baseline loop:
*   P_path_base2(COM,T) = P.L(COM);  [in baseline loop]

*==============================================================================
* 11E: EV COMPUTATION
* Formula: EV(h,t) = CBUD1(h,t) * prod_i[P0(i,t)/P1(i,t)]^beta(i,h)
*                   - CBUD0(h,t)
* Note: GAMS has no PROD() function — implement as exp(sum of logs)
*==============================================================================

* Consumption budgets under scenarios
CBUD_s10(HH,T) = (1 - mps(HH)) * (1 - ty(HH)) * Y_path_s10(HH,T);
CBUD_s11(HH,T) = (1 - mps(HH)) * (1 - ty(HH)) * Y_path_s11(HH,T);
CBUD_s12(HH,T) = (1 - mps(HH)) * (1 - ty(HH)) * Y_path_s12(HH,T);

* Price index using log-sum trick: prod_i (P0/P1)^beta = exp(sum_i beta*log(P0/P1))
Parameter
  log_price_idx_s10(HH,T)
  log_price_idx_s11(HH,T)
  log_price_idx_s12(HH,T)
;

log_price_idx_s10(HH,T) = sum(COM$(P_path_s10(COM,T)>0
                               and P_path_base2(COM,T)>0
                               and betaC(COM,HH)>0),
    betaC(COM,HH) * log(P_path_base2(COM,T) / P_path_s10(COM,T)));

log_price_idx_s11(HH,T) = sum(COM$(P_path_s11(COM,T)>0
                               and P_path_base2(COM,T)>0
                               and betaC(COM,HH)>0),
    betaC(COM,HH) * log(P_path_base2(COM,T) / P_path_s11(COM,T)));

log_price_idx_s12(HH,T) = sum(COM$(P_path_s12(COM,T)>0
                               and P_path_base2(COM,T)>0
                               and betaC(COM,HH)>0),
    betaC(COM,HH) * log(P_path_base2(COM,T) / P_path_s12(COM,T)));

* EV in SAM units (billion THB if SAM unit = 1e9 THB)
EV_s10(HH,T) = CBUD_s10(HH,T) * exp(log_price_idx_s10(HH,T))
             - CBUD_base(HH,T);

EV_s11(HH,T) = CBUD_s11(HH,T) * exp(log_price_idx_s11(HH,T))
             - CBUD_base(HH,T);

EV_s12(HH,T) = CBUD_s12(HH,T) * exp(log_price_idx_s12(HH,T))
             - CBUD_base(HH,T);

* EV as % of baseline consumption budget
EV_s10_pct(HH,T)$(CBUD_base(HH,T)>0) =
    EV_s10(HH,T) / CBUD_base(HH,T) * 100;

EV_s11_pct(HH,T)$(CBUD_base(HH,T)>0) =
    EV_s11(HH,T) / CBUD_base(HH,T) * 100;

EV_s12_pct(HH,T)$(CBUD_base(HH,T)>0) =
    EV_s12(HH,T) / CBUD_base(HH,T) * 100;

*--- 11F: Display key years (2024, 2034) for Table 4
Parameter
  EV_table4(HH,*,*)  "EV summary for Table 4 (2024 and 2034)"
;
EV_table4(HH,'S10','2024') = EV_s10_pct(HH,'2024');
EV_table4(HH,'S10','2034') = EV_s10_pct(HH,'2034');
EV_table4(HH,'S11','2024') = EV_s11_pct(HH,'2024');
EV_table4(HH,'S11','2034') = EV_s11_pct(HH,'2034');
EV_table4(HH,'S12','2024') = EV_s12_pct(HH,'2024');
EV_table4(HH,'S12','2034') = EV_s12_pct(HH,'2034');

Display "=== EV Results for Table 4 (% of baseline consumption budget) ===";
Display EV_table4;

* Cross-check: EV and %ΔY should be close for S12 (lump-sum transfer)
* and diverge more for S10 (price distortion via subsidy)
Parameter ev_dY_diff_s10(HH), ev_dY_diff_s12(HH);
ev_dY_diff_s10(HH) = EV_s10_pct(HH,'2024') - dY_s10(HH,'2024');
ev_dY_diff_s12(HH) = EV_s12_pct(HH,'2024') - dY_s12(HH,'2024');
Display "=== EV vs %dY divergence (S10 > S12 expected) ===";
Display ev_dY_diff_s10, ev_dY_diff_s12;
* Expected: |ev_dY_diff_s10| > |ev_dY_diff_s12|
* S10 divergence reflects price-distortion cost of output subsidy

*--- 11G: Export
Execute_Unload "ev_results.gdx",
  EV_s10, EV_s11, EV_s12,
  EV_s10_pct, EV_s11_pct, EV_s12_pct,
  EV_table4,
  CBUD_base, CBUD_s10, CBUD_s11, CBUD_s12,
  ev_dY_diff_s10, ev_dY_diff_s12;

Display "=== EV RESULTS EXPORTED TO ev_results.gdx ===";

