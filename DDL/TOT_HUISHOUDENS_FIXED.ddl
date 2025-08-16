create or replace view FIXDB.TOT_HUISHOUDENS_FIXED as 
select round(sum(huishoudens_fixed),-3) as tot_huishoudens_fixed , round((sum(huishoudens_fixed) * 334)/1000000,1)  as tot_energierekening_bespaard, round((sum(huishoudens_fixed) * 0.3846),-2) as tot_co2_bespaard from (
select '2025_Q2' as PERIODE, gemeente, fixteams, huishoudens_fixed, betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2025_Q1 vq1 
union
select '2025_Q1' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2025_Q2 vq2 
union
select '2024_Q4' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers  
from VRAGENLIJST_2024_Q4 vq3 
union
select '2024_Q3' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2024_Q3 vq4 
union
select '2024_Q2' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2024_Q2 vq5 
union
select '2024_Q1' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2024_Q1 vq6
union
select '2023' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2023 vq7 
union
select '2022' as PERIODE, gemeente, fixteams, huishoudens_fixed,betaalde_werknemers,vrijwilligers 
from VRAGENLIJST_2022 vq8 
) A
