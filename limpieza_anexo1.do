/*==============================================================================
  LIMPIEZA - Anexo 01: Relación de locales educativos intervenidos
  Campos ingresados por GR/GL
==============================================================================*/

clear all
set more off

* ── RUTAS ────────────────────────────────────────────────────────────────────
global reporte    "C:\Users\iraid\Documents\DISCO TERA\MINEDU\TRABAJO-MINEDU\01_MINEDU"
global rep_cons_i "${reporte}\01_input"
global rep_cons_o "${reporte}\03_output"
global rep_cons_t "${reporte}\04_temporal"

// ***************************************************************
// 1. Limpieza de la BDA Reporte de inversión
// ***************************************************************
// Importar los datos
// cellrange(A3) salta el título (fila 1) y la fila en blanco (fila 2),
// usando la fila 3 del Excel directamente como nombres de variables (firstrow).
// Esto evita depender del nombre truncado que Stata genera del título.
import excel "${rep_cons_i}\Anexo_1_avance_GR_GL.xlsx", ///
    sheet("Hoja1") cellrange(A3) firstrow clear
save "${rep_cons_i}\Anexo1_GR_GL.dta", replace

// Cargar la base de datos
use "${rep_cons_i}\Anexo1_GR_GL.dta", clear
count  // 55,712
d, short

* ── RENOMBRAR a nombres cortos operativos ────────────────────────────────────
* Stata trunca los headers a 32 chars y elimina caracteres especiales de forma
* distinta según la versión. Para evitar ese problema usamos rename por posición:
* itera todas las variables en orden y las renombra sin depender de su nombre actual.

local newnames "nro cod_local region provincia distrito nombre_ie cui tipo monto avance fecha f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec comentarios"
local i = 1
foreach v of varlist _all {
    local newname : word `i' of `newnames'
    if "`newname'" != "" rename `v' `newname'
    local ++i
}

* Eliminar columnas extra vacías que vienen después de Comentarios (cols 26-29)
capture drop var26 var27 var28 var29
capture drop _var26 _var27 _var28 _var29

* Descripción de campos originales (etiquetas)
label var cui       "Código Único de Inversión (CUI)"
label var tipo      "Tipo de inversión (PI/IOARR/IRI)"
label var monto     "Monto de inversión (S/)"
label var avance    "Avance físico (%)"
label var fecha     "Fecha de recepción de obra (o estimada)"
label var f9        "Tiene F9 (SI/NO)"
label var comp      "Componentes de la inversión"
label var unid      "Intervino todas las unidades productoras (SÍ/NO)"
label var cod_mod   "Códigos modulares intervenidos (si NO)"
label var demol     "Demolición total o parcial (SÍ/NO)"
label var nueva     "Construcción de nueva infraestructura (SÍ/NO)"
label var reforz    "Reforzamiento estructural (SÍ/NO)"
label var cerco     "Cerco perimétrico (SÍ/NO)"
label var sust      "Sustitución parcial de edificaciones (SÍ/NO)"
label var ampl      "Ampliación de área (SÍ/NO)"
label var mobil     "Reposición/dotación de mobiliario y equipamiento (SÍ/NO)"
label var agua      "Acceso a agua y desagüe (SÍ/NO)"
label var elec      "Acceso a energía eléctrica (SÍ/NO)"


* ── 1. ELIMINAR FILAS SIN CUI ─────────────────────────────────────────────────
drop if missing(cui)
count
di "Filas con CUI: `r(N)'"

* Columna de errores por campo (0 = ok, 1 = error)
gen byte err_cui   = 0
gen byte err_tipo  = 0
gen byte err_monto = 0
gen byte err_avance= 0
gen byte err_f9    = 0
gen byte err_comp  = 0
gen byte err_unid  = 0
gen byte err_demol = 0
gen byte err_nueva = 0
gen byte err_reforz= 0
gen byte err_cerco = 0
gen byte err_sust  = 0
gen byte err_ampl  = 0
gen byte err_mobil = 0
gen byte err_agua  = 0
gen byte err_elec  = 0


* ── 2. CUI: solo dígitos ──────────────────────────────────────────────────────
* Limpiar espacios
replace cui = strtrim(cui)

* Marcar como error si contiene caracteres no numéricos
replace err_cui = 1 if !regexm(cui, "^[0-9]+$") & !missing(cui)

* Intentar rescatar: extraer secuencia de 5+ dígitos
gen cui_rescued = regexs(0) if regexm(cui, "[0-9]{5,}")
replace cui = cui_rescued if err_cui == 1 & !missing(cui_rescued)
replace err_cui = 0        if err_cui == 1 & !missing(cui_rescued)
drop cui_rescued


* ── 3. TIPO DE INVERSIÓN → PI / IOARR / IRI ──────────────────────────────────
replace tipo = upper(strtrim(tipo))
replace tipo = regexr(tipo, "\s+", " ")   // espacios internos múltiples

replace tipo = "IOARR" if regexm(tipo, "IOARR?|FUR")
replace tipo = "PI"    if regexm(tipo, "^PI$|PROYECTO")   & tipo != "IOARR"
replace tipo = "IRI"   if regexm(tipo, "^IRI$")           & tipo != "IOARR" & tipo != "PI"

replace err_tipo = 1 if !inlist(tipo, "PI", "IOARR", "IRI") & !missing(tipo)


* ── 4. MONTO: convertir a numérico ───────────────────────────────────────────
* Reemplazar coma decimal por punto
replace monto = subinstr(monto, ",", ".", .)
replace monto = strtrim(monto)

destring monto, replace force
replace err_monto = 1 if missing(monto) & monto != ""   // no se pudo convertir
replace err_monto = 1 if monto < 0 & !missing(monto)    // valor negativo inválido


* ── 5. AVANCE FÍSICO: porcentaje 0–100 ───────────────────────────────────────
replace avance = subinstr(avance, "%", "", .)
replace avance = subinstr(avance, ",", ".", .)
replace avance = strtrim(avance)

destring avance, replace force

* Si vino como proporción (0–1) → convertir a porcentaje
replace avance = round(avance * 100, 0.01) if !missing(avance) & avance > 0 & avance <= 1

replace err_avance = 1 if missing(avance) & avance != ""
replace err_avance = 1 if (avance < 0 | avance > 100) & !missing(avance)


* ── 6. F9: SI / NO ───────────────────────────────────────────────────────────
replace f9 = upper(strtrim(f9))
replace f9 = subinstr(f9, "Í", "I", .)

replace f9 = "SI" if inlist(f9, "SI", "SÍ", "S", "1", "SI ")
replace f9 = "NO" if inlist(f9, "NO", "N", "0", "NO ")

replace err_f9 = 1 if !inlist(f9, "SI", "NO") & !missing(f9)


* ── 7. COMPONENTES ───────────────────────────────────────────────────────────
replace comp = upper(strtrim(comp))

* Valor 4 / Integral
replace comp = "4. Integral (1, 2 y 3)" if ///
    inlist(comp, "4", "INTEGRAL", "(1,2 Y 3)", "1,2,3", "1, 2 Y 3", "1 Y 2 Y 3") | ///
    regexm(comp, "1[,Y ]+2[,Y ]+3|INTEGR")

* Combos de dos → también integral
replace comp = "4. Integral (1, 2 y 3)" if ///
    inlist(comp, "1,2", "1, 2", "1 Y 2", "1,3", "1 Y 3", "2,3", "2 Y 3", ///
                 "2. EQUIPAMIENTO 3.MOBILIARIO") & comp != "4. Integral (1, 2 y 3)"

* Solo infraestructura
replace comp = "1. Infraestructura." if ///
    inlist(comp, "1", "1.", "1. INFRAESTRUCTURA.", "1.INFRAESTRUCTURA", "INFRAESTRUCTURA") ///
    & comp != "4. Integral (1, 2 y 3)"

* Solo equipamiento
replace comp = "2. Equipamiento." if ///
    inlist(comp, "2", "2.", "2. EQUIPAMIENTO.", "2.EQUIPAMIENTO", "2. EQUIPAMIENT") ///
    & comp != "4. Integral (1, 2 y 3)"

* Solo mobiliario
replace comp = "3. Mobiliario." if ///
    inlist(comp, "3", "3.", "3. MOBILIARIO.") ///
    & comp != "4. Integral (1, 2 y 3)"

replace err_comp = 1 if ///
    !inlist(comp, "1. Infraestructura.", "2. Equipamiento.", ///
                  "3. Mobiliario.", "4. Integral (1, 2 y 3)") & !missing(comp)


* ── 8. CAMPOS SI/NO DE INTERVENCIÓN ─────────────────────────────────────────
local sino_vars "unid demol nueva reforz cerco sust ampl mobil agua elec"

foreach v of local sino_vars {
    replace `v' = upper(strtrim(`v'))
    replace `v' = subinstr(`v', "Í", "I", .)
    
    * Normalizar SI (incluyendo "SI (PARCIAL)", "SI-TOTAL", "SI PARCIAL" → SI)
    replace `v' = "SI" if regexm(`v', "^SI") | inlist(`v', "1", "S", "SÍ")
    replace `v' = "NO" if regexm(`v', "^NO") | inlist(`v', "0", "N")
    
    replace err_`v' = 1 if !inlist(`v', "SI", "NO") & !missing(`v')
}


* ── 9. SEPARAR BASES ─────────────────────────────────────────────────────────
gen byte tiene_error = (err_cui | err_tipo | err_monto | err_avance | ///
                         err_f9 | err_comp | err_unid | err_demol | ///
                         err_nueva | err_reforz | err_cerco | err_sust | ///
                         err_ampl | err_mobil | err_agua | err_elec)

* --- BASE LIMPIA ---
preserve
    keep if tiene_error == 0
    drop err_* tiene_error
    count
    di "Registros limpios: `r(N)'"
    save "${rep_cons_o}\Anexo1_base_limpia.dta", replace
restore

* --- BASE ERRORES ---
preserve
    keep if tiene_error == 1
    count
    di "Registros con errores: `r(N)'"
    save "${rep_cons_o}\Anexo1_base_errores.dta", replace
restore

di ""
di "=== RESUMEN DE ERRORES POR CAMPO ==="
foreach v in cui tipo monto avance f9 comp unid demol nueva reforz cerco sust ampl mobil agua elec {
    qui count if err_`v' == 1
    if `r(N)' > 0 di "  err_`v': `r(N)' registros"
}

di ""
di "Archivos generados:"
di "  $out_clean"
di "  $out_errors"
