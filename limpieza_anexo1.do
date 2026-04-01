/*==============================================================================
  LIMPIEZA - Anexo 01: Relación de locales educativos intervenidos
  Campos ingresados por GR/GL

  Campos limpiados:
    - CUI (solo numérico)
    - Tipo de inversión (PI/IOARR/IRI)
    - Monto de inversión (numérico positivo)
    - Avance físico (0-100%)
    - Fecha de recepción de obra
    - Tiene F9 (SI/NO)
    - Componentes (1-4)
    - Campos SI/NO de intervención
    - Comentarios
==============================================================================*/

clear all
set more off

* ── RUTAS ────────────────────────────────────────────────────────────────────
global reporte    "C:\Users\iraid\Documents\DISCO TERA\MINEDU\TRABAJO-MINEDU\01_MINEDU"
global rep_cons_i "${reporte}\01_input"
global rep_cons_o "${reporte}\03_output"
global rep_cons_t "${reporte}\04_temporal"

// ***************************************************************
// 1. Importar datos
// ***************************************************************
// cellrange(A3) salta el título (fila 1) y la fila en blanco (fila 2),
// usando la fila 3 del Excel directamente como nombres de variables (firstrow).
import excel "${rep_cons_i}\Anexo_1_avance_GR_GL.xlsx", ///
    sheet("Hoja1") cellrange(A3) firstrow clear
save "${rep_cons_t}\Anexo1_GR_GL.dta", replace

use "${rep_cons_t}\Anexo1_GR_GL.dta", clear
count
di "Observations importadas: `r(N)'"
d, short

* ── RENOMBRAR a nombres cortos operativos ────────────────────────────────────
local newnames "nro cod_local region provincia distrito nombre_ie cui tipo monto avance fecha f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec comentarios"
local i = 1
foreach v of varlist _all {
    local newname : word `i' of `newnames'
    if "`newname'" != "" rename `v' `newname'
    local ++i
}

* Eliminar columnas extra vacías que vienen después de Comentarios (cols 26+)
capture drop var26 var27 var28 var29
capture drop _var26 _var27 _var28 _var29

* Etiquetas de variables
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


// ***************************************************************
// 2. Eliminar filas sin CUI (pendientes de registro)
// ***************************************************************
drop if missing(cui)
count
di "Filas con CUI: `r(N)'"

* Columnas de errores por campo (0 = ok, 1 = error)
foreach v in cui tipo monto avance f9 comp unid demol nueva reforz cerco sust ampl mobil agua elec {
    gen byte err_`v' = 0
}


// ***************************************************************
// 3. CUI: solo dígitos
// ***************************************************************
replace cui = strtrim(cui)

* Marcar como error si contiene caracteres no numéricos
replace err_cui = 1 if !regexm(cui, "^[0-9]+$") & !missing(cui)

* Intentar rescatar: extraer secuencia de 5+ dígitos
gen cui_rescued = regexs(0) if regexm(cui, "[0-9]{5,}")
replace cui = cui_rescued if err_cui == 1 & !missing(cui_rescued)
replace err_cui = 0        if err_cui == 1 & !missing(cui_rescued)
drop cui_rescued


// ***************************************************************
// 4. TIPO DE INVERSIÓN → PI / IOARR / IRI
// ***************************************************************
replace tipo = upper(strtrim(tipo))
replace tipo = regexr(tipo, "\s+", " ")

* Marcar montos/números que no son tipo
replace err_tipo = 1 if regexm(tipo, "^[0-9]+\.?[0-9]*$") & !missing(tipo)
replace tipo = "" if regexm(tipo, "^[0-9]+\.?[0-9]*$")

* Marcar S/N como error
replace err_tipo = 1 if tipo == "S/N"
replace tipo = "" if tipo == "S/N"

* Normalizar tipos válidos
replace tipo = "IOARR" if regexm(tipo, "IOARR?|FUR") & err_tipo == 0
replace tipo = "PI"    if regexm(tipo, "^PI$|PROYECTO")   & tipo != "IOARR" & err_tipo == 0
replace tipo = "IRI"   if regexm(tipo, "^IRI$")           & tipo != "IOARR" & tipo != "PI" & err_tipo == 0

replace err_tipo = 1 if !inlist(tipo, "PI", "IOARR", "IRI", "") & !missing(tipo)


// ***************************************************************
// 5. MONTO: convertir a numérico
// ***************************************************************
replace monto = subinstr(monto, ",", ".", .)
replace monto = strtrim(monto)

* Guardar original string para detectar conversiones fallidas vs vacíos
gen monto_orig = monto
destring monto, replace force
* Error solo si había texto pero destring no pudo convertir (no si estaba vacío)
replace err_monto = 1 if missing(monto) & monto_orig != "" & !missing(monto_orig)
replace err_monto = 1 if monto < 0 & !missing(monto)
drop monto_orig


// ***************************************************************
// 6. AVANCE FÍSICO: porcentaje 0–100
// ***************************************************************
replace avance = subinstr(avance, "%", "", .)
replace avance = subinstr(avance, ",", ".", .)
replace avance = strtrim(avance)

* Guardar original string para detectar conversiones fallidas vs vacíos
gen avance_orig = avance
destring avance, replace force

* Si vino como proporción (0–1) → convertir a porcentaje
replace avance = round(avance * 100, 0.01) if !missing(avance) & avance > 0 & avance <= 1

* Error solo si había texto pero destring no pudo convertir (no si estaba vacío)
replace err_avance = 1 if missing(avance) & avance_orig != "" & !missing(avance_orig)
replace err_avance = 1 if (avance < 0 | avance > 100) & !missing(avance)
drop avance_orig


// ***************************************************************
// 7. F9: SI / NO
//    Detecta valores de comp que se corrieron a esta columna
// ***************************************************************
replace f9 = upper(strtrim(f9))
replace f9 = subinstr(f9, "Í", "I", .)

* Detectar corrimiento de columnas (valor de comp en f9)
replace err_f9 = 1 if regexm(f9, "INFRAESTRUCTURA|EQUIPAMIENTO|MOBILIARIO|INTEGRAL|[0-9]\.")
replace f9 = "" if regexm(f9, "INFRAESTRUCTURA|EQUIPAMIENTO|MOBILIARIO|INTEGRAL|[0-9]\.")

* Normalizar SI (incluyendo variantes como "SI SECCIÓN B")
replace f9 = "SI" if inlist(f9, "SI", "SÍ", "S", "1") | regexm(f9, "^SI ")
replace f9 = "NO" if inlist(f9, "NO", "N", "0") | regexm(f9, "^NO ")

replace err_f9 = 1 if !inlist(f9, "SI", "NO", "") & !missing(f9) & err_f9 == 0


// ***************************************************************
// 8. COMPONENTES
// ***************************************************************
replace comp = upper(strtrim(comp))

* Valores claramente inválidos
replace err_comp = 1 if inlist(comp, "-", "_", "-1", "-----------------------")
replace comp = "" if inlist(comp, "-", "_", "-1", "-----------------------")

* Valor 4 / Integral
replace comp = "4. Integral (1, 2 y 3)" if ///
    inlist(comp, "4", "INTEGRAL", "(1,2 Y 3)", "(1,2,3)", "(1-2-3)", "1,2,3", "1, 2 Y 3", "1 Y 2 Y 3") | ///
    regexm(comp, "1[,Y ]+2[,Y ]+3|INTEGR") & comp != ""

* Combos de dos → integral
replace comp = "4. Integral (1, 2 y 3)" if ///
    inlist(comp, "1,2", "1, 2", "1 Y 2", "1,3", "1 Y 3", "2,3", "2 Y 3", ///
                 "2. EQUIPAMIENTO 3.MOBILIARIO", "1,2 Y 3") & comp != "4. Integral (1, 2 y 3)"

* Solo infraestructura
replace comp = "1. Infraestructura." if ///
    inlist(comp, "1", "1.", "1. INFRAESTRUCTURA.", "1. INFRAESTRUCTURA", "1.INFRAESTRUCTURA", "INFRAESTRUCTURA") ///
    & comp != "4. Integral (1, 2 y 3)"

* Solo equipamiento
replace comp = "2. Equipamiento." if ///
    inlist(comp, "2", "2.", "2. EQUIPAMIENTO.", "2.EQUIPAMIENTO", "2. EQUIPAMIENT", "2. EQUIPAMIENTO") ///
    & comp != "4. Integral (1, 2 y 3)"

* Solo mobiliario
replace comp = "3. Mobiliario." if ///
    inlist(comp, "3", "3.", "3. MOBILIARIO.", "3. MOBILIARIO") ///
    & comp != "4. Integral (1, 2 y 3)"

replace err_comp = 1 if ///
    !inlist(comp, "1. Infraestructura.", "2. Equipamiento.", ///
                  "3. Mobiliario.", "4. Integral (1, 2 y 3)", "") & !missing(comp) & err_comp == 0


// ***************************************************************
// 9. CAMPOS SI/NO DE INTERVENCIÓN
//    Mover textos largos de elec a comentarios antes de normalizar
// ***************************************************************
* Mover textos largos de elec a comentarios
replace comentarios = elec if strlen(elec) > 30 & !regexm(upper(elec), "^SI|^NO") & (missing(comentarios) | strtrim(comentarios) == "")
replace comentarios = comentarios + " | " + elec if strlen(elec) > 30 & !regexm(upper(elec), "^SI|^NO") & !missing(comentarios) & strtrim(comentarios) != ""
replace elec = "" if strlen(elec) > 30 & !regexm(upper(elec), "^SI|^NO")

local sino_vars "unid demol nueva reforz cerco sust ampl mobil agua elec"

foreach v of local sino_vars {
    replace `v' = upper(strtrim(`v'))
    replace `v' = subinstr(`v', "Í", "I", .)

    * Marcar valores claramente inválidos
    replace err_`v' = 1 if inlist(`v', "-", "_", "P", "MO")
    replace `v' = "" if inlist(`v', "-", "_", "P", "MO")

    * Marcar textos largos que no empiezan con SI/NO (>30 chars)
    replace err_`v' = 1 if strlen(`v') > 30 & !regexm(`v', "^SI|^NO") & err_`v' == 0
    replace `v' = "" if strlen(`v') > 30 & !regexm(`v', "^SI|^NO")

    * Marcar números que no son 0/1
    replace err_`v' = 1 if regexm(`v', "^[0-9]+$") & !inlist(`v', "0", "1") & err_`v' == 0
    replace `v' = "" if regexm(`v', "^[0-9]+$") & !inlist(`v', "0", "1")

    * Normalizar SI (incluyendo "SI (PARCIAL)", "SI-TOTAL", "SI PARCIAL" → SI)
    replace `v' = "SI" if regexm(`v', "^SI") | inlist(`v', "1", "S", "SI")
    replace `v' = "NO" if regexm(`v', "^NO") | inlist(`v', "0", "N")

    replace err_`v' = 1 if !inlist(`v', "SI", "NO", "") & !missing(`v') & err_`v' == 0
}


// ***************************************************************
// 10. SEPARAR BASES
// ***************************************************************
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
    save "${rep_cons_o}\Anexo1_base_limpia_stata.dta", replace
    export excel using "${rep_cons_o}\Anexo1_base_limpia_stata.xlsx", ///
        firstrow(variables) sheet("Base Limpia") replace
restore

* --- BASE ERRORES ---
preserve
    keep if tiene_error == 1
    count
    di "Registros con errores: `r(N)'"
    save "${rep_cons_o}\Anexo1_base_errores_stata.dta", replace
    export excel using "${rep_cons_o}\Anexo1_base_errores_stata.xlsx", ///
        firstrow(variables) sheet("Registros con Errores") replace
restore

di ""
di "=== RESUMEN DE ERRORES POR CAMPO ==="
foreach v in cui tipo monto avance f9 comp unid demol nueva reforz cerco sust ampl mobil agua elec {
    qui count if err_`v' == 1
    if `r(N)' > 0 di "  err_`v': `r(N)' registros"
}

// ***************************************************************
// 11. VALIDACIÓN CRUZADA CON BASE DE INVERSIONES MINEDU
// ***************************************************************
di ""
di "============================================================"
di "VALIDACIÓN CRUZADA CON BASE DE INVERSIONES MINEDU"
di "============================================================"

* Guardar base actual en temporal
tempfile anexo1_temp
save `anexo1_temp', replace

* Importar Base MINEDU
capture confirm file "${rep_cons_i}\2026.03.23 Base de Inversiones_.xlsx"
if _rc == 0 {
    preserve
        import excel "${rep_cons_i}\2026.03.23 Base de Inversiones_.xlsx", ///
            sheet("Data") firstrow clear

        * Renombrar campos clave
        rename CODIGO_UNICO cui_minedu
        rename DES_TIPO_FORMATO tipo_minedu_raw
        rename COSTO_ACTUALIZADO_BI monto_minedu
        rename TIENE_F9 f9_minedu
        rename AVANCE_FISICO_F9 avance_f9_minedu
        rename AVANCE_FISICO_F12B avance_f12b_minedu
        rename ESTADO estado_minedu
        rename SITUACION situacion_minedu
        rename NOMBRE_INVERSION nombre_inv_minedu

        * Normalizar tipo MINEDU
        gen tipo_minedu = ""
        replace tipo_minedu = "IOARR" if regexm(upper(tipo_minedu_raw), "IOARR")
        replace tipo_minedu = "PI"    if regexm(upper(tipo_minedu_raw), "PROYECTO") & tipo_minedu == ""
        replace tipo_minedu = "IRI"   if regexm(upper(tipo_minedu_raw), "IRI") & tipo_minedu == ""

        replace cui_minedu = strtrim(cui_minedu)

        keep cui_minedu tipo_minedu monto_minedu f9_minedu ///
             avance_f9_minedu avance_f12b_minedu estado_minedu ///
             situacion_minedu nombre_inv_minedu

        tempfile minedu_temp
        save `minedu_temp', replace

        count
        di "Base MINEDU cargada: `r(N)' registros"
    restore

    * Cargar base Anexo1 de nuevo
    use `anexo1_temp', clear

    * Solo filas con CUI válido (no error)
    keep if err_cui == 0 & !missing(cui)

    * Limpiar CUI para merge
    replace cui = strtrim(cui)

    * Renombrar para merge
    rename cui cui_minedu
    rename tipo tipo_anexo1
    rename monto monto_anexo1
    rename f9 f9_anexo1
    rename avance avance_anexo1

    * Merge con Base MINEDU
    merge m:1 cui_minedu using `minedu_temp'

    * Clasificar resultado del cruce
    gen status_cruce = ""
    replace status_cruce = "COINCIDE" if _merge == 3
    replace status_cruce = "SOLO EN ANEXO1 (no validado por MINEDU)" if _merge == 1
    replace status_cruce = "SOLO EN MINEDU (no declarado por GR/GL)" if _merge == 2

    * Para los que coinciden, comparar campos
    gen tipo_ok = ""
    replace tipo_ok = "OK" if tipo_anexo1 == tipo_minedu & _merge == 3
    replace tipo_ok = "DIFERENTE" if tipo_anexo1 != tipo_minedu & _merge == 3

    gen f9_ok = ""
    replace f9_ok = "OK" if upper(f9_anexo1) == upper(f9_minedu) & _merge == 3
    replace f9_ok = "DIFERENTE" if upper(f9_anexo1) != upper(f9_minedu) & _merge == 3

    * Resumen
    di ""
    qui count if _merge == 3
    di "CUIs que coinciden: `r(N)'"
    qui count if _merge == 1
    di "CUIs solo en Anexo1: `r(N)'"
    qui count if _merge == 2
    di "CUIs solo en MINEDU: `r(N)'"

    qui count if tipo_ok == "DIFERENTE"
    di "Tipo discrepante: `r(N)'"
    qui count if f9_ok == "DIFERENTE"
    di "F9 discrepante: `r(N)'"

    * Exportar reporte de cruce
    rename cui_minedu cui
    keep cui status_cruce nombre_ie nombre_inv_minedu ///
         tipo_anexo1 tipo_minedu tipo_ok ///
         monto_anexo1 monto_minedu ///
         f9_anexo1 f9_minedu f9_ok ///
         avance_anexo1 avance_f9_minedu avance_f12b_minedu ///
         estado_minedu situacion_minedu

    export excel using "${rep_cons_o}\Anexo1_validacion_cruce_stata.xlsx", ///
        firstrow(variables) sheet("Validacion Cruce") replace

    di ""
    di "Reporte de cruce generado: ${rep_cons_o}\Anexo1_validacion_cruce_stata.xlsx"
}
else {
    di "AVISO: No se encontró la Base de Inversiones MINEDU en ${rep_cons_i}"
    di "  Se omite la validación cruzada."
}

* Restaurar base original
use `anexo1_temp', clear

di ""
di "Archivos generados:"
di "  ${rep_cons_o}\Anexo1_base_limpia_stata.xlsx"
di "  ${rep_cons_o}\Anexo1_base_errores_stata.xlsx"
di "  ${rep_cons_o}\Anexo1_validacion_cruce.xlsx"
