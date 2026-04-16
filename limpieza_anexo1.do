* ===========================================================================
*
* LIMPIEZA DE ANEXO 1 - Inversiones en infraestructura educativa
* Declaraciones de GR y GL
* Autor: Iraida Chavez
* Correo: diplan11@minedu.gob.pe
* Fecha de actualización: 16/04/2026
*
* ===========================================================================

clear all
set more off

global reporte "C:\Users\diplan11\Documents\00_MINEDU\TRABAJO-MINEDU\SOL_GR_GL"
global rep_cons_i "${reporte}\01_input"
global rep_cons_o "${reporte}\03_output"
global rep_cons_t "${reporte}\04_temporal"

* -------------------------------------------------------------------------
* 1. IMPORTAR ANEXO 1
* -------------------------------------------------------------------------

import excel using "${rep_cons_i}\Anexo_1_avance_GR_GL.xlsx", ///
	sheet("Hoja1") cellrange(A3) firstrow clear

save "${rep_cons_t}\anexo1_raw.dta", replace

* renombro las variables por posicion porque los nombres del excel son un desastre
local varnames "nro cod_local region provincia distrito nombre_ie cui tipo monto avance fecha f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec comentarios"
local i = 1
foreach v of varlist _all {
	local newname : word `i' of `varnames'
	if "`newname'" != "" {
		rename `v' `newname'
	}
	local i = `i' + 1
}

* a veces vienen columnas de mas al final
capture drop Z
capture drop var26
capture drop var27
capture drop var28
capture drop var29

* etiquetas
label var nro "Nro"
label var cod_local "Código de local educativo"
label var region "Región"
label var provincia "Provincia"
label var distrito "Distrito"
label var nombre_ie "Nombre de la IE"
label var cui "CUI de la inversión"
label var tipo "Tipo de inversión"
label var monto "Monto de inversión (S/)"
label var avance "Avance (%)"
label var fecha "Fecha de inicio"
label var f9 "Formato 9 aprobado (SI/NO)"
label var comp "Componente"
label var unid "Unificado (SI/NO)"
label var cod_mod "Código modular"
label var demol "Demolición (SI/NO)"
label var nueva "Construcción nueva (SI/NO)"
label var reforz "Reforzamiento (SI/NO)"
label var cerco "Cerco perimétrico (SI/NO)"
label var sust "Sustitución (SI/NO)"
label var ampl "Ampliación (SI/NO)"
label var mobil "Mobiliario (SI/NO)"
label var agua "Agua y saneamiento (SI/NO)"
label var elec "Electricidad (SI/NO)"
label var comentarios "Comentarios"

* las filas sin CUI no sirven para nada
drop if missing(cui)

* convertir todo a string por si acaso (excel importa raro a veces)
foreach v of varlist cui tipo monto avance f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec comentarios {
	capture tostring `v', replace force
	capture replace `v' = "" if `v' == "."
}

save "${rep_cons_t}\anexo1_renamed.dta", replace

* -------------------------------------------------------------------------
* 2. GENERAR FLAGS DE ERROR (todos empiezan en 0)
* -------------------------------------------------------------------------

foreach v in cui tipo monto avance f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec {
	gen byte err_`v' = 0
}

* -------------------------------------------------------------------------
* 3. LIMPIAR CODIGO LOCAL (6 dígitos, pad con ceros a la izquierda)
* -------------------------------------------------------------------------

capture tostring cod_local, replace force
replace cod_local = strtrim(cod_local)
replace cod_local = subinstr(cod_local, " ", "", .)
replace cod_local = "" if cod_local == "."

* si tiene letras => missing
replace cod_local = "" if regexm(cod_local, "[a-zA-Z]")

* quitar decimales (.0)
replace cod_local = regexr(cod_local, "\.0+$", "")

* pad con ceros a la izquierda hasta 6 dígitos
replace cod_local = "0" * (6 - strlen(cod_local)) + cod_local ///
	if strlen(cod_local) < 6 & strlen(cod_local) > 0

* -------------------------------------------------------------------------
* 4. LIMPIAR CUI (solo números)
* -------------------------------------------------------------------------

replace cui = strtrim(cui)
replace cui = subinstr(cui, " ", "", .)

* si no es numerico puro, intento rescatar la secuencia de digitos
replace err_cui = 1 if !regexm(cui, "^[0-9]+$") & cui != ""
* intento sacar algo util: secuencia de 5+ digitos
replace cui = regexs(0) if regexm(cui, "[0-9]{5,}") & err_cui == 1
* si logre rescatar, quito el error
replace err_cui = 0 if regexm(cui, "^[0-9]+$") & err_cui == 1

* -------------------------------------------------------------------------
* 5. LIMPIAR TIPO DE INVERSION
* -------------------------------------------------------------------------

replace tipo = upper(strtrim(tipo))
replace tipo = subinstr(tipo, "  ", " ", .)

* si ponen solo numeros o S/N es error
replace err_tipo = 1 if regexm(tipo, "^[0-9]+$")
replace err_tipo = 1 if tipo == "S/N" | tipo == "SN"

* normalizar los tipos validos
replace tipo = "IOARR" if regexm(tipo, "IOAR") | regexm(tipo, "FUR")
replace tipo = "PI" if regexm(tipo, "PROYECTO") | tipo == "PI"
replace tipo = "IRI" if tipo == "IRI"

* lo que quedo y no es IOARR, PI o IRI => error
replace err_tipo = 1 if !inlist(tipo, "IOARR", "PI", "IRI") & tipo != "" & err_tipo == 0

* -------------------------------------------------------------------------
* 6. LIMPIAR MONTO
* -------------------------------------------------------------------------

gen monto_orig = monto

* las comas como separador decimal
replace monto = subinstr(monto, ",", ".", .)
* a veces ponen S/ adelante
replace monto = subinstr(monto, "S/", "", .)
replace monto = subinstr(monto, "s/", "", .)
replace monto = strtrim(monto)

destring monto, replace force

* si tenia algo y se perdio en el destring, es error
replace err_monto = 1 if monto == . & monto_orig != "" & monto_orig != "."
* montos negativos tambien
replace err_monto = 1 if monto < 0 & monto != .

* -------------------------------------------------------------------------
* 7. LIMPIAR AVANCE
* -------------------------------------------------------------------------

gen avance_orig = avance

replace avance = subinstr(avance, "%", "", .)
replace avance = subinstr(avance, ",", ".", .)
replace avance = strtrim(avance)

destring avance, replace force

* si estaba entre 0 y 1 probablemente es proporcion, paso a porcentaje
replace avance = avance * 100 if avance >= 0 & avance <= 1 & avance != .

* error si no se pudo convertir pero habia dato
replace err_avance = 1 if avance == . & avance_orig != "" & avance_orig != "."
* fuera de rango
replace err_avance = 1 if (avance < 0 | avance > 100) & avance != .

* -------------------------------------------------------------------------
* 8. LIMPIAR F9 (FORMATO 9)
* -------------------------------------------------------------------------

replace f9 = upper(strtrim(f9))
replace f9 = subinstr(f9, "Í", "I", .)

* detectar valores que son de componente (se corrieron las columnas)
replace err_f9 = 1 if regexm(f9, "INFRAESTRUCTURA")
replace err_f9 = 1 if regexm(f9, "EQUIPAMIENTO")
replace err_f9 = 1 if regexm(f9, "MOBILIARIO")
replace err_f9 = 1 if regexm(f9, "INTEGRAL")

* normalizar SI/NO
replace f9 = "SI" if inlist(f9, "SI", "SÍ", "SI, CON LIQUIDACION", "SI/SNIP", "SI SECCION B")
replace f9 = "SI" if regexm(f9, "^SI,") | regexm(f9, "^SI ") | regexm(f9, "^Sí,")
replace f9 = "NO" if inlist(f9, "NO", "N", "0", "NO CULMINADA", "EN PROCESO")

* lo que no es SI, NO ni vacio es error
replace err_f9 = 1 if !inlist(f9, "SI", "NO", "") & err_f9 == 0

* -------------------------------------------------------------------------
* 9. LIMPIAR COMPONENTE
* -------------------------------------------------------------------------

replace comp = upper(strtrim(comp))
replace comp = subinstr(comp, "  ", " ", .)

* primero detecto combinaciones => integral
replace comp = "4.INTEGRAL" if regexm(comp, "1, 2, 3, 4")
replace comp = "4.INTEGRAL" if regexm(comp, "1, 2, 4") | regexm(comp, "1, 2, 3")

* combos de dos => integral
replace comp = "4.INTEGRAL" if regexm(comp, "1, 2") | regexm(comp, "1, 3")
replace comp = "4.INTEGRAL" if comp == "2, 3"

* numeros solos
replace comp = "1.INFRAESTRUCTURA" if comp == "1"
replace comp = "2.EQUIPAMIENTO" if comp == "2"
replace comp = "3.MOBILIARIO" if comp == "3"
replace comp = "4.INTEGRAL" if comp == "4"

* limpiar valores que no son componentes
replace comp = "" if regexm(comp, "^0$") | regexm(comp, "^NO$") | regexm(comp, "^SI$")

* lo que no matcheo
replace err_comp = 1 if !inlist(comp, "1.INFRAESTRUCTURA", "2.EQUIPAMIENTO", "3.MOBILIARIO", "4.INTEGRAL") & comp != ""

* -------------------------------------------------------------------------
* 10. LIMPIAR CAMPOS SI/NO
* -------------------------------------------------------------------------

* primero muevo texto largo de electricidad a comentarios
replace comentarios = comentarios + " | " + elec if strlen(elec) > 20 & elec != ""
replace elec = "" if strlen(elec) > 20

* ahora limpio todos los campos SI/NO en loop
foreach v in unid demol nueva reforz cerco sust ampl mobil agua elec {
	replace `v' = upper(strtrim(`v'))
	replace `v' = subinstr(`v', "Í", "I", .)

	* marcar valores raros como error
	replace err_`v' = 1 if inlist(`v', "-", "_", "P", "MO")
	replace err_`v' = 1 if strlen(`v') > 10 & `v' != ""
	* numeros que no son 0 o 1
	replace err_`v' = 1 if regexm(`v', "^[0-9]+$") & !inlist(`v', "0", "1")

	* normalizar
	replace `v' = "SI" if inlist(`v', "SI", "SÍ", "S", "1", "X")
	replace `v' = "NO" if inlist(`v', "NO", "N", "0")

	* lo demas
	replace err_`v' = 1 if !inlist(`v', "SI", "NO", "") & err_`v' == 0
}

* -------------------------------------------------------------------------
* 11. LIMPIAR CODIGO MODULAR (7 dígitos, separados por /)
* -------------------------------------------------------------------------

replace cod_mod = strtrim(cod_mod)

* guiones y underscores => missing
replace cod_mod = "" if inlist(cod_mod, "-", "_", "--", "---")
replace cod_mod = "" if regexm(cod_mod, "^-+$")

* estos tampoco son codigos modulares
replace cod_mod = "" if inlist(upper(cod_mod), "SI", "SÍ", "NO", "NINGUNA", "0", "SECUNDARIA")

* letras => missing (son descripciones de obras, no códigos)
replace cod_mod = "" if regexm(cod_mod, "[a-zA-Z]")

* separadores: /, -, saltos de linea => /
replace cod_mod = subinstr(cod_mod, ",", "/", .)
replace cod_mod = subinstr(cod_mod, "-", "/", .)
replace cod_mod = subinstr(cod_mod, char(10), "/", .)
replace cod_mod = subinstr(cod_mod, char(13), "/", .)
replace cod_mod = subinstr(cod_mod, " ", "", .)

* limpiar separadores dobles
replace cod_mod = subinstr(cod_mod, "//", "/", .)
replace cod_mod = regexr(cod_mod, "^/", "")
replace cod_mod = regexr(cod_mod, "/$", "")

* quitar decimales (.0) de cada código
replace cod_mod = regexr(cod_mod, "\.0+$", "")

* pad cada código a 7 dígitos con ceros a la izquierda
* para múltiples códigos separados por /, procesamos cada uno
gen cod_mod_clean = ""
gen n_codigos = 0

* contar cuantos codigos hay (separados por /)
replace n_codigos = length(cod_mod) - length(subinstr(cod_mod, "/", "", .)) + 1 if cod_mod != ""

* procesar hasta 10 codigos por celda
forvalues k = 1/10 {
	gen _cod_`k' = ""
}

* extraer cada codigo individual
split cod_mod, parse("/") gen(_parte_)

* pad cada parte a 7 dígitos
local maxparts = 10
forvalues k = 1/`maxparts' {
	capture confirm variable _parte_`k'
	if _rc == 0 {
		replace _parte_`k' = strtrim(_parte_`k')
		* pad con ceros a la izquierda
		replace _parte_`k' = "0" * (7 - strlen(_parte_`k')) + _parte_`k' ///
			if strlen(_parte_`k') < 7 & strlen(_parte_`k') > 0
	}
}

* reconstruir cod_mod con todos los códigos separados por /
replace cod_mod_clean = ""
forvalues k = 1/`maxparts' {
	capture confirm variable _parte_`k'
	if _rc == 0 {
		replace cod_mod_clean = cod_mod_clean + "/" + _parte_`k' ///
			if _parte_`k' != "" & cod_mod_clean != ""
		replace cod_mod_clean = _parte_`k' ///
			if _parte_`k' != "" & cod_mod_clean == ""
	}
}

replace cod_mod = cod_mod_clean
drop cod_mod_clean n_codigos _parte_* _cod_*

* si queda algo que no es digitos y /, es error
replace err_cod_mod = 1 if !regexm(cod_mod, "^[0-9/]*$") & cod_mod != ""

* -------------------------------------------------------------------------
* 12. SEPARAR BASES
* -------------------------------------------------------------------------

gen byte tiene_error = (err_cui | err_tipo | err_monto | err_avance | ///
	err_f9 | err_comp | err_unid | err_cod_mod | err_demol | ///
	err_nueva | err_reforz | err_cerco | err_sust | ///
	err_ampl | err_mobil | err_agua | err_elec)

* --- BASE LIMPIA ---
preserve
	keep if tiene_error == 0
	drop err_* tiene_error monto_orig avance_orig
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
foreach v in cui tipo monto avance f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec {
	qui count if err_`v' == 1
	if `r(N)' > 0 di "  err_`v': `r(N)' registros"
}

di ""
di "Archivos generados:"
di "  ${rep_cons_o}\Anexo1_base_limpia.dta"
di "  ${rep_cons_o}\Anexo1_base_errores.dta"
