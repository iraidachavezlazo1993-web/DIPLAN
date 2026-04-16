* ===========================================================================
*
* LIMPIEZA DE ANEXO 1 - Inversiones en infraestructura educativa
* Declaraciones de GR y GL
* Autor: Iraida Chavez
* Correo: diplan11@minedu.gob.pe
* Fecha de actualización: 16/04/2026
*
* Flujo:
*   1-11. Limpieza de campos del Anexo 1
*   12.   Preparar Banco de Inversiones MINEDU (2026.04.13)
*   13.   Preparar Vinculaciones (2026.04.14) - agrupa por CUI
*   14.   Cruce: Anexo1 <- Banco MINEDU (prevalece en tipo/monto/f9)
*         Cruce: Anexo1 <- Base MEF Completa (respaldo si CUI no está en MINEDU)
*         Cruce: Anexo1 <- Vinculaciones (valida y completa cod_local/cod_mod)
*   15.   Separar base limpia y base con errores, exportar a Excel
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
* 3. LIMPIAR CODIGO LOCAL (6 dígitos, separados por /)
* -------------------------------------------------------------------------

capture tostring cod_local, replace force
replace cod_local = strtrim(cod_local)
replace cod_local = subinstr(cod_local, " ", "", .)
replace cod_local = "" if cod_local == "."

* guiones y underscores => missing
replace cod_local = "" if inlist(cod_local, "-", "_", "--", "---")
replace cod_local = "" if regexm(cod_local, "^-+$")

* si tiene letras => missing
replace cod_local = "" if regexm(cod_local, "[a-zA-Z]")

* quitar decimales (.0)
replace cod_local = regexr(cod_local, "\.0+$", "")

* separadores: /, -, comas => /
replace cod_local = subinstr(cod_local, ",", "/", .)
replace cod_local = subinstr(cod_local, "-", "/", .)
replace cod_local = subinstr(cod_local, char(10), "/", .)
replace cod_local = subinstr(cod_local, char(13), "/", .)
replace cod_local = subinstr(cod_local, " ", "", .)

* limpiar separadores dobles
replace cod_local = subinstr(cod_local, "//", "/", .)
replace cod_local = regexr(cod_local, "^/", "")
replace cod_local = regexr(cod_local, "/$", "")

* pad cada código a 6 dígitos con ceros a la izquierda
* para múltiples códigos separados por /, procesamos cada uno
split cod_local, parse("/") gen(_cl_parte_)

gen cod_local_clean = ""
local maxparts = 10
forvalues k = 1/`maxparts' {
	capture confirm variable _cl_parte_`k'
	if _rc == 0 {
		replace _cl_parte_`k' = strtrim(_cl_parte_`k')
		replace _cl_parte_`k' = "0" * (6 - strlen(_cl_parte_`k')) + _cl_parte_`k' ///
			if strlen(_cl_parte_`k') < 6 & strlen(_cl_parte_`k') > 0
	}
}

forvalues k = 1/`maxparts' {
	capture confirm variable _cl_parte_`k'
	if _rc == 0 {
		replace cod_local_clean = cod_local_clean + "/" + _cl_parte_`k' ///
			if _cl_parte_`k' != "" & cod_local_clean != ""
		replace cod_local_clean = _cl_parte_`k' ///
			if _cl_parte_`k' != "" & cod_local_clean == ""
	}
}

replace cod_local = cod_local_clean
drop cod_local_clean _cl_parte_*

* -------------------------------------------------------------------------
* 4. LIMPIAR CUI (7 dígitos, separados por /)
* -------------------------------------------------------------------------

replace cui = strtrim(cui)
replace cui = subinstr(cui, " ", "", .)

* quitar decimales (.0)
replace cui = regexr(cui, "\.0+$", "")

* separadores: /, -, comas => /
replace cui = subinstr(cui, ",", "/", .)
replace cui = subinstr(cui, "-", "/", .)
replace cui = subinstr(cui, char(10), "/", .)
replace cui = subinstr(cui, char(13), "/", .)
replace cui = subinstr(cui, " ", "", .)

* limpiar separadores dobles
replace cui = subinstr(cui, "//", "/", .)
replace cui = regexr(cui, "^/", "")
replace cui = regexr(cui, "/$", "")

* si tiene letras y no se puede rescatar => error
replace err_cui = 1 if regexm(cui, "[a-zA-Z]") & cui != ""
* intento sacar algo util: secuencia de 5+ digitos
replace cui = regexs(0) if regexm(cui, "[0-9]{5,}") & err_cui == 1
replace err_cui = 0 if regexm(cui, "^[0-9/]+$") & err_cui == 1

* pad cada código a 7 dígitos con ceros a la izquierda
split cui, parse("/") gen(_cui_parte_)

gen cui_clean = ""
local maxparts = 10
forvalues k = 1/`maxparts' {
	capture confirm variable _cui_parte_`k'
	if _rc == 0 {
		replace _cui_parte_`k' = strtrim(_cui_parte_`k')
		replace _cui_parte_`k' = "0" * (7 - strlen(_cui_parte_`k')) + _cui_parte_`k' ///
			if strlen(_cui_parte_`k') < 7 & strlen(_cui_parte_`k') > 0
	}
}

forvalues k = 1/`maxparts' {
	capture confirm variable _cui_parte_`k'
	if _rc == 0 {
		replace cui_clean = cui_clean + "/" + _cui_parte_`k' ///
			if _cui_parte_`k' != "" & cui_clean != ""
		replace cui_clean = _cui_parte_`k' ///
			if _cui_parte_`k' != "" & cui_clean == ""
	}
}

replace cui = cui_clean
drop cui_clean _cui_parte_*

* si queda algo que no es digitos y /, es error
replace err_cui = 1 if !regexm(cui, "^[0-9/]+$") & cui != "" & err_cui == 0

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

* guardo el estado del anexo limpio antes del cruce
save "${rep_cons_t}\anexo1_limpio_precuce.dta", replace

* -------------------------------------------------------------------------
* 12. PREPARAR BANCO DE INVERSIONES
* -------------------------------------------------------------------------
* el header real esta en la fila 5 (A5), las primeras 4 filas son metadata

capture noisily {
	import excel using "${rep_cons_i}\2026.04.13 Base de Inversiones.xlsx", ///
		sheet("Data") cellrange(A5) firstrow clear

	* renombrar CUI para el match
	capture rename CODIGO_UNICO cui_banco
	tostring cui_banco, replace force
	replace cui_banco = strtrim(cui_banco)
	replace cui_banco = regexr(cui_banco, "\.0+$", "")

	* pad a 7 dígitos
	replace cui_banco = "0" * (7 - strlen(cui_banco)) + cui_banco ///
		if strlen(cui_banco) < 7 & strlen(cui_banco) > 0

	* conservar solo las variables que voy a usar para enriquecer
	keep cui_banco DES_TIPO_FORMATO COSTO_INV_TOTAL_BI TIENE_F9 ///
		AVANCE_FISICO_F9 AVANCE_FISICO_F12B ESTADO SITUACION ///
		COSTO_ACTUALIZADO_BI FECHA_REGISTRO FECHA_VIABILIDAD ///
		TIENE_F8 NOMBRE_INVERSION DEPARTAMENTO_CUI PROVINCIA_CUI DISTRITO

	* renombrar para evitar conflictos con anexo
	rename DES_TIPO_FORMATO banco_tipo_raw
	rename COSTO_INV_TOTAL_BI banco_monto
	rename TIENE_F9 banco_f9
	rename AVANCE_FISICO_F9 banco_avance_f9
	rename AVANCE_FISICO_F12B banco_avance_f12b
	rename ESTADO banco_estado
	rename SITUACION banco_situacion
	rename COSTO_ACTUALIZADO_BI banco_monto_actualizado
	rename FECHA_REGISTRO banco_fecha_registro
	rename FECHA_VIABILIDAD banco_fecha_viabilidad
	rename TIENE_F8 banco_tiene_f8
	rename NOMBRE_INVERSION banco_nombre_inv
	rename DEPARTAMENTO_CUI banco_departamento
	rename PROVINCIA_CUI banco_provincia
	rename DISTRITO banco_distrito

	duplicates drop cui_banco, force
	save "${rep_cons_t}\banco_prep.dta", replace

	di "Banco de Inversiones preparado: `=_N' registros únicos"
}

* -------------------------------------------------------------------------
* 13. PREPARAR VINCULACIONES (agrupar por CUI)
* -------------------------------------------------------------------------
* genera un dataset con un registro por CUI y las listas de cod_local y cod_mod

capture noisily {
	import excel using "${rep_cons_i}\2026.04.14 Vinculaciones.xlsx", ///
		sheet("Vinculaciones") firstrow clear

	* Filtrar solo estados "Vinculado"
	keep if regexm(ESTADO_VINCULACION, "^Vinculado")

	* Normalizar CUI a 7, cod_local a 6, cod_mod a 7
	tostring CUI CODIGO_LOCAL CODIGO_MODULAR, replace force

	foreach v in CUI CODIGO_LOCAL CODIGO_MODULAR {
		replace `v' = strtrim(`v')
		replace `v' = regexr(`v', "\.0+$", "")
	}

	replace CUI = "0" * (7 - strlen(CUI)) + CUI ///
		if strlen(CUI) < 7 & strlen(CUI) > 0
	replace CODIGO_LOCAL = "0" * (6 - strlen(CODIGO_LOCAL)) + CODIGO_LOCAL ///
		if strlen(CODIGO_LOCAL) < 6 & strlen(CODIGO_LOCAL) > 0
	replace CODIGO_MODULAR = "0" * (7 - strlen(CODIGO_MODULAR)) + CODIGO_MODULAR ///
		if strlen(CODIGO_MODULAR) < 7 & strlen(CODIGO_MODULAR) > 0

	rename CUI cui_banco
	rename CODIGO_LOCAL vinc_cod_local
	rename CODIGO_MODULAR vinc_cod_mod
	rename GRUPO vinc_grupo

	keep cui_banco vinc_cod_local vinc_cod_mod vinc_grupo

	* colapsar a una fila por CUI: concatenar cod_local y cod_mod únicos con /
	duplicates drop cui_banco vinc_cod_local, force
	bysort cui_banco (vinc_cod_local): gen _seq_cl = _n
	qui sum _seq_cl
	local max_cl = r(max)

	bysort cui_banco (vinc_cod_local): gen cod_local_oficial = vinc_cod_local if _n == 1
	bysort cui_banco (vinc_cod_local): replace cod_local_oficial = ///
		cod_local_oficial[_n-1] + "/" + vinc_cod_local if _n > 1
	bysort cui_banco (vinc_cod_local): replace cod_local_oficial = cod_local_oficial[_N]

	bysort cui_banco (vinc_cod_mod): gen cod_mod_oficial = vinc_cod_mod if _n == 1
	bysort cui_banco (vinc_cod_mod): replace cod_mod_oficial = ///
		cod_mod_oficial[_n-1] + "/" + vinc_cod_mod if _n > 1
	bysort cui_banco (vinc_cod_mod): replace cod_mod_oficial = cod_mod_oficial[_N]

	bysort cui_banco: keep if _n == 1
	drop _seq_cl vinc_cod_local vinc_cod_mod

	save "${rep_cons_t}\vinc_prep.dta", replace
	di "Vinculaciones preparado: `=_N' CUIs únicos"
}

* -------------------------------------------------------------------------
* 14. CRUCE CON BANCO Y VINCULACIONES
* -------------------------------------------------------------------------

use "${rep_cons_t}\anexo1_limpio_precuce.dta", clear

* para el merge uso solo el primer CUI si hay varios (no debería pasar en anexo)
gen cui_primero = cui
replace cui_primero = substr(cui, 1, strpos(cui + "/", "/") - 1) if strpos(cui, "/") > 0

rename cui_primero cui_banco

* --- merge con Banco de Inversiones ---
capture confirm file "${rep_cons_t}\banco_prep.dta"
if _rc == 0 {
	merge m:1 cui_banco using "${rep_cons_t}\banco_prep.dta", ///
		keep(master match) generate(_merge_banco)
	gen cui_en_banco = "SI" if _merge_banco == 3
	replace cui_en_banco = "NO" if _merge_banco == 1
	drop _merge_banco

	* reemplazar tipo si hay dato del banco (prevalece Banco)
	replace tipo = "IOARR" if regexm(upper(banco_tipo_raw), "IOARR|FUR") & !missing(banco_tipo_raw)
	replace tipo = "IRI"   if regexm(upper(banco_tipo_raw), "IRI") & !missing(banco_tipo_raw) & tipo != "IOARR"
	replace tipo = "PI"    if regexm(upper(banco_tipo_raw), "PROYECTO") & !missing(banco_tipo_raw) & !inlist(tipo, "IOARR", "IRI")
	replace err_tipo = 0 if cui_en_banco == "SI" & inlist(tipo, "IOARR", "IRI", "PI")

	* reemplazar monto si hay dato del banco
	capture destring banco_monto, replace force
	replace monto = banco_monto if !missing(banco_monto) & cui_en_banco == "SI"
	replace err_monto = 0 if cui_en_banco == "SI" & !missing(monto)

	* reemplazar f9 si hay dato del banco
	replace f9 = upper(strtrim(banco_f9)) if inlist(upper(strtrim(banco_f9)), "SI", "NO") & cui_en_banco == "SI"
	replace err_f9 = 0 if cui_en_banco == "SI" & inlist(f9, "SI", "NO")

	* completar avance si falta
	capture destring banco_avance_f9 banco_avance_f12b, replace force
	replace avance = banco_avance_f9 if missing(avance) & !missing(banco_avance_f9) & banco_avance_f9 > 0
	replace avance = banco_avance_f12b if missing(avance) & !missing(banco_avance_f12b)
	replace avance = avance * 100 if avance >= 0 & avance <= 1 & !missing(avance)
	replace err_avance = 0 if !missing(avance) & avance >= 0 & avance <= 100

	drop banco_tipo_raw banco_monto banco_f9
}
else {
	gen cui_en_banco = "NO VERIFICADO"
}

* --- preparar Base MEF Completa ---
capture noisily {
	import excel using "${rep_cons_i}\Rep_Inversiones_13ABR2026_MEFCOMPLETA.xlsx", ///
		sheet("INVERSIONES") firstrow clear

	capture rename CODIGO_UNICO cui_banco
	tostring cui_banco, replace force
	replace cui_banco = strtrim(cui_banco)
	replace cui_banco = regexr(cui_banco, "\.0+$", "")
	replace cui_banco = "0" * (7 - strlen(cui_banco)) + cui_banco ///
		if strlen(cui_banco) < 7 & strlen(cui_banco) > 0

	keep cui_banco TIPO_INVERSION COSTO_ACTUALIZADO MONTO_VIABLE ///
		TIENE_F9 AVANCE_FISICO AVANCE_EJECUCION ESTADO SITUACION ///
		FUNCION PROGRAMA NOMBRE_INVERSION DEPARTAMENTO PROVINCIA DISTRITO ///
		TIENE_F8 CULMINADA

	rename TIPO_INVERSION mef_tipo_raw
	rename COSTO_ACTUALIZADO mef_costo_act
	rename MONTO_VIABLE mef_monto_viable
	rename TIENE_F9 mef_f9
	rename AVANCE_FISICO mef_avance_fisico
	rename AVANCE_EJECUCION mef_avance_ejec
	rename ESTADO mef_estado
	rename SITUACION mef_situacion
	rename FUNCION mef_funcion
	rename PROGRAMA mef_programa
	rename NOMBRE_INVERSION mef_nombre_inv
	rename DEPARTAMENTO mef_departamento
	rename PROVINCIA mef_provincia
	rename DISTRITO mef_distrito
	rename TIENE_F8 mef_tiene_f8
	rename CULMINADA mef_culminada

	duplicates drop cui_banco, force
	save "${rep_cons_t}\mef_prep.dta", replace
	di "Base MEF preparada: `=_N' registros únicos"
}

* --- merge con Base MEF (respaldo para CUI no en Banco MINEDU) ---
* primero recuperar el anexo
use "${rep_cons_t}\anexo1_limpio_precuce.dta", clear
gen cui_primero2 = cui
replace cui_primero2 = substr(cui, 1, strpos(cui + "/", "/") - 1) if strpos(cui, "/") > 0
rename cui_primero2 cui_banco

* re-merge con Banco (para tener cui_en_banco)
capture confirm file "${rep_cons_t}\banco_prep.dta"
if _rc == 0 {
	merge m:1 cui_banco using "${rep_cons_t}\banco_prep.dta", ///
		keep(master match) generate(_merge_banco2)
	gen cui_en_banco2 = "SI" if _merge_banco2 == 3
	replace cui_en_banco2 = "NO" if _merge_banco2 == 1
	drop _merge_banco2
}
else {
	gen cui_en_banco2 = "NO"
}

capture confirm file "${rep_cons_t}\mef_prep.dta"
if _rc == 0 {
	merge m:1 cui_banco using "${rep_cons_t}\mef_prep.dta", ///
		keep(master match) generate(_merge_mef)
	gen cui_en_mef = "SI" if _merge_mef == 3
	replace cui_en_mef = "NO" if _merge_mef == 1
	drop _merge_mef

	* solo complementar si NO está en Banco MINEDU
	* tipo
	replace tipo = "IRI" if regexm(upper(mef_tipo_raw), "IRI") & cui_en_banco2 == "NO" & cui_en_mef == "SI" ///
		& (missing(tipo) | err_tipo == 1)
	replace tipo = "IOARR" if regexm(upper(mef_tipo_raw), "IOARR|FUR") & cui_en_banco2 == "NO" & cui_en_mef == "SI" ///
		& (missing(tipo) | err_tipo == 1)
	replace tipo = "PI" if regexm(upper(mef_tipo_raw), "PIP|PROYECTO") & cui_en_banco2 == "NO" & cui_en_mef == "SI" ///
		& (missing(tipo) | err_tipo == 1)
	replace err_tipo = 0 if inlist(tipo, "IOARR", "IRI", "PI") & cui_en_mef == "SI"

	* monto: completar si falta
	capture destring mef_costo_act, replace force
	replace monto = mef_costo_act if missing(monto) & !missing(mef_costo_act) & cui_en_banco2 == "NO"
	replace err_monto = 0 if !missing(monto) & cui_en_mef == "SI"

	* f9: completar si falta
	replace f9 = "SI" if inlist(upper(strtrim(mef_f9)), "SI", "SÍ") & cui_en_banco2 == "NO" & (missing(f9) | err_f9 == 1)
	replace f9 = "NO" if upper(strtrim(mef_f9)) == "NO" & cui_en_banco2 == "NO" & (missing(f9) | err_f9 == 1)
	replace err_f9 = 0 if inlist(f9, "SI", "NO") & cui_en_mef == "SI"

	drop mef_tipo_raw mef_costo_act mef_f9 mef_avance_fisico mef_avance_ejec
}
else {
	gen cui_en_mef = "NO VERIFICADO"
}

* reconstruir el dataset final: volver a cargar el precuce y hacer todos los merges juntos
* (esto es necesario porque Stata no permite merge sobre un dataset ya mergeado fácilmente)
drop cui_banco cui_en_banco2

* --- merge con Vinculaciones ---
* regenerar cui_banco para el merge
gen cui_banco = cui
replace cui_banco = substr(cui, 1, strpos(cui + "/", "/") - 1) if strpos(cui, "/") > 0

capture confirm file "${rep_cons_t}\vinc_prep.dta"
if _rc == 0 {
	merge m:1 cui_banco using "${rep_cons_t}\vinc_prep.dta", ///
		keep(master match) generate(_merge_vinc)
	gen cui_en_vinc = "SI" if _merge_vinc == 3
	replace cui_en_vinc = "NO" if _merge_vinc == 1
	drop _merge_vinc

	* validar si los cod_local del anexo coinciden con los oficiales
	gen cod_local_match = ""
	replace cod_local_match = "SI" if cod_local == cod_local_oficial & !missing(cod_local) & !missing(cod_local_oficial)
	replace cod_local_match = "MISMATCH" if cod_local != cod_local_oficial & !missing(cod_local) & !missing(cod_local_oficial)
	replace cod_local_match = "COMPLETADO" if missing(cod_local) & !missing(cod_local_oficial)
	replace cod_local = cod_local_oficial if missing(cod_local) & !missing(cod_local_oficial)

	* idem para cod_mod
	gen cod_mod_match = ""
	replace cod_mod_match = "SI" if cod_mod == cod_mod_oficial & !missing(cod_mod) & !missing(cod_mod_oficial)
	replace cod_mod_match = "MISMATCH" if cod_mod != cod_mod_oficial & !missing(cod_mod) & !missing(cod_mod_oficial)
	replace cod_mod_match = "COMPLETADO" if missing(cod_mod) & !missing(cod_mod_oficial)
	replace cod_mod = cod_mod_oficial if missing(cod_mod) & !missing(cod_mod_oficial)
}
else {
	gen cui_en_vinc = "NO VERIFICADO"
	gen cod_local_oficial = ""
	gen cod_mod_oficial = ""
	gen cod_local_match = ""
	gen cod_mod_match = ""
	gen vinc_grupo = ""
}

rename cui_banco cui_primero_usado

* -------------------------------------------------------------------------
* 15. SEPARAR BASES
* -------------------------------------------------------------------------

gen byte tiene_error = (err_cui | err_tipo | err_monto | err_avance | ///
	err_f9 | err_comp | err_unid | err_cod_mod | err_demol | ///
	err_nueva | err_reforz | err_cerco | err_sust | ///
	err_ampl | err_mobil | err_agua | err_elec)

* --- BASE LIMPIA ---
preserve
	keep if tiene_error == 0
	drop err_* tiene_error
	capture drop monto_orig avance_orig cui_primero_usado
	count
	di "Registros limpios: `r(N)'"
	save "${rep_cons_o}\Anexo1_base_limpia.dta", replace
	export excel using "${rep_cons_o}\Anexo1_base_limpia.xlsx", ///
		firstrow(variables) replace
restore

* --- BASE ERRORES ---
preserve
	keep if tiene_error == 1
	count
	di "Registros con errores: `r(N)'"
	save "${rep_cons_o}\Anexo1_base_errores.dta", replace
	export excel using "${rep_cons_o}\Anexo1_base_errores.xlsx", ///
		firstrow(variables) replace
restore

* --- RESUMEN DE CRUCES ---
di ""
di "=== RESUMEN DE CRUCES ==="
qui count if cui_en_banco == "SI"
di "  CUI encontrados en Banco de Inversiones: `r(N)'"
qui count if cui_en_vinc == "SI"
di "  CUI encontrados en Vinculaciones: `r(N)'"
qui count if cod_local_match == "COMPLETADO"
di "  cod_local completados desde Vinculaciones: `r(N)'"
qui count if cod_mod_match == "COMPLETADO"
di "  cod_mod completados desde Vinculaciones: `r(N)'"
qui count if cod_local_match == "MISMATCH"
di "  cod_local con MISMATCH: `r(N)'"
qui count if cod_mod_match == "MISMATCH"
di "  cod_mod con MISMATCH: `r(N)'"

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
