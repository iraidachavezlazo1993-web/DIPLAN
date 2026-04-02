* ===========================================================================
* LIMPIEZA DE ANEXO 1 - Inversiones en infraestructura educativa
* Declaraciones de GR y GL
* ===========================================================================

clear all
set more off

global reporte "C:\Users\iraid\Documents\DISCO TERA\MINEDU\TRABAJO-MINEDU\01_MINEDU"
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
* 3. LIMPIAR CUI
* -------------------------------------------------------------------------
* el CUI deberia ser solo numeros, pero ponen de todo

replace cui = strtrim(cui)
replace cui = subinstr(cui, " ", "", .)

* si no es numerico puro, intento rescatar la secuencia de digitos
replace err_cui = 1 if !regexm(cui, "^[0-9]+$") & cui != ""
* intento sacar algo util: secuencia de 5+ digitos
replace cui = regexs(0) if regexm(cui, "[0-9]{5,}") & err_cui == 1
* si logre rescatar, quito el error
replace err_cui = 0 if regexm(cui, "^[0-9]+$") & err_cui == 1

* -------------------------------------------------------------------------
* 4. LIMPIAR TIPO DE INVERSION
* -------------------------------------------------------------------------
* esto pasa porque los GR ponen cualquier cosa

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
* 5. LIMPIAR MONTO
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
* 6. LIMPIAR AVANCE
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
* 7. LIMPIAR F9 (FORMATO 9)
* -------------------------------------------------------------------------
* el f9 deberia ser SI o NO pero a veces meten datos de componente aca

replace f9 = upper(strtrim(f9))
replace f9 = subinstr(f9, "Í", "I", .)

* detectar valores que son de componente (se corrieron las columnas)
replace err_f9 = 1 if regexm(f9, "INFRAESTRUCTURA")
replace err_f9 = 1 if regexm(f9, "EQUIPAMIENTO")
replace err_f9 = 1 if regexm(f9, "MOBILIARIO")
replace err_f9 = 1 if regexm(f9, "INTEGRAL")

* normalizar SI/NO
replace f9 = "SI" if regexm(f9, "^SI$") | f9 == "SÍ" | f9 == "S" | f9 == "1"
replace f9 = "NO" if f9 == "NO" | f9 == "N" | f9 == "0"

* lo que no es SI, NO ni vacio es error
replace err_f9 = 1 if !inlist(f9, "SI", "NO", "") & err_f9 == 0

* -------------------------------------------------------------------------
* 8. LIMPIAR COMPONENTE
* -------------------------------------------------------------------------
* mapear a 4 categorias: infraestructura, equipamiento, mobiliario, integral

replace comp = upper(strtrim(comp))
replace comp = subinstr(comp, "  ", " ", .)

* primero detecto combinaciones => integral
replace comp = "4.INTEGRAL" if regexm(comp, "1,2") | regexm(comp, "1;2")
replace comp = "4.INTEGRAL" if regexm(comp, "1 Y 2") | regexm(comp, "2 Y 1")
replace comp = "4.INTEGRAL" if regexm(comp, "1,2,3") | regexm(comp, "1;2;3")
replace comp = "4.INTEGRAL" if regexm(comp, "INTEGRAL")

* las individuales
replace comp = "1.INFRAESTRUCTURA" if regexm(comp, "INFRAESTRUCTURA") & comp != "4.INTEGRAL"
replace comp = "2.EQUIPAMIENTO" if regexm(comp, "EQUIPAMIENTO") & !regexm(comp, "^[1-4]\.")
replace comp = "3.MOBILIARIO" if regexm(comp, "MOBILIARIO") & !regexm(comp, "^[1-4]\.")

* numeros solos
replace comp = "1.INFRAESTRUCTURA" if comp == "1"
replace comp = "2.EQUIPAMIENTO" if comp == "2"
replace comp = "3.MOBILIARIO" if comp == "3"
replace comp = "4.INTEGRAL" if comp == "4"

* lo que no matcheo
replace err_comp = 1 if !inlist(comp, "1.INFRAESTRUCTURA", "2.EQUIPAMIENTO", "3.MOBILIARIO", "4.INTEGRAL") & comp != ""

* -------------------------------------------------------------------------
* 9. LIMPIAR CAMPOS SI/NO
* -------------------------------------------------------------------------
* primero muevo texto largo de electricidad a comentarios
* a veces ponen comentarios en la columna de electricidad

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
* 10. LIMPIAR CODIGO MODULAR
* -------------------------------------------------------------------------

replace cod_mod = strtrim(cod_mod)

* guiones y underscores no son error, solo es que no llenaron
replace cod_mod = "" if inlist(cod_mod, "-", "_", "--", "---")
* estos tampoco son error, pero no es un codigo modular
replace cod_mod = "" if inlist(upper(cod_mod), "SI", "SÍ", "NO", "NINGUNA", "0")

* separadores: /, -, saltos de linea => coma
replace cod_mod = subinstr(cod_mod, "/", ",", .)
replace cod_mod = subinstr(cod_mod, " - ", ",", .)
replace cod_mod = subinstr(cod_mod, char(10), ",", .)
replace cod_mod = subinstr(cod_mod, char(13), ",", .)
replace cod_mod = subinstr(cod_mod, " ", "", .)

* limpiar comas dobles
replace cod_mod = subinstr(cod_mod, ",,", ",", .)
replace cod_mod = regexr(cod_mod, "^,", "")
replace cod_mod = regexr(cod_mod, ",$", "")

* si queda algo que no es digitos y comas, es error
replace err_cod_mod = 1 if !regexm(cod_mod, "^[0-9,]*$") & cod_mod != ""

* -------------------------------------------------------------------------
* 11. DUPLICADOS POR COD_LOCAL + CUI
* -------------------------------------------------------------------------

bysort cod_local cui: gen dup_n = _n
bysort cod_local cui: gen dup_N = _N
gen byte es_duplicado = (dup_N > 1)
label var es_duplicado "Registro duplicado (cod_local+cui)"
drop dup_n dup_N

* -------------------------------------------------------------------------
* 12. RENUMERAR
* -------------------------------------------------------------------------

drop nro
gen nro = _n
order nro
label var nro "Nro"

save "${rep_cons_t}\anexo1_limpio_pre.dta", replace

* -------------------------------------------------------------------------
* 13. VALIDAR CONTRA VINCULACIONES
* -------------------------------------------------------------------------
* busco el archivo de vinculaciones, tiene varios nombres posibles

local found_vinc = 0
foreach fname in "vinculaciones.xlsx" "Vinculaciones.xlsx" "VINCULACIONES.xlsx" ///
	"vinculaciones_cui.xlsx" "Vinculaciones_CUI.xlsx" "base_vinculaciones.xlsx" {
	capture confirm file "${rep_cons_i}\\`fname'"
	if _rc == 0 {
		local vinc_file = "`fname'"
		local found_vinc = 1
		continue, break
	}
}

if `found_vinc' == 1 {
	di as text ">>> Encontre archivo de vinculaciones: `vinc_file'"

	preserve
	import excel using "${rep_cons_i}\\`vinc_file'", firstrow clear

	* renombro las columnas que me interesan
	capture rename CUI cui_vinc
	capture rename CodigoLocal cod_local_vinc
	capture rename NombreIE nombre_ie_vinc
	* por si vienen con otros nombres
	capture rename CODIGOUNICO cui_vinc
	capture rename CODIGOLOCAL cod_local_vinc
	capture rename NOMBREIE nombre_ie_vinc

	* me quedo con lo que necesito
	keep cui_vinc cod_local_vinc nombre_ie_vinc
	rename cui_vinc cui

	* puede haber duplicados en la base de vinculaciones
	duplicates drop cui, force

	tempfile vinculaciones
	save `vinculaciones', replace
	restore

	* merge
	merge m:1 cui using `vinculaciones', keep(master match) gen(_merge_vinc)

	* para los que matchearon, comparo
	gen byte flag_nombre_ie = 0
	gen byte flag_cod_local = 0

	replace flag_nombre_ie = 1 if _merge_vinc == 3 & upper(strtrim(nombre_ie)) != upper(strtrim(nombre_ie_vinc)) & nombre_ie_vinc != ""
	replace flag_cod_local = 1 if _merge_vinc == 3 & strtrim(cod_local) != strtrim(cod_local_vinc) & cod_local_vinc != ""

	label var flag_nombre_ie "Nombre IE difiere de vinculaciones"
	label var flag_cod_local "Cod local difiere de vinculaciones"

	* limpio variables del merge
	capture drop nombre_ie_vinc cod_local_vinc _merge_vinc

	di as result "   Flags de vinculaciones generados"
}
else {
	di as error ">>> No encontré archivo de vinculaciones, me lo salto"
	gen byte flag_nombre_ie = 0
	gen byte flag_cod_local = 0
}

* -------------------------------------------------------------------------
* 14. VALIDAR CONTRA BASE DE INVERSIONES
* -------------------------------------------------------------------------
* la base de inversiones tambien cambia de nombre segun el dia

local found_inv = 0
foreach fname in "base_inversiones.xlsx" "Base_Inversiones.xlsx" "BASE_INVERSIONES.xlsx" ///
	"inversiones.xlsx" "Inversiones.xlsx" "base_inv.xlsx" {
	capture confirm file "${rep_cons_i}\\`fname'"
	if _rc == 0 {
		local inv_file = "`fname'"
		local found_inv = 1
		continue, break
	}
}

if `found_inv' == 1 {
	di as text ">>> Encontre base de inversiones: `inv_file'"

	preserve
	import excel using "${rep_cons_i}\\`inv_file'", firstrow clear

	* auto-detectar nombres de columna
	capture confirm variable CUI
	if _rc != 0 {
		capture confirm variable cui
		if _rc != 0 {
			capture confirm variable CODIGOUNICO
			if _rc == 0 rename CODIGOUNICO CUI
		}
		else {
			rename cui CUI
		}
	}

	capture confirm variable TIPO
	if _rc != 0 {
		capture confirm variable tipo_inversion
		if _rc == 0 rename tipo_inversion TIPO
		capture confirm variable TipoInversion
		if _rc == 0 rename TipoInversion TIPO
	}

	capture confirm variable F9
	if _rc != 0 {
		capture confirm variable formato9
		if _rc == 0 rename formato9 F9
		capture confirm variable Formato9
		if _rc == 0 rename Formato9 F9
	}

	rename CUI cui
	capture rename TIPO tipo_inv
	capture rename F9 f9_inv

	keep cui tipo_inv f9_inv
	capture tostring cui, replace force
	duplicates drop cui, force

	tempfile inversiones
	save `inversiones', replace
	restore

	merge m:1 cui using `inversiones', keep(master match) gen(_merge_inv)

	* comparo tipo y f9 con la base maestra
	gen byte flag_tipo = 0
	gen byte flag_f9 = 0

	capture confirm variable tipo_inv
	if _rc == 0 {
		replace flag_tipo = 1 if _merge_inv == 3 & upper(strtrim(tipo)) != upper(strtrim(tipo_inv)) & tipo_inv != "" & tipo != ""
		label var flag_tipo "Tipo difiere de base inversiones"
	}

	capture confirm variable f9_inv
	if _rc == 0 {
		replace flag_f9 = 1 if _merge_inv == 3 & upper(strtrim(f9)) != upper(strtrim(f9_inv)) & f9_inv != "" & f9 != ""
		label var flag_f9 "F9 difiere de base inversiones"
	}

	capture drop tipo_inv f9_inv _merge_inv

	di as result "   Flags de inversiones generados"
}
else {
	di as error ">>> No encontré base de inversiones, me lo salto"
	gen byte flag_tipo = 0
	gen byte flag_f9 = 0
}

save "${rep_cons_t}\anexo1_limpio_validado.dta", replace

* -------------------------------------------------------------------------
* 15. SEPARAR BASES
* -------------------------------------------------------------------------

* genero indicador de si tiene algun error
egen tiene_error = rowmax(err_*)
label var tiene_error "Tiene al menos un error"

* campos vacios: verifico que los campos clave no esten vacios
gen byte campos_vacios = 0
replace campos_vacios = 1 if cui == "" | tipo == "" | comp == ""
replace campos_vacios = 1 if monto == . & monto_orig == ""
replace campos_vacios = 1 if region == ""
label var campos_vacios "Le faltan campos obligatorios"

* clasifico
gen str20 estado = ""
replace estado = "LIMPIO" if tiene_error == 0 & campos_vacios == 0
replace estado = "CON ERRORES" if tiene_error == 1 & campos_vacios == 0
replace estado = "EXCLUIDO" if campos_vacios == 1
label var estado "Estado del registro"

save "${rep_cons_t}\anexo1_clasificado.dta", replace

* -------------------------------------------------------------------------
* 16. EXPORTAR
* -------------------------------------------------------------------------

* base limpia
preserve
keep if estado == "LIMPIO"
save "${rep_cons_o}\anexo1_limpia_stata.dta", replace
export excel using "${rep_cons_o}\anexo1_limpia_stata.xlsx", firstrow(varlabels) replace
local n_limpia = _N
restore

* base con errores
preserve
keep if estado == "CON ERRORES"
save "${rep_cons_o}\anexo1_errores_stata.dta", replace
export excel using "${rep_cons_o}\anexo1_errores_stata.xlsx", firstrow(varlabels) replace
local n_errores = _N
restore

* base excluidos
preserve
keep if estado == "EXCLUIDO"
save "${rep_cons_o}\anexo1_excluidos_stata.dta", replace
export excel using "${rep_cons_o}\anexo1_excluidos_stata.xlsx", firstrow(varlabels) replace
local n_excluidos = _N
restore

* resumen por region
preserve
collapse (count) nro (sum) tiene_error campos_vacios, by(region)
rename nro total_registros
rename tiene_error total_errores
rename campos_vacios total_excluidos
gen total_limpios = total_registros - total_errores - total_excluidos
export excel using "${rep_cons_o}\resumen_por_region_stata.xlsx", firstrow(variables) replace
restore

* -------------------------------------------------------------------------
* 17. RESUMEN FINAL
* -------------------------------------------------------------------------

di as text ""
di as text "=========================================="
di as result " RESUMEN DE LIMPIEZA - ANEXO 1"
di as text "=========================================="
di as text ""
di as text "Total registros:   " _N
di as result "  Limpios:         `n_limpia'"
di as result "  Con errores:     `n_errores'"
di as result "  Excluidos:       `n_excluidos'"
di as text ""

* errores por campo
di as text "------------------------------------------"
di as text " ERRORES POR CAMPO"
di as text "------------------------------------------"

foreach v in cui tipo monto avance f9 comp unid cod_mod demol nueva reforz cerco sust ampl mobil agua elec {
	quietly count if err_`v' == 1
	local nerr = r(N)
	if `nerr' > 0 {
		di as error "  err_`v':  `nerr'"
	}
	else {
		di as text "  err_`v':  0"
	}
}

di as text ""
di as text "------------------------------------------"
di as text " DUPLICADOS"
di as text "------------------------------------------"
quietly count if es_duplicado == 1
di as text "  Registros duplicados (cod_local+cui): " r(N)

di as text ""
di as text "------------------------------------------"
di as text " REGISTROS POR ESTADO"
di as text "------------------------------------------"
tab estado

di as text ""
di as result ">>> Archivos exportados en: ${rep_cons_o}"
di as text "=========================================="

* fin
