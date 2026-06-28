# Viabilidad del modelo bAV para el sistema español de pensiones — material de reproducibilidad

Este repositorio acompaña al trabajo *"Viabilidad del modelo bAV (Entgeltumwandlung) para el sistema español de pensiones: una microsimulación de cohorte cerrada mediante cadenas de Markov sobre 20,9 millones de cotizantes (2023-2053)"*. Contiene el código fuente de la microsimulación, las matrices de movilidad observada (P_obs) efectivamente utilizadas, y los ficheros de resultados completos (CSV y JSON) sobre los que se basan todas las cifras, tablas y figuras del paper.

## Estructura

```
codigo/
  microsim_bav_integrado.py   Pipeline principal: población dinámica por tramo y sexo,
                               pensión pública y complementaria, fondo privado, coste
                               fiscal, índice de Gini, análisis Estado-vs-pensionistas.
  extension_escenarios.py     Escenarios institucionales de adopción (voluntario actual,
                               auto-enrolment 85%, obligatorio 100%) y estrés financiero
                               sobre la rentabilidad real del fondo.

datos_entrada/
  03_matrices_transicion/      Matrices_Transicion_Hombres.xlsx, Matrices_Transicion_Mujeres.xlsx
                               (hoja P_obs): matrices de movilidad observada entre tramos
                               de renta, tal como las utiliza el pipeline. Ver Anexo 5 del
                               paper para los valores tabulados y la nota metodológica.

resultados/
  resultados_integrados_junio26.json   Resultados agregados completos (coste fiscal,
                                        fondo privado, Gini, análisis Estado-vs-pensionistas,
                                        sensibilidad de la rentabilidad, escenarios de
                                        compensación a la Seguridad Social).
  pipeline_detalle_tramo_año.csv       Desglose por sexo, año y tramo (N, adoptantes,
                                        IRPF perdido, cotizaciones perdidas, coste fiscal,
                                        aportación al fondo, saldo del fondo).
  resumen_anual.csv                    Serie temporal anual 2023-2053 del coste fiscal y
                                        la acumulación del fondo privado.
  resumen_individual_por_tramo_2023.csv / _2053.csv
                                        Percentiles P10/P50/P90 de la pensión total
                                        (pública y complementaria) por tramo y sexo, en
                                        los dos cortes transversales del paper (Tabla 1
                                        y Tabla 1bis).
  extension_escenarios.json            Resultados de los escenarios institucionales y de
                                        estrés financiero (Tabla 3 del paper).
```

## Qué no incluye este repositorio, y por qué

Este repositorio **no** redistribuye las tablas estadísticas oficiales de terceros sobre las que se calibra el pipeline (estadística del IRPF de la AEAT, tablas de mortalidad proyectada y proyecciones de población del INE, ni los ficheros de población dinámica por tramo derivados de ellas). No son de elaboración propia, su volumen no es trivial y su distribución exige tratar de forma explícita cada fuente. Las fuentes exactas, con enlace, están citadas en la sección de Referencias del paper; en particular:

- AEAT, Estadística del IRPF 2023 (Tablas 505 y 587).
- INE, tablas de mortalidad proyectada (tabla 36774) y Proyecciones de Población 2024-2074.
- BOE, Orden PCM/74/2023 (cotización 2023) y Ley 21/2021 (Mecanismo de Equidad Intergeneracional).

Como consecuencia, **los scripts de `codigo/` no pueden ejecutarse de extremo a extremo tal como están** sin que el usuario obtenga esas fuentes y las coloque en `datos_entrada/01_datos_base/` y `datos_entrada/02_poblacion_dinamica/` con el nombre de archivo y la estructura de hoja que el código espera (ver los comentarios al inicio de cada script). Lo que este repositorio permite verificar de forma directa, sin necesidad de obtener esas fuentes, es: (a) la lógica completa del pipeline —cómo se combinan los datos para producir cada cifra del paper—, y (b) que los ficheros de `resultados/` son exactamente los que sustentan las tablas y figuras publicadas, junto con (c) los valores reales de las matrices de movilidad P_obs, ya incluidas en `datos_entrada/03_matrices_transicion/`.

## Reproducción

Dependencias: Python 3.10+, `numpy`, `pandas`, `openpyxl`.

```bash
pip install numpy pandas openpyxl
```

Una vez completados `datos_entrada/01_datos_base/` y `datos_entrada/02_poblacion_dinamica/` con las fuentes oficiales citadas arriba (mismo nombre de archivo y estructura de hoja que usa el código):

```bash
cd codigo
python microsim_bav_integrado.py      # genera los ficheros de resultados/
python extension_escenarios.py        # genera resultados/extension_escenarios.json
```

## Limitaciones metodológicas declaradas

Las limitaciones del modelo (identificación de P_obs sin datos de panel, aproximación simplificada de la pensión pública, ausencia de estocástica de mercado, alcance restringido a España, cohorte cerrada sin reposición, entre otras) se declaran de forma explícita en el código (docstrings) y en las Conclusiones y Anexos del paper. Este repositorio no resuelve ninguna de ellas; su propósito es la transparencia y la verificación de lo ya calculado, no una versión corregida del modelo.
