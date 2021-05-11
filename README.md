# Dia_55_10mayo

```{r Protocolo, eval=FALSE, include=TRUE}
"Protocolo:
 
 1. Daniel Felipe Villa Rengifo
 
 2. Lenguaje: R
 
 3. Tema: Manejo de archivos *.xlsx
 (realice al menos dos ejercicios que
 requieran cargar archivos externos,
 leer y procesar la información del archvo leído,
 y guardar las respuestas a los ejercicios
 en archivos independientes)
 
 4. Fuentes:
 http://rpubs.com/BrendaAguilar/manual_uso_de_Paquete_xlsx_en_R" 
```

# Comentario

Profe para trabajar con archivos xlsx (y sus derivaciones= xls, excel,etc...) se hace de dos maneras:

## Leer XLSX sin JAVA en R

### `readxl`

El paquete `readxl` forma parte del paquete `tidyverse`, creado por Hadley Wickham (científico jefe en RStudio) y su equipo. Este paquete soporta XLS via la librería de C `libxls` y archivos XLSX via el paquete `RapidXML` de C++, sin la necesidad de utilizar dependencias externas.

El paquete proporciona algunos archivos Excel (XLS y XLSX) de muestra almacenados en la carpeta de instalación del paquete, por lo que con el objetivo de ofrecer un ejemplo reproducible


```{r Ejemplo1 readxl, eval=FALSE, include=TRUE}
# install.packages("readxl")
library(readxl)

# Obtener la ruta de un archivo XLSX de ejemplo del paquete
ruta_archivo <- readxl_example("Ruta del Archivo")
```

La función genérica del paquete para leer archivos de Excel en R es la función `read_excel()`, que ‘adivina’ el tipo de archivo (XLS o XLSX) según la extensión del archivo y el archivo en sí, en ese orden.

```{r Ejemplo2 readxl, eval=FALSE, include=TRUE}
read_excel(ruta_archivo)
```


### `opexlsx`

El paquete `openxlsx` usa `Rcpp` y, como no depende de JAVA, es una alternativa interesante al paquete `readxl` para leer un archivos Excel en R. Las diferencias respecto al paquete anterior son que la salida es de clase `data.frame` por defecto en lugar de `tibble` y que su uso principal no es solo la importación de archivos de Excel, ya que también proporciona una amplia variedad de funciones para escribir, diseñar y editar archivos de Excel.

```{r Ejemplo3 openxlsx, eval=FALSE, include=TRUE}
# install.packages("openxlsx")
library(openxlsx)

read.xlsx(ruta_archivo)
```


## Leer con Java: 

### `xlsx`

Aunque este paquete requiere que JAVA esté instalado en tu ordenador, es muy popular. Las funciones principales para importar archivos de Excel son `read.xlsx` y `read.xlsx2`. El segundo tiene ligeras diferencias en los argumentos predeterminados y hace más trabajo en JAVA, logrando un mejor rendimiento.

```{r Ejemplo xlsx, eval=FALSE, include=TRUE}
# install.packages("xlsx")
library(xlsx)

read.xlsx(ruta_archivo)
read.xlsx2(ruta_archivo)
```


### `XLConnect`

Una alternativa al paquete `xlsx` es `XLConnect`, que permite escribir, leer y dar formato a archivos de Excel. Para leer un Excel en R puedes usar la función `readWorksheetFromFile` como sigue.


```{r Ejemplo XLConnect, eval=FALSE, include=TRUE}
# install.packages("XLConnect")
library(XLConnect)

data <- readWorksheetFromFile(ruta_archivo, sheet = "list-column",
 startRow = 1, endRow = 5,
 startCol = 1, endCol = 2)
```

## Maneras alternativa

### Opción 1 (sin paquetes):

Convertir de antemano el archivo xlsx a csv

### Opción 2:

convertir tus archivos de Excel a formato CSV y leer el archivo CSV en R. Para este propósito usar la función `convert` del paquete `rio`.

```{r Ejemplo rio, eval=FALSE, include=TRUE}
# install.packages("rio")
library(rio)

convert(ruta_archivo, "file.csv")
```

# Ejercicio 1:

Vamos a utilizar la libreria `readxlsx`

parametros de la base de datos `base100`:

+ `Region`: Continente

+ `item type`: Productos que envia:
 
 1. Office Supplies
 2. Vegetables
 3. Fruits
 4. Cosmetics
 5. Cereal
 6. Baby Food
 7. Beverages
 8. Snacks
 9. Clothes
 10. Household
 11. Personal Care
 12. Meat

+ `Sales Channel`: Canal de la venta:
 
 1. Online
 2. Offline
 
+ `Order Priority`:Prioridad del envio
 
 1. `C`: Critico
 2. `H`: Alto
 3. `M`: Medio
 4. `L`: Bajo

+ `Unit Price`: Precio unitario (Precio de compra _[Unidades X Unit Cost]_)

+ `Unit Cost`: Costo Unitario (Cuanto cuesta la Unidad)

+ `Unit Sold`: unidad vendida (Precio de venta)

+ `Total Revenue`: Ingresos totales

+ `Total Cost`: Costos totales

+ `Total Profit`: Beneficio Total

```{r}
#Leemos una base de datos:
library(readxl)
base100 <- read_excel("100 Sales Records.xlsx")
```

El ejercicio consiste en:

Cuanto es el beneficio unitario, es decir, cuanto se gana por unidad.

```{r}
# Convirtamos a data.frame la base de datos (seleccionando solamente los datos que necesitamos)
library(dplyr)

base100$"Item Type" <- as.factor(base100$"Item Type")

## Nos guiaremos con el ID de la orden y el item Type (de una vez lo volvemos dataframe asi ir trabajando los item Type como factores):

cal_ben <- data.frame(select(base100, "Item Type","Order ID", "Units Sold", "Unit Price"))

## Lo que haremos es [Precio de venta - Precio de compra = Benefecio]
## Añadiendo una nueva columna:

## Creamos una nuevva columna que contenga [Precio de venta X (1 - 0.19) - Precio de compra = Benefecio]
## (descontando el iva 19%)

cal_ben$beneficio <- ((cal_ben$Units.Sold * (1-0.19)) - cal_ben$Unit.Price)

#observemos que hay valores negativos, es decir, hay una perdida:

perdida <- filter(cal_ben, beneficio < 0)

names(perdida)

# Eso quiere decir que la compañia de envios perdio dinero por producto unitario:
cat("\nNumero de productos donde se genera una perdida = ", length(perdida$beneficio))
cat("\n \nProdcutos que generaron perdida")
print(perdida)
```

```{r}
# Exportemos los archivos:

## 1° donde se registra el beneficio unitario por venta:

write.xlsx(cal_ben, file = "Beneficio_Unit.xlsx")

## 2° donde se registran las perdidas:

write.xlsx(perdida, file = "Perdidas_Unit.xlsx")
```


# 2° Ejercicio

Cual de los productos que envia tendra más benefecio total:

```{r}
## Saquemos la columnas que necesitamos:

baseItem <- aggregate(base100$"Total Profit" ~ base100$"Item Type", data = base100, FUN = sum)

baseItem <- baseItem %>%
  rename(
    "Beneficio_Total" = "base100$\"Total Profit\"",
    "Producto_Envio" = "base100$\"Item Type\""
    )

# Convertimos a factor una columna:

baseItem$"Producto_Envio" <- as.factor(baseItem$"Producto_Envio")
  
# Observemos varias cosas
"que los cosmeticos son el producto con mayor ingrse para la compañia e envios"

## ordenar los datos:
baseItem <- arrange(baseItem, Beneficio_Total,Producto_Envio)

cat("\nAhora el 1° termino = MENOR GANANCIA, 12° termino = MAYOR GANANCIA\n")
print(baseItem)

cat("\nen la empresa las frutas no son tan rentables como los cosmeticos\n")

# Exportamos nuestro resultado:
write.xlsx(
  baseItem,
  file = "BeneficioXProdcutoEnviado.xlsx"
)


```





