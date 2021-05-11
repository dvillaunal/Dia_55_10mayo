## ----Protocolo, eval=FALSE, include=TRUE----------------------------------------------------------------
## "Protocolo:
## 
##  1. Daniel Felipe Villa Rengifo
## 
##  2. Lenguaje: R
## 
##  3. Tema: Manejo de archivos *.xlsx
##  (realice al menos dos ejercicios que
##  requieran cargar archivos externos,
##  leer y procesar la información del archvo leído,
##  y guardar las respuestas a los ejercicios
##  en archivos independientes)
## 
##  4. Fuentes:
##  http://rpubs.com/BrendaAguilar/manual_uso_de_Paquete_xlsx_en_R"


## ----Ejemplo1 readxl, eval=FALSE, include=TRUE----------------------------------------------------------
## # install.packages("readxl")
## library(readxl)
## 
## # Obtener la ruta de un archivo XLSX de ejemplo del paquete
## ruta_archivo <- readxl_example("Ruta del Archivo")


## ----Ejemplo2 readxl, eval=FALSE, include=TRUE----------------------------------------------------------
## read_excel(ruta_archivo)


## ----Ejemplo3 openxlsx, eval=FALSE, include=TRUE--------------------------------------------------------
## # install.packages("openxlsx")
## library(openxlsx)
## 
## read.xlsx(ruta_archivo)


## ----Ejemplo xlsx, eval=FALSE, include=TRUE-------------------------------------------------------------
## # install.packages("xlsx")
## library(xlsx)
## 
## read.xlsx(ruta_archivo)
## read.xlsx2(ruta_archivo)


## ----Ejemplo XLConnect, eval=FALSE, include=TRUE--------------------------------------------------------
## # install.packages("XLConnect")
## library(XLConnect)
## 
## data <- readWorksheetFromFile(ruta_archivo, sheet = "list-column",
##  startRow = 1, endRow = 5,
##  startCol = 1, endCol = 2)


## ----Ejemplo rio, eval=FALSE, include=TRUE--------------------------------------------------------------
## # install.packages("rio")
## library(rio)
## 
## convert(ruta_archivo, "file.csv")


## -------------------------------------------------------------------------------------------------------
#Leemos una base de datos:
library(readxl)
base100 <- read_excel("100 Sales Records.xlsx")


## -------------------------------------------------------------------------------------------------------
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


## -------------------------------------------------------------------------------------------------------
# Exportemos los archivos:

## 1° donde se registra el beneficio unitario por venta:

write.xlsx(cal_ben, file = "Beneficio_Unit.xlsx")

## 2° donde se registran las perdidas:

write.xlsx(perdida, file = "Perdidas_Unit.xlsx")


## -------------------------------------------------------------------------------------------------------
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



