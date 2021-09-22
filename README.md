# Power BI + SharePoint Search

Con Power BI por default podemos consumir información de múltiples fuentes de datos , ¿Pero que pasa cuando el conector no existe o no regresa la información como la necesitamos?
Pues siempre que los datos se expongan mediate un API podemos hacer algo.
Este es el caso de SharePoint Search, podemos utilizar el API de SharePoint Search  para hacer un reporte con los resultados de su búsqueda.


¿Y por que utilizar el API en lugar del conector a las listas que ya tiene Power BI?

Imaginemos el siguiente escenario:

*Nos piden hacer un reporte de una listas de SharePoint que se repiten en varios subsitios , las listas tienen la misma estructura en todos los subsitios ,el usuario pueden crear mas subsitios (con un template) a diestra y siniestra.*

* Si utilizamos el conector a la lista, tendríamos que regresar a modificar el reporte cuando se agreguen mas sitios.
* Si utilizamos SharePoint Search podemos regresar los resultados de los N sitios.

![image](https://user-images.githubusercontent.com/50918464/134263821-d071b7ff-24db-4225-a3d2-4e9e4f83e7dd.png)

Entonces bajo la premisa de que no quisiera editar el reportoe cuando un nuevo subsitio sea creado,me inclino sobrela opción de consumir SharePoint Search.

# Datos de prueba

Voy a crear la lista de SharePoint con el nombre "ListaConMuchosItems" en nivel raíz , en algunos subsitios y capturar algunos datos.
![image](https://user-images.githubusercontent.com/50918464/134264250-35d47487-0c09-4e4c-9911-fe19589765ec.png)

Después de revisar la documentación oficial  de como hacer consultas al SharePoint Search, armamos la URL con el Query que nos ayudara a obtener los datos de las listas de todos los sitios existentes.

https://**[Tenant]**.sharepoint.com/_api/search/query?querytext='ParentLink:**[ListName]** AND SPSiteURL:https://**[Tenant]**.sharepoint.com’&rowlimit=500

Para mi ambiente seria la URL: https://midominniodev.sharepoint.com/_api/search/query?querytext='ParentLink:ListaConMuchosItems AND SPSiteURL:https://midominniodev.sharepoint.com'&rowlimit=500

Si probamos la URL podemos ver los valores de los diferentes subsitios
![image](https://user-images.githubusercontent.com/50918464/134264613-adf4c5e3-ab35-4432-b033-95e2d815358f.png)

Listo, ahora aquí tenemos un problema con la cantidad de resultados máximos que retorna el API y es que como máximo retorna 500 registros.

Para obtener mas registros debemos de agregar un query string  indicando en que fila comenzar  "…&startrow=501", tendremos que hacer tantas llamadas sean necesarias para poder regresar todos los registros.

¿Y como hacemos todo eso en Power BI? Pues con Query M podemos implementar la lógica necesaria para :

* Consumir los datos del API
* Enviar un Query String en la llamada la api
* Hacer esa llamada de forma recursiva hasta consumir la totalidad de los datos

Lo primero es agregar un query en blanco
 ![image](https://user-images.githubusercontent.com/50918464/134264763-9467c0a6-0dea-4630-89de-328ab3b695b4.png)
Agregar el siguiente Query M que nos ayudara a consumir el API  de forma recursiva hasta consumir todos los datos
![image](https://user-images.githubusercontent.com/50918464/134264840-237c2f4b-9e8f-4a49-9391-0f4f7d97ed50.png)
![image](https://user-images.githubusercontent.com/50918464/134264903-1541343d-a9e8-41f5-bb17-611170a83012.png)

Copia, Pega, Edita y Prueba !!!

```shell
let
    itemsByPage=500,  
    my_func = (startIndex) =>
    let 
        Source = OData.Feed("https://midominniodev.sharepoint.com/_api/search/query?querytext='ParentLink:ListaConMuchosItems AND SPSiteURL:https://midominniodev.sharepoint.com'&rowlimit=500", null, 
        [Implementation="2.0",Query=[#"startrow"=Number.ToText(startIndex)]]),
        PrimaryQueryResult = Source[PrimaryQueryResult],
        Rows= Source[PrimaryQueryResult][RelevantResults][Table][Rows],
        AllRows = List.Transform(Rows, each _[Cells]),
        RowsToTables = List.Transform(AllRows, each List.Transform(_, each Record.ToTable(_))),    
        SkelToList = List.Transform(RowsToTables, each Table.FromList(_, Splitter.SplitByNothing(), null, null, ExtraValues.Error)),
        CleanRows = List.Transform(SkelToList, each List.Transform(_[Column1], each Table.PromoteHeaders(Table.RemoveLastN( Table.RemoveColumns( _,{"Name"}), 1) ) ) ),
        TransposeTable = Table.FromRows(List.Transform(CleanRows, each List.Transform(_, each Record.FieldValues(_{0}){0} ))),
        ColumnRenames = List.Transform(CleanRows{0}, each { "Column" & Text.From( List.PositionOf(CleanRows{0}, _) + 1), Table.ColumnNames(_){0}}),
        RenamedTable = Table.RenameColumns(TransposeTable, ColumnRenames),
        totalItems=Table.RowCount(RenamedTable),            
        resultado= if totalItems = itemsByPage then  RenamedTable & @my_func(startIndex+500) else RenamedTable
    in
        resultado
in
    my_func(0)

```
