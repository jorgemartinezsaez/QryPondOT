Attribute VB_Name = "QryPondOT"
' ============================================= '
' Módulo: QryPondOT (Query de Ponderadores OT)
' ============================================= '
'
' Descripción:
' Este módulo contiene una colección de funciones en VBA para la consulta, transformación
' y codificación de datos relacionados con los ponderadores de rendimiento (OT).
'
' Funciones principales:
' - UrlEncodeUTF8: Codifica una cadena en formato URL UTF-8.
' - ConvertToUTF8: Convierte una cadena de texto a su representación UTF-8.
' - wsGetRendPoFromAPI: Realiza una consulta a un servicio web para obtener un valor de rendimiento.
' - CONSULTA_REND_PO: Consulta simplificada para obtener el ponderador operativo de rendimiento .
'
' Autor: Jorge Martínez S.
' Propiedad intelectual: Tracksys (https://tracksys.cl)
'
'
' Licencia:
' Este módulo de código se distribuye bajo la Licencia Pública General de GNU v3.0.
' Puede redistribuirlo y/o modificarlo bajo los términos de la GPL tal como se publica
' por la Free Software Foundation. Este código se proporciona "tal cual", sin ninguna
' garantía expresa o implícita.
'
' Copyright 2025 Tracksys.
' Tracksys y el dominio tracksys.cl son marcas registradas de Jorge Martínez Sáez.
' Todos los derechos reservados.
'
' Para obtener una copia de la licencia GNU GPL, visite:
' https://www.gnu.org/licenses/gpl-3.0.html
'
' ============================================= '
' Historial de versiones:
' ============================================= '
' Versión  Fecha        Descripción
' 1.0      2025-01-09   Versión inicial con funciones principales documentadas.
' 1.1      2025-01-31   Se agregan funciones para consultar ponderadores
' ============================================= '

' ***************************************************************************************

' ============================================= '
' Global Variables
' ============================================= '
''
' Almacena la clave de búsqueda utilizada en las consultas a la API.
'
' @var {String} API_ROC - Clave global que se usa para almacenar el identificador de la consulta en la API.
' @scope Global - Disponible en todos los módulos dentro del proyecto VBA.
' @description
'     Esta variable se asigna con el valor de "pond_id_wght" cuando se ejecuta la función
'     CONSULTA_REND_PO_ID. Se usa en otros procedimientos para mantener la consistencia
'     en las llamadas a la API y evitar la necesidad de recalcular el valor repetidamente.
'
' @throws No genera excepciones explícitas, pero su uso indebido puede generar resultados inesperados
'         si no se actualiza correctamente en cada ejecución.
'
' @author Jorge Martínez S.
Dim API_ROC As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Consulta el rendimiento por operación (RendPo) utilizando una clave API.
' Esta función realiza una llamada a la API para obtener el rendimiento y retorna su valor numérico.
'
' @method CONSULTA_REND_PO
' @param {String} API_KEY - Clave utilizada para la búsqueda en la API.
' @return {Variant} - Valor numérico del rendimiento ("rend_po").
' @throws Error si la función auxiliar `wsGetRendPoFromAPI` o el objeto JSON devuelto no contiene los datos esperados.
' @author Jorge Martínez S.
'

Public Function CONSULTA_REND_PO(API_KEY As String) As Variant
Attribute CONSULTA_REND_PO.VB_Description = "Esta función consulta el pamámetro ponderador desde el modelo PBI"
Attribute CONSULTA_REND_PO.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Declaración de variables
    Dim QRY_TO_API As Object
    
    ' Realiza la consulta a la API utilizando la función auxiliar `wsGetRendPoFromAPI`
    Set QRY_TO_API = wsGetRendPoFromAPI(UrlEncodeUTF8(API_KEY))
    
    ' Extrae y convierte el valor del campo "rend_po" a numérico
    CONSULTA_REND_PO = Val(QRY_TO_API(1)("rend_po"))
    

End Function

' ============================================= '
' Public Methods
' ============================================= '

''
' Consulta el id del grupo de ponderadores que aplica de acuerdo a una clave API.
' Esta función realiza una llamada a la API para obtener el id del RoC y retorna su identificador (String).
'
' @method CONSULTA_REND_PO_ID
' @param {String} API_KEY - Clave utilizada para la búsqueda en la API.
' @return {Variant} - Id del Ponderador("pond_id_wght").
' @throws Error si la función auxiliar `wsGetRendPoIdFromAPI` o el objeto JSON devuelto no contiene los datos esperados.
' @author Jorge Martínez S.

Public Function CONSULTA_REND_PO_ID(API_KEY As String) As Variant
    Dim QRY_TO_API As Object
    
    ' Realiza la consulta a la API utilizando la función auxiliar `wsGetRendPoFromAPI`
    Set QRY_TO_API = wsGetRendPoIdFromAPI(UrlEncodeUTF8(API_KEY))
    
    ' Manejo de excepciones
    On Error Resume Next
    
    ' Modifica las celdas según la respuesta de la API y genera las llamadas a los procedimientos
    If QRY_TO_API.Count = 0 Or QRY_TO_API Is Nothing Then
        Cells(ActiveCell.Row, 7).Validation.Delete
        Cells(ActiveCell.Row, 7).ClearContents
        Cells(ActiveCell.Row, 8).ClearContents
        CONSULTA_REND_PO_ID = "No Existe"
    
    Else
        CONSULTA_REND_PO_ID = (QRY_TO_API(1)("pond_id_wght"))
        API_ROC = (QRY_TO_API(1)("pond_id_wght"))
        Call ASIGNA_LISTA_VALIDACION_ROC
    End If
    
    If Err.Number <> 0 Then
        CONSULTA_REND_PO_ID = "No Existe" ' Devuelve #¡VALOR! si hay error
    End If
    
    
End Function

' ============================================= '
' Public Methods
' ============================================= '

''
' Consulta el id del grupo de ponderadores que aplica de acuerdo a una clave API.
' Esta función realiza una llamada a la API para obtener el id del RoC y retorna su identificador (String).
'
' @method CONSULTA_ROC
' @param {String} API_KEY - Clave utilizada para la búsqueda en la API.
' @return {Variant} - Lista de Ponderadores
' @throws Error si la función auxiliar `wsGetRoCFromAPI` o el objeto JSON devuelto no contiene los datos esperados.
' @author Jorge Martínez S.

Public Function CONSULTA_ROC(API_KEY As String) As Variant
    ' Declaración de variables
    Dim QRY_TO_API As Object
    Dim KEYS_VALUES As Object
    Dim DATA_ARRAY() As String
    Dim i As Integer
    Dim LISTA_VALIDACION As String
    
    ' Realiza la consulta a la API utilizando la función auxiliar `wsGetRendPoFromAPI`
    Set QRY_TO_API = wsGetRoCFromAPI(API_KEY)
    
    
    ' Asigna el valor devuelto y redimensiona el array dinámico según la cantidad de datos obtenidos
    Set KEY_VALUES = QRY_TO_API(1)
    ReDim DATA_ARRAY(1 To KEY_VALUES.Count)
    i = 0
    
    ' Almacenar los nombres en el array
    For Each Key In KEY_VALUES.Keys
        i = i + 1
        DATA_ARRAY(i) = CStr(KEY_VALUES(Key))
        
        LISTA_VALIDACION = LISTA_VALIDACION & DATA_ARRAY(i) & ";"
        
    Next Key
    
    ' Quitar la última coma
    If Len(LISTA_VALIDACION) > 0 Then
        LISTA_VALIDACION = Left(LISTA_VALIDACION, Len(LISTA_VALIDACION) - 1)
    End If
    
    ' Asigna la lista como valor de retorno de la función
    CONSULTA_ROC = LISTA_VALIDACION
End Function
' ============================================= '
' Private Methods
' ============================================= '
''
' Asigna una lista de validación a la celda de la columna G en la fila activa, basada en la consulta de ROC.
'
' @method ASIGNA_LISTA_VALIDACION_ROC
' @description Obtiene una lista de valores desde la función `CONSULTA_ROC`, utilizando la variable
'              global `API_ROC` como parámetro de búsqueda. Luego, asigna estos valores como una lista
'              de validación en la columna G de la fila activa.
'
' @dependencies
'    - API_ROC (Variable global utilizada como clave de búsqueda en `CONSULTA_ROC`).
'    - CONSULTA_ROC(API_ROC) (Función que devuelve los valores de la lista de validación).
'
' @return No retorna un valor directo. Modifica la validación de la celda en la columna G de la fila activa.
'
' @throws Error si la celda activa no es válida o si la función `CONSULTA_ROC` no devuelve un resultado correcto.
'
' @author Jorge Martínez S.
'
Sub ASIGNA_LISTA_VALIDACION_ROC()
    ' Asigna la respuesta a la variable
    VALORES = CONSULTA_ROC(API_ROC)
    
    ' Asigna la variable a la propiedad validate
    With Cells(ActiveCell.Row, 7).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:=VALORES
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub


' ============================================= '
' Public Methods
' ============================================= '

''
' Realiza una consulta a la API y devuelve un objeto JSON con los datos del rendimiento por operación (RendPo).
'
' @method wsGetRendPoFromAPI
' @param {String} KEY_SEARCH - Clave utilizada para realizar la búsqueda en la API.
' @return {Object} - Objeto JSON (Collection o Dictionary) que contiene los datos de la respuesta de la API.
' @throws Error si el servidor retorna un código de estado diferente a 200.
' @author Jorge Martínez S.
'

Public Function wsGetRendPoFromAPI(KEY_SEARCH As String) As Object
    ' Declaración de variables
    Dim http As Object
    Dim url As String
    Dim response As String
    
    ' Construye la URL con el parámetro de búsqueda
    url = "https://tracksys.cl/api/wsGetRendPo.php?api_key=" & KEY_SEARCH
    
    ' Inicializa el objeto HTTP para realizar la solicitud
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Configura y envía la solicitud HTTP GET
    http.Open "GET", url, False
    http.send
    
    ' Captura la respuesta del servidor
    response = http.responseText
    
    ' Verifica si la solicitud fue exitosa (código HTTP 200)
    If http.Status = 200 Then
        ' Analiza y retorna el objeto JSON
        Set wsGetRendPoFromAPI = JsonConverter.ParseJson(response)
    Else
        ' Termina la ejecución en caso de error de estado HTTP
        End
    End If
End Function

' ============================================= '
' Public Methods
' ============================================= '
''
' Realiza una consulta a la API y devuelve un objeto JSON con el ID del rendimiento por operación (RendPo).
'
' @method wsGetRendPoIdFromAPI
' @description Envía una solicitud HTTP GET a la API `wsGetRendPoId.php`, utilizando una clave de búsqueda (`KEY_SEARCH`),
'              y retorna la respuesta en formato JSON como un objeto VBA.
'
' @dependencies
'    - Se requiere acceso a `MSXML2.XMLHTTP` para la solicitud HTTP.
'    - Se debe incluir `JsonConverter` para la conversión de la respuesta JSON.
'
' @param {String} KEY_SEARCH - Clave utilizada para realizar la búsqueda en la API.
'
' @return {Object} - Objeto JSON (Collection o Dictionary) que contiene los datos de la respuesta de la API.
'
' @throws Error si el servidor retorna un código de estado diferente a 200 o si la respuesta no es válida.
'
' @author Jorge Martínez S.
'
Public Function wsGetRendPoIdFromAPI(KEY_SEARCH As String) As Object
    ' Declaración de variables
    Dim http As Object
    Dim url As String
    Dim response As String
    
    ' Construye la URL con el parámetro de búsqueda
    url = "https://tracksys.cl/api/wsGetRendPoId.php?api_key=" & KEY_SEARCH
    
    ' Inicializa el objeto HTTP para realizar la solicitud
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Configura y envía la solicitud HTTP GET
    http.Open "GET", url, False
    http.send
    
    ' Captura la respuesta del servidor
    response = http.responseText
    
    ' Se omite la verificación de la respuesta del servidor y del objeto JSON retornado para capturar el error mediante excepcion en vba
    'If http.Status = 200 Then
        ' Analiza y retorna el objeto JSON
        Set wsGetRendPoIdFromAPI = JsonConverter.ParseJson(response)
    'Else
        'Set wsGetRendPoIdFromAPI = "Error"
        ' Termina la ejecución en caso de error de estado HTTP
        'End
    'End If
End Function

' ============================================= '
' Public Methods
' ============================================= '

''
' Realiza una consulta a la API y devuelve un objeto JSON con los datos del rendimiento por operación (RendPo).
'
' @method wsGetRendPoFromAPI
' @param {String} KEY_SEARCH - Clave utilizada para realizar la búsqueda en la API.
' @return {Object} - Objeto JSON (Collection o Dictionary) que contiene los datos de la respuesta de la API.
' @throws Error si el servidor retorna un código de estado diferente a 200.
' @author Jorge Martínez S.
'

Public Function wsGetRoCFromAPI(KEY_SEARCH As String) As Object
    ' Declaración de variables
    Dim http As Object
    Dim url As String
    Dim response As String
    
    ' Construye la URL con el parámetro de búsqueda
    url = "https://tracksys.cl/api/wsGetRoC.php?api_key=" & KEY_SEARCH
    
    ' Inicializa el objeto HTTP para realizar la solicitud
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Configura y envía la solicitud HTTP GET
    http.Open "GET", url, False
    http.send
    
    ' Captura la respuesta del servidor
    response = http.responseText
    
    ' Verifica si la solicitud fue exitosa (código HTTP 200)
    If http.Status = 200 Then
        ' Analiza y retorna el objeto JSON
        Set wsGetRoCFromAPI = JsonConverter.ParseJson(response)
    Else
        ' Termina la ejecución en caso de error de estado HTTP
        End
    End If
    'Debug.Print wsGetRoCFromAPI.Count
End Function


' ============================================= '
' Public Methods
' ============================================= '

''
' Codifica una cadena de texto en formato URL utilizando la codificación UTF-8.
'
' @method UrlEncodeUTF8
' @param {String} str - Cadena de texto a codificar.
' @return {String} - Cadena codificada en formato URL con codificación UTF-8.
' @author Jorge Martínez S.
'

Public Function UrlEncodeUTF8(str As String) As String
    ' Declaración de variables
    Dim i As Long
    Dim char As String
    Dim encoded As String
    Dim byteArray() As Byte
    Dim utf8String As String
    
    ' Convertir la cadena a UTF-8 usando ADODB.Stream
    utf8String = ConvertToUTF8(str)

    ' Procesar cada carácter en la cadena UTF-8
    For i = 1 To Len(utf8String)
        char = Mid(utf8String, i, 1)
        If char Like "[A-Za-z0-9-_.~]" Then
            ' Caracteres seguros (sin codificar)
            encoded = encoded & char
        Else
            ' Codificar caracteres especiales como %XX
            encoded = encoded & "%" & Right("0" & Hex(Asc(char)), 2)
        End If
    Next i
    
    ' Retornar la cadena codificada
    UrlEncodeUTF8 = encoded
End Function



' ============================================= '
' Private Methods
' ============================================= '
'
''
' Convierte una cadena de texto en su representación UTF-8.
'
' @method ConvertToUTF8
' @param {String} str - Cadena de texto a convertir.
' @return {String} - Cadena de texto convertida a UTF-8.
' @remarks Utiliza el objeto ADODB.Stream para realizar la conversión.
' @author Jorge Martínez S.
'

Private Function ConvertToUTF8(str As String) As String
    ' Declaración de variables
    Dim stream As Object
    Dim byteArray() As Byte

    ' Crear un objeto ADODB.Stream
    Set stream = CreateObject("ADODB.Stream")
    
    ' Configurar el objeto Stream para trabajar con texto y codificación UTF-8
    stream.Type = 2 ' Modo Texto
    stream.Charset = "utf-8"
    stream.Open
    
    ' Escribir la cadena en el Stream
    stream.WriteText str
    
    ' Cambiar a modo Binario para leer los bytes
    stream.Position = 0
    stream.Type = 1 ' Modo Binario
    
    ' Leer los datos binarios en un array de bytes
    byteArray = stream.Read
    
    ' Liberar el objeto Stream
    stream.Close
    Set stream = Nothing
    
    ' Convertir los bytes a una cadena y retornar el resultado
    ConvertToUTF8 = StrConv(byteArray, vbUnicode)
End Function

' ============================================= '
' Public Methods
' ============================================= '

''
' Realiza una consulta a la API y devuelve un objeto JSON con los datos del rendimiento por operación (RendPo).
'
' @method wsGetRendPoFromAPI
' @param {String} KEY_SEARCH - Clave utilizada para realizar la búsqueda en la API.
' @return {Object} - Objeto JSON (Collection o Dictionary) que contiene los datos de la respuesta de la API.
' @throws Error si el servidor retorna un código de estado diferente a 200.
' @author Jorge Martínez S.
'

Public Function wsGetFacPonFromAPI(ROC_VALUE, PON_VALUE As String) As Object
    ' Declaración de variables
    Dim http As Object
    Dim url As String
    Dim response As String
    
    ' Construye la URL con el parámetro de búsqueda
    url = "https://tracksys.cl/api/wsGetFacPon.php?key_roc=" & ROC_VALUE & "&key_pon=" & PON_VALUE
    
    ' Inicializa el objeto HTTP para realizar la solicitud
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Configura y envía la solicitud HTTP GET
    http.Open "GET", url, False
    http.send
    
    ' Captura la respuesta del servidor
    response = http.responseText
    
    ' Verifica si la solicitud fue exitosa (código HTTP 200)
    If http.Status = 200 Then
        ' Analiza y retorna el objeto JSON
        Set wsGetFacPonFromAPI = JsonConverter.ParseJson(response)
    Else
        ' Termina la ejecución en caso de error de estado HTTP
        End
    End If
End Function


' ============================================= '
' Public Methods
' ============================================= '

''
' Consulta el rendimiento por operación (RendPo) utilizando una clave API.
' Esta función realiza una llamada a la API para obtener el rendimiento y retorna su valor numérico.
'
' @method CONSULTA_REND_PO
' @param {String} API_KEY - Clave utilizada para la búsqueda en la API.
' @return {Variant} - Valor numérico del rendimiento ("rend_po").
' @throws Error si la función auxiliar `wsGetRendPoFromAPI` o el objeto JSON devuelto no contiene los datos esperados.
' @author Jorge Martínez S.
'

Public Function CONSULTA_FAC_PON(API_ROC, API_PON As String) As Variant
    
    ' Declaración de variables
    Dim QRY_TO_API As Object
    
    ' Realiza la consulta a la API utilizando la función auxiliar `wsGetRendPoFromAPI`
    Set QRY_TO_API = wsGetFacPonFromAPI(API_ROC, UrlEncodeUTF8(API_PON))
    
    ' Extrae y convierte el valor del campo "rend_po" a numérico
    CONSULTA_FAC_PON = Val(QRY_TO_API(1)("factor_pond"))
   
End Function


