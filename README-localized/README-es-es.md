---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: El código JavaScript de este ejemplo muestra una solicitud sencilla para el asunto del mensaje de correo electrónico actual. Muestra los pasos necesarios para crear una solicitud de servicios Web Exchange (EWS) y los procedimientos recomendados para realizarla.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:32:51 PM
---
# Complemento de Outlook: Realizar una solicitud de Servicios Web Exchange desde Outlook

**Tabla de contenido**

* [Resumen](#summary)
* [Requisitos previos](#prerequisites)
* [Componentes clave del ejemplo](#components)
* [Descripción del código](#codedescription)
* [Compilar y depurar](#build)
* [Solución de problemas](#troubleshooting)
* [Preguntas y comentarios](#questions)
* [Recursos adicionales](#additional-resources)

<a name="summary"></a>
## Resumen
El código JavaScript de este ejemplo muestra una solicitud sencilla para el asunto del mensaje de correo electrónico actual. Muestra los pasos necesarios para crear una solicitud de servicios Web Exchange (EWS) y los procedimientos recomendados para realizar la solicitud.

<a name="prerequisites"></a>
## Requisitos previos ##

Este ejemplo necesita lo siguiente:  

  - Visual Studio 2013 con Update 5 o Visual Studio 2015.  
  - Un equipo que ejecute Exchange 2013 y, como mínimo, una cuenta de correo electrónico o una cuenta de Office 365. Puede [participar en el programa para desarrolladores Office 365 y obtener una suscripción gratuita durante 1 año a Office 365](https://aka.ms/devprogramsignup).
  - Cualquier explorador que admita ECMAScript 5.1, HTML5 y CSS3, como Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 o una versión posterior de estos exploradores.
  - Familiaridad con los servicios web y la programación de JavaScript.

<a name="components"></a>
## Componentes clave del ejemplo
La solución de ejemplo contiene los siguientes archivos:

- MakeEwsRequestManifest.xml: El archivo de manifiesto del complemento de Outlook.
- AppRead\Home\Home.html: La interfaz de usuario HTML para el complemento de correo electrónico de Outlook.
- AppRead\Home\Home.js: Es el archivo JavaScript que controla la solicitud y el uso de la solicitud EWS. 

<a name="codedescription"></a>
##Descripción del código

El código que crea la solicitud XML de EWS incluye dos métodos. El primer método, `getSoapEnvelope()`, encapsula un sobre SOAP en torno a una solicitud de servicio Web. Dado que el sobre SOAP es estándar para todas las solicitudes EWS, este método se puede volver a usar para encapsular cualquier solicitud EWS.

El segundo método, `getSubjectRequest()`, devuelve la solicitud EWS para obtener el campo Subject de un elemento. El parámetro id es el identificador del elemento de Exchange para el elemento solicitado. Tenga en cuenta lo siguiente sobre la solicitud:

- El elemento `ItemShape` se usa para restringir la respuesta en la forma base `IdOnly`. Esto limita la respuesta solo al identificador del elemento para este elemento y evita que se envíen demasiados datos desde el servidor. 
- El elemento `AdditionalProperties` se usa para agregar el campo Subject a la respuesta. Usando la forma base `IdOnly` y una lista de propiedades adicionales, puede limitar el tamaño de la respuesta del servidor solo a los datos que necesita el complemento. 

El método `sendRequest()` es llamado cuando se hace clic en el botón **Make EWS request** en la interfaz de usuario del complemento. Obtiene el identificador de Exchange del elemento actual y lo pasa a los métodos `getSubjectRequest()` y `getSoapEnvelope()` y, a continuación, realiza una llamada asincrónica al servidor de Exchange con el método `makeEwsRequestAsync`. El método `makeEwsRequestAsync` toma dos parámetros: la solicitud EWS encapsulada en el sobre SOAP y un método de devolución de llamada al que se llama cuando finaliza la solicitud asincrónica para EWS. Puede agregar un tercer parámetro opcional `userContext` al método `makeEwsRequestAsync` si tiene que proporcionar información adicional al método de devolución de llamada.

El método `callback()` es llamado con un solo parámetro, `asyncResult`. El objeto `asyncResult` tiene dos miembros:

- `valor` – El contenido de la respuesta de EWS. 
- `contexto` – El parámetro `userContext` pasado al método [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2). 

El método `callback` del ejemplo muestra el contenido de la respuesta en un elemento desplazable `div` en la interfaz de usuario, pero su código puede usar la respuesta de formas más sofisticadas.

<a name="build"></a>
## Compilar y depurar ##
**Nota**: El complemento de correo se activará en cualquier mensaje de correo electrónico de la bandeja de entrada del usuario. Puede hacer que sea más fácil probar el complemento enviando uno o más mensajes de correo electrónico a la cuenta de prueba antes de ejecutar el complemento de ejemplo.

1. Abra la solución en Visual Studio. Pulse F5 para compilar e implementar el complemento de ejemplo.
2. Conecte a una cuenta de Exchange proporcionando la dirección de correo electrónico y la contraseña de un servidor de Exchange 2013.
3. Permita que el servidor configure la cuenta de correo.
4. Inicie sesión en la cuenta de correo electrónico escribiendo el nombre de la cuenta y la contraseña. 
5. Selecciones un mensaje en la Bandeja de entrada
6. Espere a que aparezca la barra del complemento sobre el mensaje.
7. En la barra del complemento, haga clic en **MakeEWSRequest**.
8. Cuando aparezca el complemento de correo, haga clic en el botón **Realizar solicitud EWS** para solicitar el asunto del mensaje actual desde el servidor Exchange.
9. Revise el XML de respuesta devuelto por la solicitud.

<a name="troubleshooting"></a>
##Solución de problemas.
A continuación se enumeran los errores comunes que pueden ocurrir al usar Outlook Web App para probar un complemento de correo para Outlook:

- La barra de complemento no aparece cuando se selecciona un mensaje. Si esto ocurre, vuelva a iniciar la aplicación seleccionando **Depuración: detener depuración** en la ventana de Visual Studio y presione F5 para recompilar e implementar el complemento. 
- Es posible que los cambios en el código de JavaScript no se hayan recogido al implementar y ejecutar el complemento. Si no se han añadido los cambios, borre la memoria caché en el explorador web. Para ello, seleccione **Herramientas: opciones de Internet** y haga clic en el botón **Eliminar...**. Elimine los archivos temporales de Internet y reinicie el complemento. 

<a name="questions"></a>
##Preguntas y comentarios##

- Si tiene algún problema para ejecutar este ejemplo, [registre un problema](https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest/issues).
- Las preguntas sobre el desarrollo de complementos para Office en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-addins].


<a name="additional-resources"></a>
## Recursos adicionales ##

- [Más complementos de ejemplo](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Explore la API administrada de EWS, EWS y servicios web de Exchange](https://msdn.microsoft.com/library/office/jj536567(v=exchg.150).aspx)
- [método makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2)

## Copyright
Copyright (c) 2015 Microsoft. Todos los derechos reservados.


Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
