# WhatsApp-Messages
Software para la automatización de mensajes de WhatsApp desde una lista de Excel, con Python.

SOBRE LA BASE DE DATOS (DDBB)
En el archivo "DDBB.xls" se guardaran los registros de los clientes a los que queremos enviar los mensajes.
COLUMNA 1: En la primera columna "Confimar Cedula Cliente (Cliente)" se ingresara la cedula o ID del cliente *no obligatorio
COLUMNA 2: En la segunda columna "Nombre del titular de la cuenta" se ingresara el nombre de la persona a quien enviaremos el mensaje *opcional, pero recomendado
COLUMNA 3/4: En la columna 3 y 4 se establecen los numeros a los que se enviaran los mensajes de whatsapp. En caso de solo poseer un numero de telefono, rellenar la casilla restante con un punto "."
COLUMNA 5: En la columna 5 se ponen los dias de atraso que tiene el cliente *opcional
COLUMNA 6: En la sexta y ultima columna va el monto que esta pendiente *opcional

SOBRE EL FUNCIONAMIENTO DEL SOFTWARE
Utilizar la aplicacion resulta algo sencillo. Tras abrir la aplicacion, se escribe el mensaje que se le quiere enviar al cliente, en el espacio asignado. Luego preciona el boton "Cargar Archivo" para cargar los datos desde el archivo de excel "DDBB".
Para finalizar preciona el boton "Enviar" para comenzar el envio de los mensajes.

A TOMAR EN CUENTA
El software solo esta hecho considerando numeros telefonicos de la Republica Dominicana. Para su funcionamiento con cualquier otro numero de telefono de otro pais requiere una modificacion en el codigo, para alterar el codigo de pais.
Estoy consiente que a este software le faltan muchas actualizaciones y mejoras, las cuales las ire realizando a medida que pasa el tiempo.
