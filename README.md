# SimpleExcelFile_CSharp
Este reporsitorio sirve para crear un Reporte de Horas en Excel, el siguiente es un ejemplo: 

![alt-image](https://i.imgur.com/ggr6Rbl.png "Reporte de Horas")

## Requisitos
- NET CORE >= 2.1.0
- ClosedXML >= 0.93.1

## Estructura
El projecto se creo en Visual Studio 2017 y se estructura de la siguiente manera:

- SimpleExcelFile_CSharp
    - Program.cs
    - Email.cs
    - ExcelFile.cs
    - SimpleEscelFile.csproj
- .gitattributes
- .gitignore
- README.md
- SimpleExcelFile.sln

## Información a llenar
Los siguientes datos deben ser llenados para el correcto funcionamiento del codigo.

### Datos
- __*fromEmail*__, Correo desde el cual se enviara toda la información.
- __*fromName*__, Nombre de la persona que envia el Correo.
- __*password*__, Contraseña del correo desde donde se envia la información.
- __*toEmail*__, Correos a los cuales les llegara la información separados por comas.
- __*filename*__, nombre del archivo que se adjuntara en el Correo.
- __*subject*__, Asunto del correo.
- __*body*__, cuerpo del correo.
