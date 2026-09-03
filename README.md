![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)
# AutoDocs
Aplicación de escritorio para crear plantillas de documentos Word académicos y convertir archivos Word a PDF.

## Características
- Generación de plantillas Word con portada, índice automático y secciones/subtítulos dinámicos.
- Conversión de documentos Word (.docx/.doc) a PDF.
- Conversión de documentos PDF (.pdf) a Word.
- Modo claro y oscuro.

## Requisitos
- Python 3.10 o superior
- Microsoft Word (Necesario para transformar a PDF)
- Sistema operativo Windows

## Dependencias Utilizadas
| Dependencia   | Uso                                         |
| ------------- | ------------------------------------------- |
| CustomTKinter | Creación de la interfaz gráfica.            |
| Python-docx   | Creación de documentos Word.                |
| Docx2Pdf      | Transformar documentos Word a PDF.          |
| Pdf2Docx      | Transformar documentos PDF a Word.          |
| PyInstaller   | Compilar el proyecto en un ejecutable `.exe`|

## Instalación y ejecución
### Opción 1: Desde el código fuente
1. Clonar el repositorio:
```git clone https://github.com/Bayron-jpg/AutoDocs```
2. Instalar las dependencias:
```pip install -r requirements.txt```
3. Ejecutar la aplicación:
```python AutoDocs.py```

### Opción 2: Ejecutable (.exe)
Descargar la versión portable de [AutoDocs](https://github.com/Bayron-jpg/AutoDocs/releases/download/V1.0.0/AutoDocs.exe).
> ⚠️ **Nota:** Windows puede mostrar una advertencia de SmartScreen al ejecutar el `.exe` por primera vez, 
> ya que no está firmado digitalmente. Es normal en apps independientes — 
> haz clic en "Más información" → "Ejecutar de todas formas".

## Screenshots / Preview
- Menú
<img width="797" height="522" alt="image" src="https://github.com/user-attachments/assets/5d37c783-a0fa-41b4-9a8f-abdb0a69a2f2" />

- Convertir a PDF
<img width="597" height="403" alt="image" src="https://github.com/user-attachments/assets/502c968d-441a-427d-9f4f-3eddc14872dd" />

- Convertir a PDF (Éxito)
<img width="600" height="405" alt="image" src="https://github.com/user-attachments/assets/0fcc1654-60a0-48fc-b40c-cc7f41723802" />

- Convertir a Word
<img width="594" height="405" alt="image" src="https://github.com/user-attachments/assets/921ba1b6-12bf-4139-94fb-a2e16c22abe0" />

- Convertir a Word (Éxito)
<img width="593" height="402" alt="image" src="https://github.com/user-attachments/assets/e15aaab4-af63-422c-ba70-a58c9ad7296f" />

- Crear Plantilla Word (Parte 1)
<img width="798" height="527" alt="image" src="https://github.com/user-attachments/assets/85330396-9eef-4331-a2fa-d69e3bf34a70" />

- Crear Plantilla Word (Parte2)
<img width="796" height="523" alt="image" src="https://github.com/user-attachments/assets/af374ab7-f7c6-4002-bd77-1afc7fcd28a2" />

## Plantilla Word
- Portada
<img width="650" height="842" alt="image" src="https://github.com/user-attachments/assets/379ba084-bfe7-4b74-b905-6f4d3af1159e" />

- Índice
<img width="648" height="842" alt="image" src="https://github.com/user-attachments/assets/cf4492e2-6b3f-46ea-9f78-5e075c5ac4b2" />

- Páginas de ejemplo
<img width="489" height="635" alt="image" src="https://github.com/user-attachments/assets/44f104cd-c7e5-4ed0-837d-e9ded8fede2a" />
<img width="488" height="637" alt="image" src="https://github.com/user-attachments/assets/31502c2e-f347-4af7-8cb7-dd6c8f0cdd71" />
<img width="491" height="634" alt="image" src="https://github.com/user-attachments/assets/9669a9b0-5af9-4e49-915e-6a01e352c4ea" />

## Licencia
Este proyecto está bajo la licencia MIT. Ver [LICENSE](LICENSE) para más detalles.

## Autor
Desarrollado por Bayron Urrutia.
