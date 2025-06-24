# 📨 Carga automática diaria a SQL Server desde un correo electrónico con Power Automate y Python

Este proyecto automatiza el proceso de recepción, limpieza y carga de un reporte Excel que llega por correo electrónico. El flujo comienza con **Power Automate**, que descarga el archivo y lo almacena en **SharePoint**, seguido por un script en **Python** que realiza la extracción, transformación y carga de los datos en una base de datos **SQL Server**.

El proceso está programado para ejecutarse automáticamente todos los días antes de comenzar la jornada laboral.

---

## 🚀 Características Principales

- ✅ Automatización completa: Correo → SharePoint → Python → SQL Server
- 📥 Descarga automática desde SharePoint usando Office365 API
- 🧹 Limpieza y transformación de datos con Pandas
- 💽 Carga rápida a SQL Server con PyODBC
- ⏱ Proceso total: menos de **20 segundos**
- 👤 Seguridad mediante archivo "runETL.txt"

---

## 🛠 Tecnologías Usadas

- **Python**: Programación del proceso ETL
- **Pandas**: Manejo y transformación de datos
- **PyODBC**: Conexión y carga a SQL Server
- **Office365 / SharePoint**: Almacenamiento seguro de archivos
- **Power Automate**: Activador del proceso al recibir el correo
- **Windows Task Scheduler**: Programación diaria del script

---

## 📋 Requisitos

```bash
pip install pandas pyodbc openpyxl office365-cli
```
## Notas Adicionales
- Este proceso reemplazó un trabajo manual de ~30 minutos por uno completamente automatizado en menos de 20 segundos.
- Se evitó el uso de SSIS gracias a la integración con Python y Power Automate.
- El archivo runETL.txt actúa como una señal de seguridad para evitar procesamientos incompletos o innecesarios.

## 🦁 By Larry Humpiri (LK)
- 📧 Email: larryhumpiri@gmail.com
- 🔗 GitHub: https://github.com/LarryHumpiri
- 💼 LinkedIn: https://www.linkedin.com/in/larry-humpiri-obregon-565145189/