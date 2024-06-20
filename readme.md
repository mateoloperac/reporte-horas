# Proyecto de Exportación de Eventos de Outlook a Excel

Este proyecto permite extraer eventos del calendario de Outlook y exportarlos a un archivo Excel.

## Requisitos

- Python 3.x
- pip
- git

## Instrucciones

### Clonar el repositorio

Para clonar el repositorio, abre una terminal y ejecuta el siguiente comando:

```bash
git clone https://github.com/mateoloperac/reporte-horas
```

En la terminal navegar al proyecto

```bash
cd .\reporte-horas\
```

### Crear un entorno virtual

Primero, asegúrate de tener virtualenv instalado. Si no lo tienes, puedes instalarlo usando pip:

```python
pip install virtualenv
```

Luego, crea un entorno virtual llamado .venv dentro del directorio del proyecto:

```python
virtualenv .venv
```

### Activar el entorno virtual

Para activar el entorno virtual, utiliza el siguiente comando dependiendo de tu sistema operativo:

```python
.\.venv\Scripts\activate
```

### Instalar dependencias

Con el entorno virtual activado, instala las dependencias necesarias ejecutando:

```python
pip install -r requirements.txt
```

### Ejecutar el script

Finalmente, para ejecutar el script principal, usa el siguiente comando:

```python
python nombre_del_script.py
```
