# Usa una imagen base de Python
FROM python:3.10-slim

# Establece el directorio de trabajo
WORKDIR /app

# Copia el archivo de requerimientos
COPY requirements.txt .

# Instala las dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Copia el resto del código
COPY . .

# Expone el puerto (ej: 8000 o 5000)
EXPOSE 8000

# Comando para iniciar la aplicación (reemplaza con tu comando de inicio)
CMD ["python", "app.py"]