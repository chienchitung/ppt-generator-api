FROM python:3.9-slim

WORKDIR /app

# Install system dependencies and locales
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    locales \
    && rm -rf /var/lib/apt/lists/* \
    && sed -i '/zh_TW.UTF-8/s/^# //g' /etc/locale.gen \
    && locale-gen

# Set locale environment variables
ENV LANG zh_TW.UTF-8
ENV LANGUAGE zh_TW:zh
ENV LC_ALL zh_TW.UTF-8

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create directory for generated PPTs
RUN mkdir -p generated_ppts

# Expose port
EXPOSE 8000

# Run the application
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"] 