FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Install Korean fonts for report generation
RUN apt-get update && apt-get install -y \
    fonts-nanum \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Expose port 8080 (Cloud Run's default)
EXPOSE 8080

# Run the Streamlit application
CMD ["python", "-m", "streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
