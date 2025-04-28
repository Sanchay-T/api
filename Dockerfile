# Stage 1: Use an official Python runtime as a parent image
# Using the full image instead of slim, as it includes build tools
FROM python:3.11

# Set the working directory inside the container
WORKDIR /app

# Prevent python from writing pyc files to disc (optional)
ENV PYTHONDONTWRITEBYTECODE 1
# Ensure python output is sent straight to the terminal without buffering
ENV PYTHONUNBUFFERED 1

# Install system dependencies if needed
# Explicitly install build tools, even with the full image, to ensure packages like blis can compile from source if needed.
RUN apt-get update && apt-get install -y --no-install-recommends build-essential pkg-config && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
# First, copy only the requirements file to leverage Docker cache
COPY backend/requirements.txt backend/requirements.txt
# Upgrade pip and install dependencies
RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r backend/requirements.txt

# Copy the application code into the container
# We copy the backend directory specifically
COPY backend/ /app/backend/

# --- This is crucial ---
# Copy the CA certificate into the image
COPY backend/ca/ca-certificate.crt /app/backend/ca/ca-certificate.crt

# Expose the port the app runs on
# This informs Docker the container listens on this port, but doesn't publish it.
# Publishing is done when running the container or by App Platform.
EXPOSE 7500

# Define the command to run the application
# Use gunicorn as the process manager, specifying uvicorn workers
# It reads host/port from ENV variables (API_HOST, API_PORT) or defaults
# App Platform will inject these. For local testing, we'll set them.
CMD ["gunicorn", "--workers", "4", "--worker-class", "uvicorn.workers.UvicornWorker", "--bind", "0.0.0.0:7500", "backend.main:app"] 
