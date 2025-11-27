# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container at /app
COPY requirements.txt .

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the current directory contents into the container at /app
COPY . .

# Expose the port that the application listens on.
# Cloud Run expects the container to listen on port 8080 by default.
EXPOSE 8080

# Run the application
# We use sh -c to allow variable expansion for the PORT environment variable
CMD sh -c "streamlit run dashboard.py --server.port=${PORT:-8080} --server.address=0.0.0.0"
