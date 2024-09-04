# Use the official Python image as base image
FROM python:3.8-slim-buster

# Copy the local script directory into the container
COPY . /app

# Set the working directory
WORKDIR /app

# Install the required libraries
RUN pip install pandas xlsxwriter openpyxl
# RUN pip install --no-cache-dir -r requirements.txt

# Run the script
CMD ["python", "app.py"]
# CMD ["python", "main.py"]