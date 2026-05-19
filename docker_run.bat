@echo off
echo Building Docker image...
docker build -t fastapi-local .

echo Running Docker container on http://localhost:8080
docker run -p 8080:8080 fastapi-local
pause