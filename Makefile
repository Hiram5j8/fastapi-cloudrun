# ===== 本機模式 =====
dev:
	python run_local.py

api:
	uvicorn main:app --host 0.0.0.0 --port 8080 --reload


# ===== Docker 模擬 Cloud Run =====
build:
	docker build -t fastapi-local .

run:
	docker run -p 8080:8080 fastapi-local


# ===== 清理 =====
clean:
	docker system prune -f


# ===== 測試 API =====
test:
	curl http://localhost:8080


# ===== 一鍵全流程 =====
all: build run