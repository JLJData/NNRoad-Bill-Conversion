FROM python:3.11-slim-bookworm

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV CONVERT_HOST=127.0.0.1
ENV CONVERT_PORT=8765
EXPOSE 8765

CMD ["sh", "-c", "python -m uvicorn convert_api:app --host ${CONVERT_HOST} --port ${CONVERT_PORT}"]
