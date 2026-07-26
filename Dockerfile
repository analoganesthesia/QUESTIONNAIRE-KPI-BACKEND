FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Volume à monter pour persister study.db en dehors du conteneur (voir README)
VOLUME ["/app/data"]
ENV DATABASE_URL=sqlite:////app/data/study.db

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
