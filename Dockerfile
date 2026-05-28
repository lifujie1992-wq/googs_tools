FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY app.py xlsx_tools.py ./
COPY static ./static

RUN mkdir -p /app/work

EXPOSE 8123

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8123", "--proxy-headers"]
