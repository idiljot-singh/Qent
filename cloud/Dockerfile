# Qent web + collector image. Same image runs the web app (default) and the cron Job (`python qent.py collect`).
# Build from the repo root: docker build -f cloud/Dockerfile .  (deploy.ps1 also runs from the repo root)
FROM python:3.14-slim
ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1 PORT=8080
RUN useradd --system --uid 10001 --no-create-home qent
WORKDIR /app
COPY cloud/requirements.txt .
RUN pip install --no-cache-dir --require-hashes -r requirements.txt
COPY qent.py .
COPY fonts ./fonts
COPY static ./static
USER 10001
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s CMD python -c "import os,urllib.request;urllib.request.urlopen('http://127.0.0.1:%s/healthz' % os.environ['PORT'])"
CMD ["python", "qent.py", "serve"]
