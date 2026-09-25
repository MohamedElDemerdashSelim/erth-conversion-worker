FROM eclipse-temurin:21-jre-jammy AS epubcheck
ARG EPUBCHECK_VERSION=5.4.0
RUN apt-get update && apt-get install -y curl unzip && rm -rf /var/lib/apt/lists/* \
 && curl -fsSL -o /tmp/epubcheck.zip https://github.com/w3c/epubcheck/releases/download/v${EPUBCHECK_VERSION}/epubcheck-${EPUBCHECK_VERSION}.zip \
 && unzip /tmp/epubcheck.zip -d /opt \
 && mv /opt/epubcheck-${EPUBCHECK_VERSION} /opt/epubcheck

FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends default-jre-headless pandoc \
 && rm -rf /var/lib/apt/lists/*
COPY --from=epubcheck /opt/epubcheck /opt/epubcheck
WORKDIR /srv
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY main.py ./main.py
ENV EPUBCHECK_JAR=/opt/epubcheck/epubcheck.jar
ENV PANDOC_BIN=pandoc
EXPOSE 8080
CMD ["sh","-c","uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]
