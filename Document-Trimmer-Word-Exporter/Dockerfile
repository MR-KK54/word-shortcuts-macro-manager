FROM python:3.11-slim-bookworm

# LibreOffice is required on Render to render previews and paginate documents
# that lack MS Word pagination markers. Install the Writer suite plus fonts and
# common dependencies, then fail the build loudly if soffice is not present.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-core \
    libreoffice-common \
    libreoffice-style-colibre \
    fonts-dejavu \
    fonts-liberation \
    fontconfig \
    && rm -rf /var/lib/apt/lists/* \
    && if ! command -v soffice >/dev/null 2>&1 && [ ! -e /usr/bin/soffice ]; then \
         echo "ERROR: LibreOffice soffice binary missing after apt install"; \
         exit 1; \
       fi

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["gunicorn", "server:app", "--workers", "1", "--threads", "8", "--timeout", "900"]
