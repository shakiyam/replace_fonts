FROM python:3.14-slim-trixie
COPY --from=ghcr.io/astral-sh/uv:0.9 /uv /bin/uv
WORKDIR /opt/replace_fonts
COPY requirements.txt .
RUN uv pip install --system --no-cache-dir -r requirements.txt
COPY replace_fonts.py .
WORKDIR /work
ENTRYPOINT ["python3", "/opt/replace_fonts/replace_fonts.py"]
