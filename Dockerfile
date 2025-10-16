FROM python:3.14-slim-bookworm
COPY --from=ghcr.io/astral-sh/uv:0.9.2 /uv /bin/uv
COPY requirements.txt /requirements.txt
RUN uv pip install --system --no-cache-dir -r /requirements.txt
COPY replace_fonts.py /replace_fonts.py
WORKDIR /work
ENTRYPOINT ["python3", "/replace_fonts.py"]
