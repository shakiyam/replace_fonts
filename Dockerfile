FROM python:3.14-slim-trixie
WORKDIR /opt/replace_fonts
COPY requirements.txt .
RUN --mount=from=ghcr.io/astral-sh/uv:0.12,source=/uv,target=/bin/uv \
  uv pip install --system --no-cache-dir -r requirements.txt \
  && uv pip uninstall --system pip
COPY apply_theme_fonts.py define_theme_fonts.py logger.py replace_fonts.py ./
WORKDIR /work
ARG SOURCE_COMMIT
ENV SOURCE_COMMIT=$SOURCE_COMMIT
ENTRYPOINT ["python3", "/opt/replace_fonts/replace_fonts.py"]
