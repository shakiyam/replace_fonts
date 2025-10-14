FROM python:3.14-slim-bookworm
COPY requirements.txt /requirements.txt
# hadolint ignore=DL3013
RUN python -m pip install --no-cache-dir --upgrade pip && python -m pip install --no-cache-dir -r /requirements.txt
COPY replace_fonts.py /replace_fonts.py
WORKDIR /work
ENTRYPOINT ["python", "/replace_fonts.py"]
