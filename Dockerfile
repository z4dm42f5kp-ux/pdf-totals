# Container image of tute-lote, published to GitHub Packages (ghcr.io) with every release.
#
#   podman run --rm -v /path/to/pdfs:/data:Z ghcr.io/z4dm42f5kp-ux/tute-lote:stable /data
#
# The results are written to /path/to/pdfs/resultados_lotes. The container needs no
# network access: add --network none.
FROM docker.io/library/python:3.12-slim

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir --root-user-action=ignore -r /app/requirements.txt \
 && pip uninstall -y --root-user-action=ignore opencv-python \
 && pip install --no-cache-dir --root-user-action=ignore opencv-python-headless \
 && useradd --create-home --uid 1000 lotes

COPY lotes_auto.py diagnostico.py LICENSE /app/
COPY etapas /app/etapas

USER lotes
WORKDIR /data
ENTRYPOINT ["python", "/app/lotes_auto.py"]
CMD ["--help"]
