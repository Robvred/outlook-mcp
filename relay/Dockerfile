FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY outlook_mcp_relay.py .
RUN useradd --no-create-home relay
USER relay
EXPOSE 8080 8765
CMD ["python", "outlook_mcp_relay.py"]
