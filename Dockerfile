FROM python:3.12-slim

WORKDIR /usr/src/app/

EXPOSE 8501

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY streamlit_app.py ./
COPY pages ./pages

HEALTHCHECK CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/_stcore/health')"

ENTRYPOINT ["streamlit", "run", "streamlit_app.py", "--server.port=8501", "--server.address=0.0.0.0"]
