FROM ubuntu:latest

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y python3 python3-pip libicu-dev && rm -rf /var/lib/apt/lists/*
RUN apt-get update && apt-get install -y fontconfig
RUN apt-get update && apt-get install -y fonts-indic

# refresh system font cache
RUN fc-cache -f -v

WORKDIR /app

COPY requirements.txt .

RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt

COPY main1.py .
COPY functions.py .
COPY .env .

CMD ["python3", "main1.py"]