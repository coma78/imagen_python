FROM ubuntu:latest
RUN apt-get update && apt-get install -y python3 && apt-get install -y python3-pip && pip install python-docx
WORKDIR /home
COPY acepta.py /home