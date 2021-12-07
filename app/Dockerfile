FROM ubuntu:20.04

ENV DEBIAN_FRONTEND noninteractive

RUN apt-get update \
    && apt-get install python3 -y \
    && apt-get install python3-pip -y \
    && apt-get install nano -y

# Needed to convert Excel to PDF
RUN apt-get install -y openjdk-11-jre

# Section needed to download windows fonts
RUN apt-get install -y cabextract

RUN apt-get install -y wget

RUN wget https://www.freedesktop.org/software/fontconfig/webfonts/webfonts.tar.gz

RUN tar -xzf webfonts.tar.gz

WORKDIR /msfonts

RUN cabextract *.exe

RUN cp *.ttf *.TTF /usr/local/share/fonts/
# End Section

WORKDIR /app

COPY ./requirements.txt requirements.txt

RUN pip install -r /app/requirements.txt

COPY . .

EXPOSE 5000

CMD ["python3", "app.py"]