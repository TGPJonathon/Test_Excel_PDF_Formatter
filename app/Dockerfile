FROM ubuntu:20.04

ENV DEBIAN_FRONTEND noninteractive

RUN apt-get update \
    && apt-get install python3 -y \
    && apt-get install python3-pip -y \
    && apt-get install nano -y

RUN apt-get install -y openjdk-11-jre

WORKDIR /app

COPY ./requirements.txt requirements.txt

RUN pip install -r /app/requirements.txt

COPY . .

EXPOSE 5000

CMD ["python3", "-m", "flask", "run", "--host=0.0.0.0"]