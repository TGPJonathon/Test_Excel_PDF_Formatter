FROM python:3

COPY ./requirements.txt /app/requirements.txt

WORKDIR /app

RUN pip install --upgrade pip \
    && pip install -r /project/requirements.txt

COPY . /app

EXPOSE 5000

ENTRYPOINT ["python"]

CMD ["server.py"]