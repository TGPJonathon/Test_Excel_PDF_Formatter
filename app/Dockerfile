FROM python:3

WORKDIR /app

COPY ./requirements.txt requirements.txt

RUN pip install --upgrade pip \
    && pip install --no-cache-dir -r /app/requirements.txt

COPY . .

CMD ["python", "-m", "flask", "run", "--host=0.0.0.0"]