FROM python:3.12-windowsservercore

WORKDIR /app

COPY requirements.txt .
COPY main.py .
COPY word_parser.py .

RUN pip install -r requirements.txt
RUN pip install pyinstaller

CMD pyinstaller --name=WordParser --onefile --windowed --add-data="word_parser.py;." main.py 