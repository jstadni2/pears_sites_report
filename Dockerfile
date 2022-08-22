FROM python:3.9

WORKDIR /pears_sites_report

COPY . .

RUN pip install -r requirements.txt

CMD [ "python", "./pears_sites_report.py" ]