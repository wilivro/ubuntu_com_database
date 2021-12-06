FROM ubuntu:20.04
RUN apt-get update
RUN apt-get install python3 -y
RUN apt install python3-pip -y
ENV PYTHONUNBUFFERED 1
RUN mkdir /empreendedor_alagoas
ADD ./python /empreendedor_alagoas
WORKDIR /empreendedor_alagoas
RUN pip3 install gspread oauth2client
RUN pip3 install mysql-connector-python
RUN pip3 install python-csv
RUN pip3 install uuid


