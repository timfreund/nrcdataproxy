FROM python:3-stretch

WORKDIR /usr/local/nrc

RUN groupadd -r nrc && \
    useradd -r -g nrc nrc && \
    apt-get update && \
    apt-get install -y libpq-dev

COPY . .
RUN pip install --no-cache-dir -r ./requirements.txt
RUN python setup.py develop
RUN mkdir -p /usr/local/nrc/data && chown -R nrc:nrc /usr/local/nrc

USER nrc

CMD ["nrcdataproxy"]
