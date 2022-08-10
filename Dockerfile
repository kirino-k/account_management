FROM cdrx/pyinstaller-windows:python3
COPY ./src/requirements.txt /
RUN cd / && \
    pip install -r requirements.txt && \
    rm requirements.txt
