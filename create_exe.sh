#!/bin/bash
VERSION=0.3.0
docker run --rm --volume $(pwd)/src:/src --entrypoint /bin/sh cdrx/pyinstaller-windows:python3 \
-c "pip install -r requirements.txt && pyinstaller main.py --noconsole --onefile --clean --name account_management-${VERSION}"
