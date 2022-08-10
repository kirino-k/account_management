#!/bin/bash

# set path
DIR_CMD=$(cd $(dirname $0); pwd)

# define version automatically
VERSION=$(cat ${DIR_CMD}/src/main.py | grep 'Version' | tr -cd [0-9.])

# create exe file in docker container
docker run \
  --rm \
  --volume ${DIR_CMD}:/src \
  --entrypoint /bin/sh \
  account_management \
  -c "pyinstaller ./src/main.py --noconsole --onefile --clean --name account_management-${VERSION}"

