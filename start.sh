#!/bin/sh

docker run -it --rm \
    --volume $(pwd):/usr/src/app \
    boeingtuan/excel_price_remover:latest \
    sh