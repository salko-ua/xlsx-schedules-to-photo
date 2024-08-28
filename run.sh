#!/usr/bin/env bash

rm -rf ./screenshots ./variant
poetry run python main.py
ssh ubuntu@57.128.195.105 'rm -rf configs/data/hosted-files/schedules/*'
scp -r ./screenshots/* ubuntu@57.128.195.105:~/configs/data/hosted-files/schedules/
scp -r ./index.html ubuntu@57.128.195.105:~/configs/data/hosted-files/schedules/
