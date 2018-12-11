#!/bin/sh
apt-get update
apt-get install libffi-dev libxml2-dev libxslt1-dev
pip install -r requirements.txt
