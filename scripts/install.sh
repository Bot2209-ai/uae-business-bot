#!/bin/bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3 python3-venv git
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
