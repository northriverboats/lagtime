# Send Lagtime Report

## Build
```
git clone git@github.com:northriverboats/lagtime.git
cd lagtime
python3 -m venv venv
source venv/bin/activate
python -m pip install pip --upgrade
pip install -r requirements.txt

cp env.sample .env
# edit .env to include db / mailserver /recepients

pyinstaller --singlefile lagtime.spec

cp dist/sendlagtime <your destination>
```
