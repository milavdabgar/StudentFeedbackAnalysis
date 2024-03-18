#! /bin/sh
echo "======================================================================"
echo "Welcome to to the setup. This will setup the local virtual env." 
echo "And then it will install all the required python libraries."
echo "You can rerun this without any issues."
echo "----------------------------------------------------------------------"

if [ -d ".venv" ];
then
    echo ".venv folder exists. Installing using pip"
else
    echo "creating .pyvenv and install using pip"
    python -m venv .venv
fi

# Activate virtual env
. .venv/bin/activate

# Upgrade the PIP
pip install --upgrade pip
pip install -r requirements.txt
export FLASK_APP=app.py
# flask db init
# Work done. so deactivate the virtual env
deactivate
