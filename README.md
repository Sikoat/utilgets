Util.ini and similar files are expected to be in h:\util or else in c:\util.

*************************

Example Methods:

- repository or main branch:

git clone https://github.com/Sikoat/utilgets.git
where git is as in the portable version at https://github.com/git-for-windows/git/releases

or simply browser download zip

git clone --branch main --single-branch https://github.com/Sikoat/utilgets.git

curl -L -o utilgets-main.zip "https://github.com/Sikoat/utilgets/archive/refs/heads/main.zip"

- one file:

curl "https://raw.githubusercontent.com/Sikoat/utilgets/refs/heads/main/util.ini" -o util.ini

wget -L -O util.ini "https://raw.githubusercontent.com/Sikoat/utilgets/refs/heads/main/util.ini"

bitsadmin /transfer util_update_job /download /priority normal "https://raw.githubusercontent.com/Sikoat/utilgets/refs/heads/main/util.ini" "%CD%\util.ini"

python -c "import urllib.request; urllib.request.urlretrieve('https://raw.githubusercontent.com/Sikoat/utilgets/refs/heads/main/util.ini','util.ini')"
