Nov 24 release notes:
- Flat Finder results are now sorted in descending order by date.
- The screenshot window is now ctrl-alt-T instead of ctrl-alt-I.
- Ctrl-alt-I now calculates the multiplicative inverse, the reciprocal, of clipboard contents (if starting as a number >= 1).  It both replaces the clipboard contents with such and attempts to paste the result.
- A helper function was added, shared by the four buttons using test entry field input ... such that, if the former is blank but the clipboard contents appear probably suitable (by including two or more numerical digits), it proceeds with the latter.
- The plausibility check continues to make most of its numerical table from clipboard contents (in the four-column format received from copy all on the BAQ).  However, it now also looks at the text entry field where the price comparison qty,price (or 'qty price') should be entered and warns if the latter is absent.
- Ctrl-alt-D is an added global hotkey, opening the dayfolder like the local ctrl-D hotkey.
- The text entry field is generally made active whenever the program window comes to the foreground, unless the program is awaiting a keypress in the Quick Actions functionality.

*************************

Startup shortcut, to update now, probably in:
C:\Users\<name>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup

*************************

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


