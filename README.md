1. Download Git for Windows from [https://git-scm.com/install/windows](https://github.com/git-for-windows/git/releases/download/v2.55.0.windows.5/Git-2.55.0.5-64-bit.exe).

2. Download Python install Manager from https://www.python.org/downloads,then download the latest released version. Run the .exe file, make sure to tick the checkbox "Add python.exe to PATH" and Install Now.

3. Optional for Python code writing, testing or debug only. Download PyCharm for Windows from https://www.jetbrains.com/pycharm/download/download-thanks.html?platform=windows.

4. Open the Command prompt or Git Bash (Windows).
  
5. Download source code from Github using the command below:
   
	git clone https://github.com/kangyeechong/sqlite3-thing.git

	python --version (to make sure Python is downloaded)

	cd sqlite3-thing

	pip install -r requirements.txt

	python create_user.py --db ledger.db --email their_own_email@xekl.com --name "Their Name"

7. Check where did the file is downloaded, if it's in the Desktop enter the command cd Desktop\sqlite3-thing.

8. After showing "C:\Users\User\Desktop\sqlite3-thing>" means you are inside the folder. Then enter the command python run_server.py --db ledger.db to run the server.

9. Copy the http://127.0.0.1:5000 and open via browser.

10. In the terminal, press Ctrl + C to quit the server.

