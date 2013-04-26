--------------------------------------------------
1 : install 1 python-2.7.1 then 
--------------------------------------------------

then add 1 environment variable :

PYTHON_HOME = c:\Python27 (or wherever you installed it)

To your path variable add :

;%PYTHON_HOME%;%PYTHON_HOME%\Scripts

(DO NOT forger the first ";")

--------------------------------------------------
2 install those files
--------------------------------------------------
2 pywin32-214.win32-py2.7
3 setuptools-0.6c11.win32-py2.7
4 launchwin.64

--------------------------------------------------
3 install the last one by typing in the commande line
--------------------------------------------------

easy_install c:\xx\xx\5virtualenv-1.8.4.tar.gz

In case you did not understood what was xx\xx punch
yourself in the face, so I don't have to find you
and do it myself >:[

--------------------------------------------------
4 one last thing 
--------------------------------------------------
double clic on "C:\Python27\Lib\site-packages\win32com\client\makepy.py"
In the pop-up choose "Microsoft Excel 12.0 Object Library (1.6)"
and click ok (bottom right)
