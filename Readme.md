# ppt image processor
simple image insertion program for ppt files



``` bash
#create virtual env for development
pip3 install virtualenv 


virtualenv venv 

source venv/bin/activate

deactivate

#dependancy
pip install python-pptx
pip install -U autopep8

# pack program to exe file
pip install pyinstaller

pyinstaller --onefile main.py
```


## reference 
1. [tutorial](https://www.tutorialspoint.com/how-to-create-powerpoint-files-using-python)
2. [tutorial-1](https://zhuanlan.zhihu.com/p/291729098)
3. [pyinstall-issues](https://stackoverflow.com/questions/404744/determining-application-path-in-a-python-exe-generated-by-pyinstaller)