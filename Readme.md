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
pip install python-dotenv
pip install pylint
pip install pillow-heif

# pack program to exe file
pip install pyinstaller

pyinstaller --onefile main.py -n image-ppt-generator
```


## reference 
1. [tutorial](https://www.tutorialspoint.com/how-to-create-powerpoint-files-using-python)
2. [tutorial-1](https://zhuanlan.zhihu.com/p/291729098)
3. [tutorial-2](https://www.geeksforgeeks.org/creating-and-updating-powerpoint-presentations-in-python-using-python-pptx/)
4. [pyinstall-issues](https://stackoverflow.com/questions/404744/determining-application-path-in-a-python-exe-generated-by-pyinstaller)
5. [python-ppt lib](https://python-pptx.readthedocs.io/en/latest/index.html)
6. [python ppt font size](https://blog.csdn.net/weixin_28363123/article/details/113511188)