import os
thePip = list(os.popen('where pip'))
if len(thePip) == 0 :
    thePip = list(os.popen('which pip'))
thePip = list(os.popen('which pip'))
print(thePip)
pip = thePip[0].strip()
os.system('{} install -r requirements.txt'.format(pip))
