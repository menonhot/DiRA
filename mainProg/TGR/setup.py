import os
thePip = list(os.popen('where pip'))
pip = thePip[0].strip()
os.system('{} install -r requirements.txt'.format(pip))
