import os
user = os.getlogin()
print(user)
val = "C:/Users/" + user + "/AppData/Local/Packages/PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0/LocalCache/local-packages/Python310/site-packages/smartsheet/types.py"

f = open(val,'r')
lines = f.readlines()
f.close()

lines[17] = "import collections.abc\n"
lines[28] = "class TypedList(collections.abc.MutableSequence):"

f = open(val,'w')
f.writelines(lines)