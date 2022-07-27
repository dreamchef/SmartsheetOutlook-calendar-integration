pip3 install smartsheet
pip3 install os

echo import os > changetype.py
echo user = os.getlogin() >> changetype.py
echo print(user) >> changetype.py
echo val = "C:/Users/" + user + "/AppData/Local/Packages/PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0/LocalCache/local-packages/Python310/site-packages/smartsheet/types.py" >> changetype.py

echo f = open(val,'r') >> changetype.py
echo lines = f.readlines() >> changetype.py
echo f.close() >> changetype.py

echo lines[17] = "import collections.abc\n" >> changetype.py
echo lines[28] = "class TypedList(collections.abc.MutableSequence):\n" >> changetype.py

echo f = open(val,'w') >> changetype.py
echo f.writelines(lines) >> changetype.py

python3 changetype.py

del changetype.py

pause