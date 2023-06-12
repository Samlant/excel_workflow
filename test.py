import subprocess
import os

path = os.getcwd()
path = os.path.join(path, "test2.py")

data = path

subprocess.run(["python", path], input=data, encoding="utf-8")
