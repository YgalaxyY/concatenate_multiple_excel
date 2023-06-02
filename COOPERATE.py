import os
import pandas as pd

path = os.getcwd()
files = os.listdir(path)
print(files)

last=[]

for i in files:
     if i[-3:]=='lsx':
          last.append(i)

