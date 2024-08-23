import os

for i in os.listdir('./Type 2/ENG/'):
  os.remove('./Type 2/ENG/' + i)
for i in os.listdir('./Type 2/RUS/'):
  os.remove('./Type 2/RUS/' + i)
for i in os.listdir('./Type 1/ENG/'):
  os.remove('./Type 1/ENG/' + i)
for i in os.listdir('./Type 1/RUS/'):
  os.remove('./Type 1/RUS/' + i)
