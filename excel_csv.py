
import pandas as pd
import os 


from_directory = 'C:\\Users\\valms\\Desktop\\NOTAS-30-05-2022'
destiny_directory = 'C:\\Users\\valms\\Desktop\\NOTAS-30-05-2022'

for filename in os.listdir(from_directory):
    file = os.path.join(from_directory,filename)
    if os.path.isfile(file):
        read_file = pd.read_excel(file)
        read_file.to_csv(f'{destiny_directory}\{os.path.splitext(filename)[0]}.csv', index=None, header=True)


