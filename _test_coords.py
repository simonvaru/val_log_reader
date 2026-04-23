import re
linea = '"latitude":-65.19335174560547,"lineNumber":768,"longitude":-26.83445167541504,'
m58 = re.search(r'latitude.{0,3}(-?[0-9]+\.[0-9]+)', linea)
m59 = re.search(r'longitude.{0,3}(-?[0-9]+\.[0-9]+)', linea)
print('ID 58:', m58.group(1) if m58 else 'NO MATCH')
print('ID 59:', m59.group(1) if m59 else 'NO MATCH')
