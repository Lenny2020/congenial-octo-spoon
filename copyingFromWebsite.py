import re
import pyperclip

text = pyperclip.paste()


tableRegex = re.compile(r'(\d{6})\t([\w, ]*)\t(\.*\d)\t(\d+/\d+/\d+|.NULL.)\t(\w+)\t(\w+)\t([\d, ]*)\t([\d, ]*)'
                        r'\t([\d, ]*)\t([\d, ]*)\t([\d, ]*)\t([\d, ]*)\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)'
                        r'\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)'
                        r'\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)\t([\d. ]*)')

table = tableRegex.findall(text)
print(table[0][6])

# 0. Consult No
# 1. Name
# 2. Level
# 6. Career Sales
# 11 Personal Sales
