import re

with open("codes.txt") as codes_raw:
    codes_raw = codes_raw.readlines()

    new_codes = []
    for code in codes_raw:
        if re.search("[A-Z]\d+", code) != None: new_codes.append(code.replace("\n", ""))
    new_codes = "', '".join(new_codes)

with open("test.txt", "w") as test:
    test.write(new_codes)