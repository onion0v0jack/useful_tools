import string

def Getletterfromindex(num):
    if num <= 0:
        print('The number should greater then 0.')
    # produces a string from numbers so

    # 1->A
    # 2->B
    # 26->Z
    # 27->AA
    # 28->AB
    # 52->AZ
    # 53->BA
    # 54->BB

    num2alphadict = dict(zip(range(1, 27), string.ascii_lowercase))
    outval = ''
    numloops = (num - 1) //26

    if numloops > 0:
        outval += Getletterfromindex(numloops) # 遞迴

    remainder = num % 26
    if remainder > 0:
        outval += num2alphadict[remainder]
    else:
        outval += 'z'
    return outval.upper()  # 最後轉大寫

# 測試
# for i in range(26 * 26 + 26 + 26):
#     print('{}    {}'.format(i + 1, Getletterfromindex(i + 1)))