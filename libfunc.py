import numpy as np

def checker(arg1, arg2):
    a = arg1.shape
    b = arg2.shape
    if a[1]==b[0]:
        return True
    else:
        return False

def fail_safe(arg):
    for i in range(arg.shape[0]):
        if arg[i,i] ==0:
            return False
    return True


def multiplier(arg1, arg2):
    shaper1 = arg1.shape
    shaper2 = arg2.shape
    ary1 = []
    for row in range(shaper1[0]):
        for column in range(shaper2[1]):
            hold = []
            for i in range(shaper1[1]):
                val = arg1[row,i] * arg2[i,column]
                hold.append(val)
            ary1.append(sum(hold))
    
    ary2 = np.array(ary1)
    ary3 = ary2.reshape(shaper1[0],shaper2[1])
    return np.matrix(ary3)



def solver(arg):
    shaper = arg.shape
    rows = shaper[0]
    cols = shaper[1] -1
    for j in range(cols):
        i = rows - 1
        while i>j:
            if fail_safe(arg):
                arg[i]=(-((arg[i,j]/arg[j,j]))*arg[j]) + arg[i]
                print(arg)
                i = i - 1
            else:
                print("\n\nOne of the diagonal values turns to zero.\nNot suitable for this method.\n\n")
                return None
    j = cols - 1
    while j>0:
        i = 0
        while i<j:
            if fail_safe(arg):
                arg[i]=(-((arg[i,j]/arg[j,j]))*arg[j]) + arg[i]
                print(arg)
                i = i + 1
            else:
                print("\n\nOne of the diagonal values turns to zero.\nNot suitable for this method.\n\n")
                return None
        j = j - 1
    for i in range(rows):
        arg[i] = arg[i]/arg[i,i]
    return arg
