import math
import numpy as np
def quadratic(a, b, c):
    x1 = ((-1*b)+math.sqrt(b**2-4*a*c))/(2*a)
    x2 = (-1*b)-math.sqrt(b**2-4*a*c)
    return x1,x2

def zero_check(a):
    for i in range(len(a)):
        if a[i]>=0.005:
            return False
    return True

def carry_over(bal,mental_block_right, mental_block_left,joints):
    co = list(range(len(bal)))
    if not mental_block_right and not mental_block_left:
        i = 1
        co[1] = 0.5 * bal[0]
        co[len(bal) - 2] = 0.5 * bal[len(bal) - 1]
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            beforep = before - 1
            afterp = after + 1
            co[beforep] = 0.5 * bal[before]
            co[afterp] = 0.5 * bal[after]
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        beforep = before - 1
        afterp = after + 1
        co[beforep] = 0.5 * bal[before]
        co[afterp] = 0.5 * bal[after]
        co[0]=0.5*bal[1]
        co[len(bal)-1]=0.5*bal[len(bal)-2]
        return co
    elif mental_block_left and not mental_block_right:
        co[0]=0.0
        co[1] = 0.5 * bal[0]
        co[len(bal) - 2] = 0.5 * bal[len(bal) - 1]
        i = 1
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            beforep = before - 1
            afterp = after + 1
            co[beforep] = 0.5 * bal[before]
            co[afterp] = 0.5 * bal[after]
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        beforep = before - 1
        afterp = after + 1
        co[beforep] = 0.5 * bal[before]
        co[afterp] = 0.5 * bal[after]
        co[0] = 0.0
        co[len(bal) - 1] = 0.5 * bal[len(bal) - 2]
        return co
    elif not mental_block_left and mental_block_right:
        co[len(bal)-1]=0.0
        co[1] = 0.5 * bal[0]
        co[len(bal) - 2] = 0.5 * bal[len(bal) - 1]
        i = 1
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            beforep = before - 1
            afterp = after + 1
            co[beforep] = 0.5 * bal[before]
            co[afterp] = 0.5 * bal[after]
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        beforep = before - 1
        afterp = after + 1
        co[beforep] = 0.5 * bal[before]
        co[afterp] = 0.5 * bal[after]
        co[0] = 0.5 * bal[1]
        co[len(bal) - 1] = 0.0
        return co
    else:
        i = 1
        co[1]=0.5*bal[0]
        co[len(bal)-2]=0.5*bal[len(bal)-1]
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            beforep = before - 1
            afterp = after + 1
            co[beforep] = 0.5 * bal[before]
            co[afterp] = 0.5 * bal[after]
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        beforep = before - 1
        afterp = after + 1
        co[beforep] = 0.5 * bal[before]
        co[afterp] = 0.5 * bal[after]
        co[0] = 0.0
        co[len(bal) - 1] = 0.0
        return co

def balance(co,joints,df,mental_block_right, mental_block_left):
    if not mental_block_right and not mental_block_left:
        bal = []
        bal.append(co[0] * 0)
        i = 1
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            bal.append((-1 * (co[before] + co[after])) * df[before])
            bal.append((-1 * (co[before] + co[after])) * df[after])
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        bal.append((-1.0 * (co[before] + co[after])) * df[before])
        bal.append((-1.0 * (co[before] + co[after])) * df[after])
        bal.append(co[len(co) - 1] * 0)
        return bal
    elif mental_block_left and not mental_block_right:
        bal = []
        bal.append(co[0] * -1)
        i = 1
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            bal.append((-1 * (co[before] + co[after])) * df[before])
            bal.append((-1 * (co[before] + co[after])) * df[after])
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        bal.append((-1.0 * (co[before] + co[after])) * df[before])
        bal.append((-1.0 * (co[before] + co[after])) * df[after])
        bal.append(co[len(co) - 1] * 0)
        return bal
    elif not mental_block_left and mental_block_right:
        bal = []
        bal.append(co[0] * 0)
        i = 1
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            bal.append((-1 * (co[before] + co[after])) * df[before])
            bal.append((-1 * (co[before] + co[after])) * df[after])
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        bal.append((-1.0 * (co[before] + co[after])) * df[before])
        bal.append((-1.0 * (co[before] + co[after])) * df[after])
        bal.append(co[len(co) - 1] * -1)
        return bal
    else:
        bal = []
        bal.append(co[0] * -1)
        i = 1
        while i < len(joints) - 2:
            before = (2 * i) - 1
            after = (2 * i)
            bal.append((-1.0 * (co[before] + co[after])) * df[before])
            bal.append((-1.0 * (co[before] + co[after])) * df[after])
            i = i + 1
        before = (2 * i) - 1
        after = (2 * i)
        bal.append((-1.0 * (co[before] + co[after])) * df[before])
        bal.append((-1.0 * (co[before] + co[after])) * df[after])
        bal.append(co[len(co) - 1] * -1)
        return bal

def moment_dist(w,end_conditions,spans,sections):
    mental_block_left = False
    mental_block_right = False
    fems=[]
    for i in range(len(spans)):
        l=spans[i]
        fem = -1*((w*l**2)/12)
        fems.append(fem)
        fem = ((w*l**2)/12)
        fems.append(fem)
    if len(end_conditions)!=len(spans)+1 or len(end_conditions)!=len(sections)+1 or len(sections)!=len(spans):
        print("\n\nError. Bad inputs.")
        return
    k=[]
    for i in range(len(end_conditions)-1):
        if end_conditions[i]=="pinned" or end_conditions[i+1]=="pinned":
            ei = sections[i]
            l = spans[i]
            k.append((3*ei)/l)
        else:
            ei = sections[i]
            l = spans[i]
            k.append((4 * ei) / l)
    df=[]
    for i in range(len(end_conditions)-1):
        if i==0:
            if end_conditions[i]=="fixed":
                df.append(0)
            elif end_conditions[i]=="pinned":
                df.append(1)
                mental_block_left = True
            else:
                print("\n\nError. Bad inputs.")
                return
        elif i<len(k)-1:
            kt = k[i-1]+k[i]
            rel_k = k[i-1]/kt
            df.append(rel_k)
            rel_k = k[i]/kt
            df.append(rel_k)
        else:
            kt = k[i - 1] + k[i]
            rel_k = k[i - 1] / kt
            df.append(rel_k)
            rel_k = k[i] / kt
            df.append(rel_k)
            if end_conditions[i+1]=="fixed":
                df.append(0)
            elif end_conditions[i+1]=="pinned":
                df.append(1)
                mental_block_right = True
            else:
                print("\n\nError. Bad inputs.")
                return
    joints = list(range(len(end_conditions)))
    members = range(len(spans))
    BAL=[]
    CO = []
    n=1
    bal = balance(fems,joints,df,mental_block_right, mental_block_left)
    co = carry_over(bal,mental_block_right,mental_block_left,joints)
    for i in range(len(bal)):
        BAL.append(bal[i])
    for i in range(len(co)):
        CO.append(co[i])
    while not zero_check(bal):
        bal = balance(co, joints, df, mental_block_right, mental_block_left)
        co = carry_over(bal, mental_block_right, mental_block_left, joints)
        for i in range(len(bal)):
            BAL.append(bal[i])
        for i in range(len(co)):
            CO.append(co[i])
        n = n + 1
    CO = np.array(CO)
    BAL = np.array(BAL)
    BAL = BAL.reshape(n, len(bal))
    CO = CO.reshape(n, len(co))
    end_moments = []
    for i in range(len(bal)):
        value = fems[i] + float(sum(BAL[:, i])) + float(sum(CO[:, i]))
        end_moments.append(value)
    # Convert flat list of end moments [Mab, Mba, Mbc, Mcb, ...]
    # into list of tuples [(Mab, Mba), (Mbc, Mcb), ...]
    end_moments_pairs = [(end_moments[i], end_moments[i + 1]) for i in range(0, len(end_moments), 2)]
    return end_moments_pairs

def example():
    w = 4.0
    end_conditions = ["pinned", "continuous", "continuous", "pinned"]
    spans = [18.0, 20.0, 18.0]
    sections = [1.0, 1.0, 1.0]
    val = moment_dist(w, end_conditions, spans, sections)
    return val


if __name__ == '__main__':
    print('Running example...')
    from pprint import pprint
    pprint(example())