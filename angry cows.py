def has(a,b,c):
    return(bool(list(filter(lambda x:b<=x<=c,a))))
n=int(input())
hloc=[]
for i in range(n):
    x=int(input())
    hloc.append(x)
hloc.sort()
mx=0
for i in hloc:
    a=hloc[:]
    flag=1
    l=i
    r=i
    cnt=1
    prel=l
    prer=r
    l-=flag
    r+=flag
    ld=False
    rd=False
    while True:
        mnl=l
        mxl=r
        flag+=1
        if not ld:
            if l<0:
                l=0
                for j in range(l,prel):
                    if j in a:
                        cnt+=1
                        del a[a.index(j)]
            elif has(a,l,prel-1)==False:
                ld=True
            else:
                mnl=1e18
                for j in range(prel-1,l-1,-1):
                    if j in a:
                        cnt+=1
                        mnl=min(mnl,j)
                        del a[a.index(j)]
                if mnl==1e18:
                    mnl=l
        if not rd:
            if r>1000000000:
                r=1000000000
                for j in range(prer,r+1):
                    if j in a:
                        cnt+=1
                        del a[a.index(j)]
                rd=True
            elif has(a,prer,r)==False:
                rd=True
            else:
                mxl=-1
                for j in range(prer+1,r+1):
                    if j in a:
                        cnt+=1
                        mxl=max(mxl,j)
                        del a[a.index(j)]
                if mxl==-1:
                    mxl=r
        if ld and rd:
            break
        l=max(mnl,l)
        r=min(mxl,r)
        prel=l
        prer=r
        l-=flag
        r+=flag
    mx=max(mx,cnt)
print(mx)