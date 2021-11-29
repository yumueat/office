sum=0
n=int(input("请输入n的值"))#估计是输入n，代表有n组输入
for i in range(n):#利用for循环进行多组输入
    x=int(input("请输入x的值"))#输入
    if (x%8//2)==2:#根据if里面的规则判断这个数是否符合
        sum=sum+x#对符合的数求和
print(sum)#输出总和