import math

a=''
b=''

pages=input('请输入文件页数：')
pages=int(pages)

n=input('请输入每张A4纸的打印页数：')
n=int(n)

if pages%n!=0:
    pages=pages+4-pages%4
z=math.ceil(pages/n/2)

print('共计打印页数为：',pages)
print('共计打印A4纸张数为：',z)

print('打印次序为：')

if pages%(n)==0:
    for i in range(z):
        print('第',i+1,'张A4纸正面为：',pages-n*i,',',1+n*i,',',pages-n*i-2,',',1+n*i+2,',',sep='')
        a=a+str(pages-n*i)+','+str(1+n*i)+','+str(pages-n*i-2)+','+str(1+n*i+2)+','
    for i in range(z):
        print('第',i+1,'张A4纸反面为：',n*i+2,',',pages-1-n*i,',',n*i+2+2,',',pages-1-n*i-2,',',sep='')
        b=b+str(n*i+2)+','+str(pages-1-n*i)+','+str(n*i+2+2)+','+str(pages-1-n*i-2)+','

else:
    for i in range(z-1):
        print('*第',i+1,'张A4纸正面为：',pages-n*i,',',1+n*i,',',pages-n*i-2,',',1+n*i+2,',',sep='')
        a=a+str(pages-n*i)+','+str(1+n*i)+','+str(pages-n*i-2)+','+str(1+n*i+2)+','
    for i in range(z-1):
        print('*第',i+1,'张A4纸反面为：',n*i+2,',',pages-1-n*i,',',n*i+2+2,',',pages-1-n*i-2,',',sep='')
        b=b+str(n*i+2)+','+str(pages-1-n*i)+','+str(n*i+2+2)+','+str(pages-1-n*i-2)+','
    print('*第',z,'张A4纸正面为：',pages-n*2*(n-1),',',1+n*2*(n-1),sep='')
    print('*第',z,'张A4纸反面为：',n*2*(n-1)+2,',',pages-1-n*2*(n-1),sep='')
    a=a+str(pages-n*(n-1))+','+str(1+n*(n-1))+','
    b=b+str(n*(n-1)+2)+','+str(pages-1-n*(n-1))+','

print('正面拼接：',a[0:-1])
print('反面拼接：',b[0:-1])

print('————————————————————————')

