import sys
import time

'''下载进度条函数'''
def progress_bar(name,i,total=100,width=50):
    p=int(i* width / total)
    percent= "["+str(int(p * 100 / width)) + "%"+"]"+"["+ str(i) +"/" + str(total) +"]"
    bar_str = name + " [" + "#" * p + "_" * (width-p) + "]" + percent
    sys.stdout.write("\r%s" %bar_str)
    sys.stdout.flush()
    time.sleep(0.1)

total = 10
for i in range(1,total+1):
    progress_bar('demo1:',i,total,10)

print('\n')
total = 10
for i in range(1,total+1):
    progress_bar('demo2:',i,total,20)

print('\n')
total = 20
for i in range(1,total+1):
    progress_bar('demo3:',i,total,20)



