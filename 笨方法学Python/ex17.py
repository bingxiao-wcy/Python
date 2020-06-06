from sys import argv
from os.path import exists

script, from_file, to_file = argv

print("copying from %s to %s" % (from_file,to_file))

#we chould do these two on one line too, how?

raw_input = open(from_file) #将第一个文件的内容保存到raw_input变量中

indata = raw_input.read() #将变量raw_input的内容读到indata中

print("The input file is %d bytes long" % len(indata)) #输出indata的字节长度

print("Does the output file exist? %r " % exists(to_file))#判断to_file文件是否存在
print("Ready, hit RETURN to continue, CTRL-C to abort.")
input()

output = open(to_file,'w') #以写入方式打开文件，文件标识符为output
output.write(indata)#将变量indata中的内容写入文件中

print("Alright, all done")

output.close()#关闭文件
raw_input.close()#关闭文件
