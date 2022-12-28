dim program1,name,Msg  '定义变量并分配内存
'Inputbox()函数进行输入
name=Inputbox("请输入你要给谁发送消息")    
Msg=Inputbox("请输入你发送消息的内容")
num=Inputbox("请输入重复发送多少次消息")
tim=Inputbox("请输入延迟多少毫秒")
'program1表示WeChat.exe的位置
'program1="D:\other-systems\WeChat\WeChat.exe"
'program1="D:\other-systems\QQ\Bin\QQ.exe"
program1=Inputbox("请输入WeChat.exe或QQ.exe的存储位置")
wscript.sleep tim
set wshshell=CreateObject("wscript.shell")  '创建Windows的shell对象打开shell窗口
'在Windows的shell窗口执行cmd.exe /c echo " & Name & " | clip.exe"将name复制到剪切板中
wshshell.Run "cmd.exe /c echo " & Name & " | clip.exe",0,True
wshshell.Run "mshta javascript:window.execScript('window.close','vbs')",0,True
'通过绝对路径的方式打开微信窗口
set oexec=wshshell.exec(program1)
wscript.sleep 500  '停留500毫秒
wshshell.sendKeys "^f" '执行ctrl + F快捷键在微信窗口唤醒搜索功能
wscript.sleep 500  
wshshell.sendKeys "^v"  '粘贴要搜索的name
wscript.sleep 500
wshshell.sendKeys "{ENTER}"  '按键盘的enter键，进入要发送消息的name联系人窗口
wscript.sleep 500
'在Windows的shell窗口执行cmd.exe /c echo " & Name & " | clip.exe"将发送的消息内容复制到剪切板中
wshshell.Run "cmd.exe /c echo " & Msg & " | clip.exe",0,True
wscript.sleep 500
for i=1 to num   '循环发送num次
wshshell.sendKeys "^v"  '粘贴要发送的msg消息内容
wscript.sleep 500
wshshell.sendKeys "{ENTER}"    '按enter键进行发送
next
wscript.quit
