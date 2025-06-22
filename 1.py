# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token':1.AVUAGvsxClxDXkyk_6JPR75l_Xn8g-6caT5DucpQ6--ygW6_AONVAA.AgABAwEAAABVrSpeuWamRam2jAF1XRQEAwDs_wUA9P_Q66iLwiodRepU7eiObeSNvRRTN5iqrc2uygr4GzOFlDhSn2QWyIAgsfrR6sH_cCfirOsI3WJUiLaKepo13Q2Zp2EdfR6fm2bB13uYPf7gCxEddb-jgcSJq7P1HKdKHiGzD1nZ5oT1vJ7QC0O6piXbZG38MTVIoaUdRenpEcHSmNgaUmqXkoK5UjDLp5EpQ_5Ux-BOULeorEZ8eq-jF_gFfUN8xtyq4P7VrdiaV0bBowVJIVwSO-7N4NifrMBuCeWPL2vVC3v6KntzHkGmlMoRHYIMOAcSbsfpIaTHB1E1VGIFSRUJLxkyEdi88OVnHX9uwxdB8VTY-jiJHv5ANOKfFkuZ8w531bajTfz554cxT9Nz6aEZBszz6OsOnmaIMAFF28-R-gdYcLG9FNRcktTNizDOGdXy6cTNJWmlFl-MgEzRVLX-5cypRL36IQkI5uBnoN7NAhnTEY7i5eesaYjSbgzgl5wkwZVt5IxAut561nFWgluKGogA8RUM2N2Ind_J2CGcAyhHAs9kD00ZixsFLjIyOt_gsqXgV598W2j6a8f6rtulloa9fVkOZaqMsdDjU4gz1MjKw_p5xLX5cb0LXqbJnhVC9zVhCPIlf8AWA8VXl7lzj_f0rjBxbCpxHKcNXf6XH7_xJpmZfCCM-6DnSOvR_FfE4S0NiiMYC_KI9X7jSelQiA5FPseaJuOaBQr8ZJcqf2BAQtpvxHiFNOydD7Hmzac4C5Z2Rnwk-VPMOGtjx847yZXag58VB-aTdYG19dXIJqCiIp2oQrZlWdj4aYMuWRRM7zE-FUQm2zWBW5lnLpwToQTj9r_6_SPBpn2hPgvFI3CkP7aZ44OtSimUS5hyX_rOkNbA5zT2Qzupg2SwIz-kr6nitboS2zYEVxio3f5spwFI1XRRso_j9Ee1wrfXjMb9oDOBBADEcg7BYR_jvETgL4Dq6pmJQIeBi-D1WgrH9B0BBF5QNVyA-hsPo8W9cHeKrlfU34dlHInmJFeJS4wDaTUIuquW21JyxSQHrA,
          'client_id':ee83fc79-699c-433e-b9ca-50ebefb2816e ,
          'client_secret':iVc8Q~2EcpzACcLvJ~WrvkGUCQie.XTSoTJf-dxF,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 :', localtime)
    except:
        print("pass")
        pass
for _ in range(3):
    main()
