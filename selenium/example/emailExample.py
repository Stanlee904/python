from datetime import datetime
import win32com.client


curTime = datetime.today()
year = curTime.strftime("%Y")
month = curTime.strftime("%m")
day = curTime.strftime("%d")

subjectList = []

outlook = win32com.client.Dispatch("Outlook.Application")

rxOutlook = outlook.GetNamespace("MAPI")

inbox = rxOutlook.GetDefaultFolder(6).Folders("rmsend")

messages = inbox.Items
# message = messages.GetLast()

#날짜 값 가져오기
# print("okay" if str(message.ReceivedTime)[0:10] == year+"-"+month+"-"+day else "not okay")

for message in messages:
    if (str(message.ReceivedTime)[0:10] == year+"-"+month+"-"+day) and "[ETN]" in message.Subject:
        subjectList.append(message.Subject)

print(subjectList)
print(len(subjectList))




# A수신함에서 B수신함으로 옮기기 
# rxOutlook = outlook.GetNamespace("MAPI")

# 메일 생성
# newMail =  outlook.CreateItem(0)

# 3: 휴지통 4: 보낼편지함 5: 보낸편지함 6: 받은 편지함 GetDefaultFolder
# inbox = rxOutlook.GetDefaultFolder(6).Folders("citi")
# donebox = rxOutlook.GetDefaultFolder(6).Folders("bndsnd")

# messages = inbox.Items
# message = messages.GetLast()

# print(message.Subject)

# message.Move(donebox)



# for mail in messages:
#     print(mail)


# donebox = rxOutlook.GetDefaultFolder(6).Folders("npsnice")


# message = messages.GetLast()

# message.Move(donebox)



# inbox = outlook.GetDefaultFolder(6).Folders.Item("Your_Folder_Name")

# print("전체 읽은 개수" + str(inbox.items.count))
# print("전체 읽은 개수" + str(inbox.UnReadItemCount))
# print("읽은 메일 개수  : " + str(inbox.items.count - inbox.UnReadItemCount))


# newMail.To = "dkxoxo12@naver.com"
# newMail.subject = "python 메일 테스트"

# newMail.HTMLBody = "이 메일은 python 메일 테스트 발신 전용 메일입니다. "


# attachment = r'D:\새 폴더\통합 문서1.xlsx'

# newMail.Attachments.Add(attachment)

# newMail.Send()



