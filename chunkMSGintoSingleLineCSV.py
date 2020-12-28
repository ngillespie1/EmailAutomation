import os
# -*- coding: utf-8 -*-
import win32com.client
import re
import codecs
import nltk
from nltk.corpus import stopwords
stoplist = stopwords.words('english')


inputFolder = r"C:\test\Finalized training data\phishing"
count = 1

#with codecs.open(r'C:\test\spam\spam.csv', mode='a', encoding='utf-8') as csvfile:
for file in os.listdir(inputFolder):
    if file.endswith(".msg"):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        print(file)
        filename = str(file)
        filePath = inputFolder  + '\\' + file         
        msg = outlook.OpenSharedItem(filePath)             
        body = msg.body
    #   print msg.ReceivedTime
    #   print msg.Subject
    #   print msg.Body
    #   print msg.To
    #   To = msg.To
    #   print msg.Size
    #   print msg.Attachments
        body = body.replace('\n', '').replace('\r','')
        body = re.sub(r'(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:\'".,<>?«»“”‘’]))', '', body)
        body = (str(filename) + ' ' + body)
    #    body = [word for word in body.split() if word not in stoplist] removing stop words
    #    body.split()[:500] first 500 words
        print(count)
        count +=1
 
        with codecs.open(r'C:\test\Finalized training data\phishing.csv', 'a', encoding='utf-8') as f:
            f.write(str(body) + "\n")

del outlook, msg
