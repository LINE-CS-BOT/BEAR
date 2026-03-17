@echo off
cd /d C:\Users\bear\Desktop\code\line-cs-bot
C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe scripts/sync_cust_from_web.py --auto >> data\sync_customers.log 2>&1
