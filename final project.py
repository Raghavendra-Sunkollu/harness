import pandas as pd
import smtplib as s
import numpy
import boto3
df=pd.read_excel('test.xlsx',sheet_name=0)
print(df)
ob=s.SMTP('smtp.gmail.com',587)
ob.starttls()
ob.login('snraghavendra2000@gmail.com','mfcrfmjpcoxdkcxy')
def sem1(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[4,4,3,3,3,1,1,1]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    print(k)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'st sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)
def sem2(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[4,4,3,3,3,1,1,1]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'nd sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)
def sem3(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[3,4,3,3,3,3,2,2,1]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'rd sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)
def sem4(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[3,4,3,3,3,3,2,2,1]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'th sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)
def sem5(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[3,4,4,3,3,3,2,2,1]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'th sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)
def sem6(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[4,4,4,3,3,2,2,2]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'th sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)

def sem7(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[4,4,3,3,3,2,1]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'th sem result is'+str(res)+' SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)
def sem8(k,sem,usn,name,email):
    #print(k,sem,usn,name,email)
    l=[3,3,8,1,3]
    class1=''
    c=sum(l)
    for i in range(len(k)):
        k[i]=(k[i]//10)+1
    prod=numpy.multiply(k,l)
    tot=sum(prod)
    res=tot/c
    if res>=7.75:
        class1='Distinction'
        print(class1)
    if res>=6 and res<7.75:
        class1='First class'
        print(class1)
    if res>=5 and res<6:
        class1='First class'
        print(class1)
    message='Hi '+name+'Congratulations Your '+str(sem)+'th sem result is '+str(res)+'SGPA and your class is '+class1+' This is an automated email kindly do not reply '
    ob.sendmail('snraghavendra2000@gmail.com',email,message)
    # ---------DYNAMODB-----------------
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table('project')
    response=table.put_item(
            
    Item={
            'usn':usn,
            'name':name,
            'email':email,
            'sem':str(sem),
            "Sgpa":str(res),
            'Class':class1,
            
            }
        )
    # ---------DYNAMODB-----------------
    print(res)

        
    
for i in range(0,len(df)):
    list=df.iloc[i,:].tolist()
    #print(list)
    k=list[4:]
    sem=list[3]
    usn=list[0]
    name=list[1]
    email=list[2]
    if(sem==1):
        k=list[4:12]
        sem1(k,sem,usn,name,email)
        #print(k)
    if(sem==2):
        k=list[4:12]
        sem2(k,sem,usn,name,email)
    if(sem==3):
        k=list[4:13]
        sem3(k,sem,usn,name,email)
    if(sem==4):
        k=list[4:13]
        sem4(k,sem,usn,name,email)
    if(sem==5):
        k=list[4:13]
        sem5(k,sem,usn,name,email)
    if(sem==6):
        k=list[4:12]
        sem6(k,sem,usn,name,email)
    if(sem==7):
        k=list[4:11]
        sem7(k,sem,usn,name,email)
    if(sem==8):
        k=list[4:9]
        sem8(k,sem,usn,name,email)
        
    
    

    