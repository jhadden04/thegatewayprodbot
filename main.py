import praw
from openpyxl import load_workbook
import random
import time

workbook = load_workbook("hello_world.xlsx")  # opens up the excel file you are using
# it must be a file that is in the same folder as the script
sheet = workbook.active
reddit = praw.Reddit(client_id='MGSGhbDCXWHAmA',
                     client_secret='u3JgbIqbY7kpsGCCaqerYaopm8iicA', 
                     user_agent='bot',
                     username='takingitdaybyday77',   
                     password='MitiKape42$')
n = 1
usernames = []
while True:
    x = sheet[f'A{n}']
    if x.value is None:
        break
    usernames.append(x.value)
    n += 1  # this just gets the usernames from the excel and puts them into a list

usedusernames = []  # used usernames go here


def message():
    title = "this is the title"  # title text, you should change this
    text = ["this is the first text option", "second", "third", "etc"]  # permutations of the text,
    # you can change this infinitely
    q = 0

    for i in range(len(usernames)):
        w = random.randint(1, (len(text) - 1))
        reddit.redditor(usernames[q]).message(title, text[w])  # messages all the people from the excel file

        if usernames[q] in usedusernames:  # means they can't be messaged twice
            continue
        print(f'you just messaged u/{usernames[q]}')  # notifies you if a message has been sent
        usedusernames.append(usernames[q])
        time.sleep(40)  # time interval between messages, measured in seconds, you can change this
        q += 1


message()  # calls the message function
