import os

def url_in_one_click(fileName):
    fileName = fileName + '.txt'
    file = open(fileName)

    for site in file:
        os.system('start ' + site)

    file.close()