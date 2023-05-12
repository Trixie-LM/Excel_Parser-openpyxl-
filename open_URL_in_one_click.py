import os

def url_in_one_click(fileName):
    try:
        fileName = fileName + '.txt'
        file = open(fileName)
    except FileNotFoundError:
        fileName = 'Расхождения в отчетах/' + fileName
        file = open(fileName)



    for site in file:
        os.system('start ' + site)

    file.close()

    