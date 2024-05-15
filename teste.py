from app.backups import BACK_UP
from registStatus import getDrsStatus



if __name__ == '__main__':
    data = getDrsStatus()
    print(data)
