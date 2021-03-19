import datetime

def get_hours_left():
    now = datetime.datetime.now()
    hours_dict = []
    for i in range(1, 25):
        hours_dict.append(i)
    return hours_dict


def get_all_min():
    mins = []
    for i in range(0, 60):
        mins.append(i)
    return mins


def get_total_min_to_shut_down(now_h, now_m, t2_h, t2_m):
    total_hours = 0
    total_mins = 0
    if t2_h <= now_h:
        total_mins = t2_m - now_m
    else:
        total_hours = t2_h - now_h
        total_mins = total_hours * 60 + (t2_m - now_m)
    return total_mins

