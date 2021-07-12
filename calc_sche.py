'''
write todo.txt
monitor finish.txt
'''


import os
import time
import sys
import datetime


def generate_file_content(track_number, assigned_node, backup_path, modifiy_time=None):
    if not modifiy_time:
        modifiy_time = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    content = ['Track Number = {tn}\n'.format(tn=track_number),
               'Modify Time = {mt}\n'.format(mt=modifiy_time),
               'Status = 0\n',
               'Current Node = calc\n',
               'iRun user = jenkins\n',
               'Is Ideas = 1\n',
               'Case Function = 7\n',
               'Select Type = 0\n',
               'Is Auto Close = 0\n',
               'Is Auto Calc = 1\n',
               'Suffix = ts{sf}\n'.format(sf=modifiy_time),
               'Comment = ts{cm}\n'.format(cm=modifiy_time),
               'Assigned Node = {an}\n'.format(an=assigned_node),
               'Lib Path = /data58/InputData-New\n',
               'Backup Path = {bp}\n[END]'.format(bp=backup_path)]
    return content


def monitor_finish(filepath):
    while not os.path.exists(filepath):
        time.sleep(10)
    with open(filepath, 'r') as f:
        for line in f:
            # if there's no line or no line with Is Calc Finished, it exit normally with no hit
            if 'Is Calc Finished' in line:
                finish_code = int(line.partition('=')[-1].strip(' '))
                if not finish_code:
                    print('Failed: the calc finish code is {fc}'.format(
                        fc=finish_code))
                    sys.exit(1)
                else:
                    print('libcalc calc finished!')


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print('Usage: python file.py <task_path> <track_number> <assigned_node> <backup_path> <modify_time>')
        sys.exit(1)
    task_path = sys.argv[1]
    track_number = sys.argv[2]
    idle_machine = sys.argv[3]
    backup_path = sys.argv[4]
    modifiy_time = sys.argv[5]
    todo_filepath = task_path + os.path.sep + 'todo.txt'
    print(todo_filepath)
    with open(todo_filepath, 'w') as f:
        f.writelines(generate_file_content(track_number=track_number,
                                           assigned_node=idle_machine, backup_path=backup_path, modifiy_time=modifiy_time))
    finish_filepath = task_path + os.path.sep + 'finish.txt'
    monitor_finish(finish_filepath)
