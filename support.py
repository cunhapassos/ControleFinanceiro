import os


def open_file(filename):
    ''' Load a file from the fixtures directory. '''
    path = filename
    if ('tests' in os.listdir('.')):
        path = 'tests/' + path
    return open(path, mode='rb')