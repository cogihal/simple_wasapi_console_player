"""
complete folder/file name using Tab key when input() in Python
"""

# Refer: https://stackoverflow.com/questions/5637124/tab-completion-in-pythons-raw-input

import os
import pathlib
import rlcompleter

def complete_path(text, state):
    if os.path.exists(text) and os.path.isfile(text):
        return None

    incomplete_path = pathlib.Path(text)

    if incomplete_path.is_dir():
        completions = [p.as_posix() for p in incomplete_path.iterdir()]

    elif incomplete_path.exists():
        completions = [incomplete_path]

    else:
        exists_parts = pathlib.Path('.')
        for part in incomplete_path.parts:
            test_next_part = exists_parts / part
            if test_next_part.exists():
                exists_parts = test_next_part

        completions = []
        for p in exists_parts.iterdir():
            p_str = p.as_posix()
            if p_str.startswith(text):
                completions.append(p_str)

    return completions[state]


if __name__ == '__main__':
    rlcompleter.readline.parse_and_bind("tab: complete")
    rlcompleter.readline.set_completer(complete_path)

    f = input('input filename: ')
    print(f'File: {f}')


