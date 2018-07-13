# coding: utf-8

"""Notes リンクを URI に変換する

クリップボードの内容を取得し、ノーツリンクをブラウザで開ける URI に変換する。

Require:
    pyperclip

Usage:
    Notes の文書リンクをコピーした後、本プログラムを実行してください。
    URI 変換後の文字列がクリップボードにセットされます。
"""

import os
import sys
import re

import pyperclip as clip

REDIRECT = 'https://rfgricoh.sharepoint.com/sites/rfgportal/SitePages/redirect.aspx?'
# REDIRECT = 'http://10.63.100.25/notes/?uri='


def main():
    # クリップボード値を取得
    text = clip.paste()

    # Notes リンク情報を取得する
    pathlist = []
    m = re.search(r'<NDL>(.+)', text)
    if m:
        # TITLE
        title = text.splitlines()[0].strip()
        title = title.replace('<NDL>', '')
        title = title.replace('[', '(').replace(']', ')')  # [] → ()
        # SERVER
        m = re.search(r'<HINT>CN=(\w+)(/?[A-Z]+=\w+)*</HINT>', text)
        if m:
            server = m.group(1)
            pathlist.append(server)
        # REPLICA
        m = re.search(r'<REPLICA (\w+):(\w+)>', text)
        if m:
            pathlist.append(''.join(m.groups()))
        # VIEW
        m = re.search(r'<VIEW OF(\w+):(\w+)-ON(\w+):(\w+)>', text)
        if m:
            pathlist.append(''.join(m.groups()))
        # NOTE
        m = re.search(r'<NOTE OF(\w+):(\w+)-ON(\w+):(\w+)>', text)
        if m:
            pathlist.append(''.join(m.groups()))
        # URI
        uri = 'Notes://' + '/'.join(pathlist)
        # Markdown
        # uri = '[{0}]({1})'.format(title, REDIRECT + uri)  # Teams 用
        uri = '[[{1}|{0}]]'.format(title, uri)  # DokuWiki 用
    else:
        uri = text

    # 素の Notes URI の場合
    m = re.match('Notes://', uri)
    if m:
        title = 'Notes Link'
        # uri = '[{0}]({1})'.format(title, REDIRECT + uri)  # Teams 用
        uri = '[[{1}|{0}]]'.format(title, uri)  # DokuWiki 用

    # 変換後の URI をクリップボードにセットする
    if uri:
        print('\nクリップボードにコピーしました。')
        print('\n' + uri)
        clip.copy(uri)


if __name__ == '__main__':
    print(os.path.basename(sys.argv[0]))
    main()
    input('\nPress Enter to exit.')
