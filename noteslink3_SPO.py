# coding: utf-8

"""Notes リンクを URL に変換する（NotesLink2SPO）

まだ全然できませんが↓この DB と同等の機能を実装したいです。
Notes://ZARG11/4925824B001210A2

Notes の文書リンクをコピーした後、本プログラムを実行してください。
URL 変換後の文字列がクリップボードにセットされます。
"""

import os
import sys
import re

# import win32clipboard as CB
# win32clipboard を利用するには PyWin32 が必要です。
# https://sourceforge.net/projects/pywin32/
# Anaconda なら conda install pywin32 でインストールできます。

# 2018/03/14 クリップボードのコピペは pyperclip を使えばもっと簡単
import pyperclip as clip

REDIRECT = 'https://rfgricoh.sharepoint.com/sites/rfgportal/SitePages/redirect.aspx?'
# REDIRECT = 'http://10.63.100.25/notes/?url='


def main():
    print(os.path.basename(sys.argv[0]))

    # クリップボード値を取得
    """
    CB.OpenClipboard()
    try:
        # decode() で bytes 型を str 型に変換する
        # ※ Windows の日本語文字コードは CP932（Shift_JIS の拡張）
        text = CB.GetClipboardData(1).decode('cp932')  # 1: CF_TEXT
    except TypeError:
        text = ''
    CB.CloseClipboard()
    """
    text = clip.paste()

    # Notes リンク情報を取得する
    pathlist = []
    match = re.search(r'<NDL>(.+)', text)
    if match:
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
        # URL
        url = 'Notes://' + '/'.join(pathlist)
        # マークダウン
        url = '[{0}]({1})'.format(title, REDIRECT + url)
        url = '<a href="{1}">{0}</a>'.format(title, REDIRECT + url)
        url = rtf_encode(url)
    else:
        url = text

    if url:
        print('\nクリップボードにコピーしました。')
        print(url)

    # 変換後の URL をクリップボードにセットする
    """
    CB.OpenClipboard()
    CB.EmptyClipboard()
    CB.SetClipboardText(url)
    CB.CloseClipboard()
    """
    clip.copy(url)


def rtf_encode(s):
  result = ""
  for c in s:
    if ord(c) > 0x7f:
      result += "\\'%x" % ord(c)
    else:
      if c in ['\\', '{', '}']:
        result += '\\'
      result += c
  return result


if __name__ == '__main__':
    main()
    input('\nPress Enter to exit.')
