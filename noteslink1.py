#!/usr/bin/python
# coding: utf-8

import os
import sys
import re
import time
import win32clipboard as CB
# win32clipboard を利用するには PyWin32 が必要です。
# https://sourceforge.net/projects/pywin32/

"""
MI事業 発信文書 - 【文書配布】下位手順一覧
<NDL>
<REPLICA 49257F99:001E6010>
<VIEW OF492562B7:00343D44-ON49256298:0005DADD>
<NOTE OF19578902:7492327B-ON492584B4:00047442>
<HINT>CN=ZARG11/O=RGroup</HINT>
<REM>Database 'MI事業 発信文書', View '1.最新文書ﾋﾞｭｰ\1.文書番号名別', Document '【文書配布】下位手順一覧'</REM>
</NDL>
"""
#
#noteslinkの場合
#https://rfgricoh.sharepoint.com/sites/rfgportal/SitePages/redirect.aspx?Notes://ZARG11/49257F99001E6010/492562B700343D44492562980005DADD/195789027492327B492584B400047442)
#
#NotesLink2SPOの場合（/ ではなく & でつなぐ）
#http://rgrsc01.nws.ricoh.co.jp/common/NUReSPO.nsf/NotesURL?open&//ZARG11&49257F99001E6010&492562B700343D44492562980005DADD&195789027492327B492584B400047442

srcdir = os.path.abspath(os.path.dirname(sys.argv[0]))
# htmlfile = os.path.join(srcdir, 'noteslink.html')
# starfile = os.path.join(srcdir, 'star.png')

if __name__ == '__main__':
    print os.path.basename(sys.argv[0])

    # クリップボード値を取得
    CB.OpenClipboard()
    try:
        text = CB.GetClipboardData(1)  # 1: CF_TEXT
    except TypeError:
        text = ''
    CB.CloseClipboard()

    # Notesリンク情報を取得
    pathlist = []
    m = re.search(r'<NDL>(.+)', text)
    if m:
        """
        print
        print '[0] Plain: Notes://...'
        print '[1] Markdown: [Title](Notes://...)'
        print '[2] Wiki: [[Notes://...|Title]]'
        output = int(raw_input('>>> '))
        """
        output = 0  # Plain
        output = 1  # Markdown

        # TITLE
        title = text.splitlines()[0].strip()

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

        # TITLE
##        m = re.search(r'<REM>(.+)</REM>', text)
##        if m:
##            rem = m.group(1)
##            n = re.search(r"Database '(.+)', View '(.+)', Document '(.+)'", rem)
##            if n:
##                database = n.group(1)
##                document = n.group(3)
##            else:
##                database = rem
##                document = None
##            title = document if document else database
##            title = re.split(r'\r\n', text)[0]
        if output == 1:
            uri = '[{0}]({1})'.format(title, uri)
        elif output == 2:
            uri = '[[{1}|{0}]]'.format(title, uri)

    else:
        uri = text

    # クリップボード
    CB.OpenClipboard()
    CB.EmptyClipboard()
    CB.SetClipboardText(uri)
    CB.CloseClipboard()
    print u'\nクリップボードにコピーしました。'
    print '\n' + uri

    # 終了
    # raw_input('\nPress Enter to exit.')
    sys.exit()

    # リンク集作成
    if not os.path.isfile(starfile):
        sys.exit()
    m = re.match('\[(.+)\]\((.+)\)$', text)
    if m:
        linklist = []

        # 既存リンクの取得
        if os.path.isfile(htmlfile):
            for line in open(htmlfile):
                n = re.search('<a href="(.+)">(.+)</a>', line)
                if n:
                    uri, title = n.groups()
                    print title, uri
                    linklist.append((title, uri))

        # 新規リンク登録
        title, uri = m.groups()
        if (title, uri) not in linklist:
            linklist.insert(0, (title, uri))

        # HTML作成
        html = '<p style="margin: 10 10; font-weight: bold; color: #888888;">Notes DB</p>\n'
        for i, (title, uri) in enumerate(linklist):
            icon = ' <img src="new.jpg">' if i == 0 else ''  # 新しい登録にはNEWアイコンを表示
            link = '<a href="{0}">{1}</a>'.format(uri, title)
            html += '\n<!-- {0} -->\n'.format(title)
            html += '<p style="margin: 3 10; font-size: smaller;"><img src="star.png">{0} {1}\n'.format(icon, link)

        html = """<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="ja">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="content-style-type" content="text/css">
<title>Notes DB</title>
</head>
<body>
{0}
</body>
</html>
""".format(html)

        f = open(htmlfile, 'w')
        f.write(html)
        f.close()
        os.system('start {0}'.format(htmlfile))
        os.system('start notepad {0}'.format(htmlfile))

    # time.sleep(0.1)
    # raw_input('\nPress Enter to exit.')
