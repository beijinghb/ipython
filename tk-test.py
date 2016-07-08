# -*- coding: UTF-8 -*-
#python tkinter menu
#python version 3.3.2
#EN = Window 7


from tkinter import *
from tkinter.messagebox import *
from tkinter.filedialog import *
from tkinter.colorchooser import *
'''
    在python 3.3.2中，tkinter模块可以创建一个窗口控件，如Java中的Swing
    功能描述：
        根据Python 3.3.2 IDEL的菜单，创建出一个tkinter窗口
        File-Exit    :  退出功能完成
        Help-About IDEL     ： 打印相应信息
        其他的菜单项，当点击时，会打印出相应菜单项的名称
'''

__author__ = 'Hongten'
MENU_ITEMS = ['File', 'Edit', 'Format', 'Run', 'Options', 'Windows', 'Help', 'Test']
#菜单File中的选项
MENU_FILE_ITEMS = ['New Window      Ctrl+N      ',
                   'Open...         Ctrl+O      ',
                   'Recent Files                ',
                   'Open Module...  Alt+M       ',
                   'Class Browser   Alt+C       ',
                   'Path Browser                ',
                   'Save            Ctrl+S      ',
                   'Save As...      Ctrl+Shift+S',
                   'Save Copy As... Ctrl+Alt+S  ',
                   'Print Window    Ctrl+P      ',
                   'Close           Alt+F4      ',
                   'Exit            Ctrl+Q      ']
#菜单Edit中的选项
MENU_EDIT_ITEMS = ['Undo                     Ctrl+Z        ',
                   'Redo                     Ctrl+Shift+Z  ',
                   'Cut                      Ctrl+X        ',
                   'Copy                     Ctrl+C        ',
                   'Paste                    Ctrl+V        ',
                   'Select All               Ctrl+A        ',
                   'Find...                  Ctrl+F        ',
                   'Find Again               Ctrl+G        ',
                   'Find Selections          Ctrl+F3       ',
                   'Find in Files            Alt+F3        ',
                   'Replace...               Ctrl+H        ',
                   'Go to Line               Alt+G         ',
                   'Expend Word              Alt+/         ',
                   'Show call tip            Ctrl+backslash',
                   'Show surerounding parens Ctrl+0        ',
                   'Show Completions         Ctrl+space    ']
#菜单Format中的选项
MENU_FORMAT_ITEMS = ['Check Module        Alt+X   ',
                   'Ident Region        Ctrl+]  ',
                   'Dedent Region       Ctrl+[  ',
                   'Commemt Out Region  Alt+3   ',
                   'Uncomment Region    Alt+4   ',
                   'Tabify Region       Alt+5   ',
                   'Untabify Region     Alt+6   ',
                   'Toggle Tabs         Alt+T   ',
                   'New Ident Width     Alt+U   ',
                   'Format Paragraph    Alt+Q   ',
                   'Strip trailing whitespace   ']
#菜单Run中的选项
MENU_RUN_ITEMS = [ 'Python Shell                ',
                   'Check Module    Alt+X       ',
                   'Run Module      F5          ']
#菜单Options中的选项
MENU_OPTIONS_ITEMS = ['Config IDEL...              ',
                   'Code Context                ']
#菜单Windows中的选项
MENU_WINDOWS_ITEMS = ['Zoom Height     Alt+2       ']
#菜单Help中的选项
MENU_HELP_ITEMS = ['About IDEL                  ',
                   'IDEL Help                   ',
                   'Python Docs     F1          ']
#菜单Test中的选项
MENU_TEST_ITEMS = ['about                       ',
                   'askokcancel messagebox      ',
                   'askquestion messagebox      ',
                   'askretrycancel messagebox   ',
                   'askyesno messagebox         ',
                   'showerror messagebox        ',
                   'showinfo messagebox         ',
                   'showwarning messagebox      ',
                   'open file                   ',
                   'save as file                ',
                   'colorchooser                ']

#help-About IDEL
ABOUT_MESSAGE = '''
    Author       : Hongten
    Author_email : hongtenzone@foxmail.com
    Blog         : http://www.cnblogs.com/hongten
    QQ           : 648719819
    Created      : 2013-09-05
    Version      : 1.0
    '''
def get_tk():
    '''获取一个Tk对象'''
    return Tk()

def set_tk_title(tk, title):
    '''给窗口定义title'''
##    if title is not None and title != '':
    if title:
        tk.title(title)
    else:
        tk.title('Hongten v1.0')

def set_tk_geometry(tk, size):
    '''设置窗口大小，size的格式为：widthxheight,如：size = '200x100'.'''
##    if size is not None and size != '':
    if size:
        tk.geometry(size)
    else:
        tk.geometry('670x600')

def get_menu(tk):
    '''获取一个菜单条'''
    return Menu(tk)

def menu_file(menubar):
    '''定义菜单File'''
    filemenu = Menu(menubar, tearoff=1)
    filemenu.add_command(label=MENU_FILE_ITEMS[0], command=lambda:print(MENU_FILE_ITEMS[0]))
    filemenu.add_command(label=MENU_FILE_ITEMS[1], command=lambda:print(MENU_FILE_ITEMS[1]))
    filemenu.add_command(label=MENU_FILE_ITEMS[2], command=lambda:print(MENU_FILE_ITEMS[2]))
    
    filemenu.add_command(label=MENU_FILE_ITEMS[3], command=lambda:print(MENU_FILE_ITEMS[3]))
    filemenu.add_command(label=MENU_FILE_ITEMS[4], command=lambda:print(MENU_FILE_ITEMS[4]))
    filemenu.add_command(label=MENU_FILE_ITEMS[5], command=lambda:print(MENU_FILE_ITEMS[5]))
    filemenu.add_separator()
    filemenu.add_command(label=MENU_FILE_ITEMS[6], command=lambda:print(MENU_FILE_ITEMS[6]))
    filemenu.add_command(label=MENU_FILE_ITEMS[7], command=lambda:print(MENU_FILE_ITEMS[7]))
    filemenu.add_command(label=MENU_FILE_ITEMS[8], command=lambda:print(MENU_FILE_ITEMS[8]))
    filemenu.add_separator()
    filemenu.add_command(label=MENU_FILE_ITEMS[9], command=lambda:print(MENU_FILE_ITEMS[9]))
    filemenu.add_separator()
    filemenu.add_command(label=MENU_FILE_ITEMS[10], command=lambda:print(MENU_FILE_ITEMS[10]))
    filemenu.add_command(label=MENU_FILE_ITEMS[11], command=root.destroy)
    menubar.add_cascade(label=MENU_ITEMS[0], menu=filemenu)

def menu_edit(menubar):
    '''定义菜单Edit'''
    edit_menu = Menu(menubar, tearoff=1)
    edit_menu.add_command(label=MENU_EDIT_ITEMS[0], command=lambda:print(MENU_EDIT_ITEMS[0]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[1], command=lambda:print(MENU_EDIT_ITEMS[1]))
    edit_menu.add_separator()
    edit_menu.add_command(label=MENU_EDIT_ITEMS[2], command=lambda:print(MENU_EDIT_ITEMS[2]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[3], command=lambda:print(MENU_EDIT_ITEMS[3]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[4], command=lambda:print(MENU_EDIT_ITEMS[4]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[5], command=lambda:print(MENU_EDIT_ITEMS[5]))
    edit_menu.add_separator()
    edit_menu.add_command(label=MENU_EDIT_ITEMS[6], command=lambda:print(MENU_EDIT_ITEMS[6]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[7], command=lambda:print(MENU_EDIT_ITEMS[7]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[8], command=lambda:print(MENU_EDIT_ITEMS[8]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[9], command=lambda:print(MENU_EDIT_ITEMS[9]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[10], command=lambda:print(MENU_EDIT_ITEMS[10]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[11], command=lambda:print(MENU_EDIT_ITEMS[11]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[12], command=lambda:print(MENU_EDIT_ITEMS[12]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[13], command=lambda:print(MENU_EDIT_ITEMS[13]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[14], command=lambda:print(MENU_EDIT_ITEMS[14]))
    edit_menu.add_command(label=MENU_EDIT_ITEMS[15], command=lambda:print(MENU_EDIT_ITEMS[15]))
    menubar.add_cascade(label=MENU_ITEMS[1], menu=edit_menu)

def menu_format(menubar):
    '''定义菜单Format'''
    format_menu = Menu(menubar, tearoff=1)
    format_menu.add_command(label=MENU_FORMAT_ITEMS[0], command=lambda:print(MENU_FORMAT_ITEMS[0]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[1], command=lambda:print(MENU_FORMAT_ITEMS[1]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[2], command=lambda:print(MENU_FORMAT_ITEMS[2]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[3], command=lambda:print(MENU_FORMAT_ITEMS[3]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[4], command=lambda:print(MENU_FORMAT_ITEMS[4]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[5], command=lambda:print(MENU_FORMAT_ITEMS[5]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[6], command=lambda:print(MENU_FORMAT_ITEMS[6]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[7], command=lambda:print(MENU_FORMAT_ITEMS[7]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[8], command=lambda:print(MENU_FORMAT_ITEMS[8]))
    format_menu.add_command(label=MENU_FORMAT_ITEMS[9], command=lambda:print(MENU_FORMAT_ITEMS[9]))
    format_menu.add_separator()
    format_menu.add_command(label=MENU_FORMAT_ITEMS[10], command=lambda:print(MENU_FORMAT_ITEMS[10]))
    menubar.add_cascade(label=MENU_ITEMS[2], menu=format_menu)

def menu_run(menubar):
    '''定义菜单Run'''
    run_menu = Menu(menubar, tearoff=1)
    run_menu.add_command(label=MENU_RUN_ITEMS[0], command=lambda:print(MENU_RUN_ITEMS[0]))
    run_menu.add_separator()
    run_menu.add_command(label=MENU_RUN_ITEMS[1], command=lambda:print(MENU_RUN_ITEMS[1]))
    run_menu.add_command(label=MENU_RUN_ITEMS[2], command=lambda:print(MENU_RUN_ITEMS[2]))
    menubar.add_cascade(label=MENU_ITEMS[3], menu=run_menu)

def meun_options(menubar):
    '''定义菜单Options'''
    options_menu = Menu(menubar, tearoff=1)
    options_menu.add_command(label=MENU_OPTIONS_ITEMS[0], command=lambda:print(MENU_OPTIONS_ITEMS[0]))
    options_menu.add_separator()
    options_menu.add_command(label=MENU_OPTIONS_ITEMS[1], command=lambda:print(MENU_OPTIONS_ITEMS[1]))
    menubar.add_cascade(label=MENU_ITEMS[4], menu=options_menu)

def menu_windows(menubar):
    '''定义菜单Windows'''
    windows_menu = Menu(menubar, tearoff=1)
    windows_menu.add_command(label=MENU_WINDOWS_ITEMS[0], command=lambda:print(MENU_WINDOWS_ITEMS[0]))
    windows_menu.add_separator()
    menubar.add_cascade(label=MENU_ITEMS[5], menu=windows_menu)

def meun_help(menubar):
    '''定义菜单Help'''
    help_menu = Menu(menubar, tearoff=1)
    help_menu.add_command(label=MENU_HELP_ITEMS[0], command=lambda:print(MENU_HELP_ITEMS[0]))
    help_menu.add_separator()
    help_menu.add_command(label=MENU_HELP_ITEMS[1], command=lambda:print(MENU_HELP_ITEMS[1]))
    help_menu.add_command(label=MENU_HELP_ITEMS[2], command=lambda:print(MENU_HELP_ITEMS[2]))
    menubar.add_cascade(label=MENU_ITEMS[6], menu=help_menu)

def meun_test(menubar):
    '''定义菜单Help'''
    test_menu = Menu(menubar, tearoff=1)
    test_menu.add_command(label=MENU_TEST_ITEMS[0], command=about)
    test_menu.add_separator()
    test_menu.add_command(label=MENU_TEST_ITEMS[1], command=help_test1)
    test_menu.add_command(label=MENU_TEST_ITEMS[2], command=help_test2)
    test_menu.add_command(label=MENU_TEST_ITEMS[3], command=help_test3)
    test_menu.add_command(label=MENU_TEST_ITEMS[4], command=help_test4)
    test_menu.add_command(label=MENU_TEST_ITEMS[5], command=help_test5)
    test_menu.add_command(label=MENU_TEST_ITEMS[6], command=help_test6)
    test_menu.add_command(label=MENU_TEST_ITEMS[7], command=help_test7)
    test_menu.add_separator()
    test_menu.add_command(label=MENU_TEST_ITEMS[8], command=help_test8)
    test_menu.add_command(label=MENU_TEST_ITEMS[9], command=help_test9)
    test_menu.add_separator()
    test_menu.add_command(label=MENU_TEST_ITEMS[10], command=help_test10)
    menubar.add_cascade(label=MENU_ITEMS[7], menu=test_menu)

############################################################
# Test Menu Items Functions
def about():
    '''Help-About IDEL function'''
    label = Label(root, text=ABOUT_MESSAGE, fg='red')
    label.pack(side='top')

def help_test1():
    ask = askokcancel('askokcancel messagebox','你确定要这样做吗？')
    if ask:
        # to do something
        print('你选择的是：确定')
    else:
        # to do something
        print('你选择的是：取消')

def help_test2():
    ask = askquestion('askquestion messagebox', '你很喜欢那个女孩吗?')
    if 'yes' == ask:
        #to do something
        print('是的，我很喜欢')
    elif 'no' == ask:
        #to do something
        print('不是这样的，我不喜欢她')

def help_test3():
    ask = askretrycancel('askretrycancel messagebox', '该程序被其他程序占用,请重试..')
    if ask:
        #to do something
        print('重试')
    else:
        #to do something
        print('取消重试')

def help_test4():
    ask = askyesno('askyesno messagebox', 'Can you cut down a tree with a herrign?')
    print(ask)
    if 'yes' == ask:
        #to do something
        print('yes')
    elif 'no' == ask:
        #to do something
        print('no')

def help_test5():
    error = showerror('showerror messagebox', 'ClassNotFoundException...')
    if 'ok' == error:
        #to do something
        print('你选择的是确定...')
        
def help_test6():
    info = showinfo('showinfo messagebox', 'This is an ex-parrot')
    if 'ok' == info:
        #to do something
        print('ok')

def help_test7():
    warn = showwarning('showwarning messagebox', '内存溢出...')
    if 'ok' == warn:
        #to do something
        print('确认内存溢出...')

def help_test8():
    '''打开文件'''
    #('All files', '*')
    openfilename = askopenfilename(filetypes=[('xml', '*.xml')])
    try:
        with open(openfilename, 'r') as fp:
            for line in fp:
                print(line)
            fp.close()
    except:
        print('Could not open File:%s'%openfilename)

def help_test9():
    '''打开文件'''
    saveasfilename = asksaveasfilename()
    print('saves', saveasfilename.encode('utf-8'))

def help_test10():
    '''颜色面板选择器'''
    color = askcolor(title='颜色面板')
    print(color)
    
############################################################
#init menu bar
def init_menu_bar(menubar):
    '''初始化菜单条'''
    menu_file(menubar)     #file
    menu_edit(menubar)     #edit
    menu_format(menubar)   #format
    menu_run(menubar)      #run
    meun_options(menubar)  #options
    menu_windows(menubar)  #windows
    meun_help(menubar)     #help
    meun_test(menubar)     #test

#获得窗口对象
root = get_tk()
#设置窗口大小
set_tk_geometry(root, '')
#设置窗口title
set_tk_title(root, 'Python 3.3.2 Shell')
#获取菜单对象
menubar = get_menu(root)
#初始化菜单
init_menu_bar(menubar)
#加载菜单配置
root.config(menu=menubar)


mainloop()
