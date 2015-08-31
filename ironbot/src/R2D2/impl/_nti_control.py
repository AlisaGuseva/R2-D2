# -*- coding: utf-8 -*-

from time import sleep

import clr
import logging
clr.AddReference("White.Core")
clr.AddReference("System")
clr.AddReference("System.Core")
clr.AddReference("UIAutomationClient")
clr.AddReference("UIAutomationTypes")

from White.Core.WindowsAPI.KeyboardInput import SpecialKeys

import White.Core.Application as Application
from System.Diagnostics import ProcessStartInfo, Process
import White.Core.Desktop as Desktop
from System.Windows.Automation import AutomationProperty, AutomationElement

from White.Core.UIItems.Finders import SearchCriteria

from White.Core.UIItems import Button, TextBox, RadioButton, Label, CheckBox, ListView, Panel, ListViewRow

from White.Core.UIItems.ListViewItems import ListViewHeader, ListViewFactory

from White.Core.UIItems.ListBoxItems import ListBox, ListItem, ComboBox, ListItemContainer

from White.Core.UIItems.TabItems import Tab, TabPage
# import White.Core.UIItems
# logging.warning(repr(dir(White.Core.UIItems)))
from White.Core.UIItems.TreeItems import Tree, TreeNode
from White.Core.UIItems.WindowItems import Window, DisplayState, TitleBar
from White.Core.UIItems.MenuItems import Menu, PopUpMenu
from White.Core.UIItems.WindowStripControls import ToolStrip, MenuBar
# from White.Core.UIItems.ListBoxItems import ComboBox
# from White.Core.UIItems.TableItems import Table

from System.Windows.Automation import ControlType

from System.Windows.Automation import ControlType

import clr
clr.AddReference('Win32API')
from Win32API import Win32API

from _params import Delay, fixed_val, pop, pop_re, pop_type, robot_args, pop_bool, pop_menu_path, str_2_bool
from _util import IronbotException, waiting_iterator, result_modifier, error_decorator
from _attr import AttributeDict
from _attr import attr_checker, re_checker, my_getattr, attr_reader
from _keys import pop_key, pop_key_string

NIT_GET_PARAMS = (
    (pop,), {
        'menu_path': (('menu_path', pop_menu_path),),
        'negative': (('negative', pop_type(Delay)),),
        'timeout': (('timeout', pop_type(Delay)),),
        'single': (('single', fixed_val(True)),),
        'none': (('none', fixed_val(True)),),
        'assert': (('_assert', fixed_val(True)),),
        'failure_text': (('failure_text', pop),),
    }
)

@robot_args(NIT_GET_PARAMS, insert_attr_dict=True)
@error_decorator
def nti_menu_select(app_name, **kw):
    timeout = kw.get('timeout', None)
    _assert = kw.get('assert', True)
    lst = kw.get('menu_path', None)
    for w in Desktop.Instance.Windows():
        if w.Name == app_name:
            break
    else:
        raise Exception('No app')
    for toolbar in w.GetMultiple(SearchCriteria.ByControlType(ToolStrip)):
        print toolbar.Name
        if toolbar.Name == 'Menu':
            break
    else:
        raise Exception('No toolbar')
    lst.reverse()
    f = toolbar.MenuItem(lst.pop())
    f.Click()
    lst.reverse()
    ls = iter(lst)
    if 'timeout' in kw:
        del kw['timeout']
    for l in lst:
        print(l)
        for _ in waiting_iterator(timeout):
            menu_ae = Desktop.Instance.GetElement(SearchCriteria.ByControlType(ControlType.Menu).AndByText("Menu"))
            mnu = MenuBar(menu_ae, Desktop.Instance.ActionListener)
        mnu.MenuItem(ls.next()).Click()

LIST_GET_PARAMS = (
    (pop,), {
        'window_name': (('window_name', pop),),
        'list_item_path': (('list_item_path', pop_menu_path),),
        'negative': (('negative', pop_type(Delay)),),
        'timeout': (('timeout', pop_type(Delay)),),
        'single': (('single', fixed_val(True)),),
        'none': (('none', fixed_val(True)),),
        'assert': (('_assert', fixed_val(True)),),
        'failure_text': (('failure_text', pop),),
    }
)

@robot_args(LIST_GET_PARAMS, insert_attr_dict=True)
@error_decorator
def nti_list_select(app_name, **kw):
    timeout = kw.get('timeout', None)
    _assert = kw.get('assert', True)
    wnd_name = kw.get('window_name',None)
    list_item_path = kw.get('list_item_path', None)
    wnd = kw.get('window_name', True)
    for w in Desktop.Instance.Windows():
        if w.Name == app_name:
            break
    else:
        raise Exception('No app')
    l = list_item_path
    print(l)
    lst = iter(list_item_path)
    for _ in waiting_iterator(timeout):
        try:
             w1 = w.GetElement(SearchCriteria.ByControlType(ControlType.Window).AndByText(wnd_name))
             if w1:
                 break
        except:
            print('Can not find ', wnd_name)
    else:
        raise Exception('Can not find ', wnd_name)

    for lstview in w.GetMultiple(SearchCriteria.ByControlType(ListView)):
        lis_t = list(lstview.Items)
        lst_item = lst.next()
        for item_l in lis_t:
            if item_l == lst_item:
                ind = lis_t.index(item_l)
                lstview.Select(ind)
                break
        else:
            raise Exception('List Item not found:', lst_item)
