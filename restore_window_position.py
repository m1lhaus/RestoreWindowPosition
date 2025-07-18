# -*- coding: utf-8 -*-

# RestoreWindowPosition - simple script for Windows that can remember and restore window position.
# Copyright (C) 2019 Milan Herbig
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import re
import sys
import argparse
import time
import codecs
import msvcrt
import threading
import configparser
import traceback
from collections import defaultdict
from operator import sub

import win32gui
import win32con
import win32api


# if True, all found window names are printed to console and window enumeration wont get interrupted
DEBUG_MODE = False


def get_window_by_name(target_name, is_name_regex, lowercase_only, is_child_window):
    # helper custom exception to stop pywin32 enumeration
    class StopEnumerateWindows(Exception):
        pass

    def compare_window_title(target_wname, current_wname):
        nonlocal is_name_regex, lowercase_only
        if lowercase_only:
            target_wname = target_wname.lower()
            current_wname = current_wname.lower()
        return (is_name_regex and re.match(target_wname, current_wname)) or (target_wname == current_wname)

    def enum_child_windows_callback(i_hwnd, _):
        nonlocal found_window_hwnd, target_name, is_name_regex, lowercase_only

        child_wname = win32gui.GetWindowText(i_hwnd)
        if DEBUG_MODE: print("Child:", child_wname)
        if child_wname and compare_window_title(target_name, child_wname):
            found_window_hwnd = i_hwnd
            if not DEBUG_MODE: raise StopEnumerateWindows()

    def enum_windows_callback(i_hwnd, _):
        nonlocal found_window_hwnd, target_name, is_name_regex, lowercase_only, is_child_window

        window_name = win32gui.GetWindowText(i_hwnd)
        if DEBUG_MODE: print("Top-level:", window_name)
        if window_name and compare_window_title(target_name, window_name):
            found_window_hwnd = i_hwnd
            if not DEBUG_MODE: raise StopEnumerateWindows()

        if is_child_window:
            try:
                win32gui.EnumChildWindows(i_hwnd, enum_child_windows_callback, None)
            except StopEnumerateWindows:        # this is enumeration stop flag => pass it to upper levels
                raise
            # might happen whatever, but KeyboardInterrupt won't get caught
            except Exception:
                pass        # usually lot of "access denied" errors might be triggered, so skip exception traceback print

    found_window_hwnd = 0
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except StopEnumerateWindows:
        pass         # this is enumeration stop flag
    # because exception might be cause by win32 api, script functionality might be still OK => continue
    except Exception:
        # print exception because there should be none
        traceback.print_exc()

    return found_window_hwnd


def read_ini_file(filename):
    if not os.path.isfile(filename):
        raise Exception("Unable to find file: %s" % os.path.abspath(filename))

    parser = configparser.ConfigParser()
    with codecs.open(filename, "r", "utf8") as f:
        parser.read_file(f)

    return parser


def read_ini_parser(parser):
    def has_quotes(string):
        return string and ((string[0] == '"' and string[-1] == '"') or (string[0] == "'" and string[-1] == "'"))

    def remove_quotes(string):
        if has_quotes(string):
            string = string[1:] if string[0] in ('"', "'") else string
            string = string[:-1] if string[-1] in ('"', "'") else string
        return string

    config = defaultdict(dict)

    config["DEFAULT"]["RefreshRateInSec"] = parser.getfloat("DEFAULT", "RefreshRateInSec", fallback=1)
    config["DEFAULT"]["SaveRateInMin"] = parser.getfloat("DEFAULT", "SaveRateInMin", fallback=1)

    for section in parser.sections():
        config[section]["WindowTitle"] = remove_quotes(parser.get(section, "WindowTitle", fallback=""))
        config[section]["UseRegEx"] = parser.getboolean(section, "UseRegEx", fallback=False)
        config[section]["CaseSensitive"] = parser.getboolean(section, "CaseSensitive", fallback=True)
        config[section]["ChildWindow"] = parser.getboolean(section, "ChildWindow", fallback=False)
        config[section]["OnTop"] = parser.getboolean(section, "CaseSensitive", fallback=False)
        config[section]["PosX0"] = parser.getint(section, "PosX0", fallback=-1)
        config[section]["PosY0"] = parser.getint(section, "PosY0", fallback=-1)
        config[section]["PosY1"] = parser.getint(section, "PosY1", fallback=-1)
        config[section]["PosX1"] = parser.getint(section, "PosX1", fallback=-1)
        config[section]["WindowActive"] = False
        config[section]["Minimized"] = False

    return config


def write_ini_parser(parser, config):
    for section in parser.sections():
        parser[section]["PosX0"] = str(config[section]["PosX0"])
        parser[section]["PosY0"] = str(config[section]["PosY0"])
        parser[section]["PosX1"] = str(config[section]["PosX1"])
        parser[section]["PosY1"] = str(config[section]["PosY1"])


def write_ini_file(filename, parser):
    with codecs.open(filename, "w", "utf8") as configfile:
        text = ["; Using this config file you can add new windows to be tracked/saved/restored.\n"
                "; The config file need to have DEFAULT section with following items:\n"
                "; \t- RefreshRateInSec - how fast the script should check opened windows, their positions, etc. (default 1 sec)\n"
                "; \t- SaveRateInMin - how often should be current window positions written to config file (default 1 min)\n"
                "; \n"
                "; New window to track:\n"
                "; --------------------\n"
                "; [any_name_of_the_record]\n"
                "; windowtitle = name/title of the window or regular expression (string)\n"
                "; useregex = whether or not should be windowtitle handled as regex (boolean)\n"
                "; ontop = whether or not put everytime window on top (boolean)\n"
                "; casesensitive = whether or not to ignore case in window title (boolean)\n"
                "; childwindow = search also child windows for every top-level windows - BEAWARE of performance impact (boolean)\n"
                "; ==================================================================================================================\n"
                "\n"]

        configfile.writelines(text)
        parser.write(configfile)


def find_all_windows(config):
    for win_alias, win_properties in config.items():
        if win_alias == "DEFAULT":
            continue

        win_hwnd = get_window_by_name(win_properties["WindowTitle"], is_name_regex=win_properties["UseRegEx"],
                                      lowercase_only=not win_properties["CaseSensitive"], is_child_window=win_properties["ChildWindow"])
        if win_hwnd:
            if win_hwnd != win_properties.get("HWND", -1):
                win_properties["WindowActive"] = False      # we need to detect case the window reopens between two calls
            win_properties["HWND"] = win_hwnd
            win_properties["RealWindowTitle"] = win32gui.GetWindowText(win_hwnd)

            # if window is opened and not hidden (transparent), restore its position
            if (not win_properties["WindowActive"] and win32gui.IsWindowEnabled(win_hwnd) and not win32gui.IsIconic(win_hwnd)): 
                x0, y0, x1, y1 = win_properties["PosX0"], win_properties["PosY0"], win_properties["PosX1"], win_properties["PosY1"]
                if is_position_valid(x0, y0, x1, y1):
                    try:
                        restore_window_position(win_hwnd, x0, y0, x1, y1, win_properties["OnTop"])
                    except Exception:
                        print("Error when restoring window position")
                win_properties["WindowActive"] = True
        else:
            win_properties["HWND"] = None
            win_properties["RealWindowTitle"] = None
            win_properties["WindowActive"] = False

    return config


def update_positions(config):
    for win_alias, win_properties in config.items():
        if win_alias == "DEFAULT":
            continue

        # skip processing if window does not exists (no hwnd) or it is not activate (hidden or disabled)
        if win_properties["HWND"] is None or not win_properties["WindowActive"]:
            continue

        try:
            x0, y0, x1, y1 = win32gui.GetWindowRect(win_properties["HWND"])      # (left, top, right, bottom)
        except Exception:
            pass        # might happen whatever during win32 call
        else:
            win_properties["Minimized"] = win32gui.IsIconic(win_properties["HWND"])
            if not win_properties["Minimized"] and is_position_valid(x0, y0, x1, y1):
                win_properties["PosX0"] = x0
                win_properties["PosY0"] = y0
                win_properties["PosX1"] = x1
                win_properties["PosY1"] = y1

    return config


def is_position_valid(x0, y0, x1, y1):
    # Get the total virtual screen dimensions  
    virtual_screen_left = win32api.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN  
    virtual_screen_top = win32api.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN  
    virtual_screen_width = win32api.GetSystemMetrics(78) # SM_CXVIRTUALSCREEN  
    virtual_screen_height = win32api.GetSystemMetrics(79) # SM_CYVIRTUALSCREEN  
    x0_min, y0_min = virtual_screen_left, virtual_screen_top
    x1_max, y1_max = virtual_screen_left + virtual_screen_width, virtual_screen_top + virtual_screen_height
    is_valid = (x0 >= x0_min) and (y0 >= y0_min) and (x1 <= x1_max) and (y1 <= y1_max)
    return is_valid


def print_summary(config):
    for win_record, win_details in config.items():
        if win_record == "DEFAULT":
            continue

        real_window_title = win_details["RealWindowTitle"] if win_details["RealWindowTitle"] else "N/A"
        hwnd = win_details["HWND"] if win_details["HWND"] else "N/A"
        is_window_active = win_details["WindowActive"]
        is_minimized = win_details["Minimized"]
        na_flag = "N/A " if not win_details["WindowActive"] or is_minimized else ""
        position = "%s(%d, %d, %d, %d)" % (na_flag, win_details["PosX0"], win_details["PosY0"], win_details["PosX1"], win_details["PosY1"])

        print("[%s]" % win_record)
        print("RealWindowTitle =", real_window_title)
        print("HWND =", hwnd)
        print("WindowActive =", is_window_active)
        print("Minimized =", is_minimized)
        print("Position =", position)
        print("")


def restore_window_position(hwnd, x0, y0, x1, y1, on_top):
    if on_top:
        hwnd_insert_after = win32con.HWND_TOPMOST
        uflags = win32con.SWP_SHOWWINDOW
    else:
        hwnd_insert_after = win32con.HWND_TOP
        uflags = win32con.SWP_NOZORDER

    print("Restoring window: %s, on top:%s, pos: %s" % (hwnd, on_top, (x0, y0, x1, y1)))

    # I assume there might be race condition when window is opened, displayed and moved - so if we would engage
    # before it is moved, final position might be set by the app and not by our restore (but that is really a guess)
    max_num_tries = 10
    for i in range(max_num_tries):
        width, height = x1 - x0, y1 - y0
        win32gui.SetWindowPos(hwnd, hwnd_insert_after, x0, y0, width, height, uflags)
        time.sleep(0.05)
        x0_new, y0_new, x1_new, y1_new = win32gui.GetWindowRect(hwnd)  # (left, top, right, bottom)

        diff = tuple(map(sub, (x0_new, y0_new, x1_new, y1_new), (x0, y0, x1, y1)))      # new - old
        diff = tuple(map(abs, diff))        # make a diff as abs so we can make a sum
        diff = sum(diff)        # if position was set, sum should be zero
        if diff == 0:
            break


def restore_window_position_worker(config_sile, stop_event):
    cfg_parser = read_ini_file(config_sile)
    cfg = read_ini_parser(cfg_parser)
    refresh_rate = cfg["DEFAULT"]["RefreshRateInSec"]  # seconds
    save_rate = cfg["DEFAULT"]["SaveRateInMin"]  # minutes
    save_every = max(int(round((60 * save_rate) / refresh_rate)), 1)  # cycles
    last_summary = time.time()

    print("Restore window position launched")

    while not stop_event.is_set():

        i = 0
        while i < save_every and not stop_event.is_set():
            print("\n"*50)

            i += 1
            cfg = find_all_windows(cfg)
            cfg = update_positions(cfg)

            if time.time() - last_summary > 0.5:
                print("Restore window position is running...")
                print("")
                print_summary(cfg)
                print("")
                print("Press 'q' key to exit")
                print("--> ", end="")
                sys.stdout.flush()
                last_summary = time.time()

            time.sleep(refresh_rate)

        write_ini_parser(cfg_parser, cfg)
        write_ini_file(config_sile, cfg_parser)

    write_ini_parser(cfg_parser, cfg)
    write_ini_file(config_sile, cfg_parser)

    print("")
    print("Restore window position closed")


if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser(description="RestoreWindowPosition is simple script for Windows that can remember and restore window position")
    arg_parser.add_argument('-c', "--config", type=str, default="config.ini", help='Path to config file')
    args = arg_parser.parse_args()

    args.config = os.path.abspath(args.config)
    if not os.path.isfile(args.config):
        raise Exception("Config file '%s' was not found!" % args.config)

    # function is executed in own thread since main loop waits for q-key to be pressed
    stop_event = threading.Event()
    worker_thread = threading.Thread(name='worker_thread', target=restore_window_position_worker, args=(args.config, stop_event))
    worker_thread.start()

    try:
        while True:
            cin = msvcrt.getch()
            if cin.strip() == b"q":
                break
    finally:
        stop_event.set()        # stops the thread
        worker_thread.join()
