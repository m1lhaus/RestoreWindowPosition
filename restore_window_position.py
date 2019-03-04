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
from collections import defaultdict

import win32api
import win32gui
import win32con


def get_window_by_name(target_name, is_name_regex, lowercase_only, find_only_one):
    # helper custom exception to stop pywin32 enumeration
    class StopEnumerateWindows(Exception):
        pass

    def enum_windows_callback(i_hwnd, _):
        nonlocal found_window_hwnd, target_name, is_name_regex, lowercase_only, find_only_one

        window_name = win32gui.GetWindowText(i_hwnd)
        if not window_name:
            return

        # print(window_name)

        if lowercase_only:
            target_name = target_name.lower()
            window_name = window_name.lower()

        if is_name_regex:
            if re.match(target_name, window_name):
                found_window_hwnd = i_hwnd
                if find_only_one:
                    raise StopEnumerateWindows()                # C-version returns true or false, pywin32 is kind-of broken here
        else:
            if target_name in window_name:
                found_window_hwnd = i_hwnd
                if find_only_one:
                    raise StopEnumerateWindows()

    found_window_hwnd = 0
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except StopEnumerateWindows:
        pass

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

    config["DEFAULT"]["CaseInsensitive"] = parser.getboolean("DEFAULT", "CaseInsensitive", fallback=False)
    config["DEFAULT"]["RefreshRateInSec"] = parser.getfloat("DEFAULT", "RefreshRateInSec", fallback=1)
    config["DEFAULT"]["SaveRateInMin"] = parser.getfloat("DEFAULT", "SaveRateInMin", fallback=1)

    for section in parser.sections():
        config[section]["WindowTitle"] = remove_quotes(parser.get(section, "WindowTitle", fallback=""))
        config[section]["UseRegEx"] = parser.getboolean(section, "UseRegEx", fallback=False)
        config[section]["CaseSensitive"] = parser.getboolean(section, "CaseSensitive", fallback=True)
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
        parser.write(configfile)


def find_all_windows(config):
    for window_record, details in config.items():
        if window_record == "DEFAULT":
            continue

        win_hwnd = get_window_by_name(details["WindowTitle"], is_name_regex=details["UseRegEx"], lowercase_only=not details["CaseSensitive"], find_only_one=True)
        if win_hwnd:
            config[window_record]["HWND"] = win_hwnd
            config[window_record]["RealWindowTitle"] = win32gui.GetWindowText(win_hwnd)

            # window is opened, restore its position
            if not config[window_record]["WindowActive"]:
                restore_window_position(win_hwnd, config[window_record]["PosX0"], config[window_record]["PosY0"], config[window_record]["PosX1"], config[window_record]["PosY1"], config[window_record]["OnTop"])

            config[window_record]["WindowActive"] = True

        else:
            config[window_record]["HWND"] = None
            config[window_record]["RealWindowTitle"] = None
            config[window_record]["WindowActive"] = False

    return config


def update_positions(config):
    for window_record, details in config.items():
        if window_record == "DEFAULT":
            continue

        win_hwnd = details["HWND"]
        if win_hwnd is None:
            return config

        try:
            x0, y0, x1, y1 = win32gui.GetWindowRect(win_hwnd)      # (left, top, right, bottom)
        except Exception:
            pass        # might happen whatever
        else:
            is_minimized = is_windows_minimized(x0, y0, x1, y1)
            details["Minimized"] = is_minimized

            if not is_minimized:
                details["PosX0"] = x0
                details["PosY0"] = y0
                details["PosX1"] = x1
                details["PosY1"] = y1

    return config


def is_windows_minimized(x0, y0, x1, y1):
    return (x0 < 0) and (y0 < 0) and (x1 < 0) and (y1 < 0)


def print_summary(config):
    for win_record, win_details in config.items():
        if win_record == "DEFAULT":
            continue

        real_window_title = win_details["RealWindowTitle"] if win_details["RealWindowTitle"] else "N/A"
        hwnd = win_details["HWND"] if win_details["HWND"] else "N/A"
        is_window_active = win_details["WindowActive"]
        is_minimized = win_details["Minimized"]
        if win_details["WindowActive"] and not is_minimized:
            position = "(%d, %d, %d, %d)" % (win_details["PosX0"], win_details["PosY0"], win_details["PosX1"], win_details["PosY1"])
        else:
            position = "N/A"

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

        width, height = x1 - x0, y1 - y0

        print("Restoring window: %s, %s" % (hwnd, on_top))

        if not is_windows_minimized(x0, y0, x1, y1):
            win32gui.SetWindowPos(hwnd, hwnd_insert_after, x0, y0, width, height, uflags)


def restore_window_position_worker(config_sile, stop_event):
    cfg_parser = read_ini_file(config_sile)
    cfg = read_ini_parser(cfg_parser)
    refresh_rate = cfg["DEFAULT"]["RefreshRateInSec"]  # seconds
    save_rate = cfg["DEFAULT"]["SaveRateInMin"]  # minutes
    save_every = max(int(round((60 * save_rate) / refresh_rate)), 1)  # cycles

    print("Restore window position laucnhed")

    while not stop_event.is_set():

        i = 0
        while i < save_every and not stop_event.is_set():
            # print("\n"*50)
            os.system("cls")

            i += 1
            cfg = find_all_windows(cfg)
            cfg = update_positions(cfg)

            print("Restore window position is running...")
            print("")
            print_summary(cfg)
            print("")
            print("Press 'q' key to exit")
            print("--> ", end="")
            sys.stdout.flush()

            time.sleep(refresh_rate)

        write_ini_parser(cfg_parser, cfg)
        write_ini_file(config_sile, cfg_parser)

    print("")
    print("Restore window position closed")


if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser(description="")
    arg_parser.add_argument('-c', "--config", type=str, default="config.ini", help='')
    args = arg_parser.parse_args()

    args.config = os.path.abspath(args.config)
    if not os.path.join(args.config):
        raise Exception("Config file '%s' was not found!" % args.config)

    stop_event = threading.Event()
    worker_thread = threading.Thread(name='worker_thread', target=restore_window_position_worker, args=(args.config, stop_event))

    worker_thread.start()

    try:
        while True:
            cin = msvcrt.getch()
            if cin.strip() == b"q":
                break
    finally:
        stop_event.set()
        worker_thread.join()
