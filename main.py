#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import json
import configparser
from collections import defaultdict

import win32api
import win32gui
import win32con
import msvcrt
import sys
import time
import threading
import pywintypes
import argparse
import re
import weakref
import pickle

# NAME = "Woofer"
# LOWERCASE = True
# REGEX = False


def get_window_by_name(target_name, is_name_regex, lowercase_only, find_only_one):
    # helper custom exception to stop pywin32 enumeration
    class StopEnumerateWindows(Exception):
        pass

    def enum_windows_callback(i_hwnd, _):
        nonlocal found_windows, target_name, is_name_regex, lowercase_only, find_only_one

        window_name = win32gui.GetWindowText(i_hwnd)
        if not window_name:
            return

        if lowercase_only:
            target_name = target_name.lower()
            window_name = window_name.lower()

        if is_name_regex:
            if re.match(target_name, window_name):
                found_windows.append(i_hwnd)
                if find_only_one:
                    raise StopEnumerateWindows()                # C-version returns true or false, pywin32 is kind-of broken here
        else:
            if target_name in window_name:
                found_windows.append(i_hwnd)
                if find_only_one:
                    raise StopEnumerateWindows()

    found_windows = []
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except StopEnumerateWindows:
        pass
    return found_windows


def read_ini_file(filename):
    if not os.path.isfile(filename):
        raise Exception("Unable to find file: %s" % os.path.abspath(filename))

    parser = configparser.ConfigParser()
    parser.read(filename)

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
    # config["DEFAULT"]["FindOnlyOne"] = parser.getboolean("DEFAULT", "FindOnlyOne", fallback=False)

    for section in parser.sections():
        config[section]["WindowTitle"] = remove_quotes(parser.get(section, "WindowTitle", fallback=""))
        config[section]["UseRegEx"] = parser.getboolean(section, "UseRegEx", fallback=False)
        config[section]["CaseSensitive"] = parser.getboolean(section, "CaseSensitive", fallback=True)
        config[section]["PosX0"] = parser.getint(section, "PosX0", fallback=-1)
        config[section]["PosY0"] = parser.getint(section, "PosY0", fallback=-1)
        config[section]["PosY1"] = parser.getint(section, "PosY1", fallback=-1)
        config[section]["PosX1"] = parser.getint(section, "PosX1", fallback=-1)

    return config


def write_ini_parser(parser, config):
    for section in parser.sections():
        parser[section]["PosX0"] = str(config[section]["PosX0"])
        parser[section]["PosY0"] = str(config[section]["PosY0"])
        parser[section]["PosX1"] = str(config[section]["PosX1"])
        parser[section]["PosY1"] = str(config[section]["PosY1"])


def write_ini_file(filename, parser):
    with open(filename, 'w') as configfile:
        parser.write(configfile)


def find_all_windows(config):
    for window_record, details in config.items():
        if window_record == "DEFAULT":
            continue

        window = get_window_by_name(details["WindowTitle"], is_name_regex=details["UseRegEx"], lowercase_only=not details["CaseSensitive"], find_only_one=True)
        if window:
            win_hwnd = window[0]
            cfg[window_record]["HWND"] = win_hwnd
            cfg[window_record]["RealWindowTitle"] = win32gui.GetWindowText(win_hwnd)

        else:
            cfg[window_record]["HWND"] = None
            cfg[window_record]["RealWindowTitle"] = None

    return config


def update_positions(config):
    for window_record, details in config.items():
        if window_record == "DEFAULT":
            continue

        win_hwnd = details["HWND"]
        if win_hwnd is not None:
            try:
                x0, y0, x1, y1 = win32gui.GetWindowRect(win_hwnd)      # (left, top, right, bottom)
            except Exception:
                pass        # might happen whatever

            details["PosX0"] = x0
            details["PosY0"] = y0
            details["PosX1"] = x1
            details["PosY1"] = y1

    return config


if __name__ == '__main__':

    filepath = "config.ini"
    cfg_parser = read_ini_file(filepath)
    cfg = read_ini_parser(cfg_parser)
    refresh_rate = cfg["DEFAULT"]["RefreshRateInSec"]       # seconds
    save_rate = cfg["DEFAULT"]["SaveRateInMin"]         # minutes
    save_every = max(int(round((60 * save_rate) / refresh_rate)), 1)        # cycles

    for i in range(save_every):
        cfg = find_all_windows(cfg)
        cfg = update_positions(cfg)

        print('\n'*20) # prints 80 line breaks
        print(json.dumps(cfg, indent="  "))
        time.sleep(refresh_rate)

    write_ini_parser(cfg_parser, cfg)
    write_ini_file(filepath, cfg_parser)