# -*- coding: utf-8 -*-
import xlrd
import json


def loadConfig(path):
    try:
        file = open(path, "r")
    except Exception:
        print(Exception)
        return ""
    else:
        config = json.load(file)
        file.close()
        return config["path"]


def saveConfig(path):
    with open(path, "w") as file:
        info = json.dumps(path)
        file.write(info)
        return path
