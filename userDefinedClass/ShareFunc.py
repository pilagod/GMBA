__author__ = 'pilagod'

from enum import Enum

class Level(Enum):
    Under = "u"
    Graduate = "g"

def caseAndSpaceIndif(string):
    return string.replace(" ", "").lower()