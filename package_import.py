# -*- coding:utf-8  -*-
# @Time     : 2022/10/2 1:14
# @Author   : BGLB
# @Software : PyCharm
import os
import stat
import sys
import time
import traceback
import zipfile
from contextlib import closing
from copy import copy
from functools import reduce

import requests
from loguru import logger
from openpyxl import load_workbook
from openpyxl.cell import Cell
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait

import json
