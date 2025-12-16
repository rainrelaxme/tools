#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : result.py 
@Author  : Shawn
@Date    : 2025/11/7 11:03 
@Info    : Description of this file
"""

import json


def res_format(msg: str = "success", code: int = 200, info: str = None, data: dict = None):
    return json.dumps({
        "msg": msg,
        "code": code,
        "info": info,
        "data": data,
    })
