#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : log.py 
@Author  : Shawn
@Date    : 2025/10/30 16:32 
@Info    : Description of this file
"""

import os
import sys
import logging


def setup_logger(log_dir, name, log_filename='info.log', level=logging.INFO):
    try:
        os.makedirs(log_dir)
    except OSError:
        pass

    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Create logging format
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    if not logger.handlers:  # 避免重复日志
        # Add file handler
        file_handler = logging.FileHandler(os.path.join(log_dir, log_filename), mode='w')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Add console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    logger.info('Log directory: %s', log_dir)
    return logger


# if __name__ == '__main__':
#     print(f"********************start********************")
#     logger = setup_logger(log_dir='logs', name='logs', level=logging.INFO)
#     logger.info("this is a info message")
#
#     print(f"********************end**********************")
