# -*- coding:utf-8  -*-
# @Time     : 2022/10/1 20:26
# @Author   : BGLB
# @Software : PyCharm
import os
import traceback
from distutils.core import setup
from shutil import copyfile, rmtree
from uuid import uuid4

import colorama
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad
from Cython.Build import cythonize

colorama.init(autoreset=True, wrap=True)

key_ = '123456789'
while len(key_)%16 != 0:
    key_ += chr(0)
key_ = key_.encode('utf-8')

ignore_file_list = []
build_file_list = ['core.py', ]
basedir = os.path.abspath(os.path.dirname(__file__))


# 本函数需要与util中的aes_encrypt保持一致！
def aes_encrypt(key, target):
    '''
    传入byte类型key及待加密数据，返回加密后数据
    Arguments:
        key {bytes} -- 16位长度byte类型秘钥
        target {bytes} -- 明文
    Returns:
        bytes -- 密文
    '''
    bs = AES.block_size
    mode = AES.MODE_CFB
    aes = AES.new(key, mode, key)
    pad_data = pad(target, bs)
    # aes.encrypt(pad_data).hex()
    return aes.encrypt(pad_data)


def get_py_file_list(dir_path):
    """
    从保护文件夹里找到所有需要加密的py文件
    :param dir_path:
    :return:
    """
    file_list = []
    f_name_list = os.listdir(dir_path)
    for f_name in f_name_list:
        f_path = os.path.join(dir_path, f_name)
        if os.path.isfile(f_path):
            if (os.path.splitext(f_path)[1].lower() == '.py'
                or os.path.splitext(f_path)[1].lower() == '.js') \
                    and f_name.lower() not in ignore_file_list:
                file_list.append(f_path)
    return file_list


def exec_setup(module_list):
    """
    编译
    :param module_list: 文件路径，单个或者list都行，建议单个
    :return:
    """
    try:
        # -b和-t是的缩写形式--build-lib和--build-temp
        setup(ext_modules=cythonize(module_list, nthreads=20, compiler_directives={'language_level': 3}),
              script_args=["build_ext", "--inplace"])  # , "-b", "./bbb", "-t", "./ttt"
    except Exception as e:
        print(colorama.Fore.RED+'编译文件失败！({})'.format(module_list))
        print(colorama.Fore.RED+traceback.format_exc())


def clean_files(module_file_path):
    """
    删除此次编译后，不需要的文件
    目标是abc.py，生成了abc.c,abc.pyd
    则删掉 abc.py和abc.c
    :param module_file_path: 目标文件
    :return:
    """
    # 删除 py文件
    if os.path.exists(module_file_path):
        os.remove(module_file_path)
    # 删除 .c文件
    c_file_path = module_file_path.replace('.py', '.c')
    if os.path.exists(c_file_path):
        os.remove(c_file_path)


def build_a_file(module_file_path):
    if os.path.exists(module_file_path):
        print(colorama.Fore.BLUE+'building {} ...'.format(os.path.split(module_file_path)[1]))
        exec_setup(module_file_path)
        # clean_files(module_file_path)


def build():
    for f_name in build_file_list:
        module_file_path = os.path.join(basedir, f_name)
        build_a_file(module_file_path)

    # 加密JS
    # py_file_list = get_py_file_list(os.path.join(basedir, ''))
    # for module_file_path in py_file_list:
    #     encryption_js_file(module_file_path)

    # 处理保护文件夹
    # for pd in protect_dir_list:
    #     py_file_list = get_py_file_list(os.path.join(basedir, pd))
    #     for module_file_path in py_file_list:
    #         build_a_file(module_file_path)


def package():
    """
        打包
    :return:
    """
    copyfile('main.py', 'build/main.py')
    os.system('move *.pyd build/')
    os.system(f'pyinstaller -F build/main.py --uac-admin --key {str(uuid4())}')


def clear():
    rm_ext_list = ['.c', '.pyd', '.da', '.spec']
    rm_dir = ['build']
    count = 0
    for item in rm_dir:
        dir_path = os.path.join(basedir, item)
        if os.path.exists(dir_path):
            rmtree(dir_path)
            print(colorama.Fore.BLUE+'rm {} ......'.format(dir_path))

    for root, dirs, files in os.walk(basedir):
        for file in files:
            f_path = os.path.join(root, file)
            if os.path.splitext(f_path)[1].lower() in rm_ext_list:
                count += 1
                print(colorama.Fore.BLUE+'{} rm {}'.format(count, f_path))
                os.remove(f_path)
    print(colorama.Fore.RED+'rm files count = {}'.format(count))


def main():
    """

    :return:
    """
    print('clear files of before build')
    clear()
    print('start build ......')
    build()
    print('package ......')
    package()
    print('clear')
    clear()


if __name__ == '__main__':
    main()
