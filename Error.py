# coding=utf-8
# @Author  : Edmond

class AppError(Exception):
    def __init__(self, errcode,errinfo) -> None:
        super().__init__(self)
        self.errCode = errcode
        self.errInfo = errinfo
        self.__errmsg = ':'.join([errcode,errinfo])
    def __str__(self):
        return self.__errmsg