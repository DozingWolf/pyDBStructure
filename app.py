# coding=utf-8
# @Author  : Edmond

from configparser import ConfigParser
from loguru import logger
from getTableStructure import getSqlserverTableStructure,getOracleTableStructure
from makeDoucmentFile import createExcel
from tkinter import messagebox
from Error import AppError

@logger.catch
def main():
    # get parase
    try:
        conf = ConfigParser()
        conf.read('db.ini',encoding='utf-8')
        logFile = conf.get('log','path')
        logger.add(logFile, rotation='00:00',compression='zip',level='WARNING')
        dbType = conf.get('DB','type')
    except Exception as err:
        logger.error(err)
        messagebox.showerror(title='Error Message',message=err)
        return 0
    # get db structure
    try:
        if dbType == 'MSSQL':
            rtn = getSqlserverTableStructure(serverIP=conf.get('DB','ip'),
                                            serverPort=conf.get('DB','port'),
                                            dbUser=conf.get('DB','user'),
                                            dbPasswd=conf.get('DB','pw'),
                                            dbName=conf.get('DB','ms_dbname'))
        elif dbType == 'ORACLE':
            rtn = getOracleTableStructure(dbIP=conf.get('DB','ip'),
                                          dbPort=conf.get('DB','port'),
                                          dbUser=conf.get('DB','user'),
                                          dbPasswd=conf.get('DB','pw'),
                                          dbServicename=conf.get('DB','orcl_servicename'))
        elif dbType == 'MYSQL':
            raise AppError(errcode='E000',errinfo='Dont worry, this model was empty~')
        elif dbType == 'PG':
            raise AppError(errcode='E000',errinfo='Dont worry, this model was empty~')
        else:
            raise AppError(errcode='E999',errinfo='Error db type %s'%dbType)
    except Exception as err:
        logger.error(err)
        messagebox.showerror(title='Error Message',message=err)
        return 0
    else:
        logger.debug('GET DB STRUCTURE DATA SUCCESS!')
    # build document
    try:
        createExcel(dbtype=dbType,dataset=rtn,path=conf.get('output','path'),filename=conf.get('output','filename'))
    except Exception as err:
        logger.error(err)
        messagebox.showerror(title='Error Message',message=err)
        return 0
    else:
        logger.debug('CREATE DB STRUCTURE DOCUMENT SUCCESS!')
        logger.debug('HAVE A GOOD TIME!')
        messagebox.showinfo(title='Success',message='Make database table structure document file success')
        return 1

if __name__ == '__main__':
    main()