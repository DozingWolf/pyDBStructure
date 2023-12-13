# coding=utf-8
# @Author  : Edmond

from pathlib import Path
from configparser import ConfigParser
from traceback import format_exc
from loguru import logger
from makeDoucmentFile import createExcel
from tkinter import messagebox
from Error import AppError

@logger.catch
def main():
    # get parase
    try:
        absPath = Path(__file__)
        conf = ConfigParser()
        conf.read('db.ini',encoding='utf-8')
        logFile = conf.get('log','path')
        logFilePath = absPath.parent / Path(logFile).parent
        if logFilePath.exists():
            logger.debug('log path {filepath} is exists', filepath = str(logFilePath))
            logger.add(logFile, rotation='00:00',compression='zip',level='DEBUG')
        else:
            logger.debug(str(logFilePath))
            logFilePath.mkdir(parents=True)

        outputPath = conf.get('output','path')
        absOutputPath = absPath.parent / outputPath
        if absOutputPath.exists():
            logger.debug('output path {filepath} is exists',filepath = str(absOutputPath))
        else:
            absOutputPath.mkdir(parents=True)
        dbType = conf.get('DB','type')
    except Exception as err:
        logger.error(format_exc())
        messagebox.showerror(title='Error Message',message=err)
        return 0
    # get db structure
    if dbType == 'MSSQL':
        from getTableStructure import getSqlserverTableStructure
        try:
            rtn = getSqlserverTableStructure(serverIP=conf.get('DB','ip'),
                                            serverPort=conf.get('DB','port'),
                                            dbUser=conf.get('DB','user'),
                                            dbPasswd=conf.get('DB','pw'),
                                            dbName=conf.get('DB','ms_dbname'))
        except Exception as err:
            logger.error(err)
            messagebox.showerror(title='Error Message',message=err)
            raise err
        else:
            logger.debug('GET MSSQL DB STRUCTURE DATA SUCCESS!')
        
    elif dbType == 'ORACLE':
        from getTableStructure import getOracleTableStructure
        from oracledb.exceptions import DatabaseError as OrclDatabaseError
        try:
            rtn = getOracleTableStructure(dbIP=conf.get('DB','ip'),
                                        dbPort=conf.get('DB','port'),
                                        dbUser=conf.get('DB','user'),
                                        dbPasswd=conf.get('DB','pw'),
                                        dbServicename=conf.get('DB','orcl_servicename'))
        except OrclDatabaseError as err:
            logger.error('THIS IS AN OrclDatabaseError ERROR!')
            logger.error(err)
            logger.error(format_exc())
            messagebox.showerror(title='Error Message',message=err)
            return 0
        except Exception as err:
            logger.error('THIS IS AN ERROR!')
            logger.error(err)
            logger.error(format_exc())
            messagebox.showerror(title='Error Message',message=err)
            return 0
        else:
            logger.debug('GET ORACLE DB STRUCTURE DATA SUCCESS!')

    elif dbType == 'MYSQL':
        raise AppError(errcode='E000',errinfo='Dont worry, this model was empty~')
    elif dbType == 'PG':
        raise AppError(errcode='E000',errinfo='Dont worry, this model was empty~')
    else:
        messagebox.showwarning(title='Warning',message='Error db type %s'%dbType)
        raise AppError(errcode='E999',errinfo='Error db type %s'%dbType)
    
    # build document
    try:
        createExcel(dbtype=dbType,dataset=rtn,path=conf.get('output','path'),filename=conf.get('output','filename'))
    except Exception as err:
        logger.error(err)
        logger.error(format_exc())
        messagebox.showerror(title='Error Message',message=err)
        return 0
    else:
        logger.debug('CREATE DB STRUCTURE DOCUMENT SUCCESS!')
        logger.debug('HAVE A GOOD TIME!')
        messagebox.showinfo(title='Success',message='Make database table structure document file success')
        return 1

if __name__ == '__main__':
    main()