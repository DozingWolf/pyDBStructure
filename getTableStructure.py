# coding=utf-8
# @Author  : Edmond

from loguru import logger
from traceback import format_exc

#@logger.catch
def getSqlserverTableStructure(serverIP,serverPort,dbUser,dbPasswd,dbName):
    import pyodbc
    try:
        connectionString='DRIVER={ODBC Driver 17 for SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s'%(serverIP,dbName,dbUser,dbPasswd)

        conn = pyodbc.connect(connectionString)
        cur = conn.cursor()

        getTableNameListSql = '''select name from sysobjects where xtype = 'U' '''
        cur.execute(getTableNameListSql)
        tableNameRecd = cur.fetchall()
        tableNameList = []
        for tName in tableNameRecd:
            tableNameList.append(tName[0])

        getTableColumnsDescSql = '''
        SELECT 
            字段序号 = A.colorder,
            字段名 = A.name,
            字段说明 = cast(isnull(G.[value], '') as varchar(200)),
            标识 = Case When COLUMNPROPERTY(A.id, A.name, 'IsIdentity') = 1 Then 'Y' Else '' End,
            主键 = Case When exists (SELECT 1 FROM sysobjects Where xtype = 'PK' and parent_obj = A.id and name in
                                        (SELECT name FROM sysindexes WHERE indid in (SELECT indid FROM sysindexkeys WHERE id = A.id AND colid = A.colid))) then 'PK' else '' end,
            类型 = B.name,
            占用字节数 = A.Length,
            长度 = COLUMNPROPERTY(A.id, A.name, 'PRECISION'),
            小数位数 = isnull(COLUMNPROPERTY(A.id, A.name, 'Scale'), 0),
            允许空 = Case When A.isnullable = 1 Then 'Y' Else 'N' End,
            默认值 = isnull(E.Text, '')
            FROM syscolumns A
            Left Join systypes B On A.xusertype = B.xusertype
            Inner Join sysobjects D On A.id = D.id and D.xtype = 'U' and D.name <> 'dtproperties'
            Left Join syscomments E on A.cdefault = E.id
            Left Join sys.extended_properties G	on A.id = G.major_id and A.colid = G.minor_id
            Left Join sys.extended_properties F	On D.id = F.major_id and F.minor_id = 0
        where d.name = ?
        Order By A.id, A.colorder
        '''
        getTableColumnsConstraintsSql = '''
                select coc.CONSTRAINT_NAME, coc.COLUMN_NAME, con.CONSTRAINT_TYPE
                from user_cons_columns coc 
                inner join user_constraints con on coc.TABLE_NAME = con.table_name 
                    and coc.constraint_name = con.constraint_name 
                    and con.constraint_type <> 'P'
                where 1=1
                and coc.table_name = ?
                '''
        # 这段sql要改一下
        getTableColumnsIndexSql = '''
                SELECT
                IndexName=IDX.Name,
                IndexType=ISNULL(KC.type_desc,'Index'),
                UNIQUENESS=CASE WHEN IDX.is_unique=1 THEN N'UNIQUE'ELSE N'NONUNIQUE' END,
                COMPRESSION= ''
                FROM sys.indexes IDX INNER JOIN sys.index_columns IDXC ON IDX.[object_id]=IDXC.[object_id] 
                AND IDX.index_id=IDXC.index_id
                LEFT JOIN sys.key_constraints KC ON IDX.[object_id]=KC.[parent_object_id] 
                AND IDX.index_id=KC.unique_index_id
                INNER JOIN sys.objects O ON O.[object_id]=IDX.[object_id]
                INNER JOIN sys.columns C ON O.[object_id]=C.[object_id] 
                AND O.type='U' AND O.is_ms_shipped=0 
                AND IDXC.Column_id=C.Column_id 
                where O.name=?
                '''

        tableStructureList = []

        # list structure demo
        # [
        #     {
        #         tableindex:dd,
        #         tablename:xxxx,
        #         tablestructure:[]
        #     }
        # ]

        for id,exeTableName in enumerate(tableNameList):
            # logger.debug('%d table name is : %s'%(id,exeTableName))
            cur.execute(getTableColumnsDescSql,exeTableName)
            tableStructure = cur.fetchall()
            # logger.debug(tableStructure)
            cur.execute(getTableColumnsIndexSql,exeTableName)
            tableColIndexRtn = cur.fetchall()
            logger.debug(tableColIndexRtn)

            tableStructureList.append({'id' : id , 
                                    'TableName' : exeTableName , 
                                    'TableStructureData' : tableStructure ,
                                    'TableConstraint' : '',
                                    'TableColumnIndex' : tableColIndexRtn})
            

        # logger.debug(tableStructureList)
    except Exception as err:
        cur.close()
        conn.close()
        raise err
    else:
        cur.close()
        conn.close()
        return (tableStructureList)
    

#@logger.catch
def getOracleTableStructure(dbIP,dbPort,dbUser,dbPasswd,dbServicename):

    import oracledb
    from oracledb.exceptions import DatabaseError as OrclDatabaseError
    
    # conn = oracledb.connect(host=dbIP,port=dbPort,user=dbUser,password=dbPasswd,service_name=dbServicename)
    # list structure demo
    # 
    # [
    #     {
    #         tableindex:dd,
    #         tablename:xxxx,
    #         tablestructure:[],
    #         tablecolumnconstraint:[],
    #         tablecolumnindex:[],
    #         tabledesc:xxxxxx
    #     }
    # ]
    try:
        oracledb.init_oracle_client()
        conn = oracledb.connect(host=dbIP,port=dbPort,user=dbUser,password=dbPasswd,service_name=dbServicename)
        cur = conn.cursor()

        getOrclTableNameListSql = '''select table_name , comments from user_tab_comments where table_type = 'TABLE' '''
        cur.execute(getOrclTableNameListSql)
        tableNameRtn = cur.fetchall()
        # logger.debug(tableNameRtn)
        getTableColumnsDescSql = '''
        select 
        COLUMN_ID, ta.COLUMN_NAME, trim(ta.COMMENTS) as COMMENTS, ta.IsIdentity, tb.CONSTRAINT_TYPE, 
        DATA_TYPE, DATA_LENGTH, DATA_PRECISION, DATA_SCALE, NULLABLE, data_default
        from 
        (
            select 
            tc.COLUMN_ID, tc.COLUMN_NAME, tc.TABLE_NAME,cc.COMMENTS,'' as IsIdentity,
            tc.DATA_TYPE, tc.DATA_LENGTH, tc.DATA_PRECISION, tc.DATA_SCALE, tc.NULLABLE,
            tc.data_default
            from user_tab_cols tc 
            inner join user_col_comments cc on tc.TABLE_NAME = cc.table_name 
                    and tc.COLUMN_NAME = cc.column_name
            where 1=1
            and tc.TABLE_NAME = :table_name
        ) Ta
        left join 
        (
            select coc.CONSTRAINT_NAME, coc.TABLE_NAME, coc.COLUMN_NAME, con.CONSTRAINT_TYPE
            from user_cons_columns coc 
            inner join user_constraints con on coc.TABLE_NAME = con.table_name 
                    and coc.constraint_name = con.constraint_name 
                    and con.constraint_type = 'P'
            where 1=1
            and coc.table_name = :table_name
        ) Tb
        on ta.table_name = tb.table_name and ta.column_name = tb.column_name
        order by COLUMN_ID asc
        '''
        getTableColumnsConstraintsSql = '''
        select coc.CONSTRAINT_NAME, coc.COLUMN_NAME, con.CONSTRAINT_TYPE
        from user_cons_columns coc 
        inner join user_constraints con on coc.TABLE_NAME = con.table_name 
            and coc.constraint_name = con.constraint_name 
            and con.constraint_type <> 'P'
        where 1=1
        and coc.table_name = :table_name
        '''
        getTableColumnsIndexSql = '''
        select 
        INDEX_NAME, INDEX_TYPE, UNIQUENESS, COMPRESSION 
        from user_indexes ta
        where 1=1
        and ta.table_name = :table_name
        '''
        tableStructureList = []
        for idx,tableData in enumerate(tableNameRtn):
            tableDesc = tableData[1]
            logger.debug(tableData[0])
            
            # get table structure
            cur.execute(getTableColumnsDescSql,table_name = tableData[0])
            tableStruRtn = cur.fetchall()
            logger.debug(tableStruRtn)

            # get table columns constraint
            cur.execute(getTableColumnsConstraintsSql,table_name = tableData[0])
            tableColumnsConstRtn = cur.fetchall()
            logger.debug(tableColumnsConstRtn)

            # get table columns index
            cur.execute(getTableColumnsIndexSql,table_name = tableData[0])
            tableColIndexRtn = cur.fetchall()
            logger.debug(tableColIndexRtn)

            # build json structure
            tableStructureList.append({'id' : idx ,
                                    'TableName' : tableData[0] , 
                                    'TableDesc' : tableDesc, 
                                    'TableStructureData' : tableStruRtn,
                                    'TableConstraint' : tableColumnsConstRtn,
                                    'TableColumnIndex' : tableColIndexRtn})
        logger.debug(tableStructureList)
    except OrclDatabaseError as err:
        try:
            cur
        except NameError:
            pass
        else:
            cur.close()
        try:
            conn
        except NameError:
            pass
        else:
            conn.close()
        raise err
    except Exception as err:
        cur.close()
        conn.close()
        raise err
    else:
        cur.close()
        conn.close()
        return(tableStructureList)
        

