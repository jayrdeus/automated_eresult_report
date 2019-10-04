import pyodbc as sql
from datetime import date



def get_results():
    connection = sql.connect(
            "Driver={SQL Server Native Client 11.0};"
            "Server=192.168.1.4;"
            "Database=ON_LINE;"
            "uid=sa;"
            "pwd=Jc3s@r1991;"
    )
    today = date.today()
    date_format = today.strftime("%m/%d/%y")
    print(f'Fetching results from {date_format}')
    cursor =connection.cursor()
    query = f"""SELECT (b.PALNAME + ' ' + b.PAFNAME + ' ' + b.PAFNAME) as 'Patient Name',a.NORMAL_VALUE as 'Procedure',b.PAMOBILENO1 as 'Mobile Number',
                        CONVERT(VARCHAR(10), cast(a.DT_UPLOAD as date), 101) as 'Date Registered',  convert(varchar(10), cast(a.RESULT_DATE as date), 101) as 'Result date', datediff(day,a.RESULT_DATE,a.DT_UPLOAD) as 'Variance',
                        CONVERT(VARCHAR(10), cast(a.DT_SENT as date), 101) as 'Date SMS Sent',CONVERT(VARCHAR(10), cast(a.dl_datetime as date), 101) as 'Date Downloaded',
                        CASE WHEN a.IS_DL is not null THEN 'YES' ELSE '--' END as 'Downloaded'

                FROM X_RESULT a
                        INNER JOIN mkt_fmc_db.dbo.M_PATIENT b on a.PAPIN = b.PAPIN
                        WHERE a.DT_UPLOAD between getdate() - 5 and  getdate()
                        ORDER BY a.DT_UPLOAD DESC

            """ 
    cursor.execute(query)
    results = []
    for row in cursor:
        results.append(row)
    connection.close()
    return results