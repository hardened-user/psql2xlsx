Postgres SQL to .xlsx
======================

Utility for saving Postgres SQL querys results to .xlsx file 


Required
=======
* Python 3.x
  * psycopg2
  * xlsxwriter


Example
=======
Help
```console
user@localhost:~$ ./psql2xlsx.py -h
```
Config
```console
user@localhost:~$ nano config.ini
```
Run
```console
user@localhost:~$ ./psql2xlsx.py -f /tmp/1.xlsx
[..] Generated page :: page1 ...
[OK] PostgreSQL successfully connected
[..] Generated page :: page2 ...
[OK] PostgreSQL successfully connected
[..] Generated page :: page3 ...
[OK] PostgreSQL successfully connected
[OK] Workbook saved :: /tmp/1.xlsx

```

See also [WiKi](http://wiki.enchtex.info/handmade/postgres/psql2xlsx) page.
