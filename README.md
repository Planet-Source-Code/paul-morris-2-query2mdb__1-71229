<div align="center">

## Query2MDB


</div>

### Description

I needed to query 2 password protected Access databases with a single SQL statement.

The Microsoft web page http://support.microsoft.com/kb/113701 was helpful and gave me hope that it may be possible, but was not too clear.

I use ADO with VB6 whereas the Microsoft example was for VB3 and the old ODBC. The Microsoft example worked almost straightaway for me if the databases had no password protection, it was the password protection that made it more difficult.

Eventually after quite a bit of experimentation I cracked it and I thought I must share this with fellow coders in case they have the same requirement.
 
### More Info
 
There are 2 databases with this code: -

1. BookSale_2002.mdb  - password = ABCD

2. BS2.mdb       - password = 1234

They are the BookSale.mdb database, supplied by Microsoft in Visual Studio, with the tables split between them.

The 2 databases should be located in the same folder as the VB files.


<span>             |<span>
---                |---
**Submitted On**   |2008-10-13 09:50:02
**By**             |[Paul Morris \#2](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-morris-2.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Query2MDB21304610132008\.zip](https://github.com/Planet-Source-Code/paul-morris-2-query2mdb__1-71229/archive/master.zip)








