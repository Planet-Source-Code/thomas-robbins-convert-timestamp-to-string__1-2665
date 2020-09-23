<div align="center">

## Convert TimeStamp To String


</div>

### Description

Converts the timestamp data type of SQL server into a string that is then suitable for use in a where clause
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thomas Robbins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-robbins.md)
**Level**          |Unknown
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thomas-robbins-convert-timestamp-to-string__1-2665/archive/master.zip)





### Source Code

```
Public Function CvtTimeStamp(TimeSt As String) As String
' This function will recieve the eight byte string of binary data
' returned as the value of a timestamp column and convert it into
' as string in the format 0x000000000000000 suitable for the use in an sql statement
' where clause
Dim HexValue As String
Dim K As Integer
For K = 1 To 8
    HexValue = HexValue & Right$("00" & Hex(AscB(MidB(TimeSt, K, 1))), 2)
Next K
CvtTimeStamp = "0x" & HexValue
End Function
```

