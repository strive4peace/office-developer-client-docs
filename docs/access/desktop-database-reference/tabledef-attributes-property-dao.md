---
title: TableDef.Attributes property (DAO)
TOCTitle: Attributes Property
ms:assetid: d01588c3-e94e-06bd-6568-974873411f2d
ms:mtpsurl: https://msdn.microsoft.com/library/Ff834701(v=office.15)
ms:contentKeyID: 48547828
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
---

# TableDef.Attributes property (DAO)


**Applies to**: Access 2013, Office 2013


Sets or returns a value that indicates one or more characteristics of a **TableDef** object. Read/write **Long**.

## Syntax

*expression* .Attributes

*expression* A variable that represents a **TableDef** object.

## Remarks

For an object not yet appended to a collection, this property is read/write.

## Example

This example displays the **Attributes** property for **Field**, **Relation**, and **TableDef** objects in the Northwind database.

```vb 
Sub AttributesX() 
 
 Dim dbsNorthwind As Database 
 Dim fldLoop As Field 
 Dim relLoop As Relation 
 Dim tdfloop As TableDef 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
 With dbsNorthwind 
 
 ' Display the attributes of a TableDef object's 
 ' fields. 
 Debug.Print "Attributes of fields in " & _ 
 .TableDefs(0).Name & " table:" 
 For Each fldLoop In .TableDefs(0).Fields 
 Debug.Print " " & fldLoop.Name & " = " & _ 
 fldLoop.Attributes 
 Next fldLoop 
 
 ' Display the attributes of the Northwind database's 
 ' relations. 
 Debug.Print "Attributes of relations in " & _ 
 .Name & ":" 
 For Each relLoop In .Relations 
 Debug.Print " " & relLoop.Name & " = " & _ 
 relLoop.Attributes 
 Next relLoop 
 
 ' Display the attributes of the Northwind database's 
 ' tables. 
 Debug.Print "Attributes of tables in " & .Name & ":" 
 For Each tdfloop In .TableDefs 
 Debug.Print " " & tdfloop.Name & " = " & _ 
 tdfloop.Attributes 
 Next tdfloop 
 
 .Close 
 End With 
 
End Sub 
 
```
## Example to Document Table Attributes

The attributes property is the sum of indicators such as whether a table is linked, or if it is hidden. This example loops through all the tables in the current database, and writes each table name and attribute information to the debug window.

```vb 
Sub LoopTables_ListAttributes( _
   )
   '190329 strive4peace

   On Error GoTo Proc_Err

   Dim db As DAO.Database _
      , tdf As DAO.TableDef
      
   Dim sTablename As String _
      , nCountTables As Long _
      , nAttributes As Long _
      , i As Integer _
      , vMsg As Variant
   
   Set db = CurrentDb
   
   nCountTables = 0

   Debug.Print "*** Tables and Attributes *** " & Now()
      Debug.Print Tab(5); "Name";
      Debug.Print Tab(40); "Attributes"
         
   For Each tdf In db.TableDefs
      nCountTables = nCountTables + 1
      With tdf
         sTablename = .Name
         nAttributes = .Attributes
         Debug.Print Tab(5); sTablename; Tab(40); nAttributes;
      End With 'tdf
      
      vMsg = Null
      
      'document attributes
      If (nAttributes And dbSystemObject) <> 0& Then
         vMsg = (vMsg + ", ") & "System"
      End If
      If (nAttributes And dbHiddenObject) <> 0& Then
         vMsg = (vMsg + ", ") & "Hidden"
      End If
      If (nAttributes And dbAttachExclusive) <> 0& Then
         vMsg = (vMsg + ", ") & "Attached exclusively"
      End If
      If (nAttributes And dbAttachSavePWD) <> 0& Then
         vMsg = (vMsg + ", ") & "User ID and password saved"
      End If
      If (nAttributes And dbAttachedODBC) <> 0& Then
         vMsg = (vMsg + ", ") & "Linked ODBC table"
      End If
      If (nAttributes And dbAttachedTable) <> 0& Then
         vMsg = (vMsg + ", ") & "Linked/Attached table"
      End If

      Debug.Print " " & vMsg
      
   Next tdf
   
   vMsg = "documented " & nCountTables & " tables"
   Debug.Print "-- " & vMsg
   vMsg = vMsg & vbCrLf & "Press Ctrl-G to open the debug window and look at the results"
   MsgBox vMsg, , "Done"
   
Proc_Exit:
   On Error Resume Next
   Set tdf = Nothing
   Set db = Nothing
   Exit Sub

Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   LoopTables_ListAttributes"
   
   Resume Proc_Exit
   Resume
End Sub
```

## Example to Set Hidden Attribute of a table

Here is code to set the hidden attribute of a table to true:

```vb 
Sub SetHidden(psTablename As String)
'190329 s4p
   'set the Hidden attribute of the passed table
   Dim db As DAO.Database _
      , tdf As DAO.TableDef
   Set db = CurrentDb
   Set tdf = db.TableDefs(psTablename)
   With tdf
      .Attributes = .Attributes + dbHiddenObject
   End With
   MsgBox "done"
   Set tdf = Nothing
   Set db = Nothing
End Sub
```

## See also

- [TableDefAttributeEnum enumeration (DAO)](tabledefattributeenum-enumeration-dao.md)

