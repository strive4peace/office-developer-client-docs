---
title: FieldAttributeEnum enumeration (DAO)
TOCTitle: FieldAttributeEnum Enumeration
ms:assetid: 2dc6697c-d3e1-ce76-1b8b-fc60dc6f16a9
ms:mtpsurl: https://msdn.microsoft.com/library/Ff192118(v=office.15)
ms:contentKeyID: 48543977
ms.date: 09/18/2015
mtps_version: v=office.15
localization_priority: Normal
---

# FieldAttributeEnum enumeration (DAO)

**Applies to**: Access 2013, Office 2013, Access 2016, Office 2016, Office 365

Used with the **Attributes** property to determine attributes of a **Field** object.

|Name|Value|Description|
|:-----|:-----|:-----|
dbDescending|1|sort in descending (ZA) order, otherwise use ascending (AZ) order|
dbFixedField|1|The field size is fixed (default for Numeric fields).|
dbVariableField|2|The field size is variable (Text fields only).|
dbAutoIncrField|16|The field value for new records is automatically incremented to a unique Long integer that can't be changed (for Microsoft Access database engine database tables).|
dbUpdatableField|32|The field value can be changed.|
dbSystemField|8192|replication information; can't delete but can create a new table without it and append records.
dbHyperlinkField|32768|The field contains hyperlink information (Long Text fields only).

