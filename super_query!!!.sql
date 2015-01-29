INSERT INTO [spZEU] SELECT * FROM [MS Access;DATABASE=C:\KVPL\KVPLS.mdb;].[spZEU]

C#:

INSERT INTO [Bill_Master] SELECT * FROM [MS Access;DATABASE="+     
               "\\Data.mdb" + ";Jet OLEDB:Database Password=12345;].[Bill_Master]