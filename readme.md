# inoAccessSC

This repository holds the VBA code to export and import objects of an Access database.

After the export the code can be saved in a source code repository.

The code is located in the module [mdl_ExportDbObjects.bas](source/mdl_ExportDbObjects.bas) in the folder source.

After the import to the database the code is ready to use.

A demo is avilable in the sample folder [AccessSC.accdb](sample/AccessSC.accdb).

## Sub ExportDatabaseObjects

By default the export is stored in a default folder `sourcecode` which will be create if not present in the folder where the database file is located. 

The code exports the following:

* all tables but not linked tables with structure and data.

* all forms, reports and modules 
* all querys and macros, even the ones stored in a form. The can be identified by the "~" in the file name.

## Sub RestoreDatabaseObjectsFromFolder

By default the import is taken from the default folder `sourcecode`. If this folder or the given folder is not available the user is pompted to define the location of the folder. Otherwise the sub will be exited.

The following file will NOT be imported:

* all interal queries and macro files containing a "~"
* the module `mdl_ExportDbObjects` which is defined as constant `baseModule` at the beginning of the code.
* If the code is started from a form the name od the form should be storedin the constant `currentForm` at the beginning of the code and the varibale `blnCurrentForm` must be set to `True`.