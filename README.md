# Overview
This project is a fork from the original and completed version! I just tried to simplify some things!

All the behaviour is encapsuled inside one class module called: 'sqlLite'

## List files:
* code.base // a module containing 'howto' code
* sqlLite.cls // class module wrapper of SQLite

## List of methods:

* Dim sqlLite: Set sqlLite = New sqlLite                  		// create an instance 

* sqlLite.openDb 																							// open the database

* sqlLite.selectQry 																					// select query command
sqlLite.execute 																						'execute a query command: create, update, delete

* sqlLite.data 																								// return an array(2d) of data - all the data from select
* sqlLite.header 																							// return an array(2d) of data / header from select

## List of properties:

* sqlLite.qtyColumns 																			// quantity of columns
* sqlLite.qtyRows 																						// quantity of rows
