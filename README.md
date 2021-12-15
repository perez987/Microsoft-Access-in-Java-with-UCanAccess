# Microsoft Access in Java with 64 bits UCanAccess driver

<table>
  <tr><td align=center><img src=jdbc.png></td></tr>
  <tr><td>Java code for console: How to connect to a Microsoft Access database using UCanAccess JDBC driver, sending SQL statements to the database and showing results, running the program from command line including dependencies</td></tr>
</table><br>

Since Java 8, JDBC-ODBC bridge is no longer included. There are some proprietary JDBC drivers to connect to MS Access but UCanAccess project is currently active, open source, easy to use and provides a JDBC driver built on top of Jackcess code.

Note: Jackcess, unlike UCanAccess, is a Java code library designed to read and write MS Access databases that is not a JDBC driver but a direct implementation of the available features to interact with MS Access databases. Its license is of Apache License type.

### Database

MS Access database named 50empresas.accdb, it has a single table called _Contactos_ and this table has 3 fields: _Id_ (auto-numeric) / _Nombre_ (text) / _Telefono_ (text).

### Dependencies

You need ucanaccess.jdbc.UcanaccessDriver (driver that allows connect with MS Access from Java) and the appropriate connection string in addition to 5 JAR files that are required as dependencies. They can be downloaded from the Maven repository but they are also included into the distribution package. When you unzip UCanAccess you will see something like this:

```
ucanaccess-5.0.1-package
├── lib
│   ├── commons-lang3-3.8.1.jar
│   ├── commons-logging-1.2.jar
│   ├── hsqldb-2.5.0.jar
│   └── jackcess-3.0.1.jar
└── ucanaccess-5.0.1.jar
```

These files must be called when running the program from the command line.

### Running the code with its dependencies

AccessEnJava.java source file is in the D:\Java\JDBC folder. The 5 JAR files plus 50empresas.accdb database are in the D:\Java\JDBC\lib folder. To run the program from console, you have to include the main class, as is usually done, and also the dependencies, JAR files in this case.

To call the dependencies, use the java executable with -cp modifier followed by the path to the current folder (D:\Java\JDBC in this example) + path to each of the required JAR files + main class.\
`java -cp working-dir-path jar-files-path main-class`

The required command is different depending on the operating system. In **Windows** you have to use (everything goes on a single line):

```
java -cp D:/Java/JDBC;D:/Java/JDBC/lib/ucanaccess-5.0.1.jar;D:/Java/JDBC/lib/commons-lang3-3.8.1.jar;D:/Java/JDBC/lib/commons-logging-1.2.jar;D:/Java/JDBC/lib/hsqldb-2.5.0.jar;D:/Java/JDBC/lib/jackcess-3.0.1.jar AccessEnJava
```

Notice that paths are separated by **;** and there are no spaces between them but there are spaces between paths and the name of the class at the end. As it is cumbersome to write the whole text each time, you can create a script file with **.bat** extension from which to run it comfortably with double click. In this case the text of the text file must be:

```
@echo off
cd D:\Java\JDBC\
java -cp D:/Java/JDBC;D:/Java/JDBC/lib/ucanaccess-5.0.1.jar;D:/Java/JDBC/lib/commons-lang3-3.8.1.jar;D:/Java/JDBC/lib/commons-logging-1.2.jar;D:/Java/JDBC/lib/hsqldb-2.5.0.jar;D:/Java/JDBC/lib/jackcess-3.0.1.jar AccessEnJava
REM pause
REM exit
```
REM pause line can be left like this or uncommented by removing REM, what it does is stop the flow asking to press a key to continue.

The equivalent Terminal command in **macOS** would be (it goes all on one line):

```
java -cp .:./lib/ucanaccess-5.0.1.jar:./lib/commons-lang3-3.8.1.jar:./lib/commons-logging-1.2.jar:./lib/hsqldb-2.5.0.jar:./lib/jackcess-3.0.1.jar AccessEnJava
```

Notice that period refers to the current folder and path separator is **:** instead of **;** as in Windows.\
A script file with .sh extension can be generated to make more comfortable to run the program:

```
#! /bin/bash
java -cp.:./lib/ucanaccess-5.0.1.jar:./lib/commons-lang3-3.8.1.jar:./lib/commons-logging-1.2.jar:./lib/hsqldb-2.5 .0.jar:./lib/jackcess-3.0.1.jar AccessEnJava
```

In this case, instead of double-clicking, you have to run the .sh file directly in Terminal preceded by **./**.

### Required packages

It is necessary to import packages for the connection with MS Access in addition to the Scanner class that allows receiving from the keyboard (to stop the flow of the program until the user presses a key).

```java
// packages for the connection with MS Access and SQL commands
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.io.*;
// package to use the Scanner class that allows receiving from keyboard
import java.util.Scanner;
```

3 variables are used: connection, SQL statement and result obtained with returned rows.

```java
Connection conectar = null;
Statement sentencia = null;
ResultSet resultado = null;
// variable to detect keystrokes, used to stop
// the flow of the program until the user presses a key
Scanner s = new Scanner(System.in);
```
