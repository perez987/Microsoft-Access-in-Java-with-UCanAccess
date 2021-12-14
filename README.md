# Microsoft Access in Java with 64 bits UCanAccess driver

<table>
  <tr><td align=center><img src=jdbc.png></td></tr>
  <tr><td>Java code for console: How to connect to a Microsoft Access database using UCanAccess JDBC driver, sending SQL statements to the database and showing results, running the program from command line including dependencies</td></tr>
</table><br>

Since Java 8, JDBC-ODBC bridge is no longer included. There are some proprietary JDBC drivers to connect to MS Access but UCanAccess project is currently active, open source, easy to use and provides a JDBC driver built on top of Jackcess code.

Note: Jackcess, unlike UCanAccess, is a Java code library designed to read and write MS Access databases that is not a JDBC driver but a direct implementation of the features available to interact with MS Access databases. Its license is of Apache License type.

### Database

MS Access database named 50company.accdb, it has a single table called Contactos and this table has 3 fields: Id (auto-numeric) / Nombre (text) / Telefono (text).

### Dependencies

You need `ucanaccess.jdbc.UcanaccessDriver` (driver that allows connect with MS Access from Java) and the appropriate connection string in addition to 5 JAR files that are required as dependencies. They can be downloaded from the Maven repository but they are also included into the distribution package. When you unzip UCanAccess you will see something like this:

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

`AccessEnJava.java` source file is in the `D:\Java\JDBC` folder. The 5 JAR files plus `50company.accdb` database are in the `D:\Java\JDBC\lib` folder. To run the program from console, you have to include the main class, as is usually done, and also the dependencies, JAR files in this case. This is so because we are working in the simplest way, with text editor and command line. If you work with an IDE, the addition of dependencies is done from the tools of the environment itself and the result of the compilation incorporates what is necessary to be executed.

To call the dependencies, use the java executable with -cp modifier followed by the path to the current folder (D:\Java\JDBC in this example) + path to each of the required JAR files + main class.
`java -cp working-dir-path jar-files-path main-class`

The required command is different depending on the operating system. In **Windows** you have to use (everything goes on a single line):

`java -cp D:/Java/JDBC;D:/Java/JDBC/lib/ucanaccess-5.0.1.jar;D:/Java/JDBC/lib/commons-lang3-3.8.1.jar;D:/Java/JDBC/lib/commons-logging-1.2.jar;D:/Java/JDBC/lib/hsqldb-2.5.0.jar;D:/Java/JDBC/lib/jackcess-3.0.1.jar AccessEnJava`

Notice that the paths are separated by **;** and there are no spaces between them but there are spaces between paths and the name of the class at the end. As it is cumbersome to write or copy all the text each time, you can create a script file with **.bat** extension from which to run it comfortably with double click. In this case the text of the text file must be:

Notice that the paths are separated by **;** and there are no spaces between them but there are spaces between paths and the name of the class at the end. As it is cumbersome to write or copy all the text each time, you can create a script file with **.bat** extension from which to run it comfortably with double click. In this case the text of the text file must be:

```
@echo off
cd D:\Java\JDBC\
java -cp D:\Java\JDBC;D:\Java\JDBC/lib/ucanaccess-5.0.1.jar;D:\Java\JDBC/lib/commons-lang3-3.8.1.jar;D:\Java\JDBC/lib/commons-logging-1.2.jar;D:\Java\JDBC/lib/hsqldb-2.5.0.jar;D:\Java\JDBC/lib/jackcess-3.0.1.jar AccessEnJava
REM pause
REM exit
```
