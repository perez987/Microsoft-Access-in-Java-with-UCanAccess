# Microsoft Access in Java with 64 bits UCanAccess driver

<table>
  <tr><td align=center><img src=img/jdbc.png></td></tr>
  <tr><td><b>Java exercise for console: How to connect to a Microsoft Access database using the JDBC UCanAccess driver, issue SQL statements to the database and display the results, run the program from the command line along with dependencies</b></td></tr>
</table><br>

Since Java 8, JDBC-ODBC bridge is no longer included. There are some proprietary JDBC drivers to connect to MS Access but the UCanAccess project is active, is open source, works well, is simple to use and provides a JDBC driver generated over Jackcess code.

Note: Jackcess, unlike UCanAccess, is a Java code library designed to read and write MS Access databases that is not a JDBC driver but a direct implementation of the features available to interact with MS Access databases. Its license is of the Apache License type.

### Database

MS Access database named 50empresas.accdb, it has a single table called _Contactos_ and this table has 3 fields: _Id_ (auto-numeric) / _Nombre_ (text) / _Telefono_ (text).

### Dependencies

You need ucanaccess.jdbc.UcanaccessDriver (driver that allows connect with MS Access from Java) and the connection string in addition to 5 JAR files that are needed as dependencies. They can be downloaded from the Maven repository but they are also included in UCanAccess. When you unzip the distribution package you get something like this:

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

### Compile the code

This is done in the usual way, with the command `javac AccessEnJava.java` in the folder where the code file is.

### Run the program with dependencies

In this exercise the AccessEnJava source code is in the disk folder D:\Java\JDBC; the 5 JAR files and the 50empresas.accdb database are in the folder D:\Java\JDBC\lib. To run the program from console, you have to include the main class, as usual, and also the dependencies.

To call dependencies use the java executable with -cp modifier followed by the path to the current folder (D:\Java\JDBC) + path to each of the required JAR files + main class.

`java -cp working-dir-path jar-files-path main-class`

Command is different depending on the operating system.
 
<table>
	<tr><td>Windows</td></tr>
</table>

Command line in Windows is (goes all on one line):
 
```
java -cp D:/Java/JDBC;D:/Java/JDBC/lib/ucanaccess-5.0.1.jar;D:/Java/JDBC/lib/commons-lang3-3.8.1.jar;D:/Java/JDBC/lib/commons-logging-1.2.jar;D:/Java/JDBC/lib/hsqldb-2.5.0.jar;D:/Java/JDBC/lib/jackcess-3.0.1.jar AccessEnJava
```

Notice that paths are separated by **;** and there are no spaces between them but between paths and the name of the class at the end. As it is cumbersome to write the whole text each time, you can create a script file with extension **.bat** from which to run it comfortably with double click. The text of the bat file must be:

```
@echo off
cd D:\Java\JDBC\
java -cp D:/Java/JDBC;D:/Java/JDBC/lib/ucanaccess-5.0.1.jar;D:/Java/JDBC/lib/commons-lang3-3.8.1.jar;D:/Java/JDBC/lib/commons-logging-1.2.jar;D:/Java/JDBC/lib/hsqldb-2.5.0.jar;D:/Java/JDBC/lib/jackcess-3.0.1.jar AccessEnJava
REM pause
REM exit
```
REM pause line can be left like this or uncommented by removing REM, what it does is stop the flow by asking to press a key to continue.

<table>
	<tr><td>macOS</td></tr>
</table>
	
Terminal command in macOS is (goes all on one line):

```
java -cp .:./lib/ucanaccess-5.0.1.jar:./lib/commons-lang3-3.8.1.jar:./lib/commons-logging-1.2.jar:./lib/hsqldb-2.5.0.jar:./lib/jackcess-3.0.1.jar AccessEnJava
```

Notice that the dot refers to the current folder and path separator is **:** instead of **;** as in Windows.\
A script file with .sh extension can be generated to make more comfortable to run the program:

```
#! /bin/bash
java -cp.:./lib/ucanaccess-5.0.1.jar:./lib/commons-lang3-3.8.1.jar:./lib/commons-logging-1.2.jar:./lib/hsqldb-2.5 .0.jar:./lib/jackcess-3.0.1.jar AccessEnJava
```

Instead of double-clicking, you have to run the .sh file directly in Terminal preceded by **./**.

### Required packages

It is necessary to import packages for the connection with MS Access in addition to the Scanner class that allows receiving from keyboard (to stop the flow of the program until the user presses a key).

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

### Declared variables

We need 3 variables: connection, SQL statement and result obtained with returned rows.

```java
Connection conectar = null;
Statement sentencia = null;
ResultSet resultado = null;
// variable to detect keystrokes, used to stop the flow of the program until the user presses a key
Scanner s = new Scanner(System.in);
```

### Database connection

First you have to register the UCanAccess driver.

```java
Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
```

Connection string is built with `jdbc:ucanaccess://` followed by the path to the database. To avoid leaving the path set as a constant, use the `getAbsolutePath` method that returns the path to the working folder.

```java
// get the absolute path to the project's working folder
// trick: create a new file and get its path
String ruta = new File(".").getAbsolutePath();
// getAbsolutePath returns the path to the folder but adds a period at the end
// we remove that last character with substring and length
ruta = ruta.substring(0, ruta.length()-1);
// MS Access DB name is added
// File.separatorChar is used, it places \ or / depending on the operating system
ruta = ruta + "lib" + File.separatorChar + "50empresas.accdb";
// Message
System.out.println ("Path to database: " + ruta + " \nPress a key to continue...");
// String ruta = "D:/JDBC/50empresas.accdb";
String dbURL = "jdbc:ucanaccess://" + ruta;
```

Connection is created with DriverManager class, path to the data source is passed as a parameter.

```java
conectar = DriverManager.getConnection(dbURL);
```

### Press a key to continue ...

Scanner class is declared as **s** and used instantiating a new string to collect the variable **s** (the keystroke).

```java
Scanner s = new Scanner(System.in);
// ...
String una = s.nextLine();
// ...
String dos = s.nextLine();
```

### SQL command objects

SQL command objects (_statement_) have different execution methods:

- ResultSet `executeQuery()` executes a SQL query and returns a ResultSet (table with returned rows), it is used to read the contents of the table and display the records (SELECT statements)
- int `executeUpdate()` executes a SQL query that must be INSERT, UPDATE or DELETE, it is used to modify the content of the table and returns the number of  affected rows
- boolean `execute()` executes a SQL query, true if the statement returns a set of rows and false if it returns an updates count or there is no result.

Here the statement is created with the connection's `createStatement` method and returns a ResultSet, the query is of the SELECT type so the command is of the `executeQuery` type. Records are displayed by looping through the ResultSet with its `getInt()` method into which the index of each column in the table is passed and returns the value of that cell.

```java
sentencia = conectar.createStatement();
resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos");
System.out.println("Id\tTel\u00e9fono\tNombre");
System.out.println("==\t========\t======");
while(resultado.next()) {
System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
}
```
This is the output:

```
Conection open.
The 50 records of the table will be displayed.
SELECT Id, Nombre, Telefono FROM Contactos
Press a key to continue...
 
Id      Teléfono        Nombre
==      ========        ======
1       966521455       AEK Goup Inc.
2       944444102       ACE Hardware CO.
3       963325874       ADI Systems Inc.
4       986523365       Administaff Inc.
5       923656987       ADVO Inc.
6       956231487       Aeroquip - Vickers Inc.
7       933659214       LST IND. CO
8       956231487       Agway Inc.
9       912548758       Air Products and Chemicals Inc.
10      965235898       Airborne Freight Corp.
...
```

The exercise continues with new SQL statements and the presentation of the results on screen. SELECT statements are tested with BETWEEN (_executeQuery_ command), INSERT and UPDATE statements (_executeUpdate_ commands), etc. At the end, resources used are emptied and connection is closed.

### Java code

This is the code of the main class.

```java
// paquetes para la conexion con MS Access y comandos SQL
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.io.*;
// paquete para usar la clase Scanner que permite recibir desde teclado
import java.util.Scanner;

public class AccessEnJava {

    public static void main(String[] args) {

    // variables conexion con la BD, sentencias SQL y filas obtenidas
    Connection conectar = null;
    Statement sentencia = null;
    ResultSet resultado = null;
		// variable para detectar pulsaciones de teclas, se usa para detener
		// el flujo del programa hasta que el usuario pulse una tecla
		Scanner s = new Scanner(System.in);

    // registrar la clase JDBC driver de Ucanaccess
    try {
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
    }
    // si la clase no se puede utilizar
    catch(ClassNotFoundException ex) {
         ex.printStackTrace();
    }

      try {
			// obtener la ruta absoluta a la carpeta de trabajo del proyecto
			// se emplea el truco de crear nuevo archivo y obtener su ruta
			String ruta = new File(".").getAbsolutePath();
			// getAbsolutePath devuelve la ruta a la carpeta pero agrega un punto al final
			// quitamos ese ultimo caracter con substring y length
			ruta = ruta.substring(0, ruta.length()-1);
			// y se agrega el nombre de la BD de MS Access
			// se usa File.separatorChar que coloca \ o / dependiendo del sistema operativo
			ruta = ruta + "lib" + File.separatorChar + "50empresas.accdb";
			System.out.println();
			System.out.println();
			// Mensaje informativo
			System.out.println ("Ruta a la base de datos: " + ruta + " \nPulsa una tecla para continuar...");
			// se para hasta pulsar una tecla
			String una = s.nextLine();
      //String ruta = "D:/JDBC/50empresas.accdb";
      String dbURL = "jdbc:ucanaccess://" + ruta;
			// Mensaje informativo
			System.out.println("Conexi\u00f3n abierta.\nSe van a mostrar los 50 registros de la tabla.");
      System.out.println("SELECT Id, Nombre, Telefono FROM Contactos");
      System.out.println ("Pulsa una tecla para continuar...");
			// se para hasta pulsar una tecla
			String dos = s.nextLine();
			//
      // crear la conexion usando la clase DriverManager
      conectar = DriverManager.getConnection(dbURL);
			//
			// MOSTRAR TODOS LOS REGISTROS
      // crear la sentencia SQL
      sentencia = conectar.createStatement();
      // Los objetos de comando (statement) SQL tienen distintos metodos de ejecucion:
      // - boolean execute() ejecuta una sentencia SQL, es true si la instrucción devuelve un conjunto 
      // de resultados y false si devuelve un recuento de actualizaciones o no hay ningún resultado.
      // - ResultSet executeQuery() ejecuta una sentencia SQL y devuelve el ResultSet generado por la consulta
      // se usa para leer el contenido de la tabla y mostrar los registros
      // - int executeUpdate() ejecuta una sentencia SQL que ha de ser SQL INSERT, UPDATE o DELETE
      // se usa para modificar el contenido de la tabla y devuelve el numero de filas afectadas
      resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos");
      // formatear la presentacion de los registros de la tabla
      System.out.println("Id\tTel\u00e9fono\tNombre");
      System.out.println("==\t========\t======");
       while(resultado.next()) {
          System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
      }
			// MOSTRAR SOLO REGISTROS CON ID DEL 12 AL 24
			System.out.println();
			System.out.println ("Se van a mostrar los registros con Id del 40 al 50.");
      System.out.println("SELECT Id, Nombre, Telefono FROM Contactos Where Id BETWEEN 40 AND 50");
      System.out.println ("Pulsa una tecla para continuar...");
			String tres = s.nextLine();
			resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos Where Id BETWEEN 40 AND 50");
			// mostrar los registros del 12 al 24
      System.out.println("Id\tTel\u00e9fono\tNombre");
      System.out.println("==\t========\t======");
       while(resultado.next()) {
          System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
			}
			// REEMPLAZAR LOS VALORES DE NOMBRE Y TELEFONO DE  UN REGISTRO
			System.out.println();
			System.out.println ("Se va a reemplazar Bay Networks Inc. 986523365 por Colubi Corp. 935469214.\nSe van a mostrar los registros con Id a partir de 40.");
      System.out.println("UPDATE Contactos SET Nombre ='Colubi Corp.', Telefono='935469214' WHERE Nombre ='Bay Networks Inc.'");
      System.out.println("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println ("Pulsa una tecla para continuar...");
			String nueve = s.nextLine();
			sentencia.executeUpdate("UPDATE Contactos SET Nombre ='Colubi Corp.', Telefono='935469214' WHERE Nombre ='Bay Networks Inc.'");
			// mostrar todos los registros
			resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println("Id\tTel\u00e9fono\tNombre");
      System.out.println("==\t========\t======");
       while(resultado.next()) {
          System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
			}
			// REVERTIR EL CAMBIO DE NOMBRE Y TELEFONO REALIZADO
			System.out.println();
			System.out.println ("Se va a revertir el cambio realizado.\nSe van a mostrar los registros con Id a partir de 40.");
      System.out.println("UPDATE Contactos SET Nombre ='Bay Networks Inc.', Telefono='986523365' WHERE Nombre ='Colubi Corp.');");
      System.out.println("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println ("Pulsa una tecla para continuar...");
			String DIEZ = s.nextLine();
			sentencia.executeUpdate("UPDATE Contactos SET Nombre ='Bay Networks Inc.', Telefono='986523365' WHERE Nombre ='Colubi Corp.'");
			// mostrar todos los registros
			resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println("Id\tTel\u00e9fono\tNombre");
      System.out.println("==\t========\t======");
      while(resultado.next()) {
          System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
			}
			// INSERTAR 15 REGISTROS NUEVOS
			System.out.println();
			System.out.println ("Se van a insertar 15 registros nuevos.\nSe van a mostrar los registros con Id a partir de 40.");
      String insertar = "INSERT INTO Contactos (Nombre, Telefono) VALUES ('TOP Group INC', '923652547');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Orion Capital Corp.', '900125458');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('IBM', '956985447');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('SCE Hardware Corp.', '945896325');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Faxter International Inc.', '936521452');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Lockheed Martin Corp.', '903056898');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Ray Networks Inc.', '985654125');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('MCI Communications Corp.', '900326587');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('DO Seidman LLP', '978563214');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Katerpillar Inc.', '932655096');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Data General Corp.', '955561023');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Gateway Inc.', '993201145');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('Hewlett-Packard Co.', '975556912');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('ICF Kaiser International Inc.', '990023654');" +
      "\nINSERT INTO Contactos (Nombre, Telefono) VALUES ('LDI Systems Inc.', '922569687');";
      System.out.println(insertar);
      System.out.println("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println ("Pulsa una tecla para continuar...");
			String cuatro = s.nextLine();
      // comandos de insercion (insertan 15 registros nuevos en la tabla)
      //sentencia.executeUpdate(insertar);
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('TOP Group INC', '923652547')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Orion Capital Corp.', '900125458')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('IBM', '956985447')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('SCE Hardware Corp.', '945896325')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Faxter International Inc.', '936521452')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Lockheed Martin Corp.', '903056898')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Ray Networks Inc.', '985654125')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('MCI Communications Corp.', '900326587')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('DO Seidman LLP', '978563214')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Katerpillar Inc.', '932655096')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Data General Corp.', '955561023')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Gateway Inc.', '993201145')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('Hewlett-Packard Co.', '975556912')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('ICF Kaiser International Inc.', '990023654')");
			sentencia.executeUpdate("INSERT INTO Contactos (Nombre, Telefono) VALUES ('LDI Systems Inc.', '922569687')");
			// mostrar todos los registros
			resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println("Id\tTel\u00e9fono\tNombre");
      System.out.println("==\t========\t======");
      while(resultado.next()) {
          System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
			}
			// BORRAR LOS REGISTROS INSERTADOS
			System.out.println();
			System.out.println ("Se van a eliminar los registros insertados.\nSe van a mostrar los registros con Id a partir de 40.");
      System.out.println("DELETE FROM Contactos WHERE Id > 50");
      System.out.println("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println ("Pulsa una tecla para continuar...");
			String seis = s.nextLine();
			// eliminar registros con Id del 31 al 50
			sentencia.executeUpdate("DELETE FROM Contactos WHERE Id > 50");
			// mostrar todos los registros
			resultado = sentencia.executeQuery("SELECT Id, Nombre, Telefono FROM Contactos Where Id > 39");
      System.out.println("Id\tTel\u00e9fono\tNombre");
      System.out.println("==\t========\t======");
      while(resultado.next()) {
          System.out.println(resultado.getInt(1) + "\t" + resultado.getString(3) + "\t" + resultado.getString(2));
			}
		}

      catch(SQLException ex){
          ex.printStackTrace();
      }
      finally {
          try {
              if(null != conectar) {
              // limpiar los recursos relacionados con la conexion
              resultado.close();
              sentencia.close();
              // cerrar la conexion
              conectar.close();
              System.out.println();
              // Mensaje informativo
              System.out.println("Conexi\u00f3n cerrada. \nPulsa una tecla para salir...");
              // se para hasta pulsar una tecla
              String ocho = s.nextLine();
              }
          }
          catch (SQLException ex) {
              ex.printStackTrace();
          }
        }
    }
}
```
