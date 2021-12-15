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

You need ucanaccess.jdbc.UcanaccessDriver (driver that allows connect with MS Access from Java) and the connection string in addition to 5 JAR files that are required as dependencies. They can be downloaded from the Maven repository but they are also included into the distribution package. When you unzip UCanAccess you will see something like this:

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

AccessEnJava.java source file is in the D:\Java\JDBC folder. The 5 JAR files plus 50empresas.accdb database are in the D:\Java\JDBC\lib folder. To run the program from console, you have to include the main class, as is usually done, and also the dependencies.

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

We need 3 variables: connection, SQL statement and result obtained with returned rows.

```java
Connection conectar = null;
Statement sentencia = null;
ResultSet resultado = null;
// variable to detect keystrokes, used to stop
// the flow of the program until the user presses a key
Scanner s = new Scanner(System.in);
```

### Database connection

First you need to register the UCanAccess driver.

```java
Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
```

Connection string is built with `jdbc:ucanaccess://` followed by the path to the database. In order not to leave the path configured as a constant, `getAbsolutePath` method is used, it returns the path to the working folder.

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

### Java code

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

(Work in progress...)
