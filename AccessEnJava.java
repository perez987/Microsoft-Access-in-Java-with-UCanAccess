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