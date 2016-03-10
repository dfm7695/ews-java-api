package database;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import utility.Messenger;

public class DBUtil {
  private static final String USER = "cywusr";
  private static final String CONN_DEV = "jdbc:oracle:thin:@[ccpdbcscomp.isis.unc.edu]:1521:"
      + "cywprd";
  private static final String PW = "cywusr12#";

  public static Connection getConnection(DBList db) {
    switch (db) {
    case DEV:
      try {
        return DriverManager.getConnection(USER, CONN_DEV, PW);
      } catch (SQLException e) {
        processException(e, "DBUtil CONN_DEV:\n" + e.getMessage());
        return null;
      }
    default: // PROD
      try {
        return DriverManager.getConnection(USER, CONN_DEV, PW);
      } catch (SQLException e) {
        processException(e, "DBUtil CONN_PROD:\n" + e.getMessage());
        return null;
      }
    }
  }

  public static void processException(SQLException e, String body) {
    e.printStackTrace();
    System.err.println("Error message: " + e.getMessage());
    System.err.println("Error code: " + e.getErrorCode());
    System.err.println("SQL state: " + e.getSQLState());
    new Messenger().email(body);
  }
}