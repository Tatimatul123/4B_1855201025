/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Pertemuan_10;

import com.mysql.jdbc.Connection;
import java.sql.DriverManager;
import javax.swing.JOptionPane;

/**
 *
 * @author Bu Tika
 */
public class Koneksi {

    public static void main(String[] args) {

        Koneksi tesKoneksi = new Koneksi();
        tesKoneksi.koneksi();

    }

    Connection con = null;
    String statusKoneksi;

    public void koneksi() {
        try {
            String connectionURL = "jdbc:mysql://localhost/nilaikelasb";
            String username = "root";
            String password = "";
            Class.forName("com.mysql.jdbc.Driver");
            con = (Connection) DriverManager.getConnection(connectionURL, username, password);
            //JOptionPane.showMessageDialog(null, "Sukses Koneksi");
            statusKoneksi = "Berhasil Tersambung";
        } catch (Exception e) {
            //JOptionPane.showMessageDialog(null, e);
            //statusKoneksi = "Gagal Tersambung";
//System.exit(0);
        }
    }
}
