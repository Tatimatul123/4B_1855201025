/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Pertemuan_10;

import java.sql.Statement;
import javax.swing.JOptionPane;

/**
 *
 * @author Bu Tika
 */
public class Update {

    Koneksi konek = new Koneksi();

    public void update(int nim, String nama, String alamat, String jenis) {

        try {
            konek.koneksi();
            Statement statement = konek.con.createStatement();

// data akan diambil lewat nama textField.getText
            String sql_nama = "update datamahasiswa set nama ='" + nama + "' where nim= '" + nim + "'";
            String sql_alamat = "update datamahasiswa set alamat ='" + alamat + "' where nim= '" + nim + "'";

            String sql_jenis = "update datamahasiswa set jenis ='" + jenis + "' where nim= '" + nim + "'";
            
            statement.executeUpdate(sql_nama);
            statement.executeUpdate(sql_alamat);
            statement.executeUpdate(sql_jenis);
            statement.close();

//// kemudian data yang disisipkan akan saya tampilkan lewat textArea 
//                menampilkan.setText(mah_nim.getText() + "\n" + mah_nama.getText() + "\n" + mah_alamat.getText() + "\n" + mah_fak_id.getText() + "\n" + jenis);
            JOptionPane.showMessageDialog(null, "Berhasil Diubah");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex);
        }
    }
}
