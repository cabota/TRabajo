import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class front {



    public JTextField RutaRead;
    public JTextField RutaSave;
    public JButton ejecutar;
    public JPanel jpanel;




    public void Run () {
        excelreader excel = new excelreader();
        front front = new front();
        String rutasave = front.RutaSave.getText();
        String rutaleer = front.RutaRead.getText();

        excel.EjecutarLectura(rutasave, rutaleer);
    }

    public static void main(String args[]) {
        final front front = new front();
        final JFrame frame =new JFrame("Aplicacion Exportar");
        frame.setContentPane(new front().jpanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);





        }


    }



