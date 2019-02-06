import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class front {



    public JTextField RutaRead;
    public JTextField RutaSave;
    public JButton ejecutar;
    public JPanel jpanel;
    private JLabel Rutasavetext;


    public front() {
        ejecutar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                front front = new front();
                front.Run();
            }
        });
    }

    public void Run () {
        excelreader excel = new excelreader();

        String rutasave = this.RutaSave.getText();
        String rutaleer = this.RutaRead.getText();

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



