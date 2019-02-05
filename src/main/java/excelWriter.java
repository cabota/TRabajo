
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class excelWriter {

    public  void Ecribirdatos (String []data , int totalfilas,String save) throws FileNotFoundException, IOException {
        String[][] datosSpliteados = new String[totalfilas][];


        String filename = save+"\\test.xlsx";
        OutputStream salida = new FileOutputStream(filename);

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja1 = libro.createSheet("Resultado");
        //cabecera de la hoja de excel
        String[] header = new String[]{"Clientes", "Ftra.Nº", "Tipo Hab.", "Nº.Pax", "Regimen", "Entrada", "Salida", "Importe Ftra", "Días", "Importe Base", "Diferencia", "D Negativa"};
        CellStyle style = libro.createCellStyle();
        Font font = libro.createFont();
        font.setBold(true);
        style.setFont(font);



        XSSFRow row = hoja1.createRow(0);
        for (int j = 0; j < header.length; j++) {


            XSSFCell cell = row.createCell(j);//se crea las celdas para la cabecera, junto con la posición
            cell.setCellStyle(style); // se añade el style crea anteriormente
            cell.setCellValue(header[j]);//se añade el contenido
        }

         for (int a = 0; a < totalfilas; a++) {

                datosSpliteados[a] = data[a].split("_");


                XSSFRow row1 = hoja1.createRow(a+1);
                for (int k = 0; k < header.length; k++) {

                XSSFCell cell = row1.createCell(k);//se crea las celdas para la contenido, junto con la posición
                if (k == 0) {
                    cell.setCellValue(datosSpliteados[a][0]);

                }
                if (k == 1) {
                    cell.setCellValue(datosSpliteados[a][1]);


                }
                if (k == 2) {
                    cell.setCellValue(datosSpliteados[a][2]);

                }
                if (k == 3) {
                    cell.setCellValue(datosSpliteados[a][3]);

                }
                if (k == 4) {
                    cell.setCellValue(datosSpliteados[a][4]);

                }
                if (k == 5) {
                    cell.setCellValue(datosSpliteados[a][5]);

                }
                if (k == 6) {
                    cell.setCellValue(datosSpliteados[a][6]);

                }
                if (k == 7) {
                    cell.setCellValue(datosSpliteados[a][7]);

                }
                if (k == 8) {
                    cell.setCellValue(datosSpliteados[a][8]);

                }
                if (k == 9) {
                    cell.setCellValue(datosSpliteados[a][9]);

                }
                if (k == 10) {
                    cell.setCellValue(datosSpliteados[a][10]);

                }
                if (k == 11) {
                    cell.setCellValue(datosSpliteados[a][11]);

                }


            }
        }




            File file;
            file = new File(filename);
            try {
                FileOutputStream fileOuS = new FileOutputStream(file);

                libro.write(fileOuS);
                fileOuS.flush();
                fileOuS.close();
                System.out.println("Archivo Creado");

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


    }



    public static void main(String args[]) throws IOException, FileNotFoundException {

    }


}


