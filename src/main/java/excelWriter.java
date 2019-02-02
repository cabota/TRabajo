
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class excelWriter {

    public  void Ecribirdatos (String []data , int totalfilas) throws FileNotFoundException, IOException {
        String[][] datosSpliteados = new String[totalfilas][];


        String filename = "C:\\Users\\luisc\\IdeaProjects\\guillemexcel\\result\\test.xlsx";
        OutputStream salida = new FileOutputStream(filename);

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja1 = libro.createSheet("hoja1");
        //cabecera de la hoja de excel
        String[] header = new String[]{"Clientes", "Ftra.Nº", "Tipo Hab.", "Nº.Pax", "Regimen", "Entrada", "Salida", "Importe Ftra", "Días", "Importe Base", "Diferencia", "D Negativa"};
        for (int a = 0; a < totalfilas; a++) {
            datosSpliteados[a] = data[a].split("_");

            //contenido de la hoja de excel

            String[][] document = new String[][]{{datosSpliteados[a][0], datosSpliteados[a][1], datosSpliteados[a][2], datosSpliteados[a][3], datosSpliteados[a][4], datosSpliteados[a][5], datosSpliteados[a][6], datosSpliteados[a][7], datosSpliteados[a][8], datosSpliteados[a][9], datosSpliteados[a][10], datosSpliteados[a][11]}
            };

            //poner negrita a la cabecera
            CellStyle style = libro.createCellStyle();
            Font font = libro.createFont();
            font.setBold(true);
            style.setFont(font);


            //generar los datos para el documento
            for (int i = 0; i <= document.length; i++) {
                XSSFRow row = hoja1.createRow(i);//se crea las filas
                for (int j = 0; j < header.length; j++) {
                    if (i == 0) {//para la cabecera
                        XSSFCell cell = row.createCell(j);//se crea las celdas para la cabecera, junto con la posición
                        cell.setCellStyle(style); // se añade el style crea anteriormente
                        cell.setCellValue(header[j]);//se añade el contenido
                    } else {//para el contenido
                        XSSFCell cell = row.createCell(j);//se crea las celdas para la contenido, junto con la posición
                        cell.setCellValue(document[i - 1][j]); //se añade el contenido
                    }
                }
            }
        }
            File file;
            file = new File(filename);
            try {
                FileOutputStream fileOuS = new FileOutputStream(file);
                if (file.exists()) {// si el archivo existe se elimina
                    file.delete();
                    System.out.println("Archivo eliminado");
                }
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


