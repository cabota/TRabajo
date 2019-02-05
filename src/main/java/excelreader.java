
import org.apache.poi.hssf.usermodel.HSSFRow;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;



public class excelreader {
    String document = null;
    int posicionarray = 0;

    public String[] getData() {
        return data;
    }

    String [] data = new String[10000];

    int columnas = 13;

    public String conseguirsalida (String entrada , String dias) throws ParseException {
        String salida;
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yy");
        Date calcularfecha = formatter.parse(entrada);
        Calendar calendar = Calendar.getInstance();


        calendar.setTime(calcularfecha); // Configuramos la fecha que se recibe

        calendar.add(Calendar.DATE, +Integer.valueOf(dias));

        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM//yy");

        salida = sdf.format(calendar.getTime());


        return salida;
    }

    public String CalculoDispo ( int primerAdulto, int segundoAdulto ,int tercerAdulto,int primerNino,int segundoNino,int tercerNino){
        String dispo=null;
        int sumaadultos = primerAdulto+segundoAdulto+tercerAdulto;

        int sumaninos = primerNino+segundoNino+tercerNino;

        dispo = sumaadultos+"AD" + sumaninos +"N";

        return dispo;
    }

    public void escribirdatos (String [] datos,String save) throws IOException {
        excelWriter excelWriter = new excelWriter();
        excelWriter.Ecribirdatos(datos ,this.posicionarray,save);
    }

    public String []  guardarArray (String data, int posicionarray ){


        this.data[posicionarray] =data;


        return this.data;



    }


    public String CalcularDiferencia (double importeBAse , double importetotal){
        String diferencia = null;

        double calculo = importeBAse - importetotal;

        diferencia = String.valueOf(calculo);
       diferencia= diferencia.replace(".",",");
        String[] datoslipt = new String[1];

        datoslipt = diferencia.split(",");


        if(datoslipt[1].length()<2) {
            return datoslipt[0] + "." + datoslipt[1].substring(0, 1);
        }else{
            return datoslipt[0] + "." + datoslipt[1].substring(0, 2);
        }
    }
    public void EjecutarLectura (String rutaleer ,String save){
        excelreader excelreader = new excelreader();
        String numerofactura=null;
        String Clientes=null;
        String tipohabitacion=null;
        String distribucion=null;
        String regimen=null;
        String entrada=null;
        String salida=null;
        String dias=null;
        String importeBase=null;
        String importeCompleto = null;
        String diferencia=null;
        String diferencianega=null;
        String primerAdulto = null;
        String segundoAdulto =null;
        String tercerAdulto=null;
        String primerNino = null;
        String segundoNino=null;
        String tercerNino=null;
        String diferenciacalculada=null;
        String difernciapositiva = null;
        String diferncianegativa = null;


        try {

            File f = new File(rutaleer);
            File[] ficheros = f.listFiles();
            HSSFRow hssfRow;
            int cols = 0;
            for (int x = 0; x < ficheros.length; x++) {

                String rutaArchivoExcel = rutaleer + "\\" + ficheros[x].getName();

                FileInputStream inputStream = new FileInputStream(rutaArchivoExcel);
                Workbook workbook = new HSSFWorkbook(inputStream);
                Sheet firstSheet = workbook.getSheetAt(0);
                int rowlast = firstSheet.getLastRowNum();

                DataFormatter formatter = new DataFormatter();
                for (int r = 0; r <= rowlast; r++) {

                    hssfRow = (HSSFRow) firstSheet.getRow(r);
                    if (hssfRow == null) {
                        break;
                    } else {
                        for (int c = 0; c <= (cols = hssfRow.getLastCellNum()); c++) {


                            String contenidoCelda = formatter.formatCellValue(hssfRow.getCell(c));
                            // System.out.println("celda: " + c + " valor " + contenidoCelda);
                            switch (c){

                                case 0 : numerofactura = contenidoCelda;
                                case 3 : tipohabitacion = contenidoCelda;
                                case 7: primerAdulto = contenidoCelda;
                                case 8 : segundoAdulto = contenidoCelda;
                                case 9 : tercerAdulto = contenidoCelda;
                                case 13 : primerNino = contenidoCelda;
                                case 14 : segundoNino = contenidoCelda;
                                case 15 : tercerNino =contenidoCelda;
                                case 16 : regimen =contenidoCelda;
                                case 18 : Clientes =contenidoCelda;
                                case 24 : entrada = contenidoCelda;
                                case 25: dias = contenidoCelda;
                                case 31: importeBase = contenidoCelda;
                                case 38 : importeCompleto = contenidoCelda;

                            }


                        }



                    }
                    String fechsalida =excelreader.conseguirsalida(entrada,dias);
                    String dispo = excelreader.CalculoDispo(Integer.valueOf(primerAdulto),Integer.valueOf(segundoAdulto),Integer.valueOf(tercerAdulto),Integer.valueOf(primerNino),Integer.valueOf(segundoNino),Integer.valueOf(tercerNino));
                    diferenciacalculada = excelreader.CalcularDiferencia(Double.valueOf(importeBase.replace(".","").replace(",",".")),Double.valueOf(importeCompleto.replace(".","").replace(",",".")));
                    //aqui se añade el  array
                    if(Double.valueOf(diferenciacalculada)<0){
                        diferncianegativa = diferenciacalculada;
                        difernciapositiva = Double.toString(0);

                    }else{
                        difernciapositiva = diferenciacalculada;
                        diferncianegativa= Double.toString(0);

                    }



                    excelreader.document= Clientes+"_"+ numerofactura+"_" +tipohabitacion+"_" + dispo+"_" +regimen +"_" +entrada+"_" +fechsalida+"_" +importeBase+"_" + dias+"_" + importeCompleto+"_" +difernciapositiva+"_" +diferncianegativa;

                    excelreader.guardarArray(excelreader.document,excelreader.posicionarray);
                    excelreader.posicionarray = excelreader.posicionarray + 1;
                }



            }
            excelreader.escribirdatos(excelreader.getData(),save);
        }
        catch (Exception e) {
            e.printStackTrace();
        }

    }


    public static void main(String args[]) throws IOException, FileNotFoundException {


    }
}


