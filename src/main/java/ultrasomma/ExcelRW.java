package ultrasomma;


import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.TextField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.Normalizer;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.regex.Pattern;

public class ExcelRW {

     private static XSSFWorkbook wb;
     private static Sheet sh;
     private static FileInputStream fis;
     private static FileOutputStream fos;
     private static Row row;
     private static Cell cell;
     private static String medValeria;
     private static String medGerusa;
     private static String medLaurise;
     private static String procedimentos;
     private static String unimedPFV;
     private static String unimedPFG;
     private static String afrafep;
     private static String caixa;
     private static String cassi;
     private static String embrapa;
     private static String unimed;
     private static String driver;
     private static String url;
     private static String user;
     private static String password;
     private static Connection con;
     private static Statement stm;
     private static String sql;
     private static ResultSet query;
     private static int cont;
     private static int rowEnd;
     private static Date data;
     private static String strData;
     private static SimpleDateFormat format;
     private static String nome;
     private static String procedimento;
     private static Double valor;
     private static String strValor;
     private static String medica;
     private static String isError;
     private static boolean Vtrue;
     private static boolean Ltrue;
     private static boolean Gtrue;
     private static boolean Ptrue;
     private static boolean Etrue;
     private static File file;
     private static FileWriter wr;


     //Afrafep
     private static List<String> listAfrafep;
     private static String quantG1;
     private static String totalG1;
     private static String quantL1;
     private static String totalL1;
     private static String quantV1;
     private static String totalV1;
     private static String quantP1;
     private static String totalP1;
     private static String  erros1;

     //Caixa
     private static List<String> listCaixa;
     private static String quantG2;
     private static String totalG2;
     private static String quantL2;
     private static String totalL2;
     private static String quantV2;
     private static String totalV2;
     private static String quantP2;
     private static String totalP2;
     private static String erros2;

     //Cassi
     private static List<String> listCassi;
     private static String quantG3;
     private static String totalG3;
     private static String quantL3;
     private static String totalL3;
     private static String quantV3;
     private static String totalV3;
     private static String quantP3;
     private static String totalP3;
     private static String erros3;

     //Embrapa
     private static List<String> listEmbrapa;
     private static String quantG4;
     private static String totalG4;
     private static String quantL4;
     private static String totalL4;
     private static String quantV4;
     private static String totalV4;
     private static String quantP4;
     private static String totalP4;
     private static String erros4;

     //Unimed
     private static List<String> listUnimed;
     private static String quantG5;
     private static String filmeG5;
     private static String medicG5;
     private static String materG5;
     private static String totalG5;
     private static String quantL5;
     private static String filmeL5;
     private static String medicL5;
     private static String materL5;
     private static String totalL5;
     private static String quantV5;
     private static String filmeV5;
     private static String medicV5;
     private static String materV5;
     private static String totalV5;
     private static String quantP5;
     private static String totalP5;
     private static String erros5;

     //Unimed  PF 
     private static List<String> listUnimedPF;
     private static String quantG6;
     private static String filmeG6;
     private static String medicG6;
     private static String materG6;
     private static String totalG6;
     private static String quantL6;
     private static String filmeL6;
     private static String medicL6;
     private static String materL6;
     private static String totalL6;
     private static String quantV6;
     private static String filmeV6;
     private static String medicV6;
     private static String materV6;
     private static String totalV6;
     private static String quantP6;
     private static String totalP6;
     private static String erros6;
     

     //Ultrasomma
     private static List<String> listUltrasomma;
     private static String quantG7;
     private static String totalG7;
     private static String quantL7;
     private static String totalL7;
     private static String quantV7;
     private static String totalV7;
     private static String quantP7;
     private static String totalP7;
     private static String erros7;


     public static void lerExcels(ObservableList<String> arquivos) throws SQLException, ClassNotFoundException, IOException {

          for (String arquivo : arquivos) {
               if(arquivo.contains("unimed pf valeria")){
                    unimedPFV = arquivo;
               } else if(arquivo.contains("unimed pf gerusa")){
                    unimedPFG = arquivo;
               } else if (arquivo.contains("valeria")) {
                    medValeria = arquivo;
               } else if (arquivo.contains("gerusa")) {
                    medGerusa = arquivo;
               } else if (arquivo.contains("laurise")) {
                    medLaurise = arquivo;
               } else if (arquivo.contains("afrafep")) {
                    afrafep = arquivo;
               } else if (arquivo.contains("caixa")) {
                    caixa = arquivo;
               } else if (arquivo.contains("cassi")) {
                    cassi = arquivo;
               } else if (arquivo.contains("embrapa")) {
                    embrapa = arquivo;
               } else if (arquivo.contains("unimed")) {
                    unimed = arquivo;
               } else if(arquivo.contains("procedimentos")) {
                    procedimentos = arquivo;
               }
          }



          //Prints Files
//          printPath(medValeria);
//          printPath(medGerusa);
//          printPath(medLaurise);
//          printPath(afrafep);
//          printPath(caixa);
//          printPath(cassi);
//          printPath(embrapa);
//          printPath(unimed);

          //Setting Database
          driver = "org.h2.Driver";
          url = "jdbc:h2:mem:";
          user = "Thorinr";
          password = "vaporubi";
          Class.forName(driver);
          System.out.println("Conectando Banco de Dados...");
          con = DriverManager.getConnection(url,user, password);
          System.out.println("Banco de Dados Conectado!");
          stm = con.createStatement();

          //Criando Tabelas das Médicas
          if(medValeria != null ){
               tabelasMedicas(medValeria, "MedValeria");
          }
          if(medGerusa != null ){
               tabelasMedicas(medGerusa, "MedGerusa");
          }
          if(medLaurise != null ){
               tabelasMedicas(medLaurise, "MedLaurise");
          }
          if(procedimentos != null ){
               tabelasMedicas(procedimentos, "Procedimentos");
          }

//          printTabela("MedValeria");
//          printTabela("MedGerusa");
//          printTabela("MedLaurise");
//          printTabela("Procedimentos");


          sql =   "CREATE TABLE Exames ( " +
                  "   data DATE , " +
                  "   nome VARCHAR(255) , " +
                  "   procedimento VARCHAR(255) , " +
                  "   valor DOUBLE , " +
                  "   convenio VARCHAR(255) ," +
                  "   medica VARCHAR(255) ," +
                  "   id int NOT NULL AUTO_INCREMENT," +
                  "   PRIMARY KEY (id)" +
                  ");";

          stm.executeUpdate(sql);
          System.out.println("\nTabela Exames Criada!");

          //Alimentando Tabelas de Exames
          if(afrafep != null ){
               tabelaExames(afrafep, "Afrafep");
          }
          if(caixa != null ){
               tabelaExames(caixa, "Caixa");
          }
          if(cassi != null ){
               tabelaExames(cassi, "Cassi");
          }
          if(embrapa != null ) {
               tabelaExames(embrapa, "Embrapa");
          }
          if(unimed != null ){
               tabelaExames(unimed, "Unimed");

          }
          if(unimedPFV != null ){
               tabelaExames(unimedPFV, "Unimed PF");

          }if(unimedPFG != null ){
               tabelaExames(unimedPFG, "Unimed PF");

          }
          



//
//          Print das Tabelas
//          printTabela("MedValeria");
//          printTabela("MedGerusa");
//          printTabela("MedLaurise");
//          printExames("Exames");
          corrigirErros();
//          printExames("Exames");
          writeExcelOK();
          writeExcelErros();

//          searchExame("dagoberto");
//          printErros();

     }

     private static void tabelaExames(String file, String name) throws IOException, SQLException {

          fis = new FileInputStream(file);
          wb = new XSSFWorkbook(fis);
          sh = wb.getSheetAt(0);
          rowEnd = sh.getLastRowNum();
          format =  new SimpleDateFormat("dd/MM/yyyy");


          for (int i = 1; i <= rowEnd; i++) {
               row = sh.getRow(i);
               Iterator<Cell> cellIterator = row.cellIterator();
               if(checkIfRowIsEmpty(row)){
                    break;
               }
               cont = 0;
               Vtrue = false;
               Ltrue = false;
               Gtrue = false;
               Ptrue = false;
               Etrue = false;
               strData = null;
               nome = null;
               strValor = null;
               medica = "";
               while (cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if (cont == 0) {
                         strData = cell.getStringCellValue();
                         if(strData == null) {
                              Etrue = true;
                         }
//                         }else{
//                              strData = format.format(data);
//                         }
                    } else if (cont == 1) {
                        nome = cell.getStringCellValue();
                        if(nome == null){
                              Etrue = true;
                         }else {
                             nome = normalizeString(cell.getStringCellValue());
                        }
                    } else if (cont == 2) {
                         procedimento = cell.getStringCellValue();
                         if(procedimento == null){
                              Etrue = true;
                         }
                    } else if (cont == 3) {
                         valor = cell.getNumericCellValue();
                         if(valor == null){
                              Etrue = true;
                         }else{
                              strValor = String.format(Locale.US, "%.2f", valor);
                         }
                    }
                    cont++;

               }

               //Finding which Doctor made
               if (!procedimento.equals("Filme") && !procedimento.equals("Medicamento") && !procedimento.equals("Material")) {
                    //Valeria
                    sql = String.format("select * from MedValeria where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            "and procedimento = '%s';", strData, nome, procedimento);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Vtrue = true;
                    }
                    //Laurise
                    sql = String.format("select * from MedLaurise where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            "and procedimento = '%s';", strData, nome, procedimento);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Ltrue = true;
                    }
                    //Gerusa
                    sql = String.format("select * from MedGerusa where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            "and procedimento = '%s';", strData, nome, procedimento);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Gtrue = true;
                    }
                    //Procedimentos
                    sql = String.format("select * from Procedimentos where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            "and procedimento = '%s';", strData, nome, procedimento);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Ptrue = true;
                    }
               } else {
                    sql = String.format("select * from MedValeria where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            ";", strData, nome);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Vtrue = true;
                    }
                    //Laurise
                    sql = String.format("select * from MedLaurise where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            ";", strData, nome);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Ltrue = true;
                    }
                    //Gerusa
                    sql = String.format("select * from MedGerusa where " +
                            "data = TO_DATE('%s', 'DD/MM/YYYY') " +
                            "and nome = '%s' " +
                            ";", strData, nome);
                    query = stm.executeQuery(sql);
                    if (query.next()) {
                         Gtrue = true;
                    }
               }

               if (Vtrue && !Ltrue && !Gtrue && !Ptrue && !Etrue) {
                    medica = "Valéria";
               } else if (Ltrue && !Vtrue && !Gtrue && !Ptrue && !Etrue) {
                    medica = "Laurise";
               } else if (Gtrue && !Ltrue && !Vtrue && !Ptrue && !Etrue) {
                    medica = "Gerusa";
               }else if(Ptrue && !Vtrue && !Gtrue && !Ltrue  && !Etrue){
                    medica = "Procedimentos";
               }else{
                    medica = "error";
               }

               sql = String.format("INSERT INTO exames VALUES ( " +
                       "TO_DATE('%s', 'DD/MM/YYYY'), " +
                       "'%s', " +
                       "'%s', " +
                       "%s, " +
                       "'%s', " +
                       "'%s'," +
                       "NULL" +
                       ");", strData, nome, procedimento, strValor, name, medica);
               stm.executeUpdate(sql);
          }
          System.out.println("\nTabela Exames Alimentada Com Convênio " + name +" !");

     }

     private static void printExames(String name) throws SQLException {

          sql = String.format("select * from %s;", name);
          query = stm.executeQuery(sql);
          System.out.println("\nTabela " + name);
          while(query.next()){
               System.out.println(query.getDate(1) + "\t\t" + query.getString(2) + "\t\t" + query.getString(3) + "\t\t" + query.getDouble(4) + "\t\t" + query.getString(5) + "\t\t" + query.getString(6));
          }
     }

     private static void printTabela(String name) throws SQLException {

          sql = String.format("select * from %s;", name);
          query = stm.executeQuery(sql);
          System.out.println("\nTabela " + name);
          while(query.next()){
               System.out.println(query.getDate(1) + "\t\t" + query.getString(2) + "\t\t" + query.getString(3) + "\t\t" + query.getDouble(4) + "\t\t" + query.getString(5));
          }
     }



     private static void printPath(String file){
          System.out.println(file);
     }

     private static void tabelasMedicas(String file, String name) throws IOException, SQLException {

          fis = new FileInputStream(file);
          wb = new XSSFWorkbook(fis);
          sh = wb.getSheetAt(0);
          rowEnd = sh.getLastRowNum();
          format =  new SimpleDateFormat("dd/MM/yyyy");

          //Create Tabelas
          sql = String.format("CREATE TABLE %s ( " +
                  "   data DATE, " +
                  "   nome VARCHAR(255), " +
                  "   procedimento VARCHAR(255), " +
                  "   valor DOUBLE, " +
                  "   error VARCHAR(50)" +
                  ");",name);

          stm.executeUpdate(sql);
          System.out.println("\nTabela " + name + " Criada!");

          for (int i = 1; i <= rowEnd; i++) {
               row = sh.getRow(i);
               if(checkIfRowIsEmpty(row)){
                    break;
               }
               Iterator<Cell> cellIterator = row.cellIterator();
               cont = 0;
               isError = "false";
               strData = null;
               nome = null;
               strValor = null;
               while (cellIterator.hasNext()) {
                    cell = cellIterator.next();

                    if (cont == 0) {
                         strData = cell.getStringCellValue();
                         if (strData == null) isError = "true";
                    }
//                         }else{
//                              strData = format.format(data);
//                              System.out.println(strData);
//                         }
                    else if(cont == 1){
                         nome = cell.getStringCellValue();
                         if(nome == null){
                              isError = "true";
                         }else{
                              nome = normalizeString(cell.getStringCellValue());
//                              System.out.println(nome);
                         }
                    }else if(cont == 2){
                         procedimento = cell.getStringCellValue();
                         if(procedimento == null){
                              isError = "true";
                         }
//                         System.out.println(procedimento);
                    }else if(cont == 3){
                         valor = cell.getNumericCellValue();
//                         System.out.println(valor);
                         if(valor == null){
                              isError = "true";
                         }else{
                              strValor = String.format(Locale.US, "%.2f", valor);

                         }
                    }
                    cont++;

               }

               sql = String.format("INSERT INTO %s VALUES ( " +
                       "TO_DATE('%s', 'DD/MM/YYYY'), " +
                       "'%s', " +
                       "'%s', " +
                       "%s , " +
                       "%s" +
                       ");", name, strData, nome, procedimento, strValor, isError);
               stm.executeUpdate(sql);
          }
          System.out.println("Tabela " + name + " Alimentada!");
     }

     public static Collection<List<String>> getData() throws SQLException, IOException {

          //Afrafep
          quantG1 = getQuant("Afrafep", "Gerusa");
          totalG1 = getTotal("Afrafep", "Gerusa");
          quantL1 = getQuant("Afrafep", "Laurise");
          totalL1 = getTotal("Afrafep", "Laurise");
          quantV1 = getQuant("Afrafep", "Valéria");
          totalV1 = getTotal("Afrafep", "Valéria");
          quantP1 = getQuant("Afrafep", "Procedimentos");
          totalP1 = getTotal("Afrafep", "Procedimentos");
          erros1 = getErros("Afrafep");
          listAfrafep = new ArrayList<>();
          listAfrafep.addAll(Arrays.asList(quantG1,totalG1,quantL1,totalL1,quantV1,totalV1,quantP1,totalP1, erros1));

          //Caixa
          quantG2 = getQuant("Caixa", "Gerusa");
          totalG2 = getTotal("Caixa", "Gerusa");
          quantL2 = getQuant("Caixa", "Laurise");
          totalL2 = getTotal("Caixa", "Laurise");
          quantV2 = getQuant("Caixa", "Valéria");
          totalV2 = getTotal("Caixa", "Valéria");
          quantP2 = getQuant("Caixa", "Procedimentos");
          totalP2 = getTotal("Caixa", "Procedimentos");
          erros2 = getErros("Caixa");
          listCaixa = new ArrayList<>();
          listCaixa.addAll(Arrays.asList(quantG2,totalG2,quantL2,totalL2,quantV2,totalV2,quantP2,totalP2, erros2));


          //Cassi
          quantG3 = getQuant("Cassi", "Gerusa");
          totalG3 = getTotal("Cassi", "Gerusa");
          quantL3 = getQuant("Cassi", "Laurise");
          totalL3 = getTotal("Cassi", "Laurise");
          quantV3 = getQuant("Cassi", "Valéria");
          totalV3 = getTotal("Cassi", "Valéria");
          quantP3 = getQuant("Cassi", "Procedimentos");
          totalP3 = getTotal("Cassi", "Procedimentos");
          erros3 = getErros("Cassi");
          listCassi = new ArrayList<>();
          listCassi.addAll(Arrays.asList(quantG3,totalG3,quantL3,totalL3,quantV3,totalV3,quantP3,totalP3, erros3));

          //Embrapa
          quantG4 = getQuant("Embrapa", "Gerusa");
          totalG4 = getTotal("Embrapa", "Gerusa");
          quantL4 = getQuant("Embrapa", "Laurise");
          totalL4 = getTotal("Embrapa", "Laurise");
          quantV4 = getQuant("Embrapa", "Valéria");
          totalV4 = getTotal("Embrapa", "Valéria");
          quantP4 = getQuant("Embrapa", "Procedimentos");
          totalP4 = getTotal("Embrapa", "Procedimentos");
          erros4 = getErros("Embrapa");
          listEmbrapa = new ArrayList<>();
          listEmbrapa.addAll(Arrays.asList(quantG4,totalG4,quantL4,totalL4,quantV4,totalV4,quantP4,totalP4, erros4));

          //Unimed
          quantG5 = getQuant("Unimed", "Gerusa");
          filmeG5 = getFilme("Unimed", "Gerusa");
          medicG5 = getMedicamento("Unimed", "Gerusa");
          materG5 =getMaterial("Unimed", "Gerusa");
          totalG5 = getTotal("Unimed", "Gerusa");
          quantL5 = getQuant("Unimed", "Laurise");
          filmeL5 = getFilme("Unimed", "Laurise");
          medicL5 =getMedicamento("Unimed", "Laurise");
          materL5 =getMaterial("Unimed", "Laurise");
          totalL5 = getTotal("Unimed", "Laurise");
          quantV5 = getQuant("Unimed", "Valéria");
          filmeV5 = getFilme("Unimed", "Valéria");
          medicV5 =getMedicamento("Unimed", "Valéria");
          materV5 =getMaterial("Unimed", "Valéria");
          totalV5 = getTotal("Unimed", "Valéria");
          quantP5 = getQuant("Unimed", "Procedimentos");
          totalP5 = getTotal("Unimed", "Procedimentos");
          erros5 = getErros("Unimed");

          listUnimed = new ArrayList<>();
          listUnimed.addAll(Arrays.asList(quantG5,filmeG5,medicG5,materG5,totalG5,quantL5,filmeL5,medicL5,materL5,totalL5,quantV5,filmeV5,medicV5,materV5,totalV5,quantP5,totalP5, erros5));
          
          //Unimed PF
          //Unimed
          quantG6 = getQuant("Unimed PF", "Gerusa");
          filmeG6 = getFilme("Unimed PF", "Gerusa");
          medicG6 = getMedicamento("Unimed PF", "Gerusa");
          materG6 =getMaterial("Unimed PF", "Gerusa");
          totalG6 = getTotal("Unimed PF", "Gerusa");
          quantL6 = getQuant("Unimed PF", "Laurise");
          filmeL6 = getFilme("Unimed PF", "Laurise");
          medicL6 =getMedicamento("Unimed PF", "Laurise");
          materL6 =getMaterial("Unimed PF", "Laurise");
          totalL6 = getTotal("Unimed PF", "Laurise");
          quantV6 = getQuant("Unimed PF", "Valéria");
          filmeV6 = getFilme("Unimed PF", "Valéria");
          medicV6 =getMedicamento("Unimed PF", "Valéria");
          materV6 =getMaterial("Unimed PF", "Valéria");
          totalV6 = getTotal("Unimed PF", "Valéria");
          quantP6 = getQuant("Unimed PF", "Procedimentos");
          totalP6 = getTotal("Unimed PF", "Procedimentos");
          erros6 = getErros("Unimed PF");

          listUnimedPF = new ArrayList<>();
          listUnimedPF.addAll(Arrays.asList(quantG6,filmeG6,medicG6,materG6,totalG6,quantL6,filmeL6,medicL6,materL6,totalL6,quantV6,filmeV6,medicV6,materV6,totalV6,quantP6,totalP6, erros6));
//
//          //Ultrasomma
          quantG7 = addQuant(quantG1, quantG2, quantG3, quantG4, quantG5, quantG6);
          totalG7 = addTotal(totalG1,totalG2,totalG3,totalG4,totalG5 ,totalG6);
          quantL7 = addQuant(quantL1, quantL2, quantL3, quantL4, quantL5, quantL6);
          totalL7 = addTotal(totalL1,totalL2,totalL3,totalL4,totalL5, totalL6);
          quantV7 = addQuant(quantV1, quantV2, quantV3, quantV4, quantV5, quantV6);
          totalV7 = addTotal(totalV1,totalV2,totalV3,totalV4,totalV5, totalV6);
          quantP7 = addQuant(quantP1, quantP2, quantP3, quantP4, quantP5, quantP6);
          totalP7 = addTotal(totalP1,totalP2,totalP3,totalP4,totalP5, totalP6);
          erros7 = addQuant(erros1, erros2, erros3,erros4,erros5, erros6);

          listUltrasomma = new ArrayList<>();
          listUltrasomma.addAll(Arrays.asList(quantG7,totalG7,quantL7,totalL7,quantV7,totalV7,quantP7,totalP7,erros7));

          writeResume(listAfrafep, listCaixa, listCassi, listEmbrapa, listUnimed, listUnimedPF, listUltrasomma);


          return Arrays.asList( listAfrafep, listCaixa, listCassi, listEmbrapa, listUnimed, listUnimedPF, listUltrasomma);


     }

     private static String getQuant(String convenio, String medica) throws SQLException {
          sql = String.format("select * from exames" +
                  " where procedimento != 'Filme' and " +
                  "procedimento != 'Material' and" +
                  " procedimento != 'Medicamento' and" +
                  " convenio = '%s' and" +
                  " medica = '%s';", convenio, medica);
          query = stm.executeQuery(sql);
          int quant = 0;
          while(query.next()){
               quant++;
          }
          return String.valueOf(quant);
     }

     private static String  getFilme(String convenio, String medica) throws SQLException {
          sql = String.format("select * from exames" +
                  " where procedimento = 'Filme' and convenio = '%s' and medica = '%s';", convenio, medica);
          query = stm.executeQuery(sql);
          double filme = 0.0;
          while(query.next()){
               filme += query.getDouble(4);
          }
          return String.format("%.2f", filme);
     }

     private static String getTotal(String convenio, String medica) throws SQLException {
          sql = String.format("select * from exames" +
                  " where convenio = '%s' and medica = '%s';", convenio, medica);
          query = stm.executeQuery(sql);
          double total = 0.0;
          while(query.next()){
               total += query.getDouble(4);

          }
          return String.format("%.2f", total);
     }

     private static String getMedicamento(String convenio, String medica) throws SQLException {
          sql = String.format("select * from exames" +
                  " where  procedimento = 'Medicamento' and convenio = '%s' and medica = '%s';", convenio, medica);
          query = stm.executeQuery(sql);
          double total = 0.0;
          while(query.next()){
               total += query.getDouble(4);

          }
          return String.format("%.2f", total);
     }

     private static String getMaterial(String convenio, String medica) throws SQLException {
          sql = String.format("select * from exames" +
                  " where  procedimento = 'Material' and convenio = '%s' and medica = '%s';", convenio, medica);
          query = stm.executeQuery(sql);
          double total = 0.0;
          while(query.next()){
               total += query.getDouble(4);

          }
          return String.format("%.2f", total);
     }

     private static String addQuant(String num1, String num2, String num3, String num4, String num5, String num6){
          int soma = 0;
          soma = Integer.parseInt(num1) + Integer.parseInt(num2) + Integer.parseInt(num3) + Integer.parseInt(num4) + Integer.parseInt(num5) + Integer.parseInt(num6);
          return String.format("%d", soma);
     }

     private static String addTotal(String num1, String num2, String num3, String num4, String num5, String num6){
          double soma = 0.0;
          String num1a = num1.replace(",",".");
          String num2a = num2.replace(",",".");
          String num3a = num3.replace(",",".");
          String num4a = num4.replace(",",".");
          String num5a = num5.replace(",",".");
          String num6a = num5.replace(",",".");
          soma = Double.parseDouble(num1a) + Double.parseDouble(num2a) + Double.parseDouble(num3a) + Double.parseDouble(num4a) + Double.parseDouble(num5a) + Double.parseDouble(num6a);
          return String.format("%.2f", soma);
     }

     private static String getErros(String convenio) throws SQLException {

          sql = String.format("select * from exames" +
                  " where convenio = '%s' and medica = 'error';", convenio);
          query = stm.executeQuery(sql);
          int quant = 0;
          while(query.next()){
               quant ++;
          }
          return String.format("%d", quant);
     }

     public static void writeResume(List<String> list1, List<String> list2, List<String> list3, List<String> list4, List<String> list5, List<String> list6 , List<String> list7) throws SQLException, IOException {



          String[] columns= {"Data Inicial", "Data Final", "Convênio", "Medica", "Quantidade de Exames", "Valor Arrecadado", "Valor Arrecadado em Filme", "Valor Arrecadado em Medicamentos", "Valor Arrecadado em Material"};
          String[] convenios = {"Afrafep", "Caixa", "Cassi", "Embrapa", "Unimed", "Unimed PF", "Geral"};
          String[] medicas = {"Gerusa", "Laurise", "Valéria", "Procedimentos"};
          int[] totais = {4,9,14,16};
          ArrayList<List<String>> listas = new ArrayList<>(Arrays.asList(list1, list2, list3, list4, list5, list6, list7));
          ArrayList<String> periodos = getMinsMax(convenios);

          Workbook wb3 = new XSSFWorkbook();

          Font headerFont = wb3.createFont();
          headerFont.setBold(true);
          headerFont.setFontHeightInPoints((short) 14);
          headerFont.setColor(IndexedColors.RED.getIndex());

          Font titleFont = wb3.createFont();
          titleFont.setBold(true);
          titleFont.setFontHeightInPoints((short) 16);

          CellStyle titleCellStyle = wb3.createCellStyle();
          titleCellStyle.setFont(titleFont);

          CellStyle headerCellStyle = wb3.createCellStyle();
          headerCellStyle.setFont(headerFont);

          sh = wb3.createSheet("BalanceUS - Resumo");

          Row titleRow = sh.createRow(0);
          Row headerRow = sh.createRow(1);

          cell = titleRow.createCell(0);
          cell.setCellValue("Resumo dos Dados");
          cell.setCellStyle(titleCellStyle);

          for (int i = 0; i < columns.length; i++) {
               cell = headerRow.createCell(i);
               cell.setCellValue(columns[i]);
               cell.setCellStyle(headerCellStyle);
          }

          int numRow = 2;
          int index;
          int numPeriodo = 0;


          //Afrafep - Embrapa
          for(int j = 0; j <  4; j++){
               index = 0;
               for (int k = 0; k < 4; k++){
                    Row medRow = sh.createRow(numRow++);
                    medRow.createCell(0).setCellValue(periodos.get(numPeriodo));
                    medRow.createCell(1).setCellValue(periodos.get(numPeriodo+1));
                    medRow.createCell(2).setCellValue(convenios[j]);
                    medRow.createCell(3).setCellValue(medicas[k]);
                    medRow.createCell(4).setCellValue(listas.get(j).get(index));
                    index++;
                    medRow.createCell(5).setCellValue(listas.get(j).get(index));
                    index++;
               }
               numPeriodo += 2;
          }

          //Unimed

          index = 0;

          for (int l = 0; l < 4; l++){
               Row unimedRow = sh.createRow(numRow++);
               unimedRow.createCell(0).setCellValue(periodos.get(numPeriodo));
               unimedRow.createCell(1).setCellValue(periodos.get(numPeriodo+1));
               unimedRow.createCell(2).setCellValue(convenios[4]);
               unimedRow.createCell(3).setCellValue(medicas[l]);
               unimedRow.createCell(4).setCellValue(listas.get(4).get(index));
               index++;
               unimedRow.createCell(5).setCellValue(listas.get(4).get(totais[l]));
               if(l == 3){
                    break;
               }
               unimedRow.createCell(6).setCellValue(listas.get(4).get(index));
               index++;
               unimedRow.createCell(7).setCellValue(listas.get(4).get(index));
               index++;
               unimedRow.createCell(8).setCellValue(listas.get(4).get(index));
               index++;
               index++;
          }

          //Unimed PF

          index = 0;

          for (int l = 0; l < 4; l++){
               Row unimedPFRow = sh.createRow(numRow++);
               unimedPFRow.createCell(0).setCellValue(periodos.get(numPeriodo));
               unimedPFRow.createCell(1).setCellValue(periodos.get(numPeriodo+1));
               unimedPFRow.createCell(2).setCellValue(convenios[5]);
               unimedPFRow.createCell(3).setCellValue(medicas[l]);
               unimedPFRow.createCell(4).setCellValue(listas.get(5).get(index));
               index++;
               unimedPFRow.createCell(5).setCellValue(listas.get(5).get(totais[l]));
               if(l == 3){
                    break;
               }
               unimedPFRow.createCell(6).setCellValue(listas.get(5).get(index));
               index++;
               unimedPFRow.createCell(7).setCellValue(listas.get(5).get(index));
               index++;
               unimedPFRow.createCell(8).setCellValue(listas.get(5).get(index));
               index++;
               index++;
          }

          //Geral
          index = 0;

          for (int m = 0; m < 4; m++){
               Row geralRow = sh.createRow(numRow++);
               geralRow.createCell(0).setCellValue("");
               geralRow.createCell(1).setCellValue("");
               geralRow.createCell(2).setCellValue(convenios[6]);
               geralRow.createCell(3).setCellValue(medicas[m]);
               geralRow.createCell(4).setCellValue(listas.get(6).get(index));
               index++;
               geralRow.createCell(5).setCellValue(listas.get(6).get(index));
               index++;
          }

          for (int k = 0; k < columns.length; k++) {
               sh.autoSizeColumn(k);
          }

          fos = new FileOutputStream("BalanceUs - Resumo.xlsx");
          wb3.write(fos);
          fos.close();
     }

     private static ArrayList<String> getMinsMax(String[] convenios) throws SQLException {
          ArrayList<String> periodos =  new ArrayList<>();
          for(int i = 0; i < 5; i++){
               sql = String.format("select MIN(data) from exames where convenio = '%s'", convenios[i]);
               query = stm.executeQuery(sql);
               query.next();
               String minData = query.getString(1);

               sql = String.format("select MAX(data) from exames where convenio = '%s'", convenios[i]);
               query = stm.executeQuery(sql);
               query.next();
               String maxData = query.getString(1);

               periodos.add(minData);
               periodos.add(maxData);
          }
          return periodos;

     }

     public static void writeExcelOK() throws SQLException, IOException {
          String[] columns = {"Data", "Nome do Paciente", "Procedimento", "Valor", "Convênio", "Médica"};
          String[] convenios = {"Afrafep", "Caixa", "Cassi", "Embrapa", "Unimed", "Unimed PF"};
          Workbook wb2 = new XSSFWorkbook();

          //Fonts & Styles
          Font headerFont = wb2.createFont();
          headerFont.setBold(true);
          headerFont.setFontHeightInPoints((short) 14);
          headerFont.setColor(IndexedColors.RED.getIndex());

          Font titleFont = wb2.createFont();
          titleFont.setBold(true);
          titleFont.setFontHeightInPoints((short) 16);

          CellStyle titleCellStyle = wb2.createCellStyle();
          titleCellStyle.setFont(titleFont);

          CellStyle headerCellStyle = wb2.createCellStyle();
          headerCellStyle.setFont(headerFont);

          for (String convenio : convenios) {
               sh = wb2.createSheet(String.format("BalanceUS - OK - %s", convenio));

               Row titleRow = sh.createRow(0);
               Row headerRow = sh.createRow(1);

               cell = titleRow.createCell(0);
               cell.setCellValue(convenio);
               cell.setCellStyle(titleCellStyle);

               for (int i = 0; i < columns.length; i++) {
                    cell = headerRow.createCell(i);
                    cell.setCellValue(columns[i]);
                    cell.setCellStyle(headerCellStyle);
               }

               int RowNum = 2;

               sql = String.format("select * from exames" +
                       " where convenio = '%s' and" +
                       " medica != 'error';", convenio);
               query = stm.executeQuery(sql);
               while (query.next()) {
                    row = sh.createRow(RowNum++);
                    row.createCell(0).setCellValue(query.getString(1));
                    row.createCell(1).setCellValue(query.getString(2));
                    row.createCell(2).setCellValue(query.getString(3));
                    row.createCell(3).setCellValue(query.getDouble(4));
                    row.createCell((4)).setCellValue(query.getString(5));
                    row.createCell((5)).setCellValue(query.getString(6));
               }

               for (int k = 0; k < columns.length; k++) {
                    sh.autoSizeColumn(k);
               }
          }

          fos = new FileOutputStream("BalanceUs - OK.xlsx");
          wb2.write(fos);
          fos.close();

     }

     public static void writeExcelErros() throws SQLException, IOException {
          String[] columns = {"Data", "Nome do Paciente", "Procedimento", "Valor", "Convênio", "Médica", "Erro"};
          String[] convenios = {"Afrafep", "Caixa", "Cassi", "Embrapa", "Unimed", "Unimed PF"};
          Workbook wb2 = new XSSFWorkbook();

          //Fonts & Styles
          Font headerFont = wb2.createFont();
          headerFont.setBold(true);
          headerFont.setFontHeightInPoints((short) 14);
          headerFont.setColor(IndexedColors.RED.getIndex());

          Font titleFont = wb2.createFont();
          titleFont.setBold(true);
          titleFont.setFontHeightInPoints((short) 16);

          CellStyle titleCellStyle = wb2.createCellStyle();
          titleCellStyle.setFont(titleFont);

          CellStyle headerCellStyle = wb2.createCellStyle();
          headerCellStyle.setFont(headerFont);

          for (String convenio : convenios) {
               sh = wb2.createSheet(String.format("BalanceUS - Erros - %s", convenio));

               Row titleRow = sh.createRow(0);
               Row headerRow = sh.createRow(1);

               cell = titleRow.createCell(0);
               cell.setCellValue(convenio);
               cell.setCellStyle(titleCellStyle);

               for (int i = 0; i < columns.length; i++) {
                    cell = headerRow.createCell(i);
                    cell.setCellValue(columns[i]);
                    cell.setCellStyle(headerCellStyle);
               }

               int RowNum = 2;

               sql = String.format("select * from exames" +
                       " where convenio = '%s' and" +
                       " medica = 'error';", convenio);
               query = stm.executeQuery(sql);
               while (query.next()) {
                    row = sh.createRow(RowNum++);
                    row.createCell(0).setCellValue(query.getString(1));
                    row.createCell(1).setCellValue(query.getString(2));
                    row.createCell(2).setCellValue(query.getString(3));
                    row.createCell(3).setCellValue(query.getDouble(4));
                    row.createCell((4)).setCellValue(query.getString(5));
               }

               for (int k = 0; k < columns.length; k++) {
                    sh.autoSizeColumn(k);
               }
          }

          fos = new FileOutputStream("BalanceUs - Erros.xlsx");
          wb2.write(fos);
          fos.close();

     }



     private static String normalizeString(String str){
          String lower = str.toLowerCase();
          String normal = Normalizer.normalize(lower, Normalizer.Form.NFD);
          Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
          return pattern.matcher(normal).replaceAll("");
     }

     private static boolean checkIfRowIsEmpty(Row row) {
          if (row == null) {
               return true;
          }
          if (row.getLastCellNum() <= 0) {
               return true;
          }
          for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
               Cell cell = row.getCell(cellNum);
               if (cell != null && cell.getCellTypeEnum() != CellType.BLANK) {
                    return false;
               }
          }
          return true;
     }

     private static void corrigirErros() throws SQLException {
          corrigindoErrosAbreviacao("Unimed");
          corrigindoErrosAbreviacao("Unimed PF");
          corrigindoErrosPreposicao("Unimed");
     }

     private static void corrigindoErrosAbreviacao(String convenio) throws SQLException {

          sql = String.format("select * from exames" +
                  " where convenio = '%s' and medica = 'error';", convenio);
          query = stm.executeQuery(sql);
          ResultSet query1;
          Statement stm1 = con.createStatement();
          String errado = "";
          int id = 0;
          String sqlG = "";
          String sqlL = "";
          String sqlV = "";
          String sqlP = "";
          while(query.next()){
               String[] palavras = {};
               errado = query.getString(2);
               id = query.getInt(7);
               palavras =  errado.split(" ");
               sqlG = "select * from MedGerusa where";
               sqlL = "select * from MedLaurise where";
               sqlV = "select * from MedValeria where";
               sqlP = "select * from procedimentos where";
               medica = "";
               int cont = 1;
               for (String palavra : palavras){
                    if(cont < palavras.length){
                         sqlG += " nome LIKE '%" + palavra + "%' AND";
                         sqlL += " nome LIKE '%" + palavra + "%' AND";
                         sqlV += " nome LIKE '%" + palavra + "%' AND";
                         sqlP += " nome LIKE '%" + palavra + "%' AND";
                    }else{
                         sqlG += " nome LIKE '%" + palavra + "%';";
                         sqlL += " nome LIKE '%" + palavra + "%';";
                         sqlV += " nome LIKE '%" + palavra + "%';";
                         sqlP += " nome LIKE '%" + palavra + "%';";
                    }
                    cont++;

               }
               query1 = stm1.executeQuery(sqlG);
               while(query1.next()){
                    medica = "Gerusa";
               }
               query1 = stm1.executeQuery(sqlL);
               while(query1.next()){
                    medica = "Laurise";
               }
               query1 = stm1.executeQuery(sqlV);
               while(query1.next()){
                    medica = "Valéria";
               }
               query1 = stm1.executeQuery(sqlP);
               while(query1.next()){
                    medica = "Procedimentos";
               }

               switch (medica) {
                    case "Gerusa", "Laurise", "Valéria", "Procedimentos" -> {
                         sql = String.format("update exames " +
                                 "set medica = '%s' " +
                                 "where id = %d;", medica, id);
                         stm1.executeUpdate(sql);
                    }
               }

          }

     }

     private static void corrigindoErrosPreposicao(String convenio) throws SQLException {

          sql = String.format("select * from exames" +
                  " where convenio = '%s' and medica = 'error';", convenio);
          query = stm.executeQuery(sql);
          ResultSet query1;
          Statement stm1 = con.createStatement();
          String errado = "";
          int id = 0;
          String sqlG = "";
          String sqlL = "";
          String sqlV = "";
          String sqlP = "";
          while(query.next()) {
               String[] palavras = {};
               ArrayList<String> semD = new ArrayList<>();
               errado = query.getString(2);
               id = query.getInt(7);
               palavras = errado.split(" ");
               for (int i = 0; i < palavras.length; i++) {
                    if (palavras[i].equals("de")) {
                         continue;
//                         System.out.println("\nde " + palavras[i]);
                    } else if (palavras[i].equals("da")) {
                         continue;
//                         System.out.println("\nda" + palavras[i]);
                    } else if (palavras[i].equals("do")) {
                         continue;
//                         System.out.println("\ndo" + palavras[i]);
                    } else if (palavras[i].equals("dos")) {
                         continue;
//                         System.out.println("\ndos" + palavras[i]);
                    } else if (palavras[i].equals("das")) {
                         continue;
//                         System.out.println("\ndas" + palavras[i]);
                    } else {
                         semD.add(palavras[i]);
                    }
               }

               System.out.println("Essas é o nome sem Preposição:");
               for (String s : semD) {
                    System.out.println(s);
               }
               System.out.println("\n\n");

               sqlG = "select * from MedGerusa where";
               sqlL = "select * from MedLaurise where";
               sqlV = "select * from MedValeria where";
               sqlP = "select * from procedimentos where";
               medica = "";
               int cont = 1;
               for (String palavra : semD){
                    if(cont < semD.size()){
                         sqlG += " nome LIKE '%" + palavra + "%' AND";
                         sqlL += " nome LIKE '%" + palavra + "%' AND";
                         sqlV += " nome LIKE '%" + palavra + "%' AND";
                         sqlP += " nome LIKE '%" + palavra + "%' AND";
                    }else{
                         sqlG += " nome LIKE '%" + palavra + "%';";
                         sqlL += " nome LIKE '%" + palavra + "%';";
                         sqlV += " nome LIKE '%" + palavra + "%';";
                         sqlP += " nome LIKE '%" + palavra + "%';";
                    }
                    cont++;

               }
               query1 = stm1.executeQuery(sqlG);
               while(query1.next()){
                    medica = "Gerusa";
               }
               query1 = stm1.executeQuery(sqlL);
               while(query1.next()){
                    medica = "Laurise";
               }
               query1 = stm1.executeQuery(sqlV);
               while(query1.next()){
                    medica = "Valéria";
               }
               query1 = stm1.executeQuery(sqlP);
               while(query1.next()){
                    medica = "Procedimentos";
               }

               switch (medica) {
                    case "Gerusa", "Laurise", "Valéria", "Procedimentos" -> {
                         sql = String.format("update exames " +
                                 "set medica = '%s' " +
                                 "where id = %d;", medica, id);
                         stm1.executeUpdate(sql);
                    }
               }

          }

     }

     private static void searchExame(String nome) throws SQLException {
          sql = String.format("select * from Exames" +
                  " where nome LIKE '%s';", nome);
          query = stm.executeQuery(sql);
          System.out.println("\n"+ "Exames do Paciente " + nome);
          while(query.next()){
               System.out.println(query.getDate(1) + "\t\t" + query.getString(2) + "\t\t" + query.getString(3) + "\t\t" + query.getDouble(4) + "\t\t" + query.getString(5)+ "\t\t" + query.getString(6));
          }
     }



}


