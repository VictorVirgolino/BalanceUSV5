package ultrasomma;


import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.TextField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import java.io.*;

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

     //Ultrasomma
     private static List<String> listUltrasomma;
     private static String quantG6;
     private static String totalG6;
     private static String quantL6;
     private static String totalL6;
     private static String quantV6;
     private static String totalV6;
     private static String quantP6;
     private static String totalP6;
     private static String erros6;


     public static void lerExcels(ObservableList<String> arquivos) throws SQLException, ClassNotFoundException, IOException {

          for (String arquivo : arquivos) {
               if (arquivo.contains("valeria")) {
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
               } else if(arquivo.contains("procedimentos")){
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

          printTabela("MedValeria");
          printTabela("MedGerusa");
          printTabela("MedLaurise");
          printTabela("Procedimentos");


          sql =   "CREATE TABLE Exames ( " +
                  "   data DATE , " +
                  "   nome VARCHAR(255) , " +
                  "   procedimento VARCHAR(255) , " +
                  "   valor DOUBLE , " +
                  "   convenio VARCHAR(255) ," +
                  "   medica VARCHAR(255) " +
                  ");";

          stm.executeUpdate(sql);
          System.out.println("\nTabela Exames Criada!");

          //Alimentando Tabelas de Exames
          if(afrafep != null ){
               System.out.println("Fazendo afrafep");
               tabelaExames(afrafep, "Afrafep");
               System.out.println("Terminou afrafep");
               printTabela("Exames");
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
               System.out.println("Fazendo Unimed");
               tabelaExames(unimed, "Unimed");
               System.out.println("Terminou Unimed");
          }



//
//          Print das Tabelas
          printTabela("MedValeria");
          printTabela("MedGerusa");
          printTabela("MedLaurise");
          printExames("Exames");

//          writeResume();
          writeExcelErros();

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
                    System.out.println("Linha em branco");
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
                         data = cell.getDateCellValue();
                         if(data == null){
                              Etrue = true;
                         }else{
                              strData = format.format(data);
                         }
                    } else if (cont == 1) {
                        nome = cell.getStringCellValue();
                        System.out.println(nome);
                        if(nome == null){
                              Etrue = true;
                         }else {
                             nome = normalizeString(cell.getStringCellValue());
                        }
                    } else if (cont == 2) {
                         procedimento = cell.getStringCellValue();
                         System.out.println(procedimento);
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
                       "'%s'" +
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
                    System.out.println("Linha em branco");
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

                    if(cont == 0){
                         data = cell.getDateCellValue();
                         if(data == null){
                              isError = "true";
                         }else{
                              strData = format.format(data);
//                              System.out.println(strData);
                         }
                    }else if(cont == 1){
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
//
//          //Ultrasomma
          quantG6 = addQuant(quantG1, quantG2, quantG3, quantG4, quantG5);
          totalG6 = addTotal(totalG1,totalG2,totalG3,totalG4,totalG5);
          quantL6 = addQuant(quantL1, quantL2, quantL3, quantL4, quantL5);
          totalL6 = addTotal(totalL1,totalL2,totalL3,totalL4,totalL5);
          quantV6 = addQuant(quantV1, quantV2, quantV3, quantV4, quantV5);
          totalV6 = addTotal(totalV1,totalV2,totalV3,totalV4,totalV5);
          quantP6 = addQuant(quantP1, quantP2, quantP3, quantP4, quantP5);
          totalP6 = addTotal(totalP1,totalP2,totalP3,totalP4,totalP5);
          erros6 = addQuant(erros1, erros2, erros3,erros4,erros5);

          listUltrasomma = new ArrayList<>();
          listUltrasomma.addAll(Arrays.asList(quantG6,totalG6,quantL6,totalL6,quantV6,totalV6,quantP6,totalP6,erros6));

          writeResume(listAfrafep, listCaixa, listCassi, listEmbrapa, listUnimed, listUltrasomma);


          return Arrays.asList( listAfrafep, listCaixa, listCassi, listEmbrapa, listUnimed, listUltrasomma);


     }

     private static String getQuant(String convenio, String medica) throws SQLException {
          sql = String.format("select * from exames" +
                  " where procedimento != 'Filme' and convenio = '%s' and medica = '%s';", convenio, medica);
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

     private static String addQuant(String num1, String num2, String num3, String num4, String num5){
          int soma = 0;
          soma = Integer.parseInt(num1) + Integer.parseInt(num2) + Integer.parseInt(num3) + Integer.parseInt(num4) + Integer.parseInt(num5);
          return String.format("%d", soma);
     }

     private static String addTotal(String num1, String num2, String num3, String num4, String num5){
          double soma = 0.0;
          String num1a = num1.replace(",",".");
          String num2a = num2.replace(",",".");
          String num3a = num3.replace(",",".");
          String num4a = num4.replace(",",".");
          String num5a = num5.replace(",",".");
          soma = Double.parseDouble(num1a) + Double.parseDouble(num2a) + Double.parseDouble(num3a) + Double.parseDouble(num4a) + Double.parseDouble(num5a);
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

     public static void writeResume(List<String> list1, List<String> list2, List<String> list3, List<String> list4, List<String> list5, List<String> list6) throws SQLException, IOException {



          String[] columns= {"Data Inicial", "Data Final", "Convênio", "Medica", "Quantidade de Exames", "Valor Arrecadado", "Valor Arrecadado em Filme", "Valor Arrecadado em Medicamentos", "Valor Arrecadado em Material"};
          String[] convenios = {"Afrafep", "Caixa", "Cassi", "Embrapa", "Unimed", "Geral"};
          String[] medicas = {"Gerusa", "Laurise", "Valéria", "Procedimentos"};
          int[] totais = {4,9,14,16};
          ArrayList<List<String>> listas = new ArrayList<>(Arrays.asList(list1, list2, list3, list4, list5, list6));
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

          //Geral
          index = 0;

          for (int m = 0; m < 4; m++){
               Row geralRow = sh.createRow(numRow++);
               geralRow.createCell(0).setCellValue("");
               geralRow.createCell(1).setCellValue("");
               geralRow.createCell(2).setCellValue(convenios[5]);
               geralRow.createCell(3).setCellValue(medicas[m]);
               geralRow.createCell(4).setCellValue(listas.get(5).get(index));
               index++;
               geralRow.createCell(5).setCellValue(listas.get(5).get(index));
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

     public static void writeExcelErros() throws SQLException, IOException {
          String[] columns = {"Data", "Nome do Paciente", "Procedimento", "Valor", "Convênio"};
          String[] convenios = {"Afrafep", "Caixa", "Cassi", "Embrapa", "Unimed"};
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



}


