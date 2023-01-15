package org.example;

import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        try {
            readData("imena.xlsx");
            writeData();
        } catch (FileNotFoundException e) {
            System.out.println("Nevalidna putanja!");
        } catch (IOException e) {
            System.out.println("Nevalidan excel fajl!");
        }

    }

        public static void readData(String relativnaPutanjaNaSrpskomDoFajlaISadZNamStaJe) throws FileNotFoundException, IOException {
        FileInputStream inputStream = new FileInputStream(relativnaPutanjaNaSrpskomDoFajlaISadZNamStaJe);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet names = workbook.getSheet("Names");


            for (int i = 0; i < 5 ; i++) {
                XSSFRow row = names.getRow(i);
                System.out.println(row.getCell(0) + " "  + row.getCell(1));
            }
    }

    public static void writeData() throws IOException {
        XSSFWorkbook stranaZaUpis = new XSSFWorkbook();
        XSSFSheet sheet1 = stranaZaUpis.getSheet("Names");
        Faker laznjak = new Faker();
//        XSSFSheet sheet1 = workbook.createSheet("IT Bootcamp Testovi");
//        XSSFSheet sheet2 = workbook.createSheet("IT Bootcamp Domaci");
        for (int i = 5; i < 10 ; i++) {
             XSSFRow row = sheet1.createRow(i);
             row.getCell(0).setCellValue(laznjak.name().firstName());
             row.getCell(1).setCellValue(laznjak.name().lastName());

        }

        FileOutputStream adresaImeFajla = new FileOutputStream("imena.xlsx");
        stranaZaUpis.write(adresaImeFajla);
        adresaImeFajla.close();


/*        for(int i=0; i<10; i++) {
            XSSFRow row = sheet1.createRow(i);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue("Java je OOP programski jezik!");
        }

        FileOutputStream outputStream = new FileOutputStream("test.xlsx");
        workbook.write(outputStream);
        outputStream.close();*/
    }
}