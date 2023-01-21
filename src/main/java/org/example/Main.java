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
            writeData("imena.xlsx");
            System.out.println("\n \n \n---------Novo Citanje-----------\n \n \n");
            readData("imena.xlsx");
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


        for (int i = 0; i < names.getLastRowNum()+1; i++) {
            XSSFRow row = names.getRow(i);
            System.out.println(row.getCell(0) + " " + row.getCell(1));
        }


    }

    public static void writeData(String relativnaPutanjaNaSrpskomDoFajlaISadZNamStaJe) throws IOException {
        FileInputStream inputStream = new FileInputStream(relativnaPutanjaNaSrpskomDoFajlaISadZNamStaJe);
        XSSFWorkbook ObjekatExcelPrikaza = new XSSFWorkbook(inputStream);
        XSSFSheet sheet1 = ObjekatExcelPrikaza.getSheet("Names");
        Faker laznjak = new Faker();
//        XSSFSheet sheet1 = workbook.createSheet("IT Bootcamp Testovi");
//        XSSFSheet sheet2 = workbook.createSheet("IT Bootcamp Domaci");
        int cilj = sheet1.getLastRowNum()+6;
        System.out.println(cilj);
        for (int i = sheet1.getLastRowNum()+1; i < cilj; i++) {
            XSSFRow row = sheet1.createRow(i);
            row.createCell(0).setCellValue(laznjak.name().firstName());
            row.createCell(1).setCellValue(laznjak.name().lastName());

        }

        FileOutputStream adresaImeFajla = new FileOutputStream("imena.xlsx");
        ObjekatExcelPrikaza.write(adresaImeFajla);
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