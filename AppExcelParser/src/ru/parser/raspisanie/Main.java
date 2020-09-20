package ru.parser.raspisanie;

import java.io.IOException;

/**
 * Основной класс, запускает на считывание файла Excel данные
 *
 * @author Aleks
 */
public class Main {

    /**
     * основной метод для проверки корректности работы класса после включения в
     * веб приложение, закоментировать.
     *
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {

        ReadExcelData cod = new ReadExcelData();
        ReadExcelData division = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
//        ReadExcelData cod = new ReadExcelData();
        System.out.printf(cod.getData(0) + " : " + division.getData(1));  
        
//        ReadExcelData division = null;
//        division.getData(1);
//        
    }

}
