package kr.excel.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelExample {
    public static void main(String[] args) {
        //메모리에 가상의 워크북(시트를 표현)을 만들어 작업
        try{
            FileInputStream file = new FileInputStream(new File("apache.xlsx"));
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            for(Row row : sheet){
                for(Cell cell : row){
                    System.out.print(cell.toString()+"\t");
                }
                System.out.println();
            }
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}
