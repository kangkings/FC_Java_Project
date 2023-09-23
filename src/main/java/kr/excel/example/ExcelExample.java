package kr.excel.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelExample {
    public static void main(String[] args) {
        //메모리에 가상의 워크북(시트를 표현)을 만들어 작업
        try{
            FileInputStream file = new FileInputStream(new File("apache.xlsx"));
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            //셀의 데이터 타입에 맞게 값 가져오기(switch문 이용)
            for(Row row : sheet){
                for(Cell cell : row){
                    switch(cell.getCellType()){
                        //숫자일경우(날짜,정수,실수)
                        case NUMERIC:
                            //날짜일경우
                            if(DateUtil.isCellDateFormatted(cell)){
                                Date dateValue = cell.getDateCellValue();
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                String formattedDate = dateFormat.format(dateValue);
                                System.out.print(formattedDate+"\t");
                            }else{
                                //실수로 받아온 cell값이 floor한 값과 같다? => 정수이다
                                double numericValue = cell.getNumericCellValue();
                                if (numericValue == Math.floor(numericValue)){
                                    //floor한 값과 같다면 1.0 2.0과 같기 떄문에
                                    //해당 실수를 캐스팅하여 정수로 출력한다
                                    int intValue = (int) numericValue;
                                    System.out.print(intValue+"\t");
                                }else {
                                    //아닌경우 실수이기 때문에 실수형태로 출력
                                    System.out.print(numericValue+"\t");
                                }
                            }
                            break;
                        case STRING:
                            String stringValue = cell.getStringCellValue();
                            System.out.print(stringValue+"\t");
                            break;
                        case BOOLEAN:
                            boolean boolValue = cell.getBooleanCellValue();
                            System.out.print(boolValue+"\t");
                            break;
                        case FORMULA:
                            String formulaValue = cell.getCellFormula();
                            System.out.print(formulaValue+"\t");
                            break;
                        case BLANK:
                            System.out.print("\t");
                            break;
                        default:
                            System.out.print("\t");
                            break;

                    }
                }
                System.out.println();
            }
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}
