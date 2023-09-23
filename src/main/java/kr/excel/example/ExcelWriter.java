package kr.excel.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class ExcelWriter {
    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);
        List<MemberVO> members = new ArrayList<>();
        //quit이라는 문자열을 입력받아야 종료하는 프로그램
        while(true) {
            //VO에 담을 값을 입력받고 VO에 담아두기
            System.out.println("이름을 입력하세요: ");
            String name = sc.nextLine();
            if(name.equals("quit")){
                break;
            }

            System.out.println("나이를 입력하세요: ");
            int age = sc.nextInt();
            sc.nextLine();//개행문자 제거

            System.out.println("생년월일을 입력하세요(yyyy-MM-dd): ");
            String birthdate = sc.nextLine();

            System.out.println("전화번호를 입력하세요: ");
            String phone = sc.nextLine();

            System.out.println("주소를 입력하세요: ");
            String address = sc.nextLine();

            System.out.println("결혼 여부를 입력하세요: ");
            boolean isMarried = sc.nextBoolean();
            sc.nextLine();//개행문자 제거

            MemberVO member = new MemberVO(name,age,birthdate,phone,address,isMarried);
            members.add(member);
        }
        sc.close();
        
        //엑셀 파일과 통신간 예외처리를 위해 try catch 사용
        try {
            //워크북 생성
            XSSFWorkbook workbook = new XSSFWorkbook();
            //시트 생성
            Sheet sheet = workbook.createSheet("회원 정보");
            
            //헤더 생성
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("이름");
            headerRow.createCell(1).setCellValue("나이");
            headerRow.createCell(2).setCellValue("생년월일");
            headerRow.createCell(3).setCellValue("전화번호");
            headerRow.createCell(4).setCellValue("주소");
            headerRow.createCell(5).setCellValue("결혼여부");

            //데이터 생성
            for (int i = 0; i < members.size();i++){
                MemberVO member = members.get(i);
                Row row = sheet.createRow(i+1);
                row.createCell(0).setCellValue(member.getName());
                row.createCell(1).setCellValue(member.getAge());
                row.createCell(2).setCellValue(member.getBirthdate());
                row.createCell(3).setCellValue(member.getPhone());
                row.createCell(4).setCellValue(member.getAddress());
                Cell marriedCell = row.createCell(5);
                marriedCell.setCellValue(member.isMarried());
            }
            
            //엑셀 파일 저장
            String filename = "members.xlsx";
            FileOutputStream outputStream = new FileOutputStream(new File(filename));
            workbook.write(outputStream);
            workbook.close();
            System.out.println("엑셀 파일 저장 완료: "+filename);
        }catch (IOException e){
            System.out.println("엑셀 파일 저장 중 오류 발생");
            e.printStackTrace();
        }
    }
}
