package kr.excel.example;

import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.property.UnitValue;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.*;

public class PDFFile {
    public static void main(String[] args) throws MalformedURLException, IOException {
        String dest = "book_table.pdf";
        new PDFFile().createPdf(dest);
    }

    public void createPdf(String dest) throws MalformedURLException, IOException{
        List<Map<String, String>> books = createDummyData();

        PdfWriter writer = new PdfWriter(dest);
        PdfDocument pdf = new PdfDocument(writer);
        Document document = new Document(pdf, PageSize.A4);
        
        //폰트 초기화
        PdfFont headerFont = null;
        PdfFont bodyFont = null;
        try {
            headerFont = PdfFontFactory.createFont("malgunsl.ttf","Identity-H",true);
            bodyFont = PdfFontFactory.createFont("malgunsl.ttf","Identity-H",true);
        }catch (IOException e){
            e.printStackTrace();
        }

        //테이블 초기화
        float[] columWidths = {1,2,2,2,2,2};
        Table table = new Table(UnitValue.createPointArray(columWidths));
        table.setWidth(UnitValue.createPercentValue(100));

        //헤더셀 초기화
        Cell headerCell1 = new Cell().add(new Paragraph("순번")).setFont(headerFont);
        Cell headerCell2 = new Cell().add(new Paragraph("제목")).setFont(headerFont);
        Cell headerCell3 = new Cell().add(new Paragraph("저자")).setFont(headerFont);
        Cell headerCell4 = new Cell().add(new Paragraph("출판사")).setFont(headerFont);
        Cell headerCell5 = new Cell().add(new Paragraph("출판일")).setFont(headerFont);
        Cell headerCell6 = new Cell().add(new Paragraph("이미지")).setFont(headerFont);
        table.addHeaderCell(headerCell1);
        table.addHeaderCell(headerCell2);
        table.addHeaderCell(headerCell3);
        table.addHeaderCell(headerCell4);
        table.addHeaderCell(headerCell5);
        table.addHeaderCell(headerCell6);
        
        //테이블에 바디셀 추가
        int rowNum = 1;
        for (Map<String,String> book : books){
            String title = book.get("title");
            String authors = book.get("authors");
            String publisher = book.get("publisher");
            String publishDate = book.get("publishDate");
            String thumbnail = book.get("thumbnail");
            Cell rowNumCell = new Cell().add(new Paragraph(String.valueOf(rowNum))).setFont(bodyFont);
            table.addCell(rowNumCell);
            Cell titleCell = new Cell().add(new Paragraph(title)).setFont(bodyFont);
            table.addCell(titleCell);
            Cell authorsCell = new Cell().add(new Paragraph(authors)).setFont(bodyFont);
            table.addCell(authorsCell);
            Cell publisherCell = new Cell().add(new Paragraph(publisher)).setFont(bodyFont);
            table.addCell(publisherCell);
            Cell publishDateCell = new Cell().add(new Paragraph(publishDate)).setFont(bodyFont);
            table.addCell(publishDateCell);
            ImageData imageData = ImageDataFactory.create(new File(thumbnail).toURI().toURL());
            Image img = new Image(imageData);
            Cell imageCell = new Cell().add(img.setAutoScale(true));
            table.addCell(imageCell);
        }
        document.add(table);
        document.close();
    }

    private static List<Map<String, String>> createDummyData(){
        List<Map<String,String>> books = new ArrayList<>();
        Scanner sc = new Scanner(System.in);
        System.out.println("책 개수를 입력하세요: ");
        int bookCount = sc.nextInt();
        sc.nextLine();//개행문자 제거

        for(int i = 1; i<=bookCount; i++){
            Map<String, String> book = new HashMap<>();

            System.out.printf("\n[ %d번째 책 정보 입력]\n",i);
            System.out.print("제목: ");
            book.put("title",sc.nextLine());

            System.out.println("저자: ");
            book.put("authors", sc.nextLine());

            System.out.println("출판사: ");
            book.put("publisher", sc.nextLine());

            System.out.println("출판일(YYYY-MM-DD: ");
            book.put("publishDate", sc.nextLine());

            System.out.println("썸네일 URL: ");
            book.put("thumbnail", sc.nextLine());

            books.add(book);
        }
        sc.close();

        return books;

    }
}
