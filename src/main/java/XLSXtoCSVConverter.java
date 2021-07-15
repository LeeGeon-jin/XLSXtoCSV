import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Scanner;

public class XLSXtoCSVConverter
{
    public static void main(String[] args) throws IOException {
        System.out.println("Please input the file path to convert: ");
        Scanner in=new Scanner(System.in);
        String filePath=in.nextLine();
        in.close();

        File file=new File(filePath);
        FileInputStream fileInputStream=new FileInputStream(file);
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        fileInputStream.close();

        FileOutputStream outputStream=null;

        
        for(int i=0;i<workbook.getNumberOfSheets();i++)
        {
            String currentSheetName=workbook.getSheetName(i);
            String outputFileName=currentSheetName+".csv";
            outputStream=new FileOutputStream(outputFileName);

            
            XSSFWorkbook tmpWorkbook= (XSSFWorkbook) WorkbookFactory.create(new File(filePath));
            for(int j=workbook.getNumberOfSheets()-1;j>=0;j--)
            {
                XSSFSheet tmpSheet=tmpWorkbook.getSheetAt(j);
                if(!tmpSheet.getSheetName().equals(currentSheetName))
                {
                    tmpWorkbook.removeSheetAt(j);
                }
            }
            tmpWorkbook.write(outputStream);

        }
        workbook.close();
        outputStream.close();
        System.out.println("Convertion successful!");
    }


}
