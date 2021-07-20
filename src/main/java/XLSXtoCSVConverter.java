import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Scanner;

public class XLSXtoCSVConverter
{
    public static void main(String[] args) throws IOException {
        System.out.println("Please input the file path to convert: ");
        String filePath=getUserInput();

        XSSFWorkbook workbook=readFile(filePath);

        outputFile(workbook,filePath);
    }

    private static String getUserInput()
    {
        Scanner in=new Scanner(System.in);
        String inputText=in.nextLine();
        in.close();
        return inputText;
    }

    private static XSSFWorkbook readFile(String filePath) throws IOException {
        File file=new File(filePath);
        FileInputStream fileInputStream=new FileInputStream(file);
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        fileInputStream.close();
        return workbook;
    }

    private static void outputFile(XSSFWorkbook workbook,String inputFilePath) throws IOException {

        FileOutputStream outputStream=null;

        //use loop to output all sheets in the workbook
        for(int i=0;i<workbook.getNumberOfSheets();i++)
        {
            //set the output file name and output stream
            String currentSheetName=workbook.getSheetName(i);
            String outputFileName=currentSheetName+".csv";
            outputStream=new FileOutputStream(outputFileName);

            //create a temporary workbook and only keep the current sheet to output
            //XSSFWorkbook tmpWorkbook= (XSSFWorkbook) WorkbookFactory.create(new File(filePath));
            XSSFWorkbook tmpWorkbook=new XSSFWorkbook(inputFilePath);
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
        assert outputStream != null;
        outputStream.close();
        System.out.println("Convertion successful!");
    }
}
