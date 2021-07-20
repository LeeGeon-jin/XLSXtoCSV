import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Scanner;

public class XLSXtoCSVNew {
    public static void main(String[] args) throws IOException {
        Workbook workbook=inputFile();
        outputFile(workbook);
    }

    private static Workbook inputFile() throws IOException {
        System.out.println("Please input the file path to convert: ");
        String filePath=getUserInput();

        FileInputStream fileInputStream=new FileInputStream(filePath);
        Workbook workbook=WorkbookFactory.create(fileInputStream);
        fileInputStream.close();

        return workbook;
    }

    private static String getUserInput()
    {
        Scanner in=new Scanner(System.in);
        String inputText=in.nextLine();
        in.close();
        return inputText;
    }

    private static void outputFile(Workbook workbook) throws IOException {
        for(int i=0;i<workbook.getNumberOfSheets();i++)
        {
            String currentSheetName=workbook.getSheetName(i);
            String outputFileName=currentSheetName+".csv";
            writeFile(outputFileName,workbook,i);
        }

        workbook.close();

        System.out.println("Conversion successful!");
    }

    private static void writeFile(String outputFileName, Workbook workbook, int sheetIndex) throws IOException {
        String data=extractData(workbook,sheetIndex);

        FileWriter fileWriter=new FileWriter(outputFileName);
        fileWriter.write(data);
        fileWriter.close();

        System.out.println(outputFileName+" :");
        System.out.println(data);
        System.out.println("\n");
    }

    private static String extractData(Workbook inputBook, int index)
    {
        Sheet sheet=inputBook.getSheetAt(index);
        String str = "";

        for(Row row: sheet)
        {
            String rowString = "";
            for(Cell cell:row)
            {
                if(cell==null)
                {
                    rowString=rowString+" "+",";
                }
                else
                {
                    rowString=rowString+cell+",";
                }
            }
            if(!rowString.isEmpty())
                str=str+rowString.substring(0,rowString.length()-1)+"\r\n";
        }

        return str;
    }

}
