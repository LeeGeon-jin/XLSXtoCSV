
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Scanner;

public class XLSXtoCSVTest3 {
    public static void main(String[] args) throws IOException {
        System.out.println("Input file path (If more than one file, separate by SPACE) : ");
        Scanner in=new Scanner(System.in);
        String filePath=in.nextLine();
        in.close();

        if(filePath.contains(" "))
        {
            String[] filePaths=filePath.split(" ");
            for(int i=0;i<filePaths.length;i++)
            {
                convert(filePaths[i],String.valueOf(i));
            }
        }
        else
        {
            convert(filePath,"");
        }





    }

    private static void convert(String filePath, String outputDirectoryIndex) throws IOException {
        FileInputStream fileInputStream=new FileInputStream(filePath);
        Workbook workbook= WorkbookFactory.create(fileInputStream);
        fileInputStream.close();

        String outputFolderName="File"+outputDirectoryIndex;
        File outputFolder=new File(outputFolderName);
        outputFolder.mkdirs();

        System.out.println(filePath);
        for(int i=0;i<workbook.getNumberOfSheets();i++)
        {

            Sheet sheet=workbook.getSheetAt(i);
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



            String outputFileName=sheet.getSheetName();
            String outputFilePath=outputFolderName+"\\"+outputFileName+".csv";
            FileWriter fileWriter=new FileWriter(outputFilePath);
            fileWriter.write(str);
            fileWriter.close();
            System.out.println(outputFileName+":");
            System.out.println(str);
        }

        System.out.println("Conversion successful!");
        workbook.close();
    }
}
