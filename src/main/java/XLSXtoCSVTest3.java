
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;

public class XLSXtoCSVTest3 {
    public static void main(String[] args) throws IOException {
        String filePath="../tmpsave/t1.xlsx";
        String outputPath="test_out.csv";
        int currentSheetIndex=0;
        
        //File inputFile=new File(filePath);
        FileInputStream fileInputStream=new FileInputStream(filePath);
        //FileOutputStream fileOutputStream=new FileOutputStream(outputPath);
        FileWriter fileWriter=new FileWriter(outputPath);

        Workbook workbook= WorkbookFactory.create(fileInputStream);
        Sheet sheet=workbook.getSheetAt(currentSheetIndex);
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

        fileWriter.write(str);
        System.out.println(str);
        workbook.close();
        fileWriter.close();
    }
}
