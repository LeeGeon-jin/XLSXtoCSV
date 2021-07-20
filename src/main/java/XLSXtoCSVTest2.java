import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;


public class XLSXtoCSVTest2 {

    // testing the application

    public static void main(String[] args) {
        String inputFilePath="../tmpsave/t1.xlsx";
        String outputFilePath="test_output.csv";
        // reading file from desktop
        File inputFile = new File(inputFilePath);
        // writing excel data to csv
        File outputFile = new File(outputFilePath);
        convert(inputFile, outputFile);
    }

    static void convert(File inputFile, File outputFile) {
        try {
            StringBuilder data=new StringBuilder();
            FileOutputStream fos = new FileOutputStream(outputFile);
            // Get the workbook object for XLSX file
            XSSFWorkbook wBook = new XSSFWorkbook(
                    new FileInputStream(inputFile));
            // Get first sheet from the workbook
            XSSFSheet sheet = wBook.getSheetAt(0);
            Row row;
            Cell cell;
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case BOOLEAN:
                            data.append(cell.getBooleanCellValue() + ",");

                            break;
                        case NUMERIC:
                            data.append(cell.getNumericCellValue() + ",");

                            break;
                        case STRING:
                            data.append(cell.getStringCellValue() + ",");
                            break;

                        case BLANK:
                            data.append("" + ",");
                            break;
                        default:
                            data.append(cell + ",");

                    }
                }
                if(!data.toString().isEmpty())
                {
                    data.setLength(data.length()-1);
                    data.append("\r\n");
                }

            }

            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }


}
