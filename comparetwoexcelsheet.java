import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.CellStyle;

import java.io.File;
import java.io.FileOutputStream;

public class CompareExcelSheets {
    public static void main(String[] args) throws Exception {
        // Open the first Excel file
        File file1 = new File("path_to_file1.xlsx");
        XSSFWorkbook workbook1 = new XSSFWorkbook(file1);

        // Open the second Excel file
        File file2 = new File("path_to_file2.xlsx");
        XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

        // Get the first sheet from the first workbook
        XSSFSheet sheet1 = workbook1.getSheetAt(0);

        // Get the first sheet from the second workbook
        XSSFSheet sheet2 = workbook2.getSheetAt(0);

        // Create a highlight cell style
        XSSFCellStyle highlightCellStyle = workbook2.createCellStyle();
        highlightCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        highlightCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

        // Iterate over the rows in the sheets
        for (int i = 0; i < sheet1.getLastRowNum(); i++) {
            XSSFRow row1 = sheet1.getRow(i);
            XSSFRow row2 = sheet2.getRow(i);

            // Iterate over the cells in the rows
            for (int j = 0; j < row1.getLastCellNum(); j++) {
                XSSFCell cell1 = row1.getCell(j);
                XSSFCell cell2 = row2.getCell(j);

                // Compare the cell values
                if (!cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    // Highlight the cell in the second sheet if the values are different
                    cell2.setCellStyle(highlightCellStyle);
                }
            }
        }

        // Save the modified workbook
        FileOutputStream fileOut = new FileOutputStream("path_to_modified_file.xlsx");
        workbook2.write(fileOut);
        fileOut.close();
    }
}
