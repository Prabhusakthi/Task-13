package task13;


	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.IOException;

	public class ExcelPOIExample {
	    
	    public static void main(String[] args) {
	        String[] columns = {"Name", "Age", "Email"};
	        String[][] data = {
	                {"John Doe", "30", "john@test.com"},
	                {"Jane Doe", "28", "jane@test.com"},
	                {"Bob Smith","35", "jacky@example.com"},
	                {"Swapnil",  "37" ,"swapnil@example.com"}
	        
	        };

	        // Create a new workbook and sheet
	        XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet = workbook.createSheet("sheet1");

	        // Create a header row
	        Row headerRow = sheet.createRow(0);
	        for (int i = 0; i < columns.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(columns[i]);
	        }

	        // Fill data rows
	        int rowNum = 1;
	        for (String[] rowData : data) {
	            Row row = sheet.createRow(rowNum++);
	            for (int i = 0; i < rowData.length; i++) {
	                row.createCell(i).setCellValue(rowData[i]);
	            }
	        }

	        // Write the output to a file
	        try (FileOutputStream fileOut = new FileOutputStream("example.xlsx")) {
	            workbook.write(fileOut);
	        } catch (IOException e) {
	            e.printStackTrace();
	        }

	        // Closing the workbook
	        try {
	            workbook.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }

	        // Read the data back from the file and print to console
	        try (FileInputStream fileIn = new FileInputStream("example.xlsx")) {
	            Workbook workbookRead = WorkbookFactory.create(fileIn);
	            Sheet sheetRead = workbookRead.getSheetAt(0);

	            for (Row row : sheetRead) {
	                for (Cell cell : row) {
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            System.out.print(cell.getStringCellValue() + "\t");
	                            break;
	                        case NUMERIC:
	                            System.out.print((int) cell.getNumericCellValue() + "\t");
	                            break;
	                        default:
	                            break;
	                    }
	                }
	                System.out.println();
	            }

	            workbookRead.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	}



