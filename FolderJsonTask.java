package Task1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;

public class FolderJsonTask {

    // Method to create folders and JSON files
    private static void createFolderAndFile(String folderName, String fileName, ArrayList<String> data) throws IOException {
        File folder = new File(folderName);
        if (!folder.exists() && !folder.mkdirs()) {
            throw new IOException("Failed to create directory: " + folderName);
        }

        File file = new File(folder, fileName);
        try (FileWriter writer = new FileWriter(file)) {
            writer.write(data.toString());
        }
    }

    // Method to read data from Excel
    private static ArrayList<Map<String, String>> readDataFromExcel(String filePath) {
        ArrayList<Map<String, String>> entries = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath)) {
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0); // First sheet

            for (Row row : sheet) {
                // Skip the header row
                if (row.getRowNum() == 0) {
                    continue;
                }

                Map<String, String> entry = new HashMap<>();
                entry.put("s_no", getCellValue(row.getCell(0))); // First column
                entry.put("name", getCellValue(row.getCell(1))); // Second column
                entry.put("ph_no", getCellValue(row.getCell(2))); // Third column
                entries.add(entry);
            }
            workbook.close();
        } catch (Exception e) {
            System.out.println("Error reading Excel file: " + e.getMessage());
        }
        return entries;
    }

    // Helper method to get cell value as String
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    public static void main(String[] args) {
        // Path to the Excel file
        String excelFilePath = "src/main/resources/emp_table.xlsx";

        // Read data from Excel
        ArrayList<Map<String, String>> entries = readDataFromExcel(excelFilePath);

        if (entries.isEmpty()) {
            System.out.println("No data found in the Excel file.");
            return;
        }

        ArrayList<String> sNoData = new ArrayList<>();
        ArrayList<String> nameData = new ArrayList<>();
        ArrayList<String> phNoData = new ArrayList<>();

        for (Map<String, String> entry : entries) {
            sNoData.add(entry.get("s_no"));
            nameData.add(entry.get("name"));
            phNoData.add(entry.get("ph_no"));
        }

        try {
            createFolderAndFile("s_no", "s_no.json", sNoData);
            createFolderAndFile("name", "name.json", nameData);
            createFolderAndFile("ph_no", "ph_no.json", phNoData);
            System.out.println("Folders and files created successfully!");
        } catch (IOException e) {
            System.out.println("An error occurred: " + e.getMessage());
        }
    }
}
