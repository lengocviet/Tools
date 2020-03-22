package Excel;

import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

public class ConvertExcelToJson {

    public static void CreateJsonFilesFromExcel(String excelFile) throws IOException {
        XSSFWorkbook excelWorkBook = openExcelFile(excelFile);
        int rowIndexOfPage = 0;
        for (int i = 0; i < excelWorkBook.getNumberOfSheets();i++) {
            LinkedHashMap<String, int[]> pagesStructure = getPagesStructure(excelWorkBook.getSheetAt(i).getRow(rowIndexOfPage));
            createJsonDataAsObject(excelWorkBook.getSheetAt(i), pagesStructure);
        }
    }

    public static XSSFWorkbook openExcelFile(String file) {
        try {
            FileInputStream fileInputStream = new FileInputStream(file.trim());
            XSSFWorkbook excelWorkBook = new XSSFWorkbook(file);
            return excelWorkBook;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return new XSSFWorkbook();
        } catch (IOException e) {
            e.printStackTrace();
            return new XSSFWorkbook();
        }
    }

    private static void writeJsonToFile(String data, String file) throws IOException {

        FileWriter jsonFileWriter = new FileWriter(file);
        jsonFileWriter.write(data);
        jsonFileWriter.flush();
    }

    private static LinkedHashMap<String, int[]> getPagesStructure(XSSFRow page) {
        LinkedHashMap<String, int[]> pagesStructure = new LinkedHashMap<>();
        int start = 0;
        int end = 0;
        int i = 0;
        int maxColumn = page.getLastCellNum();
        try {
            while (i < maxColumn) {
                XSSFCell cell = page.getCell(i);
                start = i;
                end = i;
                boolean increase = true;
                if (!("".equals(cell.getRawValue()))) {
                    XSSFCell nextCell = page.getCell(i + 1);
                    while ("".equals(nextCell.getStringCellValue()) && i < maxColumn) {
                        i++;
                        if (i < maxColumn) {
                            nextCell = page.getCell(i);
                        }
                        end = i - 1;
                        increase = false;
                    }
                    int[] startEnd = new int[]{start, end};
                    pagesStructure.put(page.getCell(start).getStringCellValue(), startEnd);
                    i = increase ? i + 1 : i;
                }
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());

        }

        return pagesStructure;
    }

    private static void createJsonDataAsObject(XSSFSheet sheet, LinkedHashMap<String, int[]> jsonScheme) throws IOException {

        List<List<String>> ret = new ArrayList<List<String>>();
        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        int firstDataRow = 3;
        int lastDataRow = sheet.getLastRowNum();
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>>> featureObject = new LinkedHashMap<>();
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>> testCaseObject = new LinkedHashMap<>();
        /*First row is page -> the second row is header of column. Data is from the third row */
        if (lastDataRow <= 3) return;
        XSSFRow page = sheet.getRow(0);
        XSSFRow type = sheet.getRow(1);
        XSSFRow header = sheet.getRow(2);
        for (int i = firstDataRow; i <= 6; i++) {
            LinkedHashMap<String, LinkedHashMap<String, String>> pageObject = new LinkedHashMap<>();
            String testCaseID = "";
            List<String> keys = new ArrayList<String>(jsonScheme.keySet());
            for (int index = 0; index < keys.size(); index++) {
                String key = keys.get(index);
                if ("TC_ID".equalsIgnoreCase(key)) {
                    testCaseID = sheet.getRow(i).getCell(0).toString().trim();
                } else {
                    LinkedHashMap<String, String> fieldObject = new LinkedHashMap<>();
                    int firstColumn = jsonScheme.get(key)[0];
                    int lastColumn = jsonScheme.get(key)[1];
                    XSSFRow currentRow = sheet.getRow(i);
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        XSSFCell cell = currentRow.getCell(j);
                        String cellType = type.getCell(j).toString();
                        if ((cell.getRawValue() == null) || ("".equals(cell.getRawValue()))) {
                            continue;
                        }
                        try {
                            switch (cellType) {
                                case "DATE":
                                    Date dateValue = cell.getDateCellValue();
                                    SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                    fieldObject.put(header.getCell(j).toString().trim(), sdf.format(dateValue));
                                    break;
                                case "INT":

                                    fieldObject.put(header.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
                                    break;
                                case "BOOLEAN":

                                    fieldObject.put(header.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));

                                case "STRING":

                                    fieldObject.put(header.getCell(j).toString().trim(), cell.getStringCellValue().trim());
                                    break;
                                default:
                                    fieldObject.put(header.getCell(j).toString().trim(), cell.getRawValue());
                            }
                        } catch (IllegalStateException e) {
                            System.out.println("Error when read data at Cell:" + header.getCell(j).toString() + " Cell type:" + cellType);
                            System.out.println("Error type:" + e.getMessage());
                        }

                    }
                    pageObject.put(key, fieldObject);

                }
                testCaseObject.put(testCaseID, pageObject);
            }

        }
        featureObject.put(sheet.getSheetName().toString(), testCaseObject);
        String jsonString = new Gson().toJson(featureObject, LinkedHashMap.class);
        String path = "." + "\\data\\" + sheet.getSheetName().toString() + ".json";
        writeJsonToFile(jsonString, path);
    }

}




