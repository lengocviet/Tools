package Excel;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

public class ConvertExcelToJson {

    public static void CreateJsonFilesFromExcel(String excelFile) throws IOException {
        XSSFWorkbook excelWorkBook = openExcelFile(excelFile);
        int rowIndexOfPage = 0;
        for (int i = 0; i < excelWorkBook.getNumberOfSheets(); i++) {
            LinkedHashMap<String, int[]> pagesStructure = getPagesStructure(excelWorkBook.getSheetAt(i).getRow(rowIndexOfPage));
            createJsonDataAsObject(excelWorkBook.getSheetAt(i), pagesStructure);
        }
    }

    public static XSSFWorkbook openExcelFile(String file) throws IOException {
        try {
            XSSFWorkbook excelWorkBook = new XSSFWorkbook(file);
            return excelWorkBook;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw e;
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        }

    }

    private static void writeJsonToFile(String data, String file) throws IOException {

//        FileWriter jsonFileWriter = new FileWriter(file);
//        jsonFileWriter.write(data);
//        jsonFileWriter.flush();
        try {
            File fileDir = new File(file);

            Writer out = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream(fileDir), "UTF8"));

            out.write(data);

            out.flush();
            out.close();

        }
        catch (UnsupportedEncodingException e)
        {
            System.out.println(e.getMessage());
        }
        catch (IOException e)
        {
            System.out.println(e.getMessage());
        }
        catch (Exception e)
        {
            System.out.println(e.getMessage());
        }
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
            System.out.println("Error when create page structure at" + page.getSheet().getSheetName() + e.getMessage());
            throw e;

        }

        return pagesStructure;
    }

    private static int getLastRow(XSSFSheet sheet) {
        int i = 3;
        while (i <= sheet.getLastRowNum()) {
            if (sheet.getRow(i).getCell(0) == null) {
                break;
            }
            i++;
        }
        return i;
    }

    private static void createJsonDataAsObject(XSSFSheet sheet, LinkedHashMap<String, int[]> jsonScheme) throws IOException {

        List<List<String>> ret = new ArrayList<List<String>>();
        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        int firstDataRow = 2;
        int lastDataRow = getLastRow(sheet);
        if (lastDataRow <= 2) return;
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>>> featureObject = new LinkedHashMap<>();
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>> testCaseObject = new LinkedHashMap<>();
        /*First row is page -> the second row is header of column. Data is from the third row */
        //XSSFRow type = sheet.getRow(1);
        XSSFRow header = sheet.getRow(1);
        for (int i = firstDataRow; i < lastDataRow; i++) {
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
                        CellType cellType = cell.getCellType();
                        if ((cell.getRawValue() == null) || ("".equals(cell.getRawValue()))) {
                            continue;
                        }
                        try {
                            switch (cellType) {
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        Date dateValue = cell.getDateCellValue();
                                        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                        fieldObject.put(header.getCell(j).toString().trim(), sdf.format(dateValue));
                                        break;
                                    } else {

                                        fieldObject.put(header.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
                                        break;
                                    }
                                case FORMULA:
                                    switch (evaluator.evaluateFormulaCell(cell)) {
                                        case NUMERIC:
                                            if (DateUtil.isCellDateFormatted(cell)) {
                                                Date dateValue = cell.getDateCellValue();
                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                                fieldObject.put(header.getCell(j).toString().trim(), sdf.format(dateValue));
                                            } else {
                                                fieldObject.put(header.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
                                            }
                                            break;
                                        case STRING:
                                            fieldObject.put(header.getCell(j).toString().trim(), cell.getStringCellValue().trim());
                                            break;
                                        default:
                                            DataFormatter df = new DataFormatter();
                                            String value = df.formatCellValue(cell);
                                            fieldObject.put(header.getCell(j).toString().trim(), value);
                                    }
                                    break;
                                default:
                                    DataFormatter df = new DataFormatter();
                                    String value = df.formatCellValue(cell).trim();
                                    fieldObject.put(header.getCell(j).toString().trim(), value);
                            }
                        } catch (IllegalStateException e) {
                            System.out.println("Error when read data at Cell:" + header.getCell(j).toString() + " Cell type:" + cellType);
                            System.out.println("Error type:" + e.getMessage());
                            throw e;
                        }

                    }
                    pageObject.put(key, fieldObject);

                }
                testCaseObject.put(testCaseID, pageObject);
            }
        }
        featureObject.put(sheet.getSheetName(), testCaseObject);
        Gson gson = new GsonBuilder().disableHtmlEscaping().create();
        String jsonString = gson.toJson(featureObject, LinkedHashMap.class);
        //String jsonString = new Gson().toJson(featureObject, LinkedHashMap.class);
        String path = "." + "\\data\\" + sheet.getSheetName() + ".json";
        writeJsonToFile(jsonString, path);
    }

}




