package Excel;

import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tools.ant.types.selectors.SelectSelector;

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

    private static LinkedHashMap<String, String[]> annualTravelStructure;
    private static LinkedHashMap<String, String[]> singleTravelStructure;
    private static LinkedHashMap<String, String[]> motorCarStructure;
    private static LinkedHashMap<String, String[]> motorCycleStructure;


    public static void initData() {
        //Todo:  Read setting from XML or Json file and to init dataStructure

        annualTravelStructure = new LinkedHashMap<String, String[]>();
        String[] tcDataRange = new String[]{"A", "A"};
        String[] PortalTravelGetQuotePage = new String[]{"C", "J"};
        String[] PortalTravelChoosePlanPage = new String[]{"K", "P"};
        String[] PortalTravelYourDetailsPage = new String[]{"Q", "CN"};
        String[] PortalTravelSummaryPage = new String[]{"CO", "CS"};
        String[] PortalBuyPage = new String[]{"CT", "CY"};

        annualTravelStructure.put("TC_ID", tcDataRange);
        annualTravelStructure.put("PortalTravelGetQuotePage", PortalTravelGetQuotePage);
        annualTravelStructure.put("PortalTravelChoosePlanPage", PortalTravelChoosePlanPage);
        annualTravelStructure.put("PortalTravelYourDetailsPage", PortalTravelYourDetailsPage);

        annualTravelStructure.put("PortalTravelSummaryPage", PortalTravelSummaryPage);
        annualTravelStructure.put("PortalBuyPage", PortalBuyPage);
//--------------------------Motor Cycle -------------

        String[] motorCycleTestCaseID = new String[]{"A", "A"};
        String[] PortalMotorcycleGetAQuotePage = new String[]{"C", "AG"};
        String[] PortalMotorcycleYourQuotePage = new String[]{"AH", "AI"};
        String[] PortalMotorcycleFinalDetailsPage = new String[]{"AJ", "AU"};
        String[] PortalMotorCycleBuyPage = new String[]{"AV", "BA"};
        motorCycleStructure = new LinkedHashMap<>();
        motorCycleStructure.put("TC_ID", motorCycleTestCaseID);
        motorCycleStructure.put("PortalMotorcycleGetAQuotePage", PortalMotorcycleGetAQuotePage);
        motorCycleStructure.put("PortalMotorcycleYourQuotePage", PortalMotorcycleYourQuotePage);
        motorCycleStructure.put("PortalMotorcycleFinalDetailsPage", PortalMotorcycleFinalDetailsPage);
        motorCycleStructure.put("PortalMotorCycleBuyPage", PortalMotorCycleBuyPage);
//----------------------Motor Car-----------------------------
        String[] motorCarTestCaseID = new String[]{"A", "A"};
        String[] PortalMotorGetAQuotePage = new String[]{"C", "AD"};
        String[] PortalMotorYourQuote = new String[]{"AE", "AF"};
        String[] PortalMotorFinalDetailsPage = new String[]{"AG", "AR"};
        String[] PortalMotorBuyPage = new String[]{"AS", "AX"};
        motorCarStructure = new LinkedHashMap<String, String[]>();
        motorCarStructure.put("TC_ID", motorCarTestCaseID);
        motorCarStructure.put("PortalMotorGetAQuotePage", PortalMotorGetAQuotePage);
        motorCarStructure.put("PortalMotorYourQuote", PortalMotorYourQuote);
        motorCarStructure.put("PortalMotorFinalDetailsPage", PortalMotorFinalDetailsPage);
        motorCarStructure.put("PortalBuyPage", PortalMotorBuyPage);
    }


    public static void CreateJsonFromExcel(String filePath) throws IOException {

        initData();
        XSSFWorkbook excelWorkBook = openExcelFile(filePath);
        LinkedHashMap<String, int[]> scheme = getScheme(excelWorkBook.getSheet("Car").getRow(0));
        //createJsonDataAsObject(excelWorkBook.getSheet("TravelAnnual"), annualTravelStructure);
//        createJsonDataAsObject(excelWorkBook.getSheet("Motorcycle"),motorCycleStructure );
//        createJsonDataAsObject(excelWorkBook.getSheet("Car"), motorCarStructure);

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

    private static LinkedHashMap<String, int[]> getScheme(XSSFRow page) {
        LinkedHashMap<String, int[]> scheme = new LinkedHashMap<>();
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
                    scheme.put(page.getCell(start).getStringCellValue(), startEnd);
                    i = increase ? i + 1 : i;
                }
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());

        }

        return scheme;
    }

    private static void createJsonDataAsObject(XSSFSheet sheet, LinkedHashMap<String, String[]> jsonScheme) throws IOException {

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
                    int firstColumn = CellReference.convertColStringToIndex(jsonScheme.get(key)[0]);
                    int lastColumn = CellReference.convertColStringToIndex(jsonScheme.get(key)[1]);
                    XSSFRow currentRow = sheet.getRow(i);
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        XSSFCell cell = currentRow.getCell(j);
                        //CellType cellType = cell.getCellType();
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




