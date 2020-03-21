package Excel;

import com.google.gson.Gson;
import com.google.gson.stream.JsonWriter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.json.JSONArray;
import org.json.JSONObject;
import org.json.JSONWriter;
import org.junit.platform.commons.function.Try;


import java.io.*;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.*;

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
      //  createJsonDataAsObject(excelWorkBook.getSheet("TravelAnnual"), annualTravelStructure);
//        createJsonDataAsObject(excelWorkBook.getSheet("Motorcycle"),motorCycleStructure );
        createJsonDataAsObject(excelWorkBook.getSheet("Car"), motorCarStructure);

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

    private static void writeJsonToFile(JSONArray data, String file) throws IOException {

        FileWriter jsonFileWriter = new FileWriter(file);
        String jsonString = data.toString();
        jsonFileWriter.write(jsonString);
        jsonFileWriter.flush();
    }

    private static void writeJsonToFile(String data, String file) throws IOException {

        FileWriter jsonFileWriter = new FileWriter(file);
        jsonFileWriter.write(data);
        jsonFileWriter.flush();
    }


    //    private static JSONArray createJsonData(XSSFSheet sheet) {
//
//        List<List<String>> ret = new ArrayList<List<String>>();
//        JSONArray listObjects = new JSONArray();
//        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        // Get the first and last sheet row number.
//        int firstRowNum = sheet.getFirstRowNum();
//        int lastRowNum = sheet.getLastRowNum();
//        //if excel file has data
//        if (lastRowNum > 1) {
//            for (int i = 2; i <= lastRowNum; i++) {
//                JSONObject object = new JSONObject();
//                XSSFRow rowHeader = sheet.getRow(1);
//                for (String key : new ArrayList<String>(annualTravelStructure.keySet())) {
//                    if ("TC_ID".equalsIgnoreCase(key)) {
//                        object.put("TC_ID", sheet.getRow(i).getCell(0).toString().trim());
//                    } else {
//                        JSONObject subObject = new JSONObject();
//                        int firstColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[0]);
//                        int lastColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[1]);
//                        XSSFRow currentRow = sheet.getRow(i);
//                        for (int j = firstColumn; j <= lastColumn; j++) {
//                            XSSFCell cell = currentRow.getCell(j);
//                            CellType cellType = cell.getCellType();
//                            if (cellType == CellType.FORMULA) {
//                                switch (evaluator.evaluateFormulaCell(cell)) {
//                                    case BOOLEAN:
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getBooleanCellValue());
//                                        break;
//                                    case NUMERIC:
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getNumericCellValue());
//                                        break;
//                                    default:
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                }
//                            } else if (cellType == CellType.STRING) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                            } else if (cellType == CellType.NUMERIC) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getNumericCellValue());
//                            } else if (cellType == CellType.BOOLEAN) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getBooleanCellValue());
//                            } else if (cellType == CellType.BLANK) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), "None");
//                            } else {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue());
//                            }
//
//                        }
//                        object.put(key, subObject);
//                    }
//                }
//
//                listObjects.put(object);
//            }
//        }
//        return listObjects;
//
//    }
//
//    private static JSONObject createJsonDataAsObject1(XSSFSheet sheet) {
//
//        List<List<String>> ret = new ArrayList<List<String>>();
//        JSONArray listObjects = new JSONArray();
//        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        // Get the first and last sheet row number.
//        int firstRowNum = sheet.getFirstRowNum();
//        int lastRowNum = sheet.getLastRowNum();
//        //if excel file has data
//        JSONObject featureObject = new JSONObject();
//        JSONObject testCaseObject = new JSONObject();
//        if (lastRowNum > 1) {
//            for (int i = 2; i <= 5; i++) {
//                JSONObject object = new JSONObject();
//                XSSFRow rowHeader = sheet.getRow(1);
//                String testCaseID = "";
//                for (String key : new ArrayList<String>(annualTravelStructure.keySet())) {
//                    if ("TC_ID".equalsIgnoreCase(key)) {
//                        testCaseID = sheet.getRow(i).getCell(0).toString().trim();
//                    } else {
//                        JSONObject subObject = new JSONObject();
//                        int firstColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[0]);
//                        int lastColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[1]);
//                        XSSFRow currentRow = sheet.getRow(i);
//                        for (int j = firstColumn; j <= lastColumn; j++) {
//                            XSSFCell cell = currentRow.getCell(j);
//                            CellType cellType = cell.getCellType();
//                            if (cellType == CellType.FORMULA) {
//                                switch (evaluator.evaluateFormulaCell(cell)) {
//                                    case BOOLEAN:
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getBooleanCellValue());
//                                        break;
//                                    case NUMERIC:
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getNumericCellValue());
//                                        break;
//                                    default:
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                }
//                            } else if (cellType == CellType.STRING) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                            } else if (cellType == CellType.NUMERIC) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getNumericCellValue());
//                            } else if (cellType == CellType.BOOLEAN) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getBooleanCellValue());
//                            } else if (cellType == CellType.BLANK) {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), "None");
//                            } else {
//                                subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue());
//                            }
//
//                        }
//                        object.put(key, subObject);
//                    }
//                }
//                testCaseObject.put(testCaseID, object);
//                //listObjects.put(object);
//            }
//        }
//        featureObject.put("TR_AnnualPlans", testCaseObject);
//        return featureObject;
//
//    }
//
//    private static JSONObject createJsonDataAsObject2(XSSFSheet sheet) {
//
//        List<List<String>> ret = new ArrayList<List<String>>();
//        JSONArray listObjects = new JSONArray();
//        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        // Get the first and last sheet row number.
//        int firstRowNum = sheet.getFirstRowNum();
//        int lastRowNum = sheet.getLastRowNum();
//        //if excel file has data
//        JSONObject featureObject = new JSONObject();
//        JSONObject testCaseObject = new JSONObject();
//        if (lastRowNum > 1) {
//            for (int i = 2; i <= 5; i++) {
//                JSONObject object = new JSONObject();
//                XSSFRow rowHeader = sheet.getRow(1);
//                String testCaseID = "";
//                for (String key : new ArrayList<String>(annualTravelStructure.keySet())) {
//
//                    if ("TC_ID".equalsIgnoreCase(key)) {
//                        testCaseID = sheet.getRow(i).getCell(0).toString().trim();
//                    } else {
//                        JSONObject subObject = new JSONObject();
//                        int firstColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[0]);
//                        int lastColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[1]);
//                        XSSFRow currentRow = sheet.getRow(i);
//                        for (int j = firstColumn; j <= lastColumn; j++) {
//                            XSSFCell cell = currentRow.getCell(j);
//                            CellType cellType = cell.getCellType();
//                            if (cellType == CellType.FORMULA) {
//                                if (cell.toString().contains("TODAY()")) {
//                                    Date dateValue = cell.getDateCellValue();
//                                    SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                    subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(dateValue).toString());
//                                } else {
//                                    switch (evaluator.evaluateFormulaCell(cell)) {
//                                        case BOOLEAN:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getBooleanCellValue());
//                                            break;
//                                        case NUMERIC:
//                                            if (rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getNumericCellValue()).toString());
//
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Double.toString(cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            } else {
//
//                                if ((cell.getRawValue() != null) && (!("".equals(cell.getRawValue())))) {
//
//                                    if (cellType == CellType.STRING) {
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                    } else if (cellType == CellType.NUMERIC) {
//
//                                        if (DateUtil.isCellDateFormatted(cell) && rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getNumericCellValue()).toString());
//                                        } else {
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Double.toString(cell.getNumericCellValue()));
//                                        }
//                                    } else if (cellType == CellType.BOOLEAN) {
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//                                    } else {
//                                        subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            }
//
//
//                        }
//                        object.put(key, subObject);
//                    }
//
//                }
//                testCaseObject.put(testCaseID, object);
//            }
//
//        }
//        featureObject.put("TR_AnnualPlans", testCaseObject);
//
//        return featureObject;
//    }
//
//    private static void createJsonDataAsObject(XSSFSheet sheet) throws IOException {
//
//        List<List<String>> ret = new ArrayList<List<String>>();
//        JSONArray listObjects = new JSONArray();
//        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        // Get the first and last sheet row number.
//        int firstRowNum = sheet.getFirstRowNum();
//        int lastRowNum = sheet.getLastRowNum();
//        //if excel file has data
//        // JSONObject featureObject = new JSONObject();
//        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>>> featureObject = new LinkedHashMap<>();
//        //JSONObject testCaseObject = new JSONObject();
//        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>> testCaseObject = new LinkedHashMap<>();
//        if (lastRowNum > 1) {
//            for (int i = 2; i <= 5; i++) {
//                // JSONObject object = new JSONObject();
//                LinkedHashMap<String, LinkedHashMap<String, String>> object = new LinkedHashMap<>();
//                XSSFRow rowHeader = sheet.getRow(1);
//                String testCaseID = "";
//                //for (String key : new ArrayList<String>(annualTravelStructure.keySet())) {
//                List<String> keys = new ArrayList<String>(annualTravelStructure.keySet());
//                for (int index = 0; index < keys.size(); index++) {
//                    String key = keys.get(index);
//                    if ("TC_ID".equalsIgnoreCase(key)) {
//                        testCaseID = sheet.getRow(i).getCell(0).toString().trim();
//                    } else {
//                        //JSONObject subObject = new JSONObject();
//                        LinkedHashMap<String, String> subObject = new LinkedHashMap<>();
//                        int firstColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[0]);
//                        int lastColumn = CellReference.convertColStringToIndex(annualTravelStructure.get(key)[1]);
//                        XSSFRow currentRow = sheet.getRow(i);
//                        for (int j = firstColumn; j <= lastColumn; j++) {
//                            XSSFCell cell = currentRow.getCell(j);
//                            CellType cellType = cell.getCellType();
//                            if (cellType == CellType.FORMULA) {
//                                if (cell.toString().contains("TODAY()")) {
//                                    Date dateValue = cell.getDateCellValue();
//                                    SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                    subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(dateValue).toString());
//                                } else {
//                                    switch (evaluator.evaluateFormulaCell(cell)) {
//                                        case NUMERIC:
//                                            if (rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        case BOOLEAN:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//                                            break;
//                                        case STRING:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            } else {
//                                if ((cell.getRawValue() != null) && (!("".equals(cell.getRawValue())))) {
//                                    switch (cellType) {
//                                        case NUMERIC:
//                                            if (DateUtil.isCellDateFormatted(cell) && rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        case BOOLEAN:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//
//                                        case STRING:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            }
//                        }
//                        object.put(key, subObject);
//
//                    }
//                    testCaseObject.put(testCaseID, object);
//                }
//            }
//            featureObject.put("TR_AnnualPlans", testCaseObject);
//        }
//
//        String path = "." + "\\data\\Annual_travel.json";
//        FileWriter jsonFileWriter = new FileWriter(path);
//        String jsonString = new Gson().toJson(featureObject, LinkedHashMap.class);
//        jsonFileWriter.write(jsonString);
//        jsonFileWriter.flush();
//
//
//    }
//
//    private static void createJsonDataAsObjectMC(XSSFSheet sheet) throws IOException {
//
//        List<List<String>> ret = new ArrayList<List<String>>();
//        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        // Get the first and last sheet row number.
//        int firstRowNum = sheet.getFirstRowNum();
//        int lastRowNum = sheet.getLastRowNum();
//        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>>> featureObject = new LinkedHashMap<>();
//        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>> testCaseObject = new LinkedHashMap<>();
//
//        if (lastRowNum > 1) {
//            for (int i = 2; i <= lastRowNum; i++) {
//                // JSONObject object = new JSONObject();
//                LinkedHashMap<String, LinkedHashMap<String, String>> object = new LinkedHashMap<>();
//                XSSFRow rowHeader = sheet.getRow(1);
//                String testCaseID = "";
//                //for (String key : new ArrayList<String>(motorCycleStructure.keySet())) {
//                List<String> keys = new ArrayList<String>(motorCycleStructure.keySet());
//                for (int index = 0; index < keys.size(); index++) {
//                    String key = keys.get(index);
//                    if ("TC_ID".equalsIgnoreCase(key)) {
//                        testCaseID = sheet.getRow(i).getCell(0).toString().trim();
//                    } else {
//                        //JSONObject subObject = new JSONObject();
//                        LinkedHashMap<String, String> subObject = new LinkedHashMap<>();
//                        int firstColumn = CellReference.convertColStringToIndex(motorCycleStructure.get(key)[0]);
//                        int lastColumn = CellReference.convertColStringToIndex(motorCycleStructure.get(key)[1]);
//                        XSSFRow currentRow = sheet.getRow(i);
//                        for (int j = firstColumn; j <= lastColumn; j++) {
//                            XSSFCell cell = currentRow.getCell(j);
//                            CellType cellType = cell.getCellType();
//                            if (cellType == CellType.FORMULA) {
//                                if (cell.toString().contains("TODAY()")) {
//                                    Date dateValue = cell.getDateCellValue();
//                                    SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                    subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(dateValue).toString());
//                                } else {
//                                    switch (evaluator.evaluateFormulaCell(cell)) {
//                                        case NUMERIC:
//
//                                            if (rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        case BOOLEAN:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//                                            break;
//                                        case STRING:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            } else {
//                                if ((cell.getRawValue() != null) && (!("".equals(cell.getRawValue())))) {
//                                    switch (cellType) {
//                                        case NUMERIC:
//                                            if (DateUtil.isCellDateFormatted(cell) && rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        case BOOLEAN:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//
//                                        case STRING:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            }
//
//                            object.put(key, subObject);
//
//
//                        }
//
//                    }
//                    testCaseObject.put(testCaseID, object);
//                }
//            }
//            featureObject.put("MotorCycle", testCaseObject);
//        }
//
//        String path = "." + "\\data\\MotorCycle.json";
//        FileWriter jsonFileWriter = new FileWriter(path);
//        String jsonString = new Gson().toJson(featureObject, LinkedHashMap.class);
//        jsonFileWriter.write(jsonString);
//        jsonFileWriter.flush();
//    }
//
//    private static void createJsonDataAsObjectMT(XSSFSheet sheet) throws IOException {
//
//        List<List<String>> ret = new ArrayList<List<String>>();
//        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        // Get the first and last sheet row number.
//        int firstRowNum = sheet.getFirstRowNum();
//        int lastRowNum = sheet.getLastRowNum();
//        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>>> featureObject = new LinkedHashMap<>();
//        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>> testCaseObject = new LinkedHashMap<>();
//
//        if (lastRowNum > 1) {
//            for (int i = 2; i <= lastRowNum; i++) {
//                // JSONObject object = new JSONObject();
//                LinkedHashMap<String, LinkedHashMap<String, String>> object = new LinkedHashMap<>();
//                XSSFRow rowHeader = sheet.getRow(1);
//                String testCaseID = "";
//                //for (String key : new ArrayList<String>(motorCycleStructure.keySet())) {
//                List<String> keys = new ArrayList<String>(motorCarStructure.keySet());
//                for (int index = 0; index < keys.size(); index++) {
//                    String key = keys.get(index);
//                    if ("TC_ID".equalsIgnoreCase(key)) {
//                        testCaseID = sheet.getRow(i).getCell(0).toString().trim();
//                    } else {
//                        //JSONObject subObject = new JSONObject();
//                        LinkedHashMap<String, String> subObject = new LinkedHashMap<>();
//                        int firstColumn = CellReference.convertColStringToIndex(motorCarStructure.get(key)[0]);
//                        int lastColumn = CellReference.convertColStringToIndex(motorCarStructure.get(key)[1]);
//                        XSSFRow currentRow = sheet.getRow(i);
//                        for (int j = firstColumn; j <= lastColumn; j++) {
//                            XSSFCell cell = currentRow.getCell(j);
//                            CellType cellType = cell.getCellType();
//                            if (cellType == CellType.FORMULA) {
//                                if (cell.toString().contains("TODAY()")) {
//                                    Date dateValue = cell.getDateCellValue();
//                                    SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                    subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(dateValue).toString());
//                                } else {
//                                    switch (evaluator.evaluateFormulaCell(cell)) {
//                                        case NUMERIC:
//
//                                            if (rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        case BOOLEAN:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//                                            break;
//                                        case STRING:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            } else {
//                                if ((cell.getRawValue() != null) && (!("".equals(cell.getRawValue())))) {
//                                    switch (cellType) {
//                                        case NUMERIC:
//                                            if (DateUtil.isCellDateFormatted(cell) && rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
//                                                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
//                                            } else {
//                                                subObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
//                                            }
//                                            break;
//                                        case BOOLEAN:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
//
//                                        case STRING:
//                                            System.out.println(rowHeader.getCell(j));
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
//                                            break;
//                                        default:
//                                            subObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
//                                    }
//                                }
//                            }
//
//                            object.put(key, subObject);
//
//
//                        }
//
//                    }
//                    testCaseObject.put(testCaseID, object);
//                }
//            }
//            featureObject.put("MotorCar", testCaseObject);
//        }
//
//        String path = "." + "\\data\\MotorCar.json";
//        FileWriter jsonFileWriter = new FileWriter(path);
//        String jsonString = new Gson().toJson(featureObject, LinkedHashMap.class);
//        jsonFileWriter.write(jsonString);
//        jsonFileWriter.flush();
//    }
    private static void createJsonDataAsObject(XSSFSheet sheet, LinkedHashMap<String, String[]> jsonScheme) throws IOException {

        List<List<String>> ret = new ArrayList<List<String>>();
        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>>> featureObject = new LinkedHashMap<>();
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, String>>> testCaseObject = new LinkedHashMap<>();
        /*First row is page -> the second row is header of column. Data is from the third row */
        if (lastRowNum <= 2) return;
        XSSFRow rowType = sheet.getRow(1);
        XSSFRow rowHeader = sheet.getRow(1);
        for (int i = 2; i <=10; i++) {
            LinkedHashMap<String, LinkedHashMap<String, String>> pageObject = new LinkedHashMap<>();
            String testCaseID = "";
            List<String> keys = new ArrayList<String>(jsonScheme.keySet());
            for (int index = 0; index < keys.size(); index++) {
                String key = keys.get(index);
                if ("TC_ID".equalsIgnoreCase(key)) {
                    testCaseID = sheet.getRow(i).getCell(0).toString().trim();
                }
                else
                    {
                    LinkedHashMap<String, String> fieldObject = new LinkedHashMap<>();
                    int firstColumn = CellReference.convertColStringToIndex(jsonScheme.get(key)[0]);
                    int lastColumn = CellReference.convertColStringToIndex(jsonScheme.get(key)[1]);
                    XSSFRow currentRow = sheet.getRow(i);
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        XSSFCell cell = currentRow.getCell(j);
                        CellType cellType = cell.getCellType();
                        if ((cell.getRawValue() != null) && (!("".equals(cell.getRawValue())))) {
                            switch (cellType) {
                                case FORMULA:
                                    if (cell.toString().contains("TODAY()")) {
                                        Date dateValue = cell.getDateCellValue();
                                        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                        fieldObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(dateValue).toString());
                                    } else {
                                        switch (evaluator.evaluateFormulaCell(cell)) {
                                            case NUMERIC:

                                                if (rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
                                                    SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                                    fieldObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
                                                } else {
                                                    fieldObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
                                                }
                                                break;
                                            case BOOLEAN:
                                                System.out.println(rowHeader.getCell(j));
                                                fieldObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));
                                                break;
                                            case STRING:
                                                System.out.println(rowHeader.getCell(j));
                                                fieldObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
                                                break;
                                            default:
                                                fieldObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
                                        }
                                    }
                                    break;
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell) && rowHeader.getCell(j).toString().trim().toLowerCase().contains("date")) {
                                        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                                        fieldObject.put(rowHeader.getCell(j).toString().trim(), sdf.format(cell.getDateCellValue()).toString());
                                    } else {
                                         fieldObject.put(rowHeader.getCell(j).toString().trim(), Integer.toString((int) cell.getNumericCellValue()));
                                    }
                                    break;
                                case BOOLEAN:
                                    System.out.println(rowHeader.getCell(j));
                                    fieldObject.put(rowHeader.getCell(j).toString().trim(), Boolean.toString(cell.getBooleanCellValue()));

                                case STRING:
                                    System.out.println(rowHeader.getCell(j));
                                    fieldObject.put(rowHeader.getCell(j).toString().trim(), cell.getStringCellValue().trim());
                                    break;
                                default:
                                    fieldObject.put(rowHeader.getCell(j).toString().trim(), cell.getRawValue().toString());
                            }

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
        writeJsonToFile(jsonString,path);
    }

}




