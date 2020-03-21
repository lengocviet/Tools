import Excel.ConvertExcelToJson;
import Json.DataDrivenHelper;
import Json.JsonUtils;
import com.google.gson.JsonObject;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;


import javax.swing.plaf.synth.SynthOptionPaneUI;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class Main<dataFile> {

    public static  void main(String[] args) throws Exception {
//    System.out.print(System.getProperty("user.dir"));

//        long startTime = System.currentTimeMillis();
//        List<Object> TC1 = DataDrivenHelper.getDataForTC("SCE-WEB-001-039");
//        List<Object> TC2 = DataDrivenHelper.getDataForTC("SCE-WEB-001-1285");
//        long endTime   = System.currentTimeMillis();
//        long totalTime = endTime - startTime;
//        System.out.println(totalTime);
//---------------------
//        List<Object> results = DataDrivenHelper.getFieldvalue("Home", "abc");
//        System.out.print("done");


        //Test xml
//        String dataFile = "." + "\\data\\configure.xml";
//        xmlUtils.readxml(dataFile);
        //Test regrex
//        String testString = "//h4[text()='Credit Card']/following-sibling::div";
//
//        String body = testString.replaceAll("[\\w\\s]*=(.*)", "$1").trim();
//        String type = testString.replaceAll("([\\w\\s]*)=.*", "$1").trim();
//
//
//        System.out.println("Body:" + body);
//        System.out.println("Type:" + type);
//        //String type = this.locator.replaceAll("([\\w\\s]*)=.*", "$1").trim();

        //DataDrivenHelper.getData();

//Convert XML to JSON
         String path = "D:\\VVVV\\Doc\\100. Projects\\11.Direct Asia\\5.Framework\\3. Tools\\AllUtilities\\data\\Data_Driven_LG_Template3.xlsx";
        try {
            ConvertExcelToJson.CreateJsonFromExcel(path);

        }
        catch (Exception e)
        {
            System.out.print(e.getMessage());
        }



    }






}
