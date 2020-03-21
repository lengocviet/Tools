package Json;

import java.io.FileReader;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class DataDrivenHelper {
    static JSONArray testingData;
    static JSONArray metaData;

    static {
        try {
            initData();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
    }

    static private JSONArray getJsonArrayFromJson(String jsonFile) throws IOException, ParseException {
        Object object = new JSONParser().parse(new FileReader(jsonFile));
        return (JSONArray) object;

    }

    static private void initData() throws IOException, ParseException {
        String dataPath = "D:\\VVVV\\Doc\\100. Projects\\11.Direct Asia\\5.Framework\\3. Tools\\AllUtilities" + "\\data\\Car.json";
        testingData = getJsonArrayFromJson(dataPath);
        String metaDataPath = "D:\\VVVV\\Doc\\100. Projects\\11.Direct Asia\\5.Framework\\3. Tools\\AllUtilities" + "\\data\\metaData.json";
        metaData = getJsonArrayFromJson(metaDataPath);
    }

    static public List<Object> getDataForTC(String TCID) throws IOException, ParseException {
        List<Object> result = IntStream.range(0, testingData.size())
                .mapToObj(index -> ((JSONObject) testingData.get(index))).filter(x -> x.get("TC-ID").equals(TCID))
                .collect(Collectors.toList());
        return result;
    }

    static public List<Object> getFieldvalue(String page, String fieldName) throws IOException, ParseException {
        List<JSONObject> pageResult = IntStream.range(0, metaData.size())
                .mapToObj(index -> ((JSONObject) metaData.get(index))).filter(x -> x.get("Page").equals(page))
                .collect(Collectors.toList());
        if (pageResult != null)
        {
            List<JSONObject> fieldList = (List<JSONObject>) pageResult.get(0).get("Fields");
            System.out.print("done");
        }


        List<Object> result = IntStream.range(0, testingData.size())
                .mapToObj(index -> ((JSONObject) testingData.get(index))).filter(x -> x.get("Page").equals(page))
                .collect(Collectors.toList());
        return result;
    }

    static public void getAllFieldName (JSONObject jsonObject) throws IOException, ParseException {
        List<Object> fields = getFieldvalue("Home", "carBrand");
    }

}


