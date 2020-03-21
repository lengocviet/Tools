package Json;

import com.google.gson.*;
import com.google.gson.stream.JsonReader;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.lang.reflect.Type;

import java.util.*;

public class JsonUtils {

    public static <T> T getData(String jsonPath, Type type, String dateFormat) throws Exception {
        try {

            GsonBuilder gsonBuilder = new GsonBuilder();
            gsonBuilder.setDateFormat(dateFormat);
            Gson gson = gsonBuilder.create();

            JsonReader reader = getJsonReader(jsonPath);
            return gson.fromJson(reader, type);
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> T getData(String jsonPath, Type type) throws Exception {
        try {

            GsonBuilder gsonBuilder = new GsonBuilder();
            Gson gson = gsonBuilder.create();

            JsonReader reader = getJsonReader(jsonPath);
            return gson.fromJson(reader, type);
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> T getData(String jsonPath, Class<?> clazz, String dateFormat) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            gsonBuilder.setDateFormat(dateFormat);
            Gson gson = gsonBuilder.create();
            JsonReader reader = getJsonReader(jsonPath);
            return gson.fromJson(reader, clazz);
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> T getData(String jsonPath, String key, Class<T> clazz) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            Gson gson = gsonBuilder.create();
            JsonReader reader = getJsonReader(jsonPath);
            JsonObject object = gson.fromJson(reader, JsonObject.class);
            return gson.fromJson(object.get(key), clazz);
        } catch (Exception e) {
            throw e;
        }
    }

    public static JsonObject getData(String jsonPath) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            Gson gson = gsonBuilder.create();
            JsonReader reader = getJsonReader(jsonPath);
            gson.fromJson(reader, JsonObject.class);
            return new JsonObject();
        } catch (Exception e) {
            throw e;
        }
    }

    public static String getNodeFromValue(String key, String value, String jsonPath) throws Exception {
        String node = "";
        JsonObject obj = getData(jsonPath);
        if (obj != null) {
            Set<String> keys = obj.keySet();
            for (String nodeName : keys) {
                Object object = obj.get(nodeName);
                if (object instanceof JsonObject) {
                    JsonObject nodeObj = (JsonObject) object;
                    if (nodeObj != null && nodeObj.get(key).getAsString().equals(value)) {
                        node = nodeName;
                        break;
                    }
                }
            }
        }
        return node;
    }

    public static <T> List<T> getListData(String jsonPath, Type type, String dateFormat) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            gsonBuilder.setDateFormat(dateFormat);
            Gson gson = gsonBuilder.create();
            JsonReader reader = getJsonReader(jsonPath);
            List<T> lst;
            lst = gson.fromJson(reader, type);
            return lst;
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> List<T> getDataAsList(String jsonPath, Type type) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            Gson gson = gsonBuilder.create();
            JsonReader reader = getJsonReader(jsonPath);
            List<T> lst;
            lst = gson.fromJson(reader, type);
            return lst;
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> List<T> getListData(String jsonPath, Class<?> clazz, String dateFormat) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            gsonBuilder.setDateFormat(dateFormat);
            Gson gson = gsonBuilder.create();
            JsonReader reader = getJsonReader(jsonPath);
            List<T> lst;
            lst = gson.fromJson(reader, clazz);
            return lst;
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> List<T> getListFromJsonObj(JsonElement jsonElement, Type classType, String dateFormat) {
        GsonBuilder gsonBuilder = new GsonBuilder();
        gsonBuilder.setDateFormat(dateFormat);
        Gson gson = gsonBuilder.create();
        List<T> lst = new ArrayList<T>();
        lst = gson.fromJson(jsonElement, classType);
        return lst;
    }

    private static JsonReader getJsonReader(String jsonPath) {
        try {
            JsonReader reader;
            reader = new JsonReader(new FileReader(jsonPath));
            return reader;
        } catch (FileNotFoundException e) {
            return null;
        }
    }

    public static Map<String, Map<String, String>> parseJsonToMap(String jsonPath) {
        Gson gson = new Gson();
        Map<String, Map<String, String>> map = new HashMap<String, Map<String, String>>();
        JsonReader reader = getJsonReader(jsonPath);
        return gson.fromJson(reader, map.getClass());
    }

    public static JSONObject readFileJson(String link) {
        JSONParser parser = new JSONParser();
        JSONObject jsonObject = new JSONObject();
        try {
            Object obj = parser.parse(new FileReader(link));
            jsonObject = (JSONObject) obj;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return jsonObject;
    }

    public static <T> List<T> getDataAsListByJsonString(String jsonString, Type type) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            Gson gson = gsonBuilder.create();
            List<T> lst;
            lst = gson.fromJson(jsonString, type);
            return lst;
        } catch (Exception e) {
            throw e;
        }
    }

    public static <T> T getDataByJsonString(String jsonString, Type type) throws Exception {
        try {
            GsonBuilder gsonBuilder = new GsonBuilder();
            Gson gson = gsonBuilder.create();
            return gson.fromJson(jsonString, type);
        } catch (Exception e) {
            throw e;
        }
    }
}



