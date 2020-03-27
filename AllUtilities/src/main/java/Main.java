import Excel.ConvertExcelToJson;

public class Main {

    public static void main(String[] args) {

        String path = "." + "\\data\\Data_Driven_LG_Template_av.xlsx";
        try {
            ConvertExcelToJson.CreateJsonFilesFromExcel(path);

        } catch (Exception e) {
            System.out.print(e.getMessage());
        }

    }


}
