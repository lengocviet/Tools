import Excel.ConvertExcelToJson;

public class Main {

    public static void main(String[] args) {

        String path = "D:\\VVVV\\Doc\\100. Projects\\11.Direct Asia\\5.Framework\\3. Tools\\AllUtilities\\data\\Data_Driven_LG_Template3.xlsx";
        try {
            ConvertExcelToJson.CreateJsonFilesFromExcel(path);

        } catch (Exception e) {
            System.out.print(e.getMessage());
        }


    }


}
