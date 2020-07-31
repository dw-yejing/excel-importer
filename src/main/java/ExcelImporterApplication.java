import dto.ResearcherImportData;
import listener.ResearcherImportListener;
import util.ExcelUtils;

public class ExcelImporterApplication {

    public static void main(String[] args) {
        String templatePath = "F:\\workspace_exercise\\excelimporter\\src\\main\\resources\\template\\import_template.xlsx";
        ResearcherImportListener researcherImportListener = new ResearcherImportListener();
        ExcelUtils.simpleRead(templatePath, ResearcherImportData.class, researcherImportListener);
    }
}
