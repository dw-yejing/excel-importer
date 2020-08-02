import dto.ResearcherImportData;
import listener.ResearcherImportListener;
import lombok.extern.slf4j.Slf4j;
import util.ExcelUtils;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.ValidatorFactory;
import java.util.Set;

@Slf4j
public class ExcelImporterApplication {

    public static void main(String[] args) {
        String templatePath = "F:\\workspace_exercise\\excel-importer\\src\\main\\resources\\template\\import_template.xlsx";
        ResearcherImportListener researcherImportListener = new ResearcherImportListener();
        ExcelUtils.simpleRead(templatePath, ResearcherImportData.class, researcherImportListener);
        // validation
        ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
        for(ResearcherImportData researcherImportData : researcherImportListener.researcherList){
            Set<ConstraintViolation<ResearcherImportData>> validate = factory.getValidator().validate(researcherImportData);
            if(!validate.isEmpty()){
                for(ConstraintViolation<ResearcherImportData> item : validate){
                    log.error(item.getRootBean().toString()+":\t"+item.getMessage());
                }
            }
        }
        // data processing
        String dataFormatter = "get a piece of data: %s";
        for(ResearcherImportData researcherImportData : researcherImportListener.researcherList){
            String dataInfo = String.format(dataFormatter, researcherImportData.toString());
            log.info(dataInfo);
        }

    }
}
