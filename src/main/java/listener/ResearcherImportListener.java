package listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import dto.ResearcherImportData;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ResearcherImportListener extends AnalysisEventListener<ResearcherImportData> {
    public List<ResearcherImportData> researcherList = new ArrayList<>();

    @Override
    public void invoke(ResearcherImportData data, AnalysisContext context){
        researcherList.add(data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context){
        log.info("人员数据解析完成");
    }
}
