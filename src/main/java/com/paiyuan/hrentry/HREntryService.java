package com.paiyuan.hrentry;

import cn.com.weaver.services.webservices.WorkflowServicePortTypeProxy;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import weaver.workflow.webservices.*;

import javax.annotation.PostConstruct;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.rmi.RemoteException;
import java.util.HashMap;

/*
 * 使用前将配置文件中的creatorId修改为自己的id，将文件路径修改为正确的值
 * 添加新字段时，请看下方注释里面的内容
 */

@Service
public class HREntryService {

    @Value("${filePath}")
    private String filePath;

    @Value("${creatorId}")
    private String creatorId;

    @Value("${workflowId}")
    private String workflowId;

    @Value("${rowStart}")
    private int rowStart;

    @Value("${rowEnd}")
    private int rowEnd;

    @Value("${sheetId}")
    private int sheetId;

    @Value("${requestName}")
    private String requestName;
    //TO-DO
    private final HashMap<String, String> deptMap = new HashMap<String, String>() {{

        put("安全质量部", "7");
        put("巴基斯坦分公司", "99");
        put("材料能源事业部", "18");
        put("财务部", "27");
        put("采购部", "16");
        put("诚信监理", "10");
        put("党群工作部", "2");
        put("党委办公室", "1");
        put("党委组织部", "24");
        put("电控室", "17");
        put("董事会办公室", "47");
        put("法务部", "31");
        put("工艺室", "6");
        put("公司办公室", "9");
        put("公司领导", "39");
        put("管道室", "14");
        put("广州联络处", "57");
        put("国际工程部", "26");
        put("国际事业部", "25");
        put("华陆工程管理", "98");
        put("华陆实业（香港）", "38");
        put("华陆新材", "53");
        put("环保事业部", "44");
        put("基础设施环保事业部", "35");
        put("纪检监督部", "32");
        put("纪委办公室", "36");
        put("技术发展部", "5");
        put("监事会办公室", "48");
        put("经发物业", "34");
        put("考核审计部", "28");
        put("人力资源部", "20");
        put("商务部", "19");
        put("设备室", "15");
        put("施工管理部", "4");
        put("数字化工程中心", "52");
        put("土建室", "8");
        put("外部监事", "76");
        put("文印中心", "23");
        put("物业公司", "11");
        put("项目管理部", "3");
        put("项目控制部", "22");
        put("新疆办事处", "56");
        put("信息档案中心", "21");
        put("巡察办公室", "37");
        put("研发中心", "30");
        put("亿阳餐饮", "33");
        put("战略规划部", "29");
        put("咨询部", "13");
        put("综合事务部", "12");
    }};

    @PostConstruct
    public void startService() throws IOException {

        System.out.println("filePath: " + filePath + "\ncreatorId: " + creatorId + "\nworkflowId: " + workflowId);
        parseExcel();
    }

    public void parseExcel() throws IOException {

        System.out.println(requestName + "");

        Path path = Paths.get(filePath);
        InputStream is = Files.newInputStream(path);

        Workbook workbook = null;

        if (filePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(is);
        } else if (filePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(is);
        } else {
            System.out.println("[ERROR] Unsupported source file format.");
            System.exit(1);
        }

        Sheet sheet = workbook.getSheetAt(sheetId);
        int realRowEnd = rowEnd == -1 ? sheet.getLastRowNum() : rowEnd;
        Row row;
        for (int i = rowStart; i <= realRowEnd; i++) {
            row = sheet.getRow(i);
            if (row != null) {
                JSONObject user = new JSONObject();

                user.put("name", row.getCell(0).getStringCellValue());
                user.put("gender", row.getCell(1).getStringCellValue().equals("男") ? "0" : "1");
                user.put("dept", deptMap.get(row.getCell(2).getStringCellValue()));
                user.put("corp", row.getCell(3).getStringCellValue());
                user.put("loginId", row.getCell(4).getStringCellValue());
                user.put("id", Integer.toString((int)row.getCell(5).getNumericCellValue()));
                user.put("mobile", Double.toString(row.getCell(6).getNumericCellValue()));
                //在这里添加新的字段
                createWorkflow(user);
            }
        }
        System.out.println("[INFO] Submission succeeded. Please check if there is any negative requestId.");
    }

    public void createWorkflow(JSONObject user) throws RemoteException {

        WorkflowBaseInfo wbi = new WorkflowBaseInfo();
        wbi.setWorkflowId(workflowId);

        WorkflowRequestInfo wri = new WorkflowRequestInfo();
        wri.setCreatorId(creatorId);
        wri.setCanView(true);
        wri.setCanEdit(true);
        wri.setRequestName(requestName);
        wri.setRequestLevel("0");
        wri.setIsnextflow("0");
        wri.setWorkflowBaseInfo(wbi);

        WorkflowRequestTableField[] wrtf = new WorkflowRequestTableField[7];    //新增字段时，要修改这里的容器大小

        wrtf[0] = new WorkflowRequestTableField("rzry", user.getString("name"));
        wrtf[1] = new WorkflowRequestTableField("xb", user.getString("gender"));
        wrtf[2] = new WorkflowRequestTableField("rzbm", user.getString("dept"));
        wrtf[3] = new WorkflowRequestTableField("dwmc", user.getString("corp"));
        wrtf[4] = new WorkflowRequestTableField("dlzh", user.getString("loginId"));
        wrtf[5] = new WorkflowRequestTableField("ygbh", user.getString("id"));
        wrtf[6] = new WorkflowRequestTableField("yddh", user.getString("mobile"));
        //在这里添加新的字段

        WorkflowRequestTableRecord[] wrtr = new WorkflowRequestTableRecord[1];
        wrtr[0] = new WorkflowRequestTableRecord();
        wrtr[0].setWorkflowRequestTableFields(wrtf);

        WorkflowMainTableInfo wmti = new WorkflowMainTableInfo();
        wmti.setRequestRecords(wrtr);

        wri.setWorkflowMainTableInfo(wmti);

        WorkflowServicePortTypeProxy testproxy = new WorkflowServicePortTypeProxy();
        String requestId = testproxy.doCreateWorkflowRequest(wri, Integer.parseInt(creatorId));

        System.out.println(user.getString("name") + " requestId: " + requestId);
    }
}
