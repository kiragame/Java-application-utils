import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;

/**
 * Excel导出;请自行安装依赖包;可完全替换apache包
 */
public final class ExcelUtils {

    /**
     * excel版本
     */
    private static final String EXCEL_VERSION = ".xls";

    /**
     * 定义编码
     */
    private static final String ENCODE_TYPE = "UTF-8";

    /**
     * 日志
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 工具类隐藏构造器
     */
    private ExcelUtils() {
    }

    /**
     * 写入多个sheet
     *
     * @param data     Map<sheetName,数据>  数据格式为List<List<String>> 从里到外list对应行、列
     * @param response HTTP返回流
     * @param browser  浏览器类型
     * @param fileName 文件名
     */
    public static void getExcel(Map<String, List<List<String>>> data, HttpServletResponse response,
                                String browser, String fileName) {
        fileName = fileName + EXCEL_VERSION;
        fileName = encodeFileName(fileName, browser);
        String headStr = "attachment; filename=\"" + fileName + "\"";
        response.setContentType("APPLICATION/OCTET-STREAM");
        response.setHeader("Content-Disposition", headStr);
        //创建一个excel
        WritableWorkbook workbook = null;
        try (OutputStream output = response.getOutputStream()) {
            workbook = Workbook.createWorkbook(output);
            //创建一个sheet
            WritableSheet sheet;
            int sheetIndex = 0;
            for (Map.Entry<String, List<List<String>>> entry : data.entrySet()) {
                sheet = workbook.createSheet(entry.getKey(), sheetIndex);
                fillRow(entry.getValue(), sheet);
                sheetIndex++;
            }
            workbook.write();
            workbook.close();
        } catch (IOException e) {
            throw XXXX;
        } catch (WriteException e) {
            
            .error(e.getMessage());
        }
    }

    /**
     * 文件名编码
     *
     * @param fileName 文件名
     * @param browser  浏览器版本
     * @return 编码后的字符串
     */
    private static String encodeFileName(String fileName, String browser) {
        //IE浏览器
        if ("MSIE".equals(browser)) {
            try {
                return URLEncoder.encode(fileName, ENCODE_TYPE);
            } catch (UnsupportedEncodingException e) {
                LOGGER.error(e.getMessage());
            }
        }
        //google,火狐浏览器
        else if ("Mozilla".equals(browser)) {
            try {
                return new String(fileName.getBytes(), ENCODE_TYPE);
            } catch (UnsupportedEncodingException e) {
                LOGGER.error(e.getMessage());
            }
        }
        //其他浏览器
        else {
            try {
                return URLEncoder.encode(fileName, ENCODE_TYPE);
            } catch (UnsupportedEncodingException e) {
                LOGGER.error(e.getMessage());
            }
        }
        return fileName;
    }

    /**
     * 填入列
     *
     * @param data  数据
     * @param sheet 表格
     */
    private static void fillRow(List<List<String>> data, WritableSheet sheet) throws WriteException {
        int rowNum = 0;
        int cellNum;
        for (List<String> list : data) {
            cellNum = 0;
            for (String obj : list) {
                Label label = new Label(cellNum++, rowNum, obj == null ? "" : obj);
                sheet.addCell(label);
            }
            rowNum++;
        }
    }

    /**
     * 文件写入response流
     *
     * @param response 返回流
     * @param workbook 工作簿
     */
    private static void writeToHttpServletRsponse(HttpServletResponse response, XSSFWorkbook workbook) {
        OutputStream output;
        try {
            output = response.getOutputStream();
            workbook.write(output);
            workbook.close();
        } catch (IOException e) {
            throw xxx;
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                LOGGER.error(e.getMessage());
            }
        }
    }
}
