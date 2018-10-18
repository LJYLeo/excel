package com.codeshare.config;

import com.codeshare.excel.Excel;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.InputStream;
import java.util.*;

/**
 * 功能描述：
 *
 * @author 18046184刘嘉宇
 * @version 1.0.0
 * @date 2018-10-17 20:12:55
 */
public class Config {

    private static Properties properties;

    static {

        InputStream ips = null;
        try {

            ips = Config.class.getClassLoader().getResourceAsStream("config.properties");
            properties = new Properties();
            properties.load(ips);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (ips != null) {
                    ips.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    @SuppressWarnings("unchecked")
    public static void loadVillageConfig() {

        try {

            String path = Config.class.getClassLoader().getResource("json.txt").toString();
            path = path.replace("file:", "");
            String json = FileUtils.readFileToString(new File(path), "UTF-8");

            JSONArray array = JSONArray.fromObject(json);

            Map<String, Excel> modelMap = new HashMap<String, Excel>();

            JSONObject sheetToModel = array.getJSONObject(0);
            JSONArray sheets = sheetToModel.getJSONArray("sheet");
            for (int i = 0; i < sheets.size(); i++) {
                JSONObject sheet = sheets.getJSONObject(i);
                String sheetName = sheet.getString("sheetName");
                JSONArray models = sheet.getJSONArray("newExcelList");
                List<String> modelList = new ArrayList<String>();
                for (int j = 0; j < models.size(); j++) {
                    JSONObject model = models.getJSONObject(j);
                    String modelName = model.getString("excelName");
                    modelList.add(modelName);
                    Excel newExcel = new Excel();
                    newExcel.setStartRow(model.getInt("newExcelStart"));
                    newExcel.setCell(model.getJSONArray("newExcelCellArray"));
                    Constants.newExcelDataMap.put(modelName, newExcel);
                    Excel oldExcel = new Excel();
                    oldExcel.setCell(model.getJSONArray("oldExcelCellArray"));
                    modelMap.put(modelName, oldExcel);
                }
                Constants.oldToNewMap.put(sheetName, modelList);
            }

            JSONObject excelObject = array.getJSONObject(1);
            JSONArray excels = excelObject.getJSONArray("excel");
            for (int i = 0; i < excels.size(); i++) {
                JSONObject excel = excels.getJSONObject(i);
                String excelName = excel.getString("name");
                String modelName = excel.getString("modelName");
                Excel oldExcel = new Excel();
                oldExcel.setStartRow(excel.getInt("rowStart"));
                oldExcel.setEndRow(excel.getInt("rowEnd"));
                oldExcel.setCell(modelMap.get(modelName).getCell());
                if (Constants.oldExcelDataMap.containsKey(excelName)) {
                    Constants.oldExcelDataMap.get(excelName).put(modelName, oldExcel);
                } else {
                    Map<String, Excel> map = new HashMap<String, Excel>(16);
                    map.put(modelName, oldExcel);
                    Constants.oldExcelDataMap.put(excelName, map);
                }
            }


        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 读取配置
     *
     * @param key
     * @return
     */
    public static String get(String key) {
        return properties.getProperty(key);
    }

}
