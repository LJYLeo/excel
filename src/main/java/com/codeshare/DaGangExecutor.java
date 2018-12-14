package com.codeshare;

import com.codeshare.config.Config;
import com.codeshare.config.Constants;
import com.codeshare.excel.Excel;
import com.codeshare.utils.ExcelUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DaGangExecutor {

    public static String oldPath = "/Users/liujiayu/Desktop/老公专属/Excel新老表格转换/大港/农村集体资产清产核资--大港街道";
    public static String resultPath = "/Users/liujiayu/Desktop/老公专属/Excel新老表格转换/result_add";

    public static void main(String[] args) {

        Config.loadVillageConfig("dagang_add.json");
        File oldPathDirector = new File(oldPath);
        Map<String, Excel> oldCommonMap = Constants.oldExcelDataMap.get("all-cun");
        if (oldPathDirector.isDirectory()) {
            File[] villageDirector = oldPathDirector.listFiles();
            if (villageDirector != null && villageDirector.length > 0) {
                for (File village : villageDirector) {
                    if (village.isDirectory()) {
                        File[] groupDirector = village.listFiles();
                        if (groupDirector != null && groupDirector.length > 0) {
                            for (File group : groupDirector) {
                                if (group.isDirectory()) {
                                    File[] oldExcels = group.listFiles();
                                    for (File excel : oldExcels) {
                                        String wholePath = excel.getAbsolutePath();
                                        String excelName = excel.getName();
                                        Map<String, Excel> oldToNewMap = Constants.oldExcelDataMap.get(excelName);
                                        if (oldToNewMap == null) {
                                            oldToNewMap = oldCommonMap;
                                        }
                                        if (oldToNewMap != null) {
                                            HSSFWorkbook workbook = ExcelUtils.createWorkBook(wholePath);
                                            String targetFileRootPath = getNewDirector(wholePath);
                                            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                                                HSSFSheet sheet = workbook.getSheetAt(i);
                                                if (Constants.oldToNewMap.get(sheet.getSheetName()) != null) {
                                                    List<String> modelList = Constants.oldToNewMap.get(sheet.getSheetName());
                                                    for (String model : modelList) {
                                                        HSSFWorkbook modelWorkBook;
                                                        String targetFilePath = targetFileRootPath + "/" + model;
                                                        File targetFile = new File(targetFilePath);
                                                        if (targetFile.exists()) {
                                                            modelWorkBook = ExcelUtils.createWorkBook(targetFilePath);
                                                        } else {
                                                            modelWorkBook = ExcelUtils.createWorkBook(Constants.modelExcelRootPath + "/" + model);
                                                        }
                                                        if (modelWorkBook != null) {
                                                            HSSFSheet newSheet = modelWorkBook.getSheetAt(0);
                                                            Excel old = oldToNewMap.get(model);
                                                            if (old == null) {
                                                                old = oldCommonMap.get(model);
                                                            }
                                                            if (old != null && Constants.newExcelDataMap.get(model) != null) {
                                                                int end = ExcelUtils.fillValueToSheet(modelWorkBook, excelName, model, sheet, newSheet, old, Constants.newExcelDataMap.get(model));
                                                                newSheet.setForceFormulaRecalculation(true);
                                                                System.out.println(excelName + "\t" + model + "\t" + end);
                                                                Map<String, Integer> map = Constants.newExcelLastLocation.get(excelName);
                                                                if (map == null) {
                                                                    map = new HashMap<String, Integer>(1);
                                                                    Constants.newExcelLastLocation.put(excelName, map);
                                                                }
                                                                map.put(model, end);
                                                                ExcelUtils.createTargetExcel(modelWorkBook, targetFilePath);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
            }
        }

    }

    /**
     * @param wholePath
     * @return
     */
    public static String getNewDirector(String wholePath) {

        String path1 = StringUtils.substring(wholePath, oldPath.length());
        int lastIndex = StringUtils.lastIndexOf(path1, "/");

        String path = resultPath + StringUtils.substring(path1, 0, lastIndex);

        File file = new File(path);
        if (!file.exists()) {
            file.mkdirs();
        }

        return path;
    }

}
