package com.codeshare;

import com.codeshare.config.Config;
import com.codeshare.config.Constants;
import com.codeshare.utils.ExcelUtils;

public class Main {

    public static void main(String[] args) {

        Config.loadVillages();

        String[] fileNames = Config.get("jsonFile").split(",");
        if (fileNames.length != 0) {
            for (String fileName : fileNames) {
                System.out.println("正在执行：" + fileName + "...");
                Config.loadVillageConfig(fileName);
                ExcelUtils.process(Constants.oldExcelRootPath, 0);
                for (String villageName : Constants.villagesFromGroupDirector) {
                    ExcelUtils.process(Constants.oldExcelGroupRootPath + "/" + villageName + "村", 1);
                }
            }
        }

        System.out.println("正在补未生成的表...");
        ExcelUtils.supplyEmptyExcel();

    }

}
