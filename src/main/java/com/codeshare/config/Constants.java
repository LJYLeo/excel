package com.codeshare.config;

import com.codeshare.excel.Excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Constants {

    public static Map<String, List<String>> oldToNewMap = new HashMap<String, List<String>>();
    public static Map<String, Map<String, Excel>> oldExcelDataMap = new HashMap<String, Map<String, Excel>>();
    public static Map<String, Excel> newExcelDataMap = new HashMap<String, Excel>();

    public static String oldExcelRootPath;
    public static String modelExcelRootPath;
    public static String resultExcelRootPath;

    static {

        oldExcelRootPath = Config.get("villageOldExcelRootPath");
        modelExcelRootPath = Config.get("modelExcelRootPath");
        resultExcelRootPath = Config.get("resultExcelRootPath");

    }

}