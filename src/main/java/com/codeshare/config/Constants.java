package com.codeshare.config;

import com.codeshare.excel.Excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Constants {

    public static List<String> villagesFromGroupDirector = new ArrayList<String>();

    public static Map<String, List<String>> oldToNewMap = new HashMap<String, List<String>>();
    public static Map<String, Map<String, Excel>> oldExcelDataMap = new HashMap<String, Map<String, Excel>>();
    public static Map<String, Excel> newExcelDataMap = new HashMap<String, Excel>();
    public static Map<String, List<Map<String, Integer>>> newExcelDefaultMap = new HashMap<String, List<Map<String, Integer>>>();
    public static Map<String, String> newExcelDefaultValueMap = new HashMap<String, String>();
    public static Map<String, Integer> newExcelCheckNumMap = new HashMap<String, Integer>(2);

    public static Map<String, String> newExcelAddMap = new HashMap<String, String>(16);
    public static Map<String, String> newExcelCheckMap = new HashMap<String, String>(16);
    public static Map<String, Map<String, Integer>> newExcelLastLocation = new HashMap<String, Map<String, Integer>>(16);

    public static String oldExcelRootPath;
    public static String oldExcelGroupRootPath;
    public static String modelExcelRootPath;
    public static String resultExcelRootPath;

    static {

        oldExcelRootPath = Config.get("villageOldExcelRootPath");
        modelExcelRootPath = Config.get("modelExcelRootPath");
        resultExcelRootPath = Config.get("resultExcelRootPath");
        oldExcelGroupRootPath = Config.get("groupOldExcelRootPath");

    }

}
