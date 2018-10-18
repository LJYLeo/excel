package com.codeshare.utils;

import com.codeshare.config.Constants;
import com.codeshare.excel.Excel;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * Excel工具类
 */
public class ExcelUtils {

    //    public static String oldExcelRootPath = "";
//    public static String oldExcelPath = "C:/Users/18046184/Desktop/村组表格/村表格/北角村表(1).xls";
//    public static String modelExcelRootPath = "C:/Users/18046184/Desktop/系统下载模板表格";
//    public static String resultExcelRootPath = "C:/Users/18046184/Desktop/result/北角村/本级";
//    public static String resultExcelPath = "C:/Users/18046184/Desktop/result/北角村/本级/农清明细11-2应付款项清查登记表（系统下载）.xls";

    static {

    }

    public static void process() {

        File oldPath = new File(Constants.oldExcelRootPath);
        if (oldPath.isDirectory()) {
            File[] oldExcels = oldPath.listFiles();
            if (oldExcels != null && oldExcels.length != 0) {
                Map<String, Excel> oldCommonMap = Constants.oldExcelDataMap.get("all-cun");
                for (File excel : oldExcels) {
                    String wholePath = excel.getAbsolutePath();
                    String excelName = excel.getName();
                    if (Constants.oldExcelDataMap.get(excelName) != null) {
                        HSSFWorkbook workbook = createWorkBook(wholePath);
                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            HSSFSheet sheet = workbook.getSheetAt(i);
                            if (Constants.oldToNewMap.get(sheet.getSheetName()) != null) {
                                List<String> modelList = Constants.oldToNewMap.get(sheet.getSheetName());
                                for (String model : modelList) {
                                    HSSFWorkbook modelWorkBook = createWorkBook(Constants.modelExcelRootPath + "/" + model);
                                    if (modelWorkBook != null) {
                                        HSSFSheet newSheet = modelWorkBook.getSheetAt(0);
                                        Excel old = Constants.oldExcelDataMap.get(excelName).get(model);
                                        if (old == null) {
                                            old = oldCommonMap.get(model);
                                        }
                                        if (old != null && Constants.newExcelDataMap.get(model) != null) {
                                            System.out.println(excelName);
                                            fillValueToSheet(sheet, newSheet, old, Constants.newExcelDataMap.get(model));
                                            createTargetExcel(modelWorkBook, getNewDirector(excelName) + "/" + model);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /*copyDataFromOldExcel(getOldWorkBook(), Arrays.asList(new String[]{"表12"}), Constants.oldExcelDataMap.get("农清明细11-2"
                + "应付款项清查登记表（系统下载）.xls"), Constants.newExcelDataMap.get("农清明细11-2应付款项清查登记表（系统下载）.xls"));*/

    }

    private static String getNewDirector(String oldExcelName) {
        int index;
        index = StringUtils.indexOf(oldExcelName, "村");
        if (index == -1) {
            index = StringUtils.indexOf(oldExcelName, ".");
        }

        String path = Constants.resultExcelRootPath + "/" + StringUtils.substring(oldExcelName, 0, index + 1) + "/本级";
        File file = new File(path);
        if (!file.exists()) {
            file.mkdirs();
        }

        return path;
    }

    /**
     * 获得模板Excel表名称
     *
     * @param sheetName
     * @return
     */
    /*private static List<String> getModelWorkBookName(String sheetName) {
        if (StringUtils.isNoneBlank(sheetName) && CollectionUtils.isNotEmpty(Constants.oldToNewMap.get(sheetName))) {
            return Constants.oldToNewMap.get(sheetName);
        }
        return new ArrayList<String>();
    }*/

    /**
     * 获得模板Excel表对象
     *
     * @param sheetName
     * @return
     */
    /*private static List<HSSFWorkbook> getModelWorkBook(String sheetName) {

        List<HSSFWorkbook> list = new ArrayList<HSSFWorkbook>();

        if (StringUtils.isNoneBlank(sheetName) && CollectionUtils.isNotEmpty(Constants.oldToNewMap.get(sheetName))) {

            for (String each : Constants.oldToNewMap.get(sheetName)) {
                // 生成目标路径
                String path = modelExcelRootPath + "/" + each;
                list.add(createWorkBook(path));
            }

        }

        return list;

    }*/

    /**
     * 生成excel
     *
     * @param workBook
     * @param targetPath
     */
    private static void createTargetExcel(HSSFWorkbook workBook, String targetPath) {

        FileOutputStream fos = null;

        try {
            fos = new FileOutputStream(targetPath);
            workBook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            close(fos);
        }

    }

    /**
     * 将数据从老的表中拷贝到新的模板中
     *
     * @param oldEcel
     * @param sheetNameList
     * @param oldData
     * @param modelData
     */
    /*private static void copyDataFromOldExcel(HSSFWorkbook oldEcel, List<String> sheetNameList, Excel oldData, Excel modelData) {

        if (CollectionUtils.isNotEmpty(sheetNameList)) {
            for (String sheetName : sheetNameList) {
                HSSFSheet oldSheet = oldEcel.getSheet(sheetName);
                if (oldSheet != null) {
                    List<HSSFWorkbook> modelWorkBooks = getModelWorkBook(sheetName);
                    List<String> modelWorkBookNames = getModelWorkBookName(sheetName);
                    for (int i = 0; i < modelWorkBooks.size(); i++) {
                        HSSFWorkbook modelWorkBook = modelWorkBooks.get(i);
                        if (modelWorkBook != null) {
                            HSSFSheet newSheet = modelWorkBook.getSheetAt(0);
                            fillValueToSheet(oldSheet, newSheet, oldData, modelData);
                            createTargetExcel(modelWorkBook, resultExcelRootPath + "/" + modelWorkBookNames.get(i));
                        }
                    }
                }
            }
        }

    }*/

    /**
     * 往sheet中填入数据
     *
     * @param oldSheet
     * @param newSheet
     * @param oldData
     * @param modelData
     */
    private static void fillValueToSheet(HSSFSheet oldSheet, HSSFSheet newSheet, Excel oldData, Excel modelData) {

        int end = oldData.getEndRow() == -1 ? Integer.MAX_VALUE : oldData.getEndRow();
        for (int i = oldData.getStartRow(), j = modelData.getStartRow(); i <= end; i++, j++) {
            int emptyValueCount = 0;
            for (int k = 0; k < oldData.getCell().size(); k++) {
                try {
                    Object[] value = checkCellType(oldSheet.getRow(i).getCell(oldData.getCell().get(k)));
                    System.out.println(i + "\t" + oldData.getCell().get(k) + "\t" + value[1]);
                    if (value[1] == null || StringUtils.isBlank(value[1].toString()) || "0.0".equals(value[1].toString())) {
                        emptyValueCount++;
                    }
                    if (Integer.parseInt(value[0].toString()) == 0 || Integer.parseInt(value[0].toString()) == 3) {
                        newSheet.getRow(j).getCell(modelData.getCell().get(k)).setCellValue(Double.parseDouble(value[1].toString()));
                    } else {
                        if (value[1] instanceof String) {
                            newSheet.getRow(j).getCell(modelData.getCell().get(k)).setCellValue((String) value[1]);
                        }
                    }
                } catch (Exception e) {
                    System.out.println("发生异常！程序继续执行！行：" + j + "，列：" + k + "，异常：" + e.toString());
                }
            }
            if (oldData.getEndRow() == -1 && emptyValueCount == oldData.getCell().size()) {
                newSheet.removeRow(newSheet.getRow(j));
                break;
            }
        }

    }

    private static Object[] checkCellType(HSSFCell cell) {
        switch (cell.getCellType()) {
            // 数字
            case Cell.CELL_TYPE_NUMERIC:
                return new Object[]{0, cell.getNumericCellValue()};
            // 字符串
            case Cell.CELL_TYPE_STRING:
                return new Object[]{1, cell.getStringCellValue()};
            // Boolean
            case Cell.CELL_TYPE_BOOLEAN:
                return new Object[]{2, cell.getBooleanCellValue()};
            // 公式
            case Cell.CELL_TYPE_FORMULA:
                return new Object[]{3, cell.getNumericCellValue()};
            // 空值
            case Cell.CELL_TYPE_BLANK:
                String value = null;
                return new Object[]{5, value};
            // 故障
            case Cell.CELL_TYPE_ERROR:
                return new Object[]{6, ""};
            default:
                return new Object[]{4, ""};
        }
    }

    private static HSSFWorkbook createWorkBook(String path) {

        HSSFWorkbook workBook = null;
        FileInputStream fis = null;

        try {
            fis = new FileInputStream(path);
            workBook = new HSSFWorkbook(fis);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            close(fis);
        }

        return workBook;

    }

    private static void close(FileOutputStream fos) {
        try {
            if (fos != null) {
                fos.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void close(FileInputStream fis) {
        try {
            if (fis != null) {
                fis.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        /*HSSFWorkbook workBook = ExcelUtils.getOldWorkBook();
        System.out.println(workBook.getSheet("表12").getRow(7).getCell(1).getStringCellValue());*/
        ExcelUtils.process();
    }

}