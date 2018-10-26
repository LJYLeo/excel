package com.codeshare.utils;

import com.codeshare.config.Config;
import com.codeshare.config.Constants;
import com.codeshare.excel.Excel;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
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

    public static void process(String path, int type) {

        File oldPath = new File(path);
        if (oldPath.isDirectory()) {
            File[] oldExcels = oldPath.listFiles();
            if (oldExcels != null && oldExcels.length != 0) {
                Map<String, Excel> oldCommonMap = Constants.oldExcelDataMap.get("all-cun");
                for (File excel : oldExcels) {
                    String wholePath = excel.getAbsolutePath();
                    String excelName = excel.getName();
                    Map<String, Excel> oldToNewMap = Constants.oldExcelDataMap.get(excelName);
                    if (oldToNewMap == null) {
                        oldToNewMap = oldCommonMap;
                    }
                    if (oldToNewMap != null) {
                        HSSFWorkbook workbook = createWorkBook(wholePath);
                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            HSSFSheet sheet = workbook.getSheetAt(i);
                            if (Constants.oldToNewMap.get(sheet.getSheetName()) != null) {
                                List<String> modelList = Constants.oldToNewMap.get(sheet.getSheetName());
                                for (String model : modelList) {
                                    HSSFWorkbook modelWorkBook;
                                    String targetFilePath = getNewDirector(wholePath, excelName, type) + "/" + model;
                                    File targetFile = new File(targetFilePath);
                                    if (targetFile.exists()) {
                                        modelWorkBook = createWorkBook(targetFilePath);
                                    } else {
                                        modelWorkBook = createWorkBook(Constants.modelExcelRootPath + "/" + model);
                                    }
                                    if (modelWorkBook != null) {
                                        HSSFSheet newSheet = modelWorkBook.getSheetAt(0);
                                        Excel old = oldToNewMap.get(model);
                                        if (old == null) {
                                            old = oldCommonMap.get(model);
                                        }
                                        if (old != null && Constants.newExcelDataMap.get(model) != null) {
                                            System.out.println(excelName + "\t" + model);
                                            fillValueToSheet(model, sheet, newSheet, old, Constants.newExcelDataMap.get(model));
                                            createTargetExcel(modelWorkBook, targetFilePath);
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

    /**
     * @param wholePath
     * @param oldExcelName
     * @param type         0-村 1-组
     * @return
     */
    private static String getNewDirector(String wholePath, String oldExcelName, int type) {

        String path = Constants.resultExcelRootPath + "/其他";

        for (String villageName : Constants.villagesFromGroupDirector) {

            if (type == 0 && StringUtils.contains(oldExcelName, villageName)) {
                path = Constants.resultExcelRootPath + "/" + villageName + "村/本级";
                break;
            } else if (type == 1 && StringUtils.contains(wholePath, villageName)) {
                String groupName = StringUtils.substring(oldExcelName, 0, StringUtils.indexOf(oldExcelName, "."));
                path = Constants.resultExcelRootPath + "/" + villageName + "村/" + groupName;
                break;
            }

        }

        if (StringUtils.contains(path, "其他")) {
            System.out.println("其他：" + wholePath);
        }

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
    private static void fillValueToSheet(String model, HSSFSheet oldSheet, HSSFSheet newSheet, Excel oldData, Excel modelData) {

        int end = oldData.getEndRow() == -1 ? Integer.MAX_VALUE : oldData.getEndRow();
        a:
        for (int i = oldData.getStartRow(), j = modelData.getStartRow(); i <= end; i++, j++) {

            // 竖表
            if (modelData.getStartRow() == -1) {
                try {
                    // 老表也是竖表
                    if (oldData.getStartRow() == -1) {
                        for (int cell = 0; cell < oldData.getDoubleCell().size(); cell++) {
                            Object[] value = checkCellType(oldSheet.getRow(oldData.getDoubleCell().get(cell).get("row")).getCell(oldData.getDoubleCell().get(cell).get("col")));
                            if (Integer.parseInt(value[0].toString()) == 0 || Integer.parseInt(value[0].toString()) == 3) {
                                if (newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")) == null) {
                                    newSheet.getRow(j).createCell(modelData.getDoubleCell().get(cell).get("col"));
                                }
                                newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")).setCellValue(Double.parseDouble(value[1].toString()));
                            } else {
                                if (value[1] instanceof String) {
                                    if (newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")) == null) {
                                        newSheet.getRow(j).createCell(modelData.getDoubleCell().get(cell).get("col"));
                                    }
                                    newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")).setCellValue((String) value[1]);
                                }
                            }
                        }
                    } else {
                        int rowNumber = oldData.getStartRow();
                        for (int cell = 0; cell < oldData.getDoubleCell().size(); cell++) {
                            Object[] value = checkCellType(oldSheet.getRow(rowNumber).getCell(oldData.getDoubleCell().get(cell).get("col")));
                            if (Integer.parseInt(value[0].toString()) == 0 || Integer.parseInt(value[0].toString()) == 3) {
                                if (newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")) == null) {
                                    newSheet.getRow(j).createCell(modelData.getDoubleCell().get(cell).get("col"));
                                }
                                newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")).setCellValue(Double.parseDouble(value[1].toString()));
                            } else {
                                if (value[1] instanceof String) {
                                    if (newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")) == null) {
                                        newSheet.getRow(j).createCell(modelData.getDoubleCell().get(cell).get("col"));
                                    }
                                    newSheet.getRow(modelData.getDoubleCell().get(cell).get("row")).getCell(modelData.getDoubleCell().get(cell).get("col")).setCellValue((String) value[1]);
                                }
                            }
                        }
                    }
                } catch (Exception e) {
                    System.out.println("竖表值错误！");
                    e.printStackTrace();
                }
                return;

            }

            boolean isSetValue = true;
            if (StringUtils.isNotBlank(Constants.newExcelCheckMap.get(model))) {
                if (oldSheet.getRow(i).getCell(Integer.parseInt(Config.get("tagCheckCellNum"))) != null
                        && !StringUtils.equals(oldSheet.getRow(i).getCell(Integer.parseInt(Config.get("tagCheckCellNum"))).getStringCellValue(), Constants.newExcelCheckMap.get(model))) {
                    isSetValue = false;
                }
            }
            int emptyValueCount = 0;
            for (int k = 0; k < oldData.getCell().size(); k++) {
                try {
                    // -1列代表原表不存在，去新表直接新增
                    if (oldData.getCell().get(k) == -1) {
                        newSheet.getRow(j).getCell(modelData.getCell().get(k)).setCellValue(Constants.newExcelAddMap.get(model));
                        emptyValueCount++;
                        continue;
                    }
                    Object[] value = checkCellType(oldSheet.getRow(i).getCell(oldData.getCell().get(k)));
                    System.out.println(i + "\t" + oldData.getCell().get(k) + "\t" + value[1]);
                    if (value[1] == null || StringUtils.isBlank(value[1].toString()) || "0.0".equals(value[1].toString())) {
                        emptyValueCount++;
                    }
                    if (Integer.parseInt(value[0].toString()) == 0 || Integer.parseInt(value[0].toString()) == 3) {
                        if (newSheet.getRow(j).getCell(modelData.getCell().get(k)) == null) {
                            newSheet.getRow(j).createCell(modelData.getCell().get(k));
                        }
                        newSheet.getRow(j).getCell(modelData.getCell().get(k)).setCellValue(Double.parseDouble(value[1].toString()));
                    } else {
                        if (value[1] instanceof String) {
                            if ("农清明细03-应收款项清查登记表（系统下载）.xls".equals(model) && "内部往来".equals((String) value[1]) && oldSheet.getRow(i).getCell(0).getNumericCellValue() == 0.0) {
                                j--;
                                continue a;
                            }
                            if ("农清明细03-应收款项清查登记表（系统下载）.xls".equals(model) && "村民监会意见(签章):".equals((String) value[1])) {
                                j--;
                                break a;
                            }
                            if (isSetValue) {
                                if (newSheet.getRow(j).getCell(modelData.getCell().get(k)) == null) {
                                    newSheet.getRow(j).createCell(modelData.getCell().get(k));
                                }
                                newSheet.getRow(j).getCell(modelData.getCell().get(k)).setCellValue((String) value[1]);
                            }
                        }
                    }
                } catch (Exception e) {
                    System.out.println("发生异常！程序继续执行！行：" + j + "，列：" + k + "，表：" + model + "，异常：" + e.toString());
                }
            }
            if (oldData.getEndRow() == -1 && emptyValueCount == oldData.getCell().size()) {
//                newSheet.removeRow(newSheet.getRow(j));
                // 先塞的值，判断为空行在删掉
                if (oldData.getCell().contains(-1)) {
                    int index = oldData.getCell().indexOf(-1);
                    newSheet.getRow(j).getCell(modelData.getCell().get(index)).setCellValue((String) null);
                }
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
            System.out.println("错误：" + path);
            e.printStackTrace();
        } finally {
            close(fis);
        }

        return workBook;

    }

    public static void supplyEmptyExcel() {

        List<String> allModelList = new ArrayList<String>();
        for (File model : new File(Constants.modelExcelRootPath).listFiles()) {
            allModelList.add(model.getName());
        }

        for (File firstDirector : new File(Constants.resultExcelRootPath).listFiles()) {
            System.out.println(firstDirector);
            for (File secondDirector : firstDirector.listFiles()) {
                String path = secondDirector.getAbsolutePath();
                List<String> copyModelList = new ArrayList<String>();
                copyModelList.addAll(allModelList);
                for (File existExcel : secondDirector.listFiles()) {
                    if (copyModelList.contains(existExcel.getName())) {
                        copyModelList.remove(existExcel.getName());
                    }
                }
                for (String supplyName : copyModelList) {
                    System.out.println(path + "/" + supplyName);
                    createTargetExcel(createWorkBook(Constants.modelExcelRootPath + "/" + supplyName), path + "/" + supplyName);
                }
            }
        }


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

    /*public static String getMergedRegionValue(Sheet sheet ,int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();

        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){

                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell) ;
                }
            }
        }

        return null ;
    }*/


    public static void main(String[] args) {
        /*HSSFWorkbook workBook = ExcelUtils.getOldWorkBook();
        System.out.println(workBook.getSheet("表12").getRow(7).getCell(1).getStringCellValue());
        ExcelUtils.process();*/

        HSSFWorkbook workbook = createWorkBook("/Users/liujiayu/Desktop/姚桥镇/村表格/华山村本级.xls");
        System.out.println(workbook.getSheet("表12").getRow(38));
        HSSFCell cell = workbook.getSheet("表12").getRow(38).getCell(12);
//        Object[] value = checkCellType(cell);
        System.out.println(cell.getNumericCellValue());

    }

}
