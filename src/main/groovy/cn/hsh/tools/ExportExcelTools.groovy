package cn.hsh.tools

import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddressList

/**
 * Created by XinXi-001 on 2018/3/7.
 */
class ExportExcelTools {
    /**
     * 创建隐藏数据域用于保存下拉列表中的内容
     * @param workbook :         excel文件对象
     * @param hideSheetName :    隐藏数据表格名称
     * @param titleList ：       一级列表List
     * @param contentList :      二级列表list
     */
    public
    static void createHideSheet(HSSFWorkbook workbook, List firstLevelList, List secondLevelList, List thirdLevelList) {
        HSSFSheet dictionary = workbook.createSheet("dataSource");//隐藏列表信息
        //在隐藏页设置选择信息
        HSSFRow provinceRow = dictionary.createRow(0);
        ExportExcelTools.createRow(provinceRow, firstLevelList)
        for (int i = 0; i < secondLevelList.size(); i++) {
            HSSFRow contentRow = dictionary.createRow(i + 1);
            ExportExcelTools.createRow(contentRow, secondLevelList.get(i))
        }
        if (thirdLevelList.size() > 0) {
            for (int i = 0; i < thirdLevelList.size(); i++) {
                HSSFRow contentRow = dictionary.createRow(secondLevelList.size() + i + 1);
                ExportExcelTools.createRow(contentRow, thirdLevelList.get(i))
            }
        }
        //设置隐藏页标志
        workbook.setSheetHidden(workbook.getSheetIndex("dataSource"), true);
    }
    /**
     * 名称管理(在Excel里面公式设置里面)
     * @param workbook
     */
    public
    static void createExcelNameList(HSSFWorkbook workbook, List titleList, List contentList, List thirdLevelList) {
        Name name;
        name = workbook.createName();
        // 设置第一级列名
        name.setNameName("formula");

        name.setRefersToFormula("dataSource" + "!\$A\$1:\$" + this.getCellColumnFlag(titleList.size()) + "\$1");
        println("dataSource" + "!\$A\$1:\$" + this.getCellColumnFlag(titleList.size()) + "\$1")
        // 设置第一级下面的第二级
        /**
         * 需要保证第一级各个值的顺序，和第二级每个list的第一个元素的值的顺序相同
         * eg:       titleList  a,b,c,d
         *      contentList(0)  a,a1,a2,a3,a4
         *      contentList(1)  b,b1,b2,b3,b4
         *      contentList(2)  c,c1,c2,c3,c4
         *              ……    ……
         */
        for (int i = 0; i < titleList.size(); i++) {
            if (titleList.get(i).toString().equals(contentList.get(i).get(0))) {
                name = workbook.createName();
                name.setNameName(titleList.get(i).toString());
                name.setRefersToFormula("dataSource" + "!\$B\$" + (i + 2) + ":\$" + this.getCellColumnFlag(contentList.get(i).size()) + "\$" + (i + 2));
                //   println("dataSource" + "!\$B\$" + (i + 2) + ":\$" + this.getCellColumnFlag(contentList.get(i).size()) + "\$" + (i + 2))
            }
        }
        if (thirdLevelList.size() > 0) {
            // 设置第二级下面的第三级
            for (int i = 0; i < thirdLevelList.size(); i++) {
                println(contentList.size());
                name = workbook.createName();
                name.setNameName(thirdLevelList.get(i).get(0))
                name.setRefersToFormula("dataSource" + "!\$B\$" + (2 + contentList.size() + i) + ":\$" + this.getCellColumnFlag(thirdLevelList.get(i).size()) + "\$" + (2 + contentList.size() + i))
                println("dataSource" + "!\$B\$" + (2 + contentList.size() + i) + ":\$" + this.getCellColumnFlag(thirdLevelList.get(i).size()) + "\$" + (2 + contentList.size() + i))
            }
        }
    }
    /**
     * Excel公式名称管理（无级联关系）
     */
    public static boolean createExcelFormula(HSSFWorkbook workbook, List firstList, List secondList, List thirdList) {
        Name name, name1, name2;
        name = workbook.createName();
        // 设置第一级列名
        name.setNameName("formula");
        name.setRefersToFormula("dataSource" + "!\$A\$1:\$" + this.getCellColumnFlag(firstList.size()) + "\$1");
        if (secondList != null && secondList.size() > 0) {
            name1 = workbook.createName()
            name1.setNameName("formula1")
            name1.setRefersToFormula("dataSource" + "!\$A\$2:\$" + this.getCellColumnFlag(secondList.get(0).size()) + "\$2");
        }
        if (thirdList != null && thirdList.size() > 0) {
            name2 = workbook.createName()
            name2.setNameName("formula2")
            name2.setRefersToFormula("dataSource" + "!\$A\$3:\$" + this.getCellColumnFlag(thirdList.get(0).size()) + "\$3");
        }
    }
    /**
     * 根据数据值确定单元格位置（比如：28-AB）
     * @param num
     * @return
     */
    public static String getCellColumnFlag(int num) {
        String columFiled = "";
        int chuNum = 0;
        int yuNum = 0;
        if (num >= 1 && num <= 26) {
            columFiled = this.doHandle(num);
        } else {
            chuNum = num / 26;
            yuNum = num % 26;

            columFiled += this.doHandle(chuNum);
            columFiled += this.doHandle(yuNum);
        }
        return columFiled;
    }

    public static String doHandle(final int num) {
        String[] charArr = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
                            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
                            "W", "X", "Y", "Z"];
        return charArr[num - 1].toString();
    }
    /**
     * 创建一行数据
     * @param currentRow
     * @param textList
     */
    public static void createRow(HSSFRow currentRow, List textList) {
        if (textList != null && textList.size() > 0) {
            int i = 0;
            for (String cellValue : textList) {
                HSSFCell cell = currentRow.createCell(i++);
                cell.setCellValue(cellValue);
            }
        }
    }
    /**
     * 使用已定义的数据源方式设置一个数据验证
     *
     * @param formulaString
     * @param naturalRowIndex
     * @param naturalColumnIndex
     * @return
     */
    public DataValidation getDataValidationByFormula(String formulaString, int naturalRowIndex, int naturalColumnIndex, int lastColumnIndex) {
        // 加载下拉列表内容
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        int firstRow = 1;
        int lastRow = naturalRowIndex;
        int firstCol = naturalColumnIndex - 1;
        int lastCol = lastColumnIndex - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        DataValidation data_validation_list = new HSSFDataValidation(regions, constraint);
        return data_validation_list;
    }
    /**
     * 数据验证:时间
     * @param naturalRowIndex
     * @param naturalColumnIndex
     * @return
     */
    public DataValidation getDataValidationByDate(int naturalRowIndex, int naturalColumnIndex) {
        DVConstraint constraint = DVConstraint.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "2001-01-01", "2999-01-01", "yyyy-MM-dd")
        int firstRow = 1;
        int lastRow = naturalRowIndex;
        int firstCol = naturalColumnIndex - 1;
        int lastCol = naturalColumnIndex - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        DataValidation data_validation_list = new HSSFDataValidation(regions, constraint);
        return data_validation_list;
    }
    /**
     * 数据验证：数字
     * @param naturalRowIndex
     * @param naturalColumnIndex
     * @return
     */
    public DataValidation getDataValidationByInteger(int naturalRowIndex, int naturalColumnIndex) {
        DVConstraint constraint = DVConstraint.createNumericConstraint(DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.GREATER_THAN, "0", null);
        int firstRow = 1;
        int lastRow = naturalRowIndex;
        int firstCol = naturalColumnIndex - 1;
        int lastCol = naturalColumnIndex - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        DataValidation data_validation_list = new HSSFDataValidation(regions, constraint);
        return data_validation_list;
    }
    /**
     * 只含一个数据验证
     */
    def exportExcel(HSSFWorkbook workbook, HSSFSheet hssfSheet, List titles, List contentList) {
        // 创建
        hssfSheet.setDefaultColumnWidth(20);
        hssfSheet.setDefaultRowHeightInPoints(20);
        // 创建单元格样式
        HSSFCellStyle titleCellStyle = workbook.createCellStyle();
        // 指定单元格居中对齐，边框为细

        // 设置填充色
        /*titleCellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        titleCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);*/
        // 指定当单元格内容显示不下时自动换行
        //   titleCellStyle.setWrapText(true);
        // 设置单元格字体
        HSSFFont titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 12);
        //   titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        titleCellStyle.setFont(titleFont);
        HSSFRow headerRow = hssfSheet.createRow(0);
        HSSFCell headerCell = null;

        /**————————————设置表头————————————**/
        for (int c = 0; c < titles.size(); c++) {
            headerCell = headerRow.createCell(c);
            headerCell.setCellStyle(titleCellStyle);
            headerCell.setCellValue(titles.get(c));
            hssfSheet.setColumnWidth(c, (30 * 160));
        }
        // ------------------------------------------------------------------
        // 创建单元格样式
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        // 指定单元格居中对齐，边框为细
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setWrapText(false)
        // 设置单元格字体
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 33);
        titleFont.setFontHeightInPoints((short) 11);
        cellStyle.setFont(font);

        /**--------------------------------填充数据-----------------------------------**/
        def list = contentList
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row = hssfSheet.createRow(i + 1);
            row.setRowStyle(cellStyle);
            for (int j = 0; j < titles.size(); j++) {
                row.createCell(j).setCellValue(list.get(i)[j])
            }
        }
        /**——————————————设置数据验证——————————————**/
        /*      // 得到验证对象
              def firstTitleColIndex = 2
              def naturalRowIndex = list.size();//终止行参数
              if (naturalRowIndex == 0) {
                  naturalRowIndex = 1
              }
              *//**——————下拉列表——————**//*
        DataValidation data_validation_list1 = this.getDataValidationByFormula("formula", naturalRowIndex, firstTitleColIndex, titles.size());
        hssfSheet.addValidationData(data_validation_list1);*/

        /**——————设置第一列禁用————————**/
        CellStyle lockCellStyle = workbook.createCellStyle();
        lockCellStyle.setLocked(false)
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row = hssfSheet.getRow(i + 1);
            for (int j = 1; j < titles.size(); j++) {
                row.getCell(j).setCellStyle(lockCellStyle)
            }
        }
      //  hssfSheet.protectSheet(GeneConstants.SHEET_PWD) //禁用密码
    }

     static exportGeoExcel(List titleList, List contentList, List baseList, String excelSheetName, String filePath) {
        ExportExcelTools dd = new ExportExcelTools()
        HSSFWorkbook wb = new HSSFWorkbook()
        HSSFSheet excelSheet = wb.createSheet(excelSheetName);
        dd.exportExcel(wb, excelSheet, titleList, contentList)
        FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
        return true;
    }

     static exportGeoExcel(List contentList,String filePath){
        ExportExcelTools dd = new ExportExcelTools()
        HSSFWorkbook wb = new HSSFWorkbook()
        contentList.eachWithIndex { it,i->
            HSSFSheet excelSheet = wb.createSheet(it.sheetName);
            dd.exportExcel(wb,excelSheet,it.title,it.content)
        }
        FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
        return true
    }
}
