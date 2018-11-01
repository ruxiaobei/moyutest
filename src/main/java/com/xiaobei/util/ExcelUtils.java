package com.xiaobei.util;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;

/**
 * excel辅助工具， 依赖于exceTitleConfig.xml文件
 *
 */
@SuppressWarnings("rawtypes")
public class ExcelUtils {

    /**
     * 黄色
     */
    public static final short YELLOW_COLOR = HSSFColor.YELLOW.index;
    /**
     * 蓝色
     */
    public static final short BLUE_COLOR = HSSFColor.BLUE.index;
    /**
     * 红色
     */
    public static final short RED_COLOR = HSSFColor.RED.index;
    /**
     * 粉色
     */
    public static final short PINK_COLOR = HSSFColor.PINK.index;
    /**
     * 绿色
     */
    public static final short GREEN_COLOR = HSSFColor.GREEN.index;

    private Workbook wbook = new HSSFWorkbook(); // 工作簿
    private Sheet sheet = wbook.createSheet();  // sheet0
    private String fileName = "default"; // 默认文件名
    private CellStyle defaultCs = createDefaultCs();
    private int titleRows = -1; // 起始行号
    private List<int[]> pointList = new ArrayList<int[]>(); // 待合并的单元格坐标列表
    private Map<String, StringBuilder> titleMap = new TreeMap<String, StringBuilder>(); // 标题列表
    private List<Integer[]> mergeCellPointList = new ArrayList<Integer[]>();

    public ExcelUtils() {
    }

    /**
     * 将当前excel写入本地指定位置
     */
    public void writeCurrExcelToLocal(String baselocalFilePath) throws Exception {
        FileOutputStream fileOut = null;
        try {
            baselocalFilePath = baselocalFilePath.lastIndexOf("/") >= 0 ? baselocalFilePath.substring(0, baselocalFilePath.lastIndexOf("/") - 1) : baselocalFilePath;
            fileOut = new FileOutputStream(this.fileName + ".xls");
            wbook.write(fileOut);
            System.out.print("OK");
        } catch (Exception e) {
            throw e;
        } finally {
        	if (fileOut != null) {
        		fileOut.close();
        	}
        }

    }


    /**
     * 将结果集填充到当前excel
     */
    public void populateCurrExcel(List<Object> resultList, ExcelUtilsRowMapper mapper) throws Exception {

        int len = resultList.size();
        for (int i = 0; i < len; i++) {
            Row row = createNewRow(sheet, this.titleRows + i, 300);
            Object[] values = mapper.rowMapping(resultList.get(i));
            int valuesLen = values.length;
            for (int j = 0; j < valuesLen; j++) {
                createCell(row, j, null, values[j]);
            }
        }
    }

    /**
     * 创建一份新的excel
     *
     * @param fileName 文件名， 非全限定名
     * @param text     配置文件中的text值
     * @throws Exception
     */
    public void createOneNewExcel(String fileName, String text) throws Exception {
        this.fileName = fileName;
        this.populateCurrExcelTitle(text);
    }

    /**
     * 填充excel标题栏
     *
     * @param whichTable 哪张表
     * @throws Exception
     */
    @SuppressWarnings("unchecked")
    private void populateCurrExcelTitle(String whichTable) throws Exception {

        SAXReader reader = new SAXReader();
        Document document = reader.read(ExcelUtils.class.getClassLoader().getResourceAsStream("exceTitleConfig.xml"));
        Node node = document.selectSingleNode("//moduleName[@text='" + whichTable + "']");

        Element nodeEl = (Element) node;
        Integer rowNum = Integer.parseInt(nodeEl.attributeValue("rowNum"));
        Integer colNum = Integer.parseInt(nodeEl.attributeValue("colNum"));

        this.titleRows = rowNum;

        Map<String, List<Element>> map = new TreeMap<String, List<Element>>();
        for (int i = 0; i < rowNum; i++) {
            String key = "row" + i;
            List tempList = document.selectNodes("/excelTitle/moduleName[@text='" + whichTable + "']/row[@index='" + i + "']/descendant::*");
            map.put(key, tempList);
        }

        for (Map.Entry<String, List<Element>> entry : map.entrySet()) {

            String key = entry.getKey();
            if (entry.getValue().size() == 0) {
                throw new RuntimeException("[" + key + "]没有设置相应的col数据");
            }

            int colNumTotalTemp = 0;
            @SuppressWarnings("unused")
			int countIndexTemp = 0;
            List<Element> everyRow_ColList = entry.getValue();
            for (Element el : everyRow_ColList) {
                String colspan = el.attributeValue("colspan"); // 跨列
                String rowspan = el.attributeValue("rowspan"); // 跨行
                String empty = el.getName(); // 空标记

                colspan = StringUtils.isNotBlank(colspan) ? colspan : "1";
                rowspan = StringUtils.isNotBlank(rowspan) ? rowspan : "1";
                int colspanInt = Integer.parseInt(colspan);
                int rowspanInt = Integer.parseInt(rowspan);
                colNumTotalTemp += colspanInt;

                this.concatTitleStrList(key, el, colspan, empty);

//				// 行合并: 起始行 结束行 起始列 结束列
                if (rowspanInt > 1) {
                    this.calRowSpanPointMapping(colNumTotalTemp, el, colspanInt, rowspanInt);
                }
//				
//				// 列合并: 起始行 结束行 起始列 结束列
                if (colspanInt > 1) {
                    this.calColSpanPointMapping(colNumTotalTemp, el, colspanInt);
                }

                countIndexTemp++;

            }

            if (!(colNumTotalTemp == colNum)) {
                throw new RuntimeException("[" + key + "]实际设置的总col数据(" + colNumTotalTemp + ")与modelName的colNum数(" + colNum + ")不配");
            }

        }

        //设置列宽
        this.setColumnAvgWidth(sheet, colNum, 3600);

        CellStyle selfCs = this.createNullCellStyle(wbook);
        this.setCellBackgroundColor(selfCs, ExcelUtils.YELLOW_COLOR);

        int count = 0;
        for (Map.Entry<String, StringBuilder> entry : titleMap.entrySet()) {
            StringBuilder sb = entry.getValue();
            String[] columns = sb.toString().split(",");
            this.setNewExcelStaticRowByDefaultCs(wbook, sheet, count, columns, 300);
            count++;
        }

        for (Integer[] integers : mergeCellPointList) {
            this.addPoint(new int[]{integers[0], integers[1], integers[2], integers[3]});
        }

        // 合并单元格
        this.mergeCellBatch(sheet);
    }

    /**
     * 组装标题字符串列表
     *
     * @param key     行号
     * @param el      当前element
     * @param colspan 跨列数
     * @param empty   empty标签
     * @return
     */
    private void concatTitleStrList(String key, Element el,
                                    String colspan, String empty) {

        StringBuilder sb = this.titleMap.get(key);
        if (sb == null) {
            sb = new StringBuilder();
        }

        if ("empty".equals(empty)) {
            sb.append(" ").append(",");
        } else {
            sb.append(el.getTextTrim()).append(",");
        }

        int colspanInt = Integer.parseInt(colspan);

        for (int i = 0; i < (colspanInt - 1); i++) {
            sb.append(" ").append(",");
        }

        this.titleMap.put(key, sb);
    }

    /**
     * 计算rowspan设置的坐标映射
     *
     * @param colNumTotalTemp 当前列数
     * @param el              当前node节点
     * @param colspanInt      当前node设置的colspan
     * @param rowspanInt      当前node设置的rowspan
     */
    private void calRowSpanPointMapping(int colNumTotalTemp,
                                        Element el, int colspanInt, int rowspanInt) {
        // 找到父节点
        Element parentEl = el.getParent();
        // 找到索引
        int rowIndex = Integer.parseInt(parentEl.attributeValue("index"));
        int startRowIndex = rowIndex;
        int endRowIndex = rowspanInt - 1 == rowIndex ? rowspanInt : rowspanInt - 1;
        int startColIndex = colNumTotalTemp - colspanInt;
        int endColIndex = colNumTotalTemp - colspanInt;
        this.mergeCellPointList.add(new Integer[]{startRowIndex, endRowIndex, startColIndex, endColIndex});
    }

    /**
     * 计算colspan设置的坐标映射
     *
     * @param colNumTotalTemp 当前列数
     * @param el              当前node节点
     * @param colspanInt      当前node设置的colspan
     */
    private void calColSpanPointMapping(int colNumTotalTemp,
                                        Element el, int colspanInt) {
        // 找到父节点
        Element parentEl = el.getParent();
        // 找到索引
        int rowIndex = Integer.parseInt(parentEl.attributeValue("index"));
        int startRowIndex = rowIndex;
        int endRowIndex = rowIndex;
        int startColIndex = colNumTotalTemp - colspanInt;
        int endColIndex = startColIndex + colspanInt - 1;
        this.mergeCellPointList.add(new Integer[]{startRowIndex, endRowIndex, startColIndex, endColIndex});
    }

    /**
     * 创建一个空样式的CellStyle
     *
     * @param wb
     */
    public CellStyle createNullCellStyle(Workbook wb) {
        CellStyle sefCs = wb.createCellStyle();
        return sefCs;
    }

    /**
     * 设置初始化列数和均宽
     *
     * @param sheet    当前表
     * @param colCount 总列数
     * @return void
     */
    public void setColumnAvgWidth(Sheet sheet, int colCount, int avgWidth) {
        int[] temp = new int[colCount];

        for (int i = 0; i < colCount; i++) {
            temp[i] = avgWidth;
        }
        this.setColumnWidth(sheet, colCount, temp);
    }

    /**
     * 设置初始化列数和自定义宽度
     *
     * @param sheet
     * @param colCount   总列数
     * @param everyWidth 每列的宽度，长度必须大于或等于colCount相等
     */
    public void setColumnWidth(Sheet sheet, int colCount, int[] everyWidth) {

        if (colCount > everyWidth.length) {
            throw new RuntimeException("colCount小于everyWidth实际长度");
        }

        for (int i = 0; i < colCount; i++) {
            sheet.setColumnWidth(i, everyWidth[i]);//设置列的宽度，参数1：列索引(0开始) 参数2：列宽
        }
    }

    /**
     * 设置单元格背景色
     *
     * @param destCs
     * @param colorIndex 颜色索引：ExcelUtil.XX_COLOR
     */
    public void setCellBackgroundColor(CellStyle destCs, short colorIndex) {
        destCs.setFillPattern(CellStyle.SOLID_FOREGROUND); // 填充单元格
        destCs.setFillForegroundColor(colorIndex); // 填色
    }

    /**
     * 设置静态行，默认单元格样式: 居中，宋体，加粗，无背景色
     *
     * @param wbook
     * @param sheet
     * @param rowRum       行号
     * @param staticColumn 静态单元格数据数组
     * @param rowHeight    行高
     */
    public void setNewExcelStaticRowByDefaultCs(Workbook wbook,
                                                Sheet sheet,
                                                int rowRum,
                                                String[] staticColumn,
                                                int rowHeight) {

        Row row = createNewRow(sheet, rowRum, rowHeight);
        int len = staticColumn.length;

        for (int i = 0; i < len; i++) {
            createCell(row, i, defaultCs, staticColumn[i]);
        }
    }

    /**
     * 创建指定位置的行
     *
     * @param sheet
     * @param rowLocation 行位置，即： 第几行，从0开始
     * @param height      行高
     * @return org.apache.poi.ss.usermodel.Row  新行
     */
    public Row createNewRow(Sheet sheet, int rowLocation, int height) {
        Row row = sheet.createRow(rowLocation);
        row.setHeight((short) height);
        return row;
    }

    /**
     * 添加合并单元格的坐标
     *
     * @return
     */
    public ExcelUtils addPoint(int[] point) {
        pointList.add(point);
        return this;
    }

    /**
     * 创建指定位置的单元格并设置值
     *
     * @param currRow      当前行
     * @param cellLocation 单元格位置
     * @param style        单元格样式
     * @param value        单元格的值
     * @return <code> org.apache.poi.ss.usermodel.Cell </code> 当前单元格
     */
    protected Cell createCell(Row currRow, int cellLocation, CellStyle style, Object value) {
        Cell cell = currRow.createCell(cellLocation);
        if (style != null) {
            cell.setCellStyle(style);
        }
        if (value == null) {
            cell.setCellValue("");
        } else if (value.getClass() == Integer.class) {
            cell.setCellValue((Integer) value);
        } else if (value.getClass() == Float.class) {
            cell.setCellValue((Float) value);
        } else if (value.getClass() == Double.class) {
            BigDecimal b = new BigDecimal((Double) value);
            double f1 = b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
            cell.setCellValue(f1);
        } else if (value.getClass() == Date.class || value.getClass() == java.sql.Date.class/*session.get方法返回的日期是java.sql.Date类型*/) {
            Date date = (Date) value;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            cell.setCellValue(sdf.format(date));
        } else if (value.getClass() == Boolean.class) {
            cell.setCellValue((Boolean) value);
        } else if (value.getClass() == Timestamp.class) {
            Date date = (Timestamp) value;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            cell.setCellValue(sdf.format(date));
        } else if (value.getClass() == String.class) {
            cell.setCellValue((String) value);
        }

        return cell;
    }

    /**
     * 批量合并单元格 <br>
     * 注：此方法内部的坐标来源于addPoint方法
     *
     * @param sheet
     */
    public void mergeCellBatch(Sheet sheet) {
        for (int[] point : pointList) {
            mergeCellSingle(sheet, point[0], point[1], point[2], point[3]);
        }
    }

    /**
     * 合并指定单元格
     *
     * @param sheet
     * @param fromRow    起始行
     * @param toRow      结束行
     * @param fromColumn 起始列
     * @param toColumn   结束列
     * @return
     */
    public int mergeCellSingle(Sheet sheet, int fromRow, int toRow, int fromColumn, int toColumn) {
        return sheet.addMergedRegion(new CellRangeAddress(fromRow, toRow, fromColumn, toColumn));
    }

    /**
     * 设置默认的单元格样式: 居中，宋体，加粗，无背景色
     */
    public CellStyle createDefaultCs() {
        CellStyle defaultCs = wbook.createCellStyle();

        this.setCellCenterLayout(defaultCs);
        this.setCellFont(wbook, defaultCs, "宋体", 10.5f, true);
        return defaultCs;
    }

    /**
     * 设置文字居中
     *
     * @param destCs 目标样式
     */
    public void setCellCenterLayout(CellStyle destCs) {
        destCs.setAlignment(CellStyle.ALIGN_CENTER);
        destCs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
    }

    /**
     * 设置字体
     *
     * @param wb
     * @param destCs   目标样式
     * @param fontName 字体名称
     * @param size     字体大小
     * @param isBold   是否加粗
     */
    public void setCellFont(Workbook wb, CellStyle destCs, String fontName, float size, boolean isBold) {
        Font headerFont = wb.createFont(); // 字体
        headerFont.setFontName(fontName);
        headerFont.setFontHeightInPoints((short) size);
        destCs.setFont(headerFont);

        if (isBold) {
            headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
        }
    }
}
