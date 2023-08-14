package com.poi.office.converter;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;


public class ExcelConverter {

    public String mFilePath;
    public String mCachePath;
    public String mUrlPath;
    public String mHtmlData;

    public ExcelConverter(String filePath, String cachePath) {
        this.mFilePath = filePath;
        this.mCachePath = cachePath;
    }

    public void readExcelToHtml() {
        try {
            String fileExt = mFilePath.substring(mFilePath.lastIndexOf("."));
            if (fileExt.equalsIgnoreCase(".xls")) {
                HSSFWorkbook wb = (HSSFWorkbook) readExcel(mFilePath);
                excel03ToHtml(wb);
            }
            if (fileExt.equalsIgnoreCase(".xlsx")) {
                Workbook wb = readExcel(mFilePath);
                excel07ToHtml(wb);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private Workbook readExcel(String fileName) {
        Workbook wb = null;
        if (fileName == null) {
            return null;
        }
        String extString = fileName.substring(fileName.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(fileName);
            if (".xls".equalsIgnoreCase(extString)) {
                return wb = new HSSFWorkbook(is);
            } else if (".xlsx".equalsIgnoreCase(extString)) {
                return wb = new XSSFWorkbook(is);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    /**
     * excel03转html
     * filename:要读取的文件所在文件夹
     * filepath:文件名
     * htmlname:生成html名称
     * path:html存放路径
     */
    private void excel03ToHtml(HSSFWorkbook excelBook) throws ParserConfigurationException, TransformerException, IOException {
        ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        excelToHtmlConverter.processWorkbook(excelBook);// excel转html
        Document htmlDocument = excelToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();// 字节数组输出流
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);
        /** 将document中的内容写入文件中，创建html页面 */
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);

        String fileName = mFilePath.substring(mFilePath.lastIndexOf("/") + 1);
        String fileExt = fileName.substring(fileName.lastIndexOf(".") + 1);
        String htmlName = fileName.replace(fileExt, "html");
        File htmlFile = new File(mCachePath, htmlName);
        FileOutputStream fileOutputStream = new FileOutputStream(htmlFile);
        fileOutputStream.write(outStream.toByteArray());
        mUrlPath = "file:///" + htmlFile.getAbsolutePath();
        mHtmlData = outStream.toString("UTF-8");

        outStream.close();

        // return outStream.toString("UTF-8");
    }

    private Map<String, Object> map[];


    /**
     * excel07转html
     * filename:要读取的文件所在文件夹
     * filepath:文件名
     * htmlname:生成html名称
     * path:html存放路径
     */
    private void excel07ToHtml(Workbook workbook) {
        ByteArrayOutputStream baos = null;
        StringBuilder html = new StringBuilder();
        try {
            for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
                Sheet sheet = workbook.getSheetAt(numSheet);
                if (sheet == null) {
                    continue;
                }
                html.append("=======================").append(sheet.getSheetName()).append("=========================<br><br>");

                int firstRowIndex = sheet.getFirstRowNum();
                int lastRowIndex = sheet.getLastRowNum();
                //                html.append("<table style='border-collapse:collapse;width:100%;' align='left'>");
                html.append("<table style='" +
                        "        font-size:11px;" +
                        "        color:#333333;" +
                        "        border-width: 0.1px;" +
                        "        border-color: #666666;" +
                        "        border-collapse: collapse;width:100%;' align='left'>");

                map = getRowSpanColSpanMap(sheet);
                // 行
                for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
                    Row currentRow = sheet.getRow(rowIndex);
                    if (null == currentRow) {
                        html.append("<tr><td >  </td></tr>");
                        continue;
                    } else if (currentRow.getZeroHeight()) {
                        continue;
                    }
                    html.append("<tr>");
                    int firstColumnIndex = currentRow.getFirstCellNum();
                    int lastColumnIndex = currentRow.getLastCellNum();
                    // 列
                    for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
                        Cell currentCell = currentRow.getCell(columnIndex);
                        if (currentCell == null) {
                            continue;
                        }
                        String currentCellValue = getCellValue(currentCell);
                        if (map[0].containsKey(rowIndex + "," + columnIndex)) {
                            String pointString = (String) map[0].get(rowIndex + "," + columnIndex);
                            int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                            int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                            int rowSpan = bottomeRow - rowIndex + 1;
                            int colSpan = bottomeCol - columnIndex + 1;
                            if (map[2].containsKey(rowIndex + "," + columnIndex)) {
                                rowSpan = rowSpan - (Integer) map[2].get(rowIndex + "," + columnIndex);
                            }
                            html.append("<td style='border-width: 0.1px;" +
                                            "        border-style: solid;" +
                                            "        border-color: #666666;" +
                                            "        background-color: #ffffff;'")
                                    .append("rowspan= '")
                                    .append(rowSpan)
                                    .append("' colspan= '")
                                    .append(colSpan)
                                    .append("' ");
                            if (map.length > 3 && map[3].containsKey(rowIndex + "," + columnIndex)) {
                                // 此类数据首行被隐藏，value为空，需使用其他方式获取值
                                currentCellValue = getMergedRegionValue(sheet, rowIndex, columnIndex);
                            }
                        } else if (map[1].containsKey(rowIndex + "," + columnIndex)) {
                            map[1].remove(rowIndex + "," + columnIndex);
                            continue;
                        } else {
                            html.append("<td style='border-width: 0.1px;" +
                                    "        border-style: solid;" +
                                    "        border-color: #666666;" +
                                    "        background-color: #ffffff;' ");
                        }
                        CellStyle cellStyle = currentCell.getCellStyle();
                        if (cellStyle != null) {
                            html.append("align='").append(getHAlignByExcel(cellStyle.getAlignmentEnum())).append("' ");// 单元格内容的水平对齐方式
                            html.append("valign='").append(getVAlignByExcel(cellStyle.getVerticalAlignmentEnum())).append("' ");// 单元格中内容的垂直排列方式
                        }
                        html.append(">");
                        if (currentCellValue != null && !"".equals(currentCellValue)) {
                            html.append(currentCellValue.replace(String.valueOf((char) 160), " "));
                        }
                        html.append("</td>");
                    }
                    html.append("</tr>");
                }
                html.append("</table>");

                baos = new ByteArrayOutputStream();
                DOMSource domSource = new DOMSource();
                StreamResult streamResult = new StreamResult(baos);
                TransformerFactory tf = TransformerFactory.newInstance();
                Transformer serializer = tf.newTransformer();
                serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
                serializer.setOutputProperty(OutputKeys.INDENT, "yes");
                serializer.setOutputProperty(OutputKeys.METHOD, "html");
                serializer.transform(domSource, streamResult);

                // 写入文件
                String fileName = mFilePath.substring(mFilePath.lastIndexOf("/") + 1);
                String fileExt = fileName.substring(fileName.lastIndexOf(".") + 1);
                String htmlName = fileName.replace(fileExt, "html");
                File htmlFile = new File(mCachePath, htmlName);
                FileUtils.writeStringToFile(htmlFile, html.toString(), "UTF-8");
                mUrlPath = "file:///" + htmlFile.getAbsolutePath();

                mHtmlData = new String(html.toString().getBytes(), StandardCharsets.UTF_8);

                baos.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (baos != null) {
                    baos.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        // return new String(html.toString().getBytes(), StandardCharsets.UTF_8);
    }

    /**
     * 分析excel表格，记录合并单元格相关的参数，用于之后html页面元素的合并操作
     *
     * @param sheet
     * @return
     */
    private static Map<String, Object>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();    // 保存合并单元格的对应起始和截止单元格
        Map<String, String> map1 = new HashMap<String, String>();    // 保存被合并的那些单元格
        Map<String, Integer> map2 = new HashMap<String, Integer>();    // 记录被隐藏的单元格个数
        Map<String, String> map3 = new HashMap<String, String>();    // 记录合并了单元格，但是合并的首行被隐藏的情况
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        Row row = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            /**
             * 此类数据为合并了单元格的数据
             * 1.处理隐藏（只处理行隐藏，列隐藏poi已经处理）
             */
            if (topRow != bottomRow) {
                int zeroRoleNum = 0;
                int tempRow = topRow;
                for (int j = topRow; j <= bottomRow; j++) {
                    row = sheet.getRow(j);
                    if (row.getZeroHeight() || row.getHeight() == 0) {
                        if (j == tempRow) {
                            // 首行就进行隐藏，将rowTop向后移
                            tempRow++;
                            continue;// 由于top下移，后面计算rowSpan时会扣除移走的列，所以不必增加zeroRoleNum;
                        }
                        zeroRoleNum++;
                    }
                }
                if (tempRow != topRow) {
                    map3.put(tempRow + "," + topCol, topRow + "," + topCol);
                    topRow = tempRow;
                }
                if (zeroRoleNum != 0) map2.put(topRow + "," + topCol, zeroRoleNum);
            }
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, topRow + "," + topCol);
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = {map0, map1, map2, map3};
        System.err.println(map0);
        return map;
    }

    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);

                    return getCellValue(fCell);
                }
            }
        }
        return null;
    }

    /**
     * 读取单元格
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue();
    }


    private static String getVAlignByExcel(VerticalAlignment align) {
        String result = "middle";
        if (align == VerticalAlignment.BOTTOM) {
            result = "bottom";
        }
        if (align == VerticalAlignment.CENTER) {
            result = "center";
        }
        if (align == VerticalAlignment.JUSTIFY) {
            result = "justify";
        }
        if (align == VerticalAlignment.TOP) {
            result = "top";
        }
        return result;
    }

    protected static String getHAlignByExcel(HorizontalAlignment align) {
        String result = "left";
        if (align == HorizontalAlignment.LEFT) {
            result = "left";
        }
        if (align == HorizontalAlignment.RIGHT) {
            result = "right";
        }
        if (align == HorizontalAlignment.JUSTIFY) {
            result = "justify";
        }
        if (align == HorizontalAlignment.CENTER) {
            result = "center";
        }
        return result;
    }


}
