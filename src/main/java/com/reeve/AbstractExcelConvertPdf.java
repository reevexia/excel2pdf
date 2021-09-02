package com.reeve;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import com.reeve.dict.LetterValue;
import com.reeve.pdf.PdfFont;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

/**
 * 实现excel转pdf的接口实现抽象类
 *
 * @Author Reeve
 * @Date 2021/7/29 10:22
 */
public class AbstractExcelConvertPdf implements ExcelConvertPdf {

    /**
     * 转换excel为pdf
     *
     * @param excelPath
     * @param pdfPath
     */
    public void convert(String excelPath, String pdfPath) {
        try {
            HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(excelPath));
            Rectangle pageSize = new Rectangle(842.0F, 595.0F);
            Document document = new Document(pageSize);
            PdfWriter.getInstance(document, new FileOutputStream(pdfPath));
            convertPdf(workBook, document);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void convert(String excelPath, OutputStream outputStream) {
        try {
            HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(excelPath));
            Rectangle pageSize = new Rectangle(842.0F, 595.0F);
            Document document = new Document(pageSize);
            PdfWriter.getInstance(document, outputStream);
            convertPdf(workBook, document);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void convert(InputStream inputStream, OutputStream outputStream) {
        try {
            HSSFWorkbook workBook = new HSSFWorkbook(inputStream);
            Rectangle pageSize = new Rectangle(842.0F, 595.0F);
            Document document = new Document(pageSize);
            PdfWriter.getInstance(document, outputStream);
            convertPdf(workBook, document);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void convert(HSSFWorkbook workBook, OutputStream outputStream) {
        try {
            Rectangle pageSize = new Rectangle(842.0F, 595.0F);
            Document document = new Document(pageSize);
            PdfWriter.getInstance(document, outputStream);
            convertPdf(workBook, document);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void convertPdf(HSSFWorkbook workBook, Document document) {
        try {
            System.out.println(workBook.getNumberOfSheets());
            for (int a = 0; a < workBook.getNumberOfSheets(); a++) {
                document.newPage();
//                HSSFSheet sheet = workBook.getSheet("FirstPage");
                System.out.println(workBook.getSheetName(a));
                HSSFSheet sheet =workBook.getSheet(workBook.getSheetName(a));
                Map<String, PictureData> pictures = getPictures(sheet);
                float[] rowHeightArray = new float[sheet.getPhysicalNumberOfRows()];
                int cellsNum = sheet.getRow(0).getLastCellNum();
                float[] columnWidthArray = new float[cellsNum];
                for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                    HSSFRow row = sheet.getRow(i);
                    rowHeightArray[i] = row.getHeightInPoints();
                    for (int i1 = 0; i1 < cellsNum; i1++) {
                        if (i == 0) columnWidthArray[i1] = sheet.getColumnWidthInPixels(i1);
                    }
                }
                List<String> point = new ArrayList<String>();
                document.open();
                PdfPTable pdfPTable = new PdfPTable(columnWidthArray);
                // 0 边距
                pdfPTable.setWidthPercentage(100);
                for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                    HSSFRow row = sheet.getRow(i);
                    for (int i1 = 0; i1 < cellsNum; i1++) {
                        Map<String, Integer> map = this.isMergedRegion(sheet, i, i1);
                        PdfPCell pdfPCell = new PdfPCell();
                        pdfPCell.setMinimumHeight(getCellHeight(rowHeightArray, rowHeightArray[i], null));
                        HSSFCell cell = row.getCell(i1);
                        if (cell != null) {
                            HSSFCellStyle cellStyle = cell.getCellStyle();
                            HorizontalAlignment alignment = cellStyle.getAlignmentEnum();
                            if ("RIGHT".equals(alignment.name()))
                                pdfPCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
                            if ("LEFT".equals(alignment.name()))
                                pdfPCell.setHorizontalAlignment(Element.ALIGN_LEFT);
                            if ("CENTER".equals(alignment.name()))
                                pdfPCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                            if (map != null) {
                                if (compareRowNum(map, i)) {
                                    pdfPCell.setMinimumHeight(getCellHeight(rowHeightArray, rowHeightArray[i], map));
                                    pdfPCell.setRowspan(map.get("lastRow") - map.get("firstRow") + 1);
                                    pdfPCell.setColspan(map.get("lastColumn") - map.get("firstColumn") + 1);
                                    this.setBorder(pdfPCell, sheet, map);
                                    this.picSet(pictures, i + "-" + i1, pdfPCell);
                                    this.setCellValue(cell,pdfPCell,cellStyle.getFont(workBook));
                                    pdfPTable.addCell(pdfPCell);
                                    i1 = map.get("lastColumn");
                                    point.add(i + "," + i1);
                                } else {
                                    i1 = map.get("lastColumn");
                                }
                            } else {
                                this.picSet(pictures, i + "-" + i1, pdfPCell);
                                this.setBorder(cellStyle, pdfPCell);
                                this.setCellValue(cell,pdfPCell,cellStyle.getFont(workBook));
                                pdfPTable.addCell(pdfPCell);
                                point.add(i + "," + i1);
                            }
                        } else {
                            pdfPCell.disableBorderSide(15);
                            pdfPTable.addCell(pdfPCell);
                            point.add(i + "," + i1);
                        }
                    }
                }
                document.add(pdfPTable);
            }
            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 判断是否是合并单元格
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    private Map isMergedRegion(Sheet sheet, int row, int column) {
        Map<String, Integer> map = new HashMap<String, Integer>();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    map.put("firstRow", firstRow);
                    map.put("lastRow", lastRow);
                    map.put("firstColumn", firstColumn);
                    map.put("lastColumn", lastColumn);
                    return map;
                }
            }
        }
        return null;
    }

    /**
     * 比较合并单元格位置是否是在第一单元格，防止重复设置合并单元格
     *
     * @param map
     * @param rowNum
     * @return
     */
    private boolean compareRowNum(Map<String, Integer> map, int rowNum) {
        if (map == null) return false;
        int firstRow = map.get("firstRow");
        if (firstRow == rowNum) return true;
        return false;
    }

    /**
     * 设置平局宽度
     *
     * @param array
     * @param w
     * @param map
     * @return
     */
    private float getCellWidth(float[] array, float w, Map<String, Integer> map) {
        float num = 0;
        for (float v : array) {
            num += v;
        }
        return w * (700 / num);
    }

    private float getCellHeight(float[] array, float w, Map<String, Integer> map) {
        float num = 0;
        for (float v : array) {
            num += v;
        }
        return w * (510 / num);
    }

    /**
     * 设置边框
     *
     * @param cellStyle
     * @param pdfPCell
     */
    private void setBorder(HSSFCellStyle cellStyle, PdfPCell pdfPCell) {
        if (cellStyle.getBorderLeftEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(4);
        }
        if (cellStyle.getBorderRightEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(8);
        }
        if (cellStyle.getBorderTopEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(1);
        }
        if (cellStyle.getBorderTopEnum().getCode() == BorderStyle.DOTTED.getCode()) {
            pdfPCell.disableBorderSide(1);
            CustomCellTop cellTop = new CustomCellTop();
            pdfPCell.setCellEvent(cellTop);
        }
        if (cellStyle.getBorderBottomEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(2);
        }
        if (cellStyle.getBorderBottomEnum().getCode() == BorderStyle.DOTTED.getCode()) {
            pdfPCell.disableBorderSide(2);
            CustomCellBottom cellBottom = new CustomCellBottom();
            pdfPCell.setCellEvent(cellBottom);
        }

        if (cellStyle.getBorderLeftEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthLeft(2.0f);
        }

        if (cellStyle.getBorderRightEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthRight(2.0f);
        }

        if (cellStyle.getBorderTopEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthTop(2.0f);
        }

        if (cellStyle.getBorderBottomEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthBottom(2.0f);
        }
    }

    /**
     * 设置边框
     *
     * @param pdfPCell
     * @param sheet
     * @param map
     */
    private void setBorder(PdfPCell pdfPCell, HSSFSheet sheet, Map<String, Integer> map) {
        HSSFCell left = sheet.getRow(map.get("firstRow")).getCell(map.get("firstColumn"));
        HSSFCellStyle cellStyle1 = left.getCellStyle();
        HSSFCell bottom = sheet.getRow(map.get("lastRow")).getCell(map.get("lastColumn"));
        HSSFCellStyle cellStyle2 = bottom.getCellStyle();
        if (cellStyle1.getBorderLeftEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(4);
        }

        if (cellStyle2.getBorderRightEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(8);
        }
        if (cellStyle1.getBorderLeftEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthLeft(2.0f);
        }

        if (cellStyle2.getBorderRightEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthRight(2.0f);
        }

        if (cellStyle1.getBorderTopEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthTop(2.0f);
        }

        if (cellStyle2.getBorderBottomEnum().getCode() == BorderStyle.MEDIUM.getCode()) {
            pdfPCell.setBorderWidthBottom(2.0f);
        }

        if (cellStyle1.getBorderTopEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(1);
        }

        if (cellStyle1.getBorderTopEnum().getCode() == BorderStyle.DOTTED.getCode()) {
            pdfPCell.disableBorderSide(1);
            CustomCellTop cellTop = new CustomCellTop();
            pdfPCell.setCellEvent(cellTop);
        }

        if (cellStyle2.getBorderBottomEnum().getCode() == BorderStyle.NONE.getCode()) {
            pdfPCell.disableBorderSide(2);
        }

        if (cellStyle2.getBorderBottomEnum().getCode() == BorderStyle.DOTTED.getCode()) {
            pdfPCell.disableBorderSide(2);
            CustomCellBottom cellBottom = new CustomCellBottom();
            pdfPCell.setCellEvent(cellBottom);
        }

    }

    //虚线格式 顶部
    static class CustomCellTop implements PdfPCellEvent {
        public void cellLayout(PdfPCell cell, Rectangle position,
                               PdfContentByte[] canvases) {
            // TODO Auto-generated method stub
            PdfContentByte cb = canvases[PdfPTable.LINECANVAS];
            cb.saveState();
            cb.setLineWidth(0.8f);
            cb.setLineDash(new float[]{1.0f, 1.2f}, 0);
            cb.moveTo(position.getLeft(), position.getTop());
            cb.lineTo(position.getRight(), position.getTop());
            cb.stroke();
            cb.restoreState();
        }
    }

    //虚线格式 底部
    static class CustomCellBottom implements PdfPCellEvent {
        public void cellLayout(PdfPCell cell, Rectangle position,
                               PdfContentByte[] canvases) {
            PdfContentByte cb = canvases[PdfPTable.LINECANVAS];
            cb.saveState();
            cb.setLineWidth(0.8f);
            cb.setLineDash(new float[]{1.0f, 1.5f}, 0);
            cb.moveTo(position.getLeft(), position.getBottom());
            cb.lineTo(position.getRight(), position.getBottom());
            cb.stroke();
            cb.restoreState();
        }
    }

    /**
     * 获取单元格数值
     *
     * @param cell
     * @return
     */
    public static String getStringValueFromCell(HSSFCell cell) {
        SimpleDateFormat sFormat = new SimpleDateFormat("MM/dd/yyyy");
        DecimalFormat decimalFormat = new DecimalFormat("#.#");
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
            cellValue = cell.getStringCellValue();
        } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                double d = cell.getNumericCellValue();
                Date date = HSSFDateUtil.getJavaDate(d);
                cellValue = sFormat.format(date);
            } else {
                cellValue = decimalFormat.format((cell.getNumericCellValue()));
            }
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            cellValue = "";
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_ERROR) {
            cellValue = "";
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            String cellFormula = cell.getCellFormula().toString();
            String p = cellFormula.substring(cellFormula.indexOf("&") + 1, cellFormula.lastIndexOf("&"));
            int rownNum = Integer.valueOf(p.replaceAll("[a-zA-Z]", ""));
            int cellNum = LetterValue.obj2Enum(p.replaceAll("\\d+", "")).getValue();
            cellValue = "*" + getStringValueFromCell(cell.getSheet().getRow(rownNum - 1).getCell(cellNum)) + "*";
        }
        return cellValue;
    }

    /**
     * 图片设置
     *
     * @param pictures
     * @param k
     * @param pdfPCell
     * @throws IOException
     * @throws BadElementException
     */
    private void picSet(Map<String, PictureData> pictures, String k, PdfPCell pdfPCell) throws IOException, BadElementException {
        if (!pictures.containsKey(k)) return;
        for (String key : pictures.keySet()) {
            if (k.equals(key)) {
                PictureData pictureData = pictures.get(key);
                byte[] data = pictureData.getData();
                // 获取图片格式
                String ext = pictureData.suggestFileExtension();
                Image img = Image.getInstance(data);
                //图片缩小到40%
//                img.scalePercent(80);
                img.scaleToFit(40, 40);
                img.setAlignment(Element.ALIGN_CENTER);
                pdfPCell.addElement(img);
//                pdfPCell.setImage(img);
            }
        }
    }

    /**
     * 获取图片和位置 (xls)
     *
     * @param sheet
     * @return
     */
    public static Map<String, PictureData> getPictures(HSSFSheet sheet) {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
        for (HSSFShape shape : list) {
            if (shape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor cAnchor = (HSSFClientAnchor) picture.getAnchor();
                PictureData pdata = picture.getPictureData();
                String key = cAnchor.getRow1() + "-" + cAnchor.getCol1(); // 行号-列号
                map.put(key, pdata);
            }
        }
        return map;
    }

    /**
     * 根据excel的单元格内容给pdf表格赋值
     * @param cell
     * @param pdfPCell
     * @param font
     */

    private void setCellValue(HSSFCell cell,PdfPCell pdfPCell,HSSFFont font){
        String stringCellValue = this.getStringValueFromCell(cell);
        if (stringCellValue != "") {
            pdfPCell.setPadding(4);
            pdfPCell.setPaddingTop(0);
            pdfPCell.setPaddingBottom(0);
            Paragraph elements = new Paragraph(stringCellValue, PdfFont.getFont(font));
            pdfPCell.setPhrase(elements);}
    }

}
