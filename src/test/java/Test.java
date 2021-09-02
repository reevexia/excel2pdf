import com.reeve.AbstractExcelConvertPdf;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @Author Reeve
 * @Date 2021/7/29 10:49
 */
public class Test {
    public static void main(String[] args) throws IOException {
//        tttt();
//        start();
        test1();
    }

    private static void test1() {
        try {
            AbstractExcelConvertPdf abstractExcelConvertPdf = new AbstractExcelConvertPdf();
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            abstractExcelConvertPdf.convert("src/main/resources/entryPrint2.xls",outputStream);
            getFile(outputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private static void start() throws IOException {
        long time3 = new Date().getTime();
        HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream("C:\\Users\\Reeve\\Desktop\\entryPrint.xls"));
        long time4 = new Date().getTime();
        System.out.println("2单次耗时："+(time4-time3));
//        HSSFSheet sheet = workBook.getSheet("test");
        HSSFSheet sheet = workBook.getSheet("FirstPage");
//        HSSFSheet sheet = workBook.getSheet("OtherPages");
        System.out.println(sheet.getNumMergedRegions());
        int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
        HSSFRow row = sheet.getRow(0);
        int physicalNumberOfCells = row.getPhysicalNumberOfCells();
        HSSFCell cell = row.getCell(0);
        CellRangeAddress mergedRegion = sheet.getMergedRegion(2);
        Map mergedRegion1 = isMergedRegion(sheet, 2, 0);
        System.out.println(compareRowNum(mergedRegion1, 0));
        System.out.println(compareRowNum(mergedRegion1, 1));
        System.out.println(compareRowNum(mergedRegion1, 3));

        int numMergedRegions = sheet.getNumMergedRegions();
//        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
//        for (CellRangeAddress region : mergedRegions) {
//            System.out.println("startRow:"+region.getFirstRow());
//            System.out.println("endRow:"+region.getLastRow());
//            System.out.println("startColumn:"+region.getFirstColumn());
//            System.out.println("endColumn:"+region.getLastColumn());
//        }
        long time5 = new Date().getTime();
        System.out.println("总耗时："+(time5-time3));
    }

    private static void tttt() throws IOException {
        long time3 = new Date().getTime();
        HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream("C:\\Users\\Reeve\\Desktop\\entryPrint2.xls"));
        long time4 = new Date().getTime();
        System.out.println("2单次耗时："+(time4-time3));
//        HSSFSheet sheet = workBook.getSheet("FirstPage");
        HSSFSheet sheet = workBook.getSheet("OtherPages");
        HSSFRow row = sheet.getRow(0);
        System.out.println(row.getHeightInPoints());
        int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
        System.out.println(physicalNumberOfRows);
        HSSFCell cell = row.getCell(2);
        int physicalNumberOfCells = row.getPhysicalNumberOfCells();
        System.out.println(physicalNumberOfCells);
        float columnWidthInPixels = sheet.getColumnWidthInPixels(cell.getColumnIndex());
        System.out.println(columnWidthInPixels);
        HSSFCellStyle cellStyle = cell.getCellStyle();
        short lastCellNum = row.getLastCellNum();
        System.out.println(lastCellNum);
        int topRow = sheet.getLastRowNum();
        short leftCol = sheet.getLeftCol();
        System.out.println(topRow);
        System.out.println(leftCol);
        long time5 = new Date().getTime();
        System.out.println("总耗时："+(time5-time3));
    }


    private static Map isMergedRegion(Sheet sheet, int row, int column) {
        Map<String, Integer> map = new HashMap<String, Integer>();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    map.put("firstRow",firstRow);
                    map.put("lastRow",lastRow);
                    map.put("firstColumn",firstColumn);
                    map.put("lastColumn",lastColumn);
                    return map;
                }
            }
        }
        return null;
    }

    private static boolean compareRowNum(Map<String, Integer> map, int rowNum){
        if (map == null) return false;
        Integer firstRow = map.get("firstRow");
        if (firstRow==rowNum) return true;
        return false;
    }

    public static String getFile(ByteArrayOutputStream outputStream) {
        BufferedOutputStream bos = null;
        FileOutputStream fos = null;
        File file = null;
        String path = "";
        try {
            //创建临时文件的api参数 (文件前缀,文件后缀,存放目录)
            file = new File("pdf.pdf");
            fos = new FileOutputStream(file);
            bos = new BufferedOutputStream(fos);
            bos.write(outputStream.toByteArray());
            path = file.getPath();
            return path;
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("创建临时文件失败!" + e.getMessage());
            return null;
        } finally {
            if (bos != null) {
                try {
                    bos.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }
    }
}
