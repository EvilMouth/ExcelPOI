package com.zyhang.excelPoi;

import android.os.Environment;
import android.util.Log;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Locale;

/**
 * ProjectName:ExcelPoi
 * Description:
 * Created by zyhang on 2017/5/26.下午4:42
 * Modify by:
 * Modify time:
 * Modify remark:
 */

public class ExcelPoiUtil {

    private static final String TAG = "ExcelPoiUtil";

    /**
     * 读Excel
     *
     * @param filePath 文件路径
     */
    public static void read(String filePath) {
        try {
            InputStream is = new FileInputStream(filePath);
            Workbook workbook;
            Sheet sheet;
            String postfix = filePath.substring(filePath.lastIndexOf("."), filePath.length());
            if (postfix.equals(".xls")) {
                // 2003 Excel
                workbook = new HSSFWorkbook(new POIFSFileSystem(is));
                sheet = workbook.getSheetAt(0);
            } else {
                // 2007 Excel
                workbook = new XSSFWorkbook(is);
                sheet = workbook.getSheetAt(0);
            }
            Log.i(TAG, "sheet === " + sheet.getSheetName());

            int rowCount = sheet.getLastRowNum();//此sheet内容总行数
            Log.i(TAG, "rowCount === " + rowCount);
            for (int r = 0; r <= rowCount; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                int cellCount = row.getLastCellNum();//此行(row)内容总列数
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) continue;
                    String content = getCellFormatValue(cell);
                    Log.i(TAG, "r === " + r + " c === " + c + " content === " + content);
                }
            }

            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取单元格内容
     *
     * @param cell 单元格
     * @return content
     */
    private static String getCellFormatValue(Cell cell) throws Exception {
        String value = "";
        // 判断当前Cell的Type
        switch (cell.getCellType()) {
            // 如果当前Cell的Type为NUMERIC
            case Cell.CELL_TYPE_NUMERIC:
                // 判断当前的cell是否为Date
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    // 方法2：这样子的data格式是不带带时分秒的：2011-10-12
                    double date = cell.getNumericCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm", Locale.CHINA);
                    value = sdf.format(HSSFDateUtil.getJavaDate(date));
                } else {
                    // 如果是纯数字通过NumberToTextConverter.toText(double)将double转成string
                    value = NumberToTextConverter.toText(cell.getNumericCellValue());
                }
                break;
            // 如果当前Cell的Type为STRING
            case Cell.CELL_TYPE_STRING:
                // 取得当前的Cell字符串
                value = cell.getStringCellValue();
                break;
            // 如果当前Cell的Type为BOOLEAN
            case Cell.CELL_TYPE_BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
        }
        return value;
    }

    /**
     * 改Excel xlsx
     *
     * @param filePath 文件路径
     * @param s        标签
     * @param r        行
     * @param c        列
     * @param content  内容
     */
    public static void update(String filePath, int s, int r, int c, String content) {
        try {
            FileInputStream fis = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(s);
            XSSFRow row = sheet.getRow(r);

            XSSFCell cell = row.createCell(c);//这里是创建
            cell.setCellValue(content);//设置内容

            //设置颜色
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);

            fis.close();

            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 写Excel xlsx
     */
    public static void write() {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("sheet"));
            //模拟添加数据
            for (int i = 0; i < 10; i++) {
                Row row = sheet.createRow(i);
                Cell cell = row.createCell(i);
                cell.setCellValue(i);
            }
            String outFileName = "test.xlsx";
            File outFile = new File(Environment.getExternalStorageDirectory().getPath(), outFileName);
            FileOutputStream fos = new FileOutputStream(outFile.getAbsolutePath());
            workbook.write(fos);
            fos.flush();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
