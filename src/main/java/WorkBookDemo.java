import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Date;

/**
 * @Author penelopeWu
 * Date:2018-03-07 19:55
 */
public class WorkBookDemo {
    public static void main(String[] args) throws Exception {
        demo5();

    }


    /**
     * 使用poi生成工作薄
     *
     * @throws FileNotFoundException
     */
    public static void demo1() throws Exception {
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream outputStream = new FileOutputStream("D:\\用poi生成的工作薄.xls");
        Sheet sheet = workbook.createSheet("工作薄1");
        workbook.createSheet("工作薄2");

        //生成一行
        Row row = sheet.createRow(0);
        //生成单元格并赋值
        row.createCell(0).setCellValue("haha");
        row.createCell(1).setCellValue(1);
        row.createCell(2).setCellValue(new Date());

        workbook.write(outputStream);
        outputStream.close();
    }


    /**
     * 处理时间格式的单元格
     *
     * @throws Exception 异常
     */
    public static void demo2() throws Exception{
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream outputStream = new FileOutputStream("D:\\时间格式的单元格.xls");
        Sheet sheet = workbook.createSheet("工作薄1");

        //生成一行
        Row row = sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();
        CreationHelper creationHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        workbook.write(outputStream);
        outputStream.close();
    }

    /**
     * 处理不同类型内容的单元格
     * <p>单元格中可以存放各种类型的数据</p>
     *
     * @throws Exception 异常
     */
    public static void demo3() throws Exception {
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream outputStream = new FileOutputStream("D:\\不同类型内容的单元格.xls");
        Sheet sheet = workbook.createSheet("工作薄1");

        Row row = sheet.createRow(0);
        //数字，字符串，布尔值，日期
        row.createCell(0).setCellValue(0);
        row.createCell(1).setCellValue("1");
        row.createCell(2).setCellValue(true);

        workbook.write(outputStream);
        outputStream.close();
    }

    /**
     * 遍历工作薄
     *
     * @throws Exception 异常
     */
    public static void demo4() throws  Exception{
        InputStream is = new FileInputStream("D:\\中奖名单");
        POIFSFileSystem fileSystem = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(fileSystem);
        //获取第一个sheet页
        HSSFSheet sheet = wb.getSheetAt(0);

        if (sheet == null){
            return;
        }

        //遍历行
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum ++){
            HSSFRow row = sheet.getRow(rowNum);
            if (row == null){
                continue;
            }
            //遍历单元格

            for (int cellNum = 0; cellNum <= row.getLastCellNum(); cellNum ++){
                HSSFCell cell = row.getCell(cellNum);
                if (cell == null){
                    continue;
                }

                System.out.print(getValue(cell));
            }
            System.out.println();
        }
    }

    private static String getValue(HSSFCell cell){
        if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN){
            return String.valueOf(cell.getBooleanCellValue());
        }else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
            return String.valueOf(cell.getNumericCellValue());
        }else {
            return String.valueOf(cell.getStringCellValue());
        }
    }


    /**
     * 文本提取
     *
     * @throws Exception
     */
    public static void demo5() throws  Exception{
        InputStream is = new FileInputStream("D:\\用poi生成的工作薄.xls");
        POIFSFileSystem fileSystem = new POIFSFileSystem(is);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);

        ExcelExtractor excelExtractor = new ExcelExtractor(workbook);
        excelExtractor.setIncludeSheetNames(false);
        System.out.println(excelExtractor.getText());
    }













}
