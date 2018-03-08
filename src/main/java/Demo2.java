import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

/**
 * <p>设置单元格格式</p>
 *
 * @Author penelopeWu
 * Date:2018-03-08 23:00
 */
public class Demo2 {
    public static void main(String[] args) throws Exception {
        test();
    }

    private static void test() throws Exception {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("第一页");
        Row row = sheet.createRow(0);
        row.setHeightInPoints(30);

        creatCell(workbook, row, (short) 0, HSSFCellStyle.ALIGN_CENTER, HSSFCellStyle.VERTICAL_BOTTOM);
        creatCell(workbook, row, (short) 0, HSSFCellStyle.ALIGN_FILL, HSSFCellStyle.VERTICAL_CENTER);
        creatCell(workbook, row, (short) 0, HSSFCellStyle.ALIGN_LEFT, HSSFCellStyle.VERTICAL_TOP);
        creatCell(workbook, row, (short) 0, HSSFCellStyle.ALIGN_RIGHT, HSSFCellStyle.VERTICAL_TOP);

        FileOutputStream out = new FileOutputStream("D:\\xxx.xls");
        workbook.write(out);
        out.close();

    }

    private static void creatCell(Workbook wb, Row row, short colum, short halign, short valign) {
        //创建单元格
        Cell cell = row.createCell(colum);
        //设置值
        cell.setCellValue(new HSSFRichTextString("ALign It"));
        //创建单元格样式
        CellStyle cellStyle = wb.createCellStyle();
        //设置单元格水平方向对齐方式
        cellStyle.setAlignment(halign);
        //设置单元格垂直方向对齐方式                   (3.17的版本，setAlignment方法的参数不是short类型了)
        cellStyle.setVerticalAlignment(valign);
        //应用单元格样式
        cell.setCellStyle(cellStyle);
    }
}
