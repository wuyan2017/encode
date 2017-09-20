import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.UnsupportedEncodingException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;
/**
 *@Author:吴焰
 *@Date:9:48 2017/9/20
 *@Description:读取Excel表中的数据，对每行数据加密
 */
public class Main {
    public static void main(String[] args) {

        try {
            Workbook wb = new XSSFWorkbook(new FileInputStream("C:\\Users\\wuyan\\Desktop\\44.xlsx"));
            // 获取sheet数目
            for (int t = 0; t < wb.getNumberOfSheets(); t++) {
                Sheet sheet = wb.getSheetAt(t);
                Row row = null;
                int lastRowNum = sheet.getLastRowNum();
                // 循环读取
                for (int i = 0; i <= lastRowNum; i++) {
                    row = sheet.getRow(i);
                    if (row != null) {
                        // 获取第一列每一行的值
                            Cell cell = row.getCell(0);
                            String value = getCellValue(cell);
                            if (!value.equals("")) {
                                //加密
                                String s=getSHA256StrJava(value);
                                System.out.print(value);
                        }
                        System.out.println();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
        private static String getCellValue(Cell cell) {
            Object result = "";
            if (cell != null) {
                switch (cell.getCellType()) {
                    case CELL_TYPE_STRING:
                        result = cell.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        result = cell.getNumericCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        result = cell.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        result = cell.getCellFormula();
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        result = cell.getErrorCellValue();
                        break;
                    case Cell.CELL_TYPE_BLANK:
                        break;
                    default:
                        break;
                }
            }
            return result.toString();
        }
    /**
     *  利用java原生的摘要实现SHA256加密
     * @param str 加密后的报文
     * @return
     */
    public static String getSHA256StrJava(String str){
        MessageDigest messageDigest;
        String encodeStr = "";
        try {
            messageDigest = MessageDigest.getInstance("SHA-256");
            messageDigest.update(str.getBytes("UTF-8"));
            encodeStr = byte2Hex(messageDigest.digest());
        } catch (NoSuchAlgorithmException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return encodeStr;
    }

    /**
     * 将byte转为16进制
     * @param bytes
     * @return
     */
    private static String byte2Hex(byte[] bytes){
        StringBuffer stringBuffer = new StringBuffer();
        String temp = null;
        for (int i=0;i<bytes.length;i++){
            temp = Integer.toHexString(bytes[i] & 0xFF);
            if (temp.length()==1){
                //1得到一位的进行补0操作
                stringBuffer.append("0");
            }
            stringBuffer.append(temp);
        }
        return stringBuffer.toString();
    }
}