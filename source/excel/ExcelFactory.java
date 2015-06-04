package excel;
import java.io.FileOutputStream;  
import java.io.IOException;  

import org.apache.poi.hssf.usermodel.*;  
/** 
 * 利用POI工具创建Excel工作薄和工作表，并向其中写入内容 
 * 
 */  
public class ExcelFactory {  
      
    private void createExcel()throws IOException {  
        String excelFile="F:/myexcel.xls";  
        FileOutputStream fos=new FileOutputStream(excelFile);  
        HSSFWorkbook wb=new HSSFWorkbook();//创建工作薄  
        HSSFSheet sheet=wb.createSheet();//创建工作表  
        wb.setSheetName(0, "sheet0");//设置工作表名  
          
        HSSFRow row=null;  
        HSSFCell cell=null;  
        int rownum = 0;
        int column = 5;
        row = sheet.createRow(rownum++);
        for(int i = 0; i < column; i++){
        	cell = row.createCell(i);
        	cell.setCellValue("专线名称"+i);
        	
        }
        for (;rownum < 10; rownum++) {  
            row=sheet.createRow(rownum);//新增一行  
            cell=row.createCell(0);//新增一列 
            cell=row.createCell(0);  
            cell.setCellValue("第"+rownum+"行");  
        }  
        wb.write(fos);  
        fos.close();  
        wb.close();
    }  
    /** 
     * @param args 
     *2012-10-23 
     *void 
     * @throws IOException  
     * @throws DocumentException 
     */  
    public static void main(String[] args) throws IOException {  

        new ExcelFactory().createExcel();  
  
    }  
  
} 
