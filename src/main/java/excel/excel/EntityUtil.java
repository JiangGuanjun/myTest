package excel.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EntityUtil {
	
	private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    //用于匹配日期的字符串
    private static final String EXCEL_DATE = "date" ;
    //用于匹配类注释开始的字符串
    private static final String EXCEL_CLASS_START = "classNotes" ;
    //用于匹配属性的开始的字符串attributeNotes
    private static final String EXCEL_ATTRIBUTE_START = "attributeNotes" ;
    
    //类的前中标记  前 + 类名 + 中 + 后
    private static final String CLASS_BEFORE = "public class " ;
    private static final String CLASS_MIDDLE = "{" ;
    private static final String CLASS_BEHIND = "}" ;
    
    //注释的标记 前 + 中 +文字 + 中。。。。+后
    private static final String NOTE_BEFORE = "/** \n" ;
    private static final String NOTE_MIDDLE = "* \n" ;
    private static final String NOTE_BEHIND = "../ " ;
    
    //注释标识
    private static final String NOTE_SIGN = "@" ;
       
    
	
	public static String getDate(){
		Date date = new Date() ;
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd hh-mm-ss") ;
		String d = format.format(date) ;
		return d ;
	}
	
	/**
     * 判断Excel的版本,获取Workbook
     * @param file
     * @return Workbook
     * @throws IOException
     */
    public static Workbook getWorkbok(File file) throws IOException{
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if(file.getName().endsWith(EXCEL_XLS)){     //Excel&nbsp;2003
            wb = new HSSFWorkbook(in);
        }else if(file.getName().endsWith(EXCEL_XLSX)){    // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }
	
	
	public static boolean writeDate(String path) throws IOException{
		File finalXlsxFile = new File(path);
		Workbook workBook = null ;
		//判断是不是文件，如果是文件直接执行写入操作
		if(finalXlsxFile.isFile()){
			workBook = getWorkbok(finalXlsxFile);
			//获取工作页
			//int sheetCount = workBook.getNumberOfSheets();
			//System.out.println("excel一共："+sheetCount+"页");
			Sheet sheet = workBook.getSheetAt(0);
			Row row = null ;
			//获取最后以上的行号（从0开始）
			int lastLine = sheet.getLastRowNum() ;
			int sumLine = lastLine + 1 ;
			System.out.println("共" + sumLine+ "行");
			for(int i=0; i<sumLine; i++){
				row = sheet.getRow(i);
				if(row == null){
					continue ;  
				}
				short lastNum = row.getLastCellNum() ;
				int cellCount = lastNum + 1 ;
				System.out.println("第"+i+"行，一共："+cellCount + "个小格");
				for(int j=0; j< cellCount; j++){
					if(row.getCell(j) == null){
						continue ;
					}
					//String  value = row.getCell(j).getStringCellValue() ;
					Cell c = row.getCell(j) ;
					String value = getCellByString(c) ;
					if(EXCEL_DATE.equals(value)){
						System.out.println("进入：：：：：");
						row = sheet.getRow(i+1);
						System.out.println(";"+getCellByString(row.getCell(j))+";");
						row.getCell(j).setCellValue(getDate());
						System.out.println(";"+getCellByString(row.getCell(j))+";");
					}
				}
			}
			
		}
		return true ;
	}
	
	
	public static Map<String ,Map<String,Object>> getExcelContent(String ExcelPath) throws IOException{
		Map<String, Map<String ,Object>> map = new HashMap<String, Map<String,Object>>() ;
		Map<String,String> classMap = new HashMap<String,String>() ;
		Map<String,Attribute> attribute = new HashMap<String,Attribute>() ;
		
		File file = new File(ExcelPath) ;
		Workbook workbok = getWorkbok(file) ;
		Sheet sheet = workbok.getSheetAt(0) ;
		int lastNumRow = sheet.getLastRowNum() ;
		//总行数
		int rowSum = lastNumRow + 1 ;
		int i=0; 
		Row row = null ;
		Cell cell = null ;
		int clllSum = 0 ;
		int lastNumCell = 0 ;
		for(; i<rowSum; i++ ){
			boolean isBreak = false ;
			row = sheet.getRow(i) ;
			if(row == null){
				continue ;
			}
			lastNumCell =  row.getLastCellNum() ;
			//每行的总列数
			clllSum = lastNumCell + 1 ;
			for(int j=0; j<clllSum; j++){
				cell = row.getCell(j) ;
				if(cell == null){
					continue ;
				}
				String value = getCellByString(cell) ;
				//当value等于类的注释标识的时候，则下一行就是类的十几注释的字段
				if(EXCEL_CLASS_START.equals(value.trim())){
					isBreak = true ;
					i++ ;
					break ;
				}
			}
			if(isBreak){
				break ;
			}
		}
		
		row = sheet.getRow(i) ;
		i++ ;
		lastNumCell = row.getLastCellNum() ;
		clllSum = lastNumCell + 1 ;
		Row rowValue = sheet.getRow(i) ;
		for(int j=0;j<clllSum;j++){
			Cell cellKey = row.getCell(j) ;
			Cell cellValue = rowValue.getCell(j) ;
			if(cellKey != null && cellValue != null){
				String stringKey = getCellByString(cellKey) ;
				String stringValue = getCellByString(cellValue) ;
				classMap.put(stringKey, stringValue) ;
			}
		}
		i++ ;
		
		boolean isSet = false ;
		
		for(; i<rowSum; i++ ){
			row = sheet.getRow(i) ;
			if(row == null){
				continue ;
			}
			lastNumCell =  row.getLastCellNum() ;
			//每行的总列数
			clllSum = lastNumCell + 1 ;
			for(int j=0; j<clllSum; j++){
				cell = row.getCell(j) ;
				if(cell == null){
					continue ;
				}
				String value = getCellByString(cell) ;
				//当value等于类的注释标识的时候，则下一行就是类的十几注释的字段
				if(!isSet){
					if(EXCEL_ATTRIBUTE_START.equals(value.trim())){
						i++ ;
						isSet = true ;
					}
				}
				
			}
			
		}
		
		
		
		
		return null ;
	}
	
	
	public static void writeEntityByExcel(String ExcelPath, String javaPath) throws IOException{
		File f = new File(ExcelPath) ;
		Workbook workBook = getWorkbok(f);
		Sheet sheet = workBook.getSheetAt(0) ;
		int LastLine = sheet.getLastRowNum() ;
		//总行数
		int countRow = LastLine + 1 ;
		int i = 0;
		for(; i<countRow; i++ ){
			Row row = sheet.getRow(i) ;
			if(row == null){
				continue ;
			}
			int lastCell = row.getLastCellNum() ;
			//每一行的列数
			int countCell = lastCell + 1 ;
			for(int j=0; j<countCell; j++){
				Cell cell = row.getCell(j) ;
				if(cell ==null){
					continue ;
				}
				String value = getCellByString(cell) ;
				//单元格的值与类的开始标识进行匹配，如果匹配成功则它的下一行即为类的注释内容
				if(EXCEL_CLASS_START.equals(value)){
					
				}
			}
			
			
		}
		
	}
	
	public static String  getCellByString(Cell cell){
		String value = "" ;
		int type = cell.getCellType() ;
		switch(cell.getCellType()){
		case HSSFCell.CELL_TYPE_NUMERIC:
			value = new DecimalFormat("0").format(cell.getNumericCellValue());
			break ;
		case HSSFCell.CELL_TYPE_STRING: // 字符串
	        value = cell.getStringCellValue();
	        break;
	    case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
	        value = cell.getBooleanCellValue() + "";
	        break;
	    case HSSFCell.CELL_TYPE_FORMULA: // 公式
	        value = cell.getCellFormula() + "";
	        break;
	    case HSSFCell.CELL_TYPE_BLANK: // 空值
	        value = "";
	        break;
	    case HSSFCell.CELL_TYPE_ERROR: // 故障
	        value = "非法字符";
	        break;
	    default:
	        value = "未知类型";
	        break;
		}
		
	
		return value ;
	}
	
	
	
	
	
	public static void main(String[] args) throws IOException {
		writeDate("E:/Test.xlsx") ;
	}
}
