package ExcelWriter;

import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class ExcelCodeGenerator {
	/**
	
		if you run code by passing column names as space separated string you will get below result

		Example :
		columnNames = OrderDate Region Rep Item Units Unit Cost Total
		delimitor = space
		output : 

		XSSFWorkbook workbook = new XSSFWorkbook();
		String sheetName = "";//Write your sheetname here 
		XSSFSheet excelSheet = workbook.createSheet(sheetName);
		XSSFRow excelHeader = excelSheet.createRow(0);
		int count = 0;
		excelHeader.createCell(count++).setCellValue("OrderDate");
		excelHeader.createCell(count++).setCellValue("Region");
		excelHeader.createCell(count++).setCellValue("Rep");
		excelHeader.createCell(count++).setCellValue("Item");
		excelHeader.createCell(count++).setCellValue("Units");
		excelHeader.createCell(count++).setCellValue("Unit");
		excelHeader.createCell(count++).setCellValue("Cost");
		excelHeader.createCell(count++).setCellValue("Total");

		int rowIndex = 1
		for () {//Write your for loop 
			XSSFRow excelRow = excelSheet.createRow(rowIndex++);
			count = 0;
			//Write your logic and computed value here
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
			excelRow.createCell(count++).setCellValue();
		}
	
	**/
	
	
	public void generateExcelWriteCode(String columnNames,String delimitor) {
		List<String> columns = Arrays.asList(columnNames.split(delimitor));
		String excelHeaderCode[] = {"XSSFWorkbook workbook = new XSSFWorkbook();\n"+
								"String sheetName = \"\";//Write your sheetname here \n"+
								"XSSFSheet excelSheet = workbook.createSheet(sheetName);\n"+
								"XSSFRow excelHeader = excelSheet.createRow(0);\n"+
								"int count = 0;\n"};
		
		columns.forEach(column ->{
			excelHeaderCode[0] += "excelHeader.createCell(count++).setCellValue(\"" + column + "\");\n";
		});
		
		String excelRowCode[] = {"int rowIndex = 1\n"+"for () {//Write your for loop \r\n" + 
								"	XSSFRow excelRow = excelSheet.createRow(rowIndex++);\r\n" + 
								"	count = 0;\n 	//Write your logic and computed value here\n"};
		
		columns.forEach(column ->{
			excelRowCode[0] += "	excelRow.createCell(count++).setCellValue();\n";
		});
		excelRowCode[0] += "}";
		
		System.out.println(excelHeaderCode[0]);
		System.out.println(excelRowCode[0]);
		
	}
}
