package my
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.configuration.RunConfiguration
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

class WriteExcel {

	@Keyword
	public void demoKey(String name) throws IOException{
		String projectDirPath = RunConfiguration.getProjectDir()
		FileInputStream fis =
				new FileInputStream(projectDirPath + "/Epikso.xlsx")
		XSSFWorkbook workbook = new XSSFWorkbook(fis)
		XSSFSheet sheet = workbook.getSheet("Sheet1")
		int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum()
		WebUI.comment("sheet.getLastRowNum():${sheet.getLastRowNum()} - sheet.getFirstRowNum():${sheet.getFirstRowNum()} => ${rowCount}")
		Row row = sheet.createRow(rowCount+1)
		Cell cell = row.createCell(0)
		cell.setCellType(cell.CELL_TYPE_STRING)
		cell.setCellValue(name)
		FileOutputStream fos =
				new FileOutputStream(projectDirPath + "/Epikso.xlsx")
		workbook.write(fos)
		fos.close()
	}

	@Keyword
	public void init(String name) throws IOException{
		String projectDirPath = RunConfiguration.getProjectDir()
		XSSFWorkbook workbook = new XSSFWorkbook()
		XSSFSheet sheet = workbook.createSheet("Sheet1")
		Row row = sheet.createRow(0)
		Cell cell = row.createCell(0)
		cell.setCellType(cell.CELL_TYPE_STRING)
		cell.setCellValue(name)
		FileOutputStream fos =
				new FileOutputStream(projectDirPath + "/Epikso.xlsx")
		workbook.write(fos)
		fos.close()
	}
}