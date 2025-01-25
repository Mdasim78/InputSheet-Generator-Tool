package InputSheet.Generator;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Automation_InputSheet {

	static XSSFWorkbook automationInputSheets[];
	public static XSSFWorkbook[] generateAutomationInputSheet() throws IOException {
		automationInputSheets = new XSSFWorkbook[2];
		java.io.InputStream instInputfis = Automation_InputSheet.class.getResourceAsStream("/Claim_Data_Institutional.xlsx");
		automationInputSheets[0] = new XSSFWorkbook(instInputfis);
		XSSFSheet instDataSheet = automationInputSheets[0].getSheet("Data");
		clearSheetContent(instDataSheet);
		
		java.io.InputStream profInputfis = EDI_InputSheet.class.getResourceAsStream("/Claim_Data_Professional.xlsx");
		automationInputSheets[1] = new XSSFWorkbook(profInputfis);
		XSSFSheet profDataSheet = automationInputSheets[1].getSheet("Data");
		clearSheetContent(profDataSheet);

		FileInputStream UTCfis = new FileInputStream(GUI.file);
		HSSFWorkbook UTCfile = new HSSFWorkbook(UTCfis);
		HSSFSheet claimSheet = UTCfile.getSheet("claims");

		int lastRowNo = claimSheet.getPhysicalNumberOfRows()-1;

		//writing input data
		for(int i=1;i<lastRowNo;i++) {
			System.out.println(lastRowNo);
			System.out.println("started processing row :"+i);
			if(claimSheet.getRow(i).getCell(0).getCellType() != org.apache.poi.ss.usermodel.CellType.BLANK) {
				if(claimSheet.getRow(i).getCell(1).getStringCellValue().equals("HCFA")) writeProfData(claimSheet, i, lastRowNo, profDataSheet);
				else writeInstData(claimSheet, i, lastRowNo, instDataSheet);
			}
			UTCfile.close();

		}
		return automationInputSheets;
	}
	//removing existing rows to enter fresh data into the input sheet
	public static void clearSheetContent(XSSFSheet inputSheet) {
		for(int i=2;i <= inputSheet.getLastRowNum(); i++) {
			XSSFRow row = inputSheet.getRow(i);
			if(row != null) inputSheet.removeRow(row);
		}
	}
	//setting cell styles to text
	public static XSSFCell setCellToText(XSSFWorkbook wb,XSSFCell cell) {
		XSSFCellStyle textStyle = wb.createCellStyle();
		XSSFDataFormat format = wb.createDataFormat();
		textStyle.setDataFormat(format.getFormat("@"));
		cell.setCellStyle(textStyle);
		return cell;
	}
	//check cell type & return string data from cell accordingly
	public static String getCellData(HSSFCell cell) {
		try {
			switch(cell.getCellType()){
			case NUMERIC:return String.valueOf(((int)cell.getNumericCellValue())).replaceAll(" ","");
			case STRING: return cell.getStringCellValue().replaceAll(" ","");
			}
		}
		catch (NullPointerException e) {
			System.out.println("Skipped cell because cell is empty");
		}
		return null;
	}

	//check date cell type & return string data from date cell accordingly
	public static String getDateCellData(HSSFCell cell) {
		try {
			switch(cell.getCellType()) {
			case NUMERIC:return cell.getLocalDateTimeCellValue().format(java.time. format.DateTimeFormatter.ofPattern("MM/dd/yyyy"));
			case STRING:return new SimpleDateFormat("MM/dd/yyyy").format(new SimpleDateFormat("M/dd/yyyy").parse(cell.getStringCellValue()));
			}
		}
		catch (NullPointerException | ParseException e) {
			System.out.println("Skipped cell because cell is empty or unable to parse date format");
		}
		return null;
	}
	//check date cell type & return string data from amount cell accordingly
	public static String getAmountCellData(HSSFCell cell) {
		try {
			switch(cell.getCellType()) {
			case NUMERIC:return String.valueOf(((double)cell.getNumericCellValue())).replaceAll("[^0-9.]","");
			case STRING: return cell.getStringCellValue().replaceAll("[^0-9.]","");
			}
		}
		catch (NullPointerException e) {
			System.out.println("Skipped cell because cell is empty or unable to parse date format");
		}
		return null;
	}
	public static void writeProfData(HSSFSheet claimSheet, int startRowNo, int lastRowNo, XSSFSheet DataSheet) {

		for(int i=startRowNo;i <= lastRowNo; i++) {
			HSSFRow claimSheetRow = claimSheet.getRow(i);
			XSSFRow DataSheetlastRowNo= DataSheet.getRow(DataSheet.getLastRowNum());
			System.out.println("prof claim key:"+(int)claimSheetRow.getCell(0).getNumericCellValue()+"|"+"start row"+startRowNo);
			if(i == startRowNo) {
				DataSheetlastRowNo = DataSheet.createRow(DataSheet.getLastRowNum()+1);

				//Scenario name
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(0)).setCellValue(GUI.ScenarioName+new DecimalFormat("000").format((int)claimSheetRow.getCell(0).getNumericCellValue()));
				//receipt date
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(2)).setCellValue(getDateCellData(claimSheetRow.getCell(0)));
				//DX code
				String DXCode[] = getCellData(claimSheet.getRow(startRowNo).getCell(7)).split(",");
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(19)).setCellValue(DXCode[0]);
				if(DXCode.length > 1) {
					for(int k=1;k<DXCode.length;k++) {
						setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(19+k)).setCellValue(DXCode[k]);
					}
				}
			}
			else {
				//index number
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(5)).setCellValue(getCellData(claimSheet.getRow(i).getCell(9)));
				//from date
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(6)).setCellValue(getDateCellData(claimSheetRow.getCell(15)));
				//service end date
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(7)).setCellValue(getDateCellData(claimSheetRow.getCell(16)));
				//POS
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(8)).setCellValue(getCellData(claimSheetRow.getCell(10)));
				//Procedure code
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(9)).setCellValue(getCellData(claimSheetRow.getCell(13)));
				//modifier
				if(claimSheetRow.getCell(11) != null && claimSheetRow.getCell(11).getCellType() != CellType.BLANK) {
					String multipleModifer[]= getCellData(claimSheetRow.getCell(11)).split(",");
					for(int k=0;k<multipleModifer.length;k++) {
						if(k>4) break;
						setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(k+10)).setCellValue(multipleModifer[k]);
					}
				}
				//dx pointer
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(15)).setCellValue("01");
				//dx version
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(18)).setCellValue("ICD-10");
				//charges
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(16)).setCellValue(getAmountCellData(claimSheetRow.getCell(18)));
				//unit
				setCellToText(automationInputSheets[1],DataSheetlastRowNo.createCell(17)).setCellValue(getCellData(claimSheetRow.getCell(17)));
				try {
					if(claimSheet.getRow(i+1).getCell(1).getStringCellValue().equals("HCFA") || claimSheet.getRow(i+1).getCell(1).getStringCellValue().equals("UB")) break;

					else {
						System.out.println(claimSheet.getRow(i+1).getCell(1).getStringCellValue());
						DataSheet.createRow(DataSheet.getLastRowNum()+1);
					}
				}
				catch(NullPointerException e) {
					break;
				}
			}
		}
	}

	public static void writeInstData(HSSFSheet claimSheet, int startRowNo, int lastRowNo,XSSFSheet DataSheet) {


		for(int i=startRowNo;i <= lastRowNo; i++) {
			HSSFRow claimSheetRow = claimSheet.getRow(i);
			XSSFRow DataSheetlastRowNo= DataSheet.getRow(DataSheet.getLastRowNum());
			System.out.println("inst claim key:"+(int)claimSheetRow.getCell(0).getNumericCellValue()+"|"+"start row"+startRowNo);
			if(i == startRowNo) {
				DataSheetlastRowNo = DataSheet.createRow(DataSheet.getLastRowNum()+1);

				//Scenario name
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(0)).setCellValue(GUI.ScenarioName+new DecimalFormat("000").format((int)claimSheetRow.getCell(0).getNumericCellValue()));
				//receipt date
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(17)).setCellValue(getDateCellData(claimSheetRow.getCell(5)));
				//statement from date
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(18)).setCellValue(getDateCellData(claimSheetRow.getCell(6)));
				//statement to date
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(2)).setCellValue(getDateCellData(claimSheetRow.getCell(0)));
				//admission date
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(19)).setCellValue(getDateCellData(claimSheetRow.getCell(4)));
				//admission source
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(20)).setCellValue("7");
				//admission type
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(21)).setCellValue("7");
				//DX code
				String DXCode[] = getCellData(claimSheet.getRow(startRowNo).getCell(7)).split(",");
					for(int k=0;k<DXCode.length;k++) {
						setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(23+2*k)).setCellValue(DXCode[k]);
						setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(24+2*k)).setCellValue("Y");
					}
				//bill type
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(16)).setCellValue(getCellData(claimSheetRow.getCell(8)));
				//discharge status
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(22)).setCellValue("01");
				
			}
			else {
				//index number
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(5)).setCellValue(getCellData(claimSheet.getRow(i).getCell(9)));
				//revenue code
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(6)).setCellValue(getCellData(claimSheetRow.getCell(12)));
				//Procedure code
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(7)).setCellValue(getCellData(claimSheetRow.getCell(13)));
				//modifier
				if(claimSheetRow.getCell(11) != null && claimSheetRow.getCell(11).getCellType() != CellType.BLANK) {
					String multipleModifer[]= getCellData(claimSheetRow.getCell(11)).split(",");
					for(int k=0;k<multipleModifer.length;k++) {
						if(k>4) break;
						setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(k+8)).setCellValue(multipleModifer[k]);
					}
				}
				//service date
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(13)).setCellValue(getDateCellData(claimSheetRow.getCell(15)));
				//unit
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(14)).setCellValue(getCellData(claimSheetRow.getCell(17)));
				//charges
				setCellToText(automationInputSheets[0],DataSheetlastRowNo.createCell(15)).setCellValue(getAmountCellData(claimSheetRow.getCell(18)));
				
				
				try {
					if(claimSheet.getRow(i+1).getCell(1).getStringCellValue().equals("HCFA") || claimSheet.getRow(i+1).getCell(1).getStringCellValue().equals("UB")) break;

					else {
						System.out.println(claimSheet.getRow(i+1).getCell(1).getStringCellValue());
						DataSheet.createRow(DataSheet.getLastRowNum()+1);
					}
				}
				catch(NullPointerException e) {
					break;
				}
			}
		}
	}
}
