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

public class EDI_InputSheet {
	static XSSFWorkbook Ediwbfile;
	public static XSSFWorkbook generateEdiInputSheet() throws IOException {

		java.io.InputStream EdiInputfis = EDI_InputSheet.class.getResourceAsStream("/EDI InputSheet.xlsx");
		Ediwbfile = new XSSFWorkbook(EdiInputfis);
		XSSFSheet DataSheet = Ediwbfile.getSheet("Data");
		clearSheetContent(DataSheet);

		FileInputStream UTCfis = new FileInputStream(GUI.file);
		HSSFWorkbook UTCfile = new HSSFWorkbook(UTCfis);
		HSSFSheet claimSheet = UTCfile.getSheet("claims");

		int lastRowNo = claimSheet.getPhysicalNumberOfRows()-1;

		//writing input data
		for(int i=1;i<lastRowNo;i++) {
			System.out.println(lastRowNo);
			System.out.println("started processing row :"+i);
			if(claimSheet.getRow(i).getCell(0).getCellType() != org.apache.poi.ss.usermodel.CellType.BLANK) {
				if(claimSheet.getRow(i).getCell(1).getStringCellValue().equals("HCFA")) writeProfData(claimSheet, i, lastRowNo, DataSheet);
				else writeInstData(claimSheet, i, lastRowNo, DataSheet);
			}
			UTCfile.close();

		}
		return Ediwbfile;
	}
	//removing existing rows to enter fresh data into the input sheet
	public static void clearSheetContent(XSSFSheet inputSheet) {
		for(int i=1;i <= inputSheet.getLastRowNum(); i++) {
			XSSFRow row = inputSheet.getRow(i);
			if(row != null) inputSheet.removeRow(row);
		}
	}
	//setting cell styles to text
	public static XSSFCell setCellToText(XSSFCell cell) {
		XSSFCellStyle textStyle = Ediwbfile.createCellStyle();
		XSSFDataFormat format = Ediwbfile.createDataFormat();
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
			case NUMERIC:return cell.getLocalDateTimeCellValue().format(java.time. format.DateTimeFormatter.ofPattern("yyyyMMdd"));
			case STRING:return new SimpleDateFormat("yyyyMMdd").format(new SimpleDateFormat("M/dd/yyyy").parse(cell.getStringCellValue()));
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
			case NUMERIC:return String.valueOf(((double)cell.getNumericCellValue())).replaceAll(" ","");
			case STRING: return cell.getStringCellValue().replaceAll("[^\\d]","");
			}
		}
		catch (NullPointerException e) {
			System.out.println("Skipped cell because cell is empty or unable to parse date format");
		}
		return null;
	}
	public static void writeProfData(HSSFSheet claimSheet, int startRowNo, int lastRowNo, XSSFSheet DataSheet) {
		XSSFRow DataSheetfirstRowNo= null;
		double totalAmount = 0;
		for(int i=startRowNo;i <= lastRowNo; i++) {
			HSSFRow claimSheetRow = claimSheet.getRow(i);
			XSSFRow DataSheetlastRowNo= DataSheet.getRow(DataSheet.getLastRowNum());
			System.out.println("prof claim key:"+(int)claimSheetRow.getCell(0).getNumericCellValue()+"|"+"start row"+startRowNo);
			if(i == startRowNo) {
				DataSheetlastRowNo = DataSheet.createRow(DataSheet.getLastRowNum()+1);
				DataSheetfirstRowNo = DataSheetlastRowNo;

				//Scenario name
				setCellToText(DataSheetlastRowNo.createCell(1)).setCellValue(GUI.ScenarioName+new DecimalFormat("000").format((int)claimSheetRow.getCell(0).getNumericCellValue()));
				//claim type/description
				setCellToText(DataSheetlastRowNo.createCell(2)).setCellValue("HCFA");
				//p/f
				setCellToText(DataSheetlastRowNo.createCell(3)).setCellValue("P");
				//Patient Account Number
				setCellToText(DataSheetlastRowNo.createCell(19)).setCellValue("PROF00"+new DecimalFormat("000").format((int)claimSheetRow.getCell(0).getNumericCellValue()));
				//DX code
				String DXCode[] = getCellData(claimSheet.getRow(startRowNo).getCell(7)).replace(".", "").split(",");
				setCellToText(DataSheetlastRowNo.createCell(9)).setCellValue(DXCode[0]);
				if(DXCode.length > 1) {
					for(int k=1;k<DXCode.length;k++) {
						setCellToText(DataSheetlastRowNo.getCell(9)).setCellValue(DataSheetlastRowNo.getCell(9)+"-"+DXCode[k]);
					}
				}
			}
			else {
				//index number
				setCellToText(DataSheetlastRowNo.createCell(0)).setCellValue(getCellData(claimSheet.getRow(i).getCell(9)));
				//line count
				setCellToText(DataSheetfirstRowNo.createCell(6)).setCellValue(getCellData(claimSheet.getRow(i).getCell(9)));
				//total amount
				//total amount
				totalAmount = totalAmount+Double.parseDouble(getAmountCellData(claimSheetRow.getCell(18)));
				setCellToText(DataSheetfirstRowNo.createCell(10)).setCellValue(String.format("%.2f",totalAmount));
				//from date
				setCellToText(DataSheetlastRowNo.createCell(7)).setCellValue(getDateCellData(claimSheetRow.getCell(15)));
				//service end date
				setCellToText(DataSheetlastRowNo.createCell(8)).setCellValue(getDateCellData(claimSheetRow.getCell(16)));
				//POS
				setCellToText(DataSheetlastRowNo.createCell(17)).setCellValue(getCellData(claimSheetRow.getCell(10)));
				//dos
				setCellToText(DataSheetlastRowNo.createCell(12)).setCellValue(getDateCellData(claimSheetRow.getCell(15)));
				//Procedure code & modifier
				setCellToText(DataSheetlastRowNo.createCell(13)).setCellValue(getCellData(claimSheetRow.getCell(13)));
				//modifier
				if(claimSheetRow.getCell(11) != null && claimSheetRow.getCell(11).getCellType() != CellType.BLANK) {
					String multipleModifer[]= getCellData(claimSheetRow.getCell(11)).split(",");
					for(int k=0;k<multipleModifer.length;k++) {
						if(k>4) break;
						setCellToText(DataSheetlastRowNo.getCell(13)).setCellValue(DataSheetlastRowNo.getCell(13).getStringCellValue()+":"+multipleModifer[k]);
					}
				}
				//charges
				setCellToText(DataSheetlastRowNo.createCell(15)).setCellValue(String.format("%.2f",Double.parseDouble(getAmountCellData(claimSheetRow.getCell(18)))));
				//unit
				setCellToText(DataSheetlastRowNo.createCell(16)).setCellValue(getCellData(claimSheetRow.getCell(17)));
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
		XSSFRow DataSheetfirstRowNo= null;
		double totalAmount = 0;
		for(int i=startRowNo;i <= lastRowNo;i++) {
			HSSFRow claimSheetRow = claimSheet.getRow(i);
			XSSFRow DataSheetlastRowNo= DataSheet.getRow(DataSheet.getLastRowNum());
			System.out.println("inst claim key : "+(int)claimSheetRow.getCell(0).getNumericCellValue()+" | "+"start row"+startRowNo);
			if(i == startRowNo) {
				DataSheetlastRowNo = DataSheet.createRow(DataSheet.getLastRowNum()+1);
				DataSheetfirstRowNo= DataSheetlastRowNo;

				//Scenario name
				setCellToText(DataSheetlastRowNo.createCell(1)).setCellValue(GUI.ScenarioName+new DecimalFormat("000").format((int)claimSheetRow.getCell(0).getNumericCellValue()));
				//claim type/description
				setCellToText(DataSheetlastRowNo.createCell(2)).setCellValue("UB");
				//p/f
				setCellToText(DataSheetlastRowNo.createCell(3)).setCellValue("F");
				//Patient Account Number
				setCellToText(DataSheetlastRowNo.createCell(19)).setCellValue("INST00"+new DecimalFormat("000").format((int)claimSheetRow.getCell(0).getNumericCellValue()));
				//Bill type
				setCellToText(DataSheetlastRowNo.createCell(18)).setCellValue(getCellData(claimSheetRow.getCell(8)));
				//DX code
				String DXCode[]=getCellData(claimSheet.getRow(startRowNo).getCell(7)).replace(".","").split(",");
				setCellToText(DataSheetlastRowNo.createCell(9)).setCellValue(DXCode[0]);
				if(DXCode.length > 1) {
					for(int k=1;k<DXCode.length;k++) {
						setCellToText(DataSheetlastRowNo.getCell(9)).setCellValue(DataSheetlastRowNo.getCell(9)+"-"+DXCode[k]);
					}
				}
			}
			else {
				//index number
				setCellToText(DataSheetlastRowNo.createCell(0)).setCellValue(getCellData(claimSheet.getRow(i).getCell(9)));
				//line count
				setCellToText(DataSheetfirstRowNo.createCell(6)).setCellValue(getCellData(claimSheet.getRow(i).getCell(9)));
				//from date
				setCellToText(DataSheetfirstRowNo.createCell(7)).setCellValue(getDateCellData(claimSheetRow.getCell(15)));
				//service end date
				setCellToText(DataSheetfirstRowNo.createCell(8)).setCellValue(getDateCellData(claimSheetRow.getCell(16)));
				//total amount
				totalAmount = totalAmount+Double.parseDouble(getAmountCellData(claimSheetRow.getCell(18)));
				setCellToText(DataSheetfirstRowNo.createCell(10)).setCellValue(String.format("%.2f",totalAmount));
				//RevenueCode
				setCellToText(DataSheetlastRowNo.createCell(14)).setCellValue(getCellData(claimSheetRow.getCell(12)));
				//Procedure code & modifier
				setCellToText(DataSheetlastRowNo.createCell(13)).setCellValue(getCellData(claimSheetRow.getCell(13)));
				//dos
				setCellToText(DataSheetlastRowNo.createCell(12)).setCellValue(getDateCellData(claimSheetRow.getCell(15)));
				//modifier
				if(claimSheetRow.getCell(11) != null && claimSheetRow.getCell(11).getCellType() != CellType.BLANK) {
					String multipleModifer[]= getCellData(claimSheetRow.getCell(11)).split(",");
					for(int k=0;k<multipleModifer.length;k++) {
						if(k>4) break;
						setCellToText(DataSheetlastRowNo.getCell(13)).setCellValue(DataSheetlastRowNo.getCell(13).getStringCellValue()+":"+multipleModifer[k]);
					}
				}
				//charges
				setCellToText(DataSheetlastRowNo.createCell(15)).setCellValue(String.format("%.2f",Double.parseDouble(getAmountCellData(claimSheetRow.getCell(18)))));
				//unit
				setCellToText(DataSheetlastRowNo.createCell(16)).setCellValue(getCellData(claimSheetRow.getCell(17)));
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
