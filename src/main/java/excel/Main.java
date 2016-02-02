package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		args = new String[]{"test data.xlsx"};
		if(args.length < 1) {
			printHelp();
		}
		String fileName = args[0];
		File f = new File(fileName);
		if(!f.exists()) {
			System.err.println("Input file "+ f.getAbsolutePath()+" is not exists");
		}
		
		int sheetNo = 0;
		boolean isHeaderRow = true;
		for(int i=1; i<args.length; ++i) {
			String option = args[i];
			String nameValue[] = option.split("=");
			if(nameValue.length < 2) {
				System.err.println("Invalid option "+ option);
				printHelp() ;
			}
			try {
				if(nameValue[0].equalsIgnoreCase("sheetNumber")) {
					sheetNo = Integer.parseInt(nameValue[1]);
				} else if(nameValue[0].equalsIgnoreCase("headerRow")){
					isHeaderRow = Boolean.parseBoolean(nameValue[1]);
				}
			} catch(Exception e) {
				System.err.println("Invalid option "+ option);
				printHelp();
			}
		}
		
		try {
			InputStream is = new FileInputStream(f);
			Collection c = readExcel(is, sheetNo, isHeaderRow);
			writeExcel(new File(f.getParent(), "output-"+f.getName()+".xlsx"), c);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static void writeExcel(File f, Collection<MyRow> rows) {
		OutputStream os = null;
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Output");
			
			int rowNum = -1;
			Row header = sheet.createRow(++rowNum);
			header.createCell(0).setCellValue("FirstName");
			header.createCell(1).setCellValue("LastName");
			header.createCell(2).setCellValue("Email");
			header.createCell(3).setCellValue("Group");
			for(MyRow r : rows) {
				Row row = sheet.createRow(++rowNum);
				row.createCell(0).setCellValue(r.fistName);
				row.createCell(1).setCellValue(r.lastName);
				row.createCell(2).setCellValue(r.email);
				row.createCell(3).setCellValue(r.group);
			}
			
			os = new FileOutputStream(f);
			workbook.write(os);
		    os.close();
		    workbook.close();
		    System.out.println("Excel written successfully into file "+ f.getAbsolutePath());
		} catch(Exception e) {
			System.err.println("Errow while writing results into output file "+ f.getAbsolutePath());
			e.printStackTrace();
		} finally {
			if(os != null) {
				try {
					os.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
	
	
	private static Collection<MyRow> readExcel(InputStream is, int sheetNo, boolean isHeaderRow) {
		int rowNum = 0;
//		Map<String, MyRow> myrows = new LinkedHashMap<String, MyRow>();
		Map<String, Set<String>> emailGroups = new HashMap<String, Set<String>>();
		List<MyRow> myrows = new ArrayList<MyRow>();
		List<MyRow> rowsToReturn = new ArrayList<MyRow>();
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(sheetNo);
			
			Iterator<Row> rowIterator = sheet.iterator();
			if(isHeaderRow) {
				rowIterator.next();
			}
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				rowNum = row.getRowNum();
				String email = row.getCell(2).getStringCellValue();
//				if(!myrows.containsKey(email.toLowerCase())) {
					MyRow r = new MyRow();
					r.email = email;
					r.fistName = row.getCell(0).getStringCellValue();
					r.lastName = row.getCell(1).getStringCellValue();
					r.group = row.getCell(3).getStringCellValue();
					myrows.add(r);
					
					Set<String> groups = emailGroups.get(r.email);
					if(groups == null) {
						groups = new HashSet<String>();
						emailGroups.put(r.email, groups);
					}
					groups.add(r.group);
//				}
			}
			workbook.close();
			
			
			
			for(MyRow row : myrows) {
				Set<String> groups = emailGroups.get(row.email);
				if(groups == null || groups.size() <= 1 || ( groups.size() > 1 && !groups.contains("Non Procuring Realtor"))) {
					rowsToReturn.add(row);
				}
			}
		} catch(Exception e) {
			System.err.println("Error while parsing excel sheet row number "+ rowNum);
			e.printStackTrace();
		} finally {
			if(is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		System.out.println("excel parsed total rows "+ rowNum);
		return rowsToReturn;
	}
	
	private static void printHelp() {
		System.out.println("java -jar excel-dup-filter.jar excel-file-name [options]");
		System.out.println("Options:");
		System.out.println("	sheetNumber=0  (sheet number contains data to be filtered)");
		System.out.println("	headerRow=true or false (first row is the header. Specify either true or false. Default value true)");
		System.exit(-1);
	}
	

}


class MyRow {
	String fistName;
	String lastName;
	String email;
	String group;
}