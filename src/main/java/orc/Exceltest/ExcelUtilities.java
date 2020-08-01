package orc.Exceltest;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.logging.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

public class ExcelUtilities {

	java.util.logging.Logger logger = java.util.logging.Logger.getLogger(ExcelUtilities.class.getName());

	private static final String CELL_DATA_DOES_NOT_MATCH = "is updated to";
	private static final String CELL_DATA_MATCH = "is equal to";

	/***
	 *
	 * Method to Compare Excel Sheets
	 * 
	 * @param dataMap
	 * @throws Exception
	 */

	public void compareExcelWorkbooks(HashMap<String, String> dataMap) throws Exception {

		String fileName = dataMap.get("fileName");
		String logFilePath = dataMap.get("fileName") + ".txt";

		if (fileName.matches("^[A-Z]:.*")) {
			System.out.println("Changing the logFile Path to " + fileName);
			logFilePath = fileName + ".txt";
		}
		// Setting up the logger
		logger.setLevel(Level.FINE);
		FileHandler filehandle = new FileHandler(logFilePath);
		filehandle.setFormatter(new customHandler());
		logger.addHandler(filehandle);

		System.out.println("Location of Log File" + logFilePath);

		// checking if Source and Target File are Present
		System.out.println(dataMap.get("expectedFilePath"));
		System.out.println(dataMap.get("actualFilePath"));

		File wrkbk1 = new File(dataMap.get("expectedFilePath"));
		File wrkbk2 = new File(dataMap.get("actualFilePath"));
		if (!(wrkbk1.exists() && wrkbk2.exists())) {
			addMessage("Expected File at " + dataMap.get("expectedFilePath") + " is present :" + wrkbk1.exists());
			addMessage("Actual File at " + dataMap.get("actualFilePath") + " is present :" + wrkbk2.exists());
			throw new Exception("Unable to find the Excel files to Compare");
		}

		// Starting of Comparision
		try {
			Workbook wb1 = WorkbookFactory.create(new FileInputStream(wrkbk1));
			Workbook wb2 = WorkbookFactory.create(new FileInputStream(wrkbk2));

			compare(wb1, wb2, dataMap);

		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("Is log File Created " + (new File(logFilePath)).exists());
		// Creating Report
		new ExcelHandler().CreateReport(dataMap);

	}

	public List<String> validateExcelCondition(HashMap<String, String> dataMap) throws Exception {

		try {

			Workbook wb1 = WorkbookFactory.create(new FileInputStream(new File(dataMap.get("excelFilePath"))));
			return checkForMismatchesinWb(wb1, dataMap);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;

	}

	private static class Locator {
		Workbook workbook;
		Sheet sheet;
		org.apache.poi.ss.usermodel.Row row;
		Cell cell;
	}

	private static class customHandler extends SimpleFormatter {

		public String format(LogRecord record) {
			return record.getMessage() + "\n";
		}
	}

	private class Columns {
		HashMap<String, Integer> sheet1 = new HashMap<String, Integer>();
		HashMap<String, Integer> sheet2 = new HashMap<String, Integer>();
		Set<String> common = new HashSet<String>();
		Set<String> extraInSheet1 = new HashSet<String>();
		Set<String> extraInSheet2 = new HashSet<String>();
		List<String> primaryColumns;
		boolean isPrimaryKeysProvided = false;
		int startingRowNum = 0;

		Columns(Locator loc1, Locator loc2) {

			// code to get the starting point
			loopOut: for (int i = 0; i < loc1.sheet.getLastRowNum(); i++) {
				System.out.println("Running for " + i);
				for (int j = 0; j < loc1.sheet.getRow(i).getLastCellNum(); j++) {
					System.out.println("running for J" + j);
					System.out.println(loc1.sheet.getRow(i).getCell(j));
					if (!isCellEmpty(loc1.sheet.getRow(i).getCell(j))) {
						if (!loc1.sheet.getRow(i).getCell(j).getStringCellValue().trim().isEmpty()) {
							this.startingRowNum = i;
							System.out.println("Staring Row Num " + startingRowNum);
							break loopOut;
						}

					}
				}
			}

			for (int i = 0; i < loc1.sheet.getRow(this.startingRowNum).getLastCellNum(); i++) {
				if (!isCellEmpty(loc1.sheet.getRow(this.startingRowNum).getCell(i)))
					sheet1.put(loc1.sheet.getRow(this.startingRowNum).getCell(i).getStringCellValue(), i);
			}
			for (int i = 0; i < loc2.sheet.getRow(this.startingRowNum).getLastCellNum(); i++) {
				if (!isCellEmpty(loc2.sheet.getRow(this.startingRowNum).getCell(i)))
					sheet2.put(loc2.sheet.getRow(this.startingRowNum).getCell(i).getStringCellValue(), i);

			}

			for (String column : sheet1.keySet()) {
				if (sheet2.containsKey(column)) {
					common.add(column);
				} else {
					extraInSheet1.add(column);
				}
			}

			for (String column : sheet2.keySet()) {
				if (!sheet1.containsKey(column)) {
					extraInSheet2.add(column);
				}
			}

			primaryColumns = new ArrayList<String>(common);

			/*
			 * System.out.println("Extra in sheet 1" + extraInSheet1);
			 * System.out.println("Extra in sheet 2" + extraInSheet2);
			 */

		}

		Columns(Locator loc1, HashMap<String, String> sheetParameters) {

			List<String> columnsToExcludeArrayList = new ArrayList<String>();
			if (sheetParameters.containsKey("columnsToExclude")) {
				columnsToExcludeArrayList = Arrays.asList(sheetParameters.get("columnsToExclude").split(";"));
				System.out.println("Columns to Exculde " + columnsToExcludeArrayList);
			}

			for (int i = 0; i < loc1.sheet.getRow(0).getLastCellNum(); i++) {
				Cell cell = loc1.sheet.getRow(0).getCell(i);
				if (isCellEmpty(cell))
					continue;

				String columnName = getStringCellValue(cell);
				if (!columnsToExcludeArrayList.contains(columnName))
					sheet2.put(columnName, i);
			}

		}

		Columns(Locator loc1, Locator loc2, HashMap<String, String> sheetParameters) {

			List<String> columnsToExcludeArrayList = new ArrayList<String>();
			if (sheetParameters.containsKey("columnsToExclude")) {
				columnsToExcludeArrayList = Arrays.asList(sheetParameters.get("columnsToExclude").split(";"));
				// System.out.println("Columns to Exculde "+columnsToExcludeArrayList);
			}
			if (sheetParameters.containsKey("primaryColumns")) {
				isPrimaryKeysProvided = true;
				primaryColumns = Arrays.asList(sheetParameters.get("primaryColumns").split(";"));
			}

			for (int i = 0; i < loc1.sheet.getRow(0).getLastCellNum(); i++) {
				Cell cell = loc1.sheet.getRow(0).getCell(i);
				if (isCellEmpty(cell))
					continue;

				String columnName = getStringCellValue(cell);
				if (!columnsToExcludeArrayList.contains(columnName))
					sheet1.put(columnName, i);
			}
			for (int i = 0; i < loc2.sheet.getRow(0).getLastCellNum(); i++) {
				Cell cell = loc2.sheet.getRow(0).getCell(i);
				if (isCellEmpty(cell))
					continue;

				String columnName = getStringCellValue(cell);
				if (!columnsToExcludeArrayList.contains(columnName))
					sheet2.put(columnName, i);

			}

			for (String column : sheet1.keySet()) {
				if (sheet2.containsKey(column)) {
					common.add(column);
				} else {
					extraInSheet1.add(column);
				}
			}

			for (String column : sheet2.keySet()) {
				if (!sheet1.containsKey(column)) {
					extraInSheet2.add(column);
				}
			}

			// System.out.println("Extra in sheet 1" + extraInSheet1);
			// System.out.println("Extra in sheet 2" + extraInSheet2);

		}

	}

	/**
	 * Utility to compare Excel File Contents cell by cell for all sheets.
	 *
	 * @param wb1
	 *            the workbook1
	 * @param wb2
	 *            the workbook2
	 * @return the Excel file difference containing a flag and a list of differences
	 */
	public static void compare(Workbook wb1, Workbook wb2, HashMap<String, String> parameters) {
		Locator loc1 = new Locator();
		Locator loc2 = new Locator();
		loc1.workbook = wb1;
		loc2.workbook = wb2;

		ExcelUtilities excelComparator = new ExcelUtilities();
		excelComparator.compareSheetData(loc1, loc2, parameters);

	}

	/**
	 * Compare sheet data.
	 */
	private void compareSheetData(Locator loc1, Locator loc2, HashMap<String, String> parameters) {
		compareDataInAllSheets(loc1, loc2, parameters);

	}

	/**
	 * Compare data in all sheets.
	 */
	private void compareDataInAllSheets(Locator loc1, Locator loc2, HashMap<String, String> parameters) {

		addMessage("Start comparision of Workbook");
		for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {

			loc1.sheet = loc1.workbook.getSheetAt(i);
			int sheetNumber = getSheetWithName(loc2, loc1.sheet.getSheetName());
			if (sheetNumber == -1) {
				logger.log(Level.FINE, "Sheet with Name " + loc1.sheet.getSheetName() + " could not be found");
				continue;
			}

			loc2.sheet = loc2.workbook.getSheetAt(sheetNumber);

			// System.out.println("Sheet Name "+loc2.sheet.getSheetName()+" is parameter
			// sheet Present "+parameters.containsKey(loc2.sheet.getSheetName()+" keySet
			// "+parameters.keySet()));
			// code to get parameters of a Sheet
			if (parameters.containsKey(loc2.sheet.getSheetName())) {

				HashMap<String, String> sheetParameters = getParameterHashMapFromString(
						parameters.get(loc2.sheet.getSheetName()),
						"primaryColumns;columnsToExclude;filters".split(";"));
				compareDataInSheet(loc1, loc2, sheetParameters);
			} else
				compareDataInSheet(loc1, loc2);
		}

		addMessage("End Comparision of Workbook");
	}

	private void compareDataInSheet(Locator loc1, Locator loc2, HashMap<String, String> sheetParameters) {

		addMessage("Starting Comparision of Sheet " + loc1.sheet.getSheetName());

		Columns sheetColumns = new Columns(loc1, loc2, sheetParameters);

		if (sheetParameters.containsKey("filters") && !sheetParameters.get("filters").isEmpty())
			applyFilters(loc2, sheetColumns, sheetParameters.get("filters"));

		for (int j = 0; j <= loc1.sheet.getLastRowNum(); j++) {

			loc1.row = loc1.sheet.getRow(j);
			if (isRowEmpty(loc1)) {
				continue;
			}

			int rowNum = getClosestPossibleRow(loc1, loc2, sheetColumns);
			if (rowNum == -1) {
				isExtraRowinSheet(loc1, 1, sheetColumns);
			} else {
				loc2.row = loc2.sheet.getRow(rowNum);

				if ((loc1.row == null) || (loc2.row == null)) {
					continue;
				}

				compareDataInRow(loc1, loc2, sheetColumns, true);

				removeRow(loc2, rowNum, -1);
			}

		}
		for (int k = 0; k <= loc2.sheet.getLastRowNum(); k++) {

			loc2.row = loc2.sheet.getRow(k);

			if (!isRowEmpty(loc2))
				isExtraRowinSheet(loc2, 2, sheetColumns);

		}

		addMessage("End Comparion of Sheet");
	}

	private void applyFilters(Locator loc, Columns column, String filterString) {

		String[][] filter = getFilterArray(filterString);
		for (int i = loc.sheet.getFirstRowNum() + 1; i <= loc.sheet.getLastRowNum(); i++) {
			loc.row = loc.sheet.getRow(i);
			// System.out.println("Should Row be Filtered"+shouldRowBeFiltered(loc, column,
			// filter));
			if (isRowEmpty(loc))
				continue;
			else if (!shouldRowBeFiltered(loc, column, filter)) {
				addMessage("Row " + i + " is Filtered");
				removeRow(loc, i, -1);
			}
		}
	}

	private boolean shouldRowBeFiltered(Locator loc, Columns column, String[][] Filter) {

		for (String[] filterData : Filter) {

			Cell cell = loc.row.getCell(column.sheet2.get(filterData[0].trim()));

			// System.out.println("Should cell be Filtered"+shouldCellBeFiltered(cell,
			// filterData[1], filterData[2]));

			if (isCellEmpty(cell))
				continue;
			else if (!shouldCellBeFiltered(cell, filterData[1], filterData[2]))
				return false;

		}

		return true;
	}

	private boolean shouldCellBeFiltered(Cell cell, String operator, String value) {
		String cellValue = getStringCellValue(cell);

		boolean result = false;
		if (cellValue.contains("-->")) {
			if (cellValue.trim().equalsIgnoreCase("-->")) {
				result = false;
			} else if (Arrays.asList(value.split("\\|\\|")).contains(cellValue.split("-->")[0])
					|| Arrays.asList(value.split("\\|\\|")).contains(cellValue.split("-->")[1])) {
				result = true;
			}
		} else {
			result = Arrays.asList(value.split("\\|\\|")).contains(cellValue);
		}

		
		if (operator.equalsIgnoreCase("!="))
			return (!result);

		return result;
	}

	private String getStringCellValue(Cell cell) {
		if (cell.getCellType() == CellType.BLANK) {
			return "BLANK";
		} else if (cell.getCellType() == CellType.NUMERIC) {
			return Double.toString(cell.getNumericCellValue());
		} else {
			return cell.getStringCellValue();
		}
	}

	private String[][] getFilterArray(String filter) {
		String[] filterArray = filter.split(";");
		String[][] filterParams = new String[filterArray.length][3];

		for (int i = 0; i < filterArray.length; i++) {
			if (filterArray[i].contains("!=")) {
				filterParams[i][0] = filterArray[i].split("!=")[0];
				filterParams[i][1] = "!=";
				filterParams[i][2] = filterArray[i].split("!=")[1];
			} else if (filterArray[i].contains("=")) {
				filterParams[i][0] = filterArray[i].split("=")[0];
				filterParams[i][1] = "=";
				filterParams[i][2] = filterArray[i].split("=")[1];
			}
		}

		return filterParams;
	}

	private void compareDataInSheet(Locator loc1, Locator loc2) {

		addMessage("Starting Comparision of Sheet " + loc1.sheet.getSheetName());

		Columns sheetColumns = new Columns(loc1, loc2);
		for (int j = 0; j <= loc1.sheet.getLastRowNum(); j++) {

			loc1.row = loc1.sheet.getRow(j);
			if (isRowEmpty(loc1)) {
				continue;
			}
			// loc2.row = loc2.sheet.getRow(j);
			int rowNum = getClosestPossibleRow(loc1, loc2, sheetColumns);
			if (rowNum == -1) {
				isExtraRowinSheet(loc1, 1, sheetColumns);
			} else {
				loc2.row = loc2.sheet.getRow(rowNum);

				if ((loc1.row == null) || (loc2.row == null)) {
					continue;
				}

				compareDataInRow(loc1, loc2, sheetColumns, true);

				removeRow(loc2, rowNum, -1);
			}

		}
		// System.out.println("Loc 2 Last " + loc2.sheet.getLastRowNum());
		for (int k = 0; k <= loc2.sheet.getLastRowNum(); k++) {

			loc2.row = loc2.sheet.getRow(k);
			// System.out.println("K value" + k + isRowEmpty(loc2));
			if (!isRowEmpty(loc2))
				isExtraRowinSheet(loc2, 2, sheetColumns);
		}

		addMessage("End Comparion of Sheet");
	}

	private int getClosestPossibleRow(Locator loc1, Locator loc2, Columns column) {

		double maxPercentage = 0;
		int rowNum = -1;

		for (int j = 0; j <= loc2.sheet.getLastRowNum(); j++) {
			loc2.row = loc2.sheet.getRow(j);
			double comparePercentage = getPercentageOfMatch(loc1, loc2, column);
			// addMessage("Percentage "+maxPercentage+" row Num"+rowNum+" j "+j);
			if (comparePercentage == 100) {
				maxPercentage = comparePercentage;
				rowNum = j;
				break;
			}
			if (maxPercentage < comparePercentage && comparePercentage > 50) {
				maxPercentage = comparePercentage;
				rowNum = j;

			}

		}

		if (column.isPrimaryKeysProvided && maxPercentage != 100)
			return -1;

		return rowNum;
	}

	private double getPercentageOfMatch(Locator loc1, Locator loc2, Columns column) {

		int percentage = 0;

		if (isRowEmpty(loc2)) {
			return 0;
		}
		for (String columnName : column.primaryColumns) {
			if (column.sheet1.containsKey(columnName) && column.sheet2.containsKey(columnName)) {
				loc1.cell = loc1.row.getCell(column.sheet1.get(columnName));
				loc2.cell = loc2.row.getCell(column.sheet2.get(columnName));

				if ((loc1.cell == null) || (loc2.cell == null)) {
					continue;
				}

				if (compareDataInCell(loc1, loc2, false) == 0) {
					percentage++;

				}
				;
			} else
				addMessage("Column: " + columnName + " Found in Workbook1 " + column.sheet1.containsKey(columnName)
						+ " and Found in Workbook2 " + column.sheet2.containsKey(columnName));
		}
		return (percentage * 100) / column.primaryColumns.size();
	}

	private double compareDataInRow(Locator loc1, Locator loc2, Columns column, boolean shouldLog) {

		int percentage = 0;

		if (loc1.row == null || loc2.row == null) {
			return 0;
		}

		if (shouldLog)
			logger.log(Level.FINE, "Starting row comparision");
		for (String columnName : column.common) {

			loc1.cell = loc1.row.getCell(column.sheet1.get(columnName));
			loc2.cell = loc2.row.getCell(column.sheet2.get(columnName));

			if (loc1.cell == null && loc2.cell != null) {
				isExtracellinSheet(loc2, 2, shouldLog);
				continue;
			} else if (loc1.cell != null && loc2.cell == null) {
				isExtracellinSheet(loc1, 1, shouldLog);
				continue;
			} else if ((loc1.cell == null) && (loc2.cell == null)) {
				addMessage("workbook1 ->" + column.sheet1.get(columnName) + " [ ] is equal to workbook2 -> "
						+ column.sheet2.get(columnName) + "[ ]");
				continue;
			}

			if (compareDataInCell(loc1, loc2, shouldLog) == 0) {
				percentage++;
			}
			;
		}

		for (String extraColumn : column.extraInSheet1) {
			// code to handle Extra column in sheet 1
			loc1.cell = loc1.row.getCell(column.sheet1.get(extraColumn));
			if ((loc1.cell == null) || (loc2.cell == null)) {
				continue;
			}
			isExtracellinSheet(loc1, 1, shouldLog);
		}

		for (String extraColumn : column.extraInSheet2) {
			// code to handle Extra column in sheet 2

			loc2.cell = loc2.row.getCell(column.sheet2.get(extraColumn));
			if ((loc1.cell == null) || (loc2.cell == null)) {
				continue;
			}
			isExtracellinSheet(loc2, 2, shouldLog);
		}
		if (shouldLog) {
			logger.log(Level.FINE, "Ending row comparision");
		}
		return (percentage * 100) / loc1.row.getLastCellNum();
	}

	private void isExtraRowinSheet(Locator loc, int workbookNum, Columns columns) {

		if (isRowEmpty(loc))
			return;

		logger.log(Level.FINE, "Starting row comparision");

		for (String columnName : columns.common) {
			int columnNumber;
			if (workbookNum == 1) {
				columnNumber = columns.sheet1.get(columnName);
			} else
				columnNumber = columns.sheet2.get(columnName);

			loc.cell = loc.row.getCell(columnNumber);
			isExtracellinSheet(loc, workbookNum, true);
		}

		for (String columnName : columns.extraInSheet1) {
			if (workbookNum == 2) {
				addMessage("workbook1 " + columnName + "-> [ ] is removed");
			} else {
				loc.cell = loc.row.getCell(columns.sheet1.get(columnName));
				isExtracellinSheet(loc, workbookNum, true);
			}
		}
		for (String columnName : columns.extraInSheet2) {
			if (workbookNum == 1) {
				addMessage("workbook2 " + columnName + "-> [ ] is added");
			} else {
				loc.cell = loc.row.getCell(columns.sheet2.get(columnName));
				isExtracellinSheet(loc, workbookNum, true);
			}
		}

		logger.log(Level.FINE, "Ending row comparision");
	}

	private void isExtracellinSheet(Locator loc, int workbookNum, boolean shouldLog) {

		String str;
		String decide = null;
		if (workbookNum == 1) {
			decide = "removed";
		} else if (workbookNum == 2) {
			decide = "added";
		}
		if (isCellEmpty(loc.cell)) {

			str = String.format(Locale.ROOT, "workbook%d -> [ ] is %s", workbookNum, decide);

		} else {

			str = String.format(Locale.ROOT, "workbook%d %s-> [%s] is %s", workbookNum,
					new CellReference(loc.cell).formatAsString(), getStringCellValue(loc.cell), decide);
		}

		addMessage(str, shouldLog);
	}

	private int compareDataInCell(Locator loc1, Locator loc2, boolean shouldLog) {
		if (isCellTypeMatches(loc1, loc2, shouldLog)) {
			CellType loc1cellType = loc1.cell.getCellType();
			switch (loc1cellType) {
			case BLANK:
			case STRING:
			case ERROR:
				return isCellContentMatches(loc1, loc2, shouldLog);

			case BOOLEAN:
				return isCellContentMatchesForBoolean(loc1, loc2, shouldLog);

			case FORMULA:
				return isCellContentMatchesForFormula(loc1, loc2, shouldLog);

			case NUMERIC:
				if (DateUtil.isCellDateFormatted(loc1.cell)) {
					return isCellContentMatchesForDate(loc1, loc2, shouldLog);
				} else {
					return isCellContentMatchesForNumeric(loc1, loc2, shouldLog);
				}

			default:
				throw new IllegalStateException("Unexpected cell type: " + loc1cellType + " Row Num "
						+ loc1.row.getRowNum() + "cell " + loc1.cell.getColumnIndex());
			}
		}

		return 1;
	}

	/**
	 * Checks if cell content matches.
	 */
	private int isCellContentMatches(Locator loc1, Locator loc2, boolean shouldLog) {
		System.out.println("Cell value " + loc1.cell.toString().trim());
		String str1 = loc1.cell.toString().trim();
		String str2 = loc2.cell.toString().trim();
		System.out.println("cell Value after replace " + str1);
		if (!str1.equalsIgnoreCase(str2)) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, str1, str2, shouldLog);
			return 1;
		}
		addMessage(loc1, loc2, CELL_DATA_MATCH, str1, str2, shouldLog);
		return 0;
	}

	/**
	 * Checks if cell content matches for boolean.
	 */
	private int isCellContentMatchesForBoolean(Locator loc1, Locator loc2, boolean shouldLog) {
		boolean b1 = loc1.cell.getBooleanCellValue();
		boolean b2 = loc2.cell.getBooleanCellValue();
		if (b1 != b2) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Boolean.toString(b1), Boolean.toString(b2), shouldLog);
			return 1;
		}
		addMessage(loc1, loc2, CELL_DATA_MATCH, Boolean.toString(b1), Boolean.toString(b2), shouldLog);
		return 0;
	}

	/**
	 * Checks if cell content matches for date.
	 */
	private int isCellContentMatchesForDate(Locator loc1, Locator loc2, boolean shouldLog) {
		Date date1 = loc1.cell.getDateCellValue();
		Date date2 = loc2.cell.getDateCellValue();
		if (!date1.equals(date2)) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, date1.toString(), date2.toString(), shouldLog);
			return 1;
		}
		addMessage(loc1, loc2, CELL_DATA_MATCH, date1.toString(), date2.toString(), shouldLog);
		return 0;
	}

	/**
	 * Checks if cell content matches for formula.
	 */
	private int isCellContentMatchesForFormula(Locator loc1, Locator loc2, boolean shouldLog) {

		String form1 = loc1.cell.getCellFormula();
		String form2 = loc2.cell.getCellFormula();
		if (!form1.equals(form2)) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, form1, form2, shouldLog);
			return 1;
		}
		addMessage(loc1, loc2, CELL_DATA_MATCH, form1, form2, shouldLog);
		return 0;
	}

	/**
	 * Checks if cell content matches for numeric.
	 */
	private int isCellContentMatchesForNumeric(Locator loc1, Locator loc2, boolean shouldLog) {

		double num1 = loc1.cell.getNumericCellValue();
		double num2 = loc2.cell.getNumericCellValue();
		if (num1 != num2) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2), shouldLog);
			return 1;
		}
		addMessage(loc1, loc2, CELL_DATA_MATCH, Double.toString(num1), Double.toString(num2), shouldLog);
		return 0;
	}

	/**
	 * Checks if cell type matches.
	 */
	private boolean isCellTypeMatches(Locator loc1, Locator loc2, boolean shouldLog) {
		CellType type1 = loc1.cell.getCellType();
		CellType type2 = loc2.cell.getCellType();
		if (type1 == type2) {
			return true;
		}
		String cellValue1 = String.valueOf(type1), cellValue2 = String.valueOf(type2);
		cellValue1 = getStringCellValue(loc1.cell);
		cellValue2 = getStringCellValue(loc2.cell);
		if (cellValue1.equalsIgnoreCase("blank") && !cellValue2.equalsIgnoreCase("blank")) {
			isExtracellinSheet(loc2, 2, true);
		} else if (!cellValue1.equalsIgnoreCase("blank") && cellValue2.equalsIgnoreCase("blank")) {
			isExtracellinSheet(loc1, 1, true);
		} else {
			addMessage(loc1, loc2, "Cell Data-Type does not Match in :: ", cellValue1, cellValue2, shouldLog);
		}

		return false;
	}

	private void removeRow(Locator loc, int rowIndex, int shiftBy) {
		// int lastRowNum = loc.sheet.getLastRowNum();

		Row removingRow = loc.sheet.getRow(rowIndex);
		if (removingRow != null) {
			loc.sheet.removeRow(removingRow);
		}
		

	}

	/**
	 * Helper Methods
	 *
	 * @param messageStart
	 * @param shouldLog
	 */

	private void addMessage(String messageStart, boolean shouldLog) {
		if (shouldLog)
			addMessage(messageStart);
	}

	private void addMessage(String messageString) {
		logger.log(Level.FINE, messageString);
	}

	/**
	 * Formats the message.
	 */
	private void addMessage(Locator loc1, Locator loc2, String messageStart, String value1, String value2) {
		String str = String.format(Locale.ROOT, "workbook1 -> %s -> %s [%s] %s workbook2 -> %s -> %s [%s]",
				loc1.sheet.getSheetName(), new CellReference(loc1.cell).formatAsString(), value1, messageStart,
				loc2.sheet.getSheetName(), new CellReference(loc2.cell).formatAsString(), value2);
		logger.log(Level.FINE, str);

	}

	private void addMessage(Locator loc1, Locator loc2, String messageStart, String value1, String value2,
			boolean shouldLog) {
		if (shouldLog) {
			addMessage(loc1, loc2, messageStart, value1, value2);
		}
	}

	private boolean isRowEmpty(Locator loc) {
		if (loc.row == null) {
			return true;
		}
		for (int i = 0; i < loc.row.getLastCellNum(); i++) {
			Cell cell = loc.row.getCell(i);

			if (!isCellEmpty(cell))
				return false;
		}

		return true;
	}

	private boolean isCellEmpty(Cell cell) {

		if (cell != null && cell.getCellType() != CellType.BLANK)
			return false;

		return true;
	}

	private int getSheetWithName(Locator loc, String sheetName) {

		for (int i = 0; i < loc.workbook.getNumberOfSheets(); i++) {
			if (loc.workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
				return i;
			}
		}

		return -1;
	}

	private HashMap<String, String> getParameterHashMapFromString(String parameterString, String[] keys) {
		HashMap<String, String> dataMap = new HashMap<String, String>();
		// String[] keys= {"primaryColumns","columnsToExclude","filters"};
		for (String key : keys) {
			if (parameterString.contains(key)) {

				String splitString = parameterString.split(key + "=")[1];
				int index = splitString.indexOf("&&");
				if (index == -1) {
					dataMap.put(key, splitString);
				} else {
					dataMap.put(key, splitString.substring(0, index));
				}

			}
		}

		for (String key : dataMap.keySet()) {
			System.out.println("Key " + key + " value " + dataMap.get(key));
		}

		return dataMap;
	}

	private static List<String> checkForMismatchesinWb(Workbook wb, HashMap<String, String> parameters)
			throws Exception {
		Locator loc1 = new Locator();
		loc1.workbook = wb;

		ExcelUtilities excelComparator = new ExcelUtilities();
		return excelComparator.checkForMismatchesInSheets(loc1, parameters);
	}

	private List<String> checkForMismatchesInSheets(Locator loc1, HashMap<String, String> parameters) throws Exception {

		List<String> logs = new ArrayList<String>();

		for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {

			loc1.sheet = loc1.workbook.getSheetAt(i);
			if (parameters.containsKey(loc1.sheet.getSheetName())) {

				String sheetParams = parameters.get(loc1.sheet.getSheetName());
				HashMap<String, String> sheetParameters = getParameterHashMapFromString(sheetParams,
						"filters;columnsToCheck".split(";"));
				logs.add("Starting Comparision of Sheet " + loc1.sheet.getSheetName());
				logs.addAll(checkforMismatchInSheet(loc1, sheetParameters));
				logs.add("Completed Comparision of Sheet " + loc1.sheet.getSheetName());
			}

		}

		return logs;

	}

	private List<String> checkforMismatchInSheet(Locator loc1, HashMap<String, String> sheetParameters)
			throws Exception {
		Columns sheetColumns = new Columns(loc1, sheetParameters);
		List<String> logs = new ArrayList<String>();
		if (sheetParameters.containsKey("filters") && !sheetParameters.get("filters").isEmpty())
			applyFilters(loc1, sheetColumns, sheetParameters.get("filters"));

		for (int j = 1; j <= loc1.sheet.getLastRowNum(); j++) {

			loc1.row = loc1.sheet.getRow(j);
			if (isRowEmpty(loc1)) {
				continue;
			}

			logs.add("Starting Comparision of Row " + j);
			logs.addAll(checkforMismatchInRow(loc1, sheetColumns, sheetParameters));
			logs.add("Completed Comparision of Row " + j);
		}
		return logs;
	}

	private List<String> checkforMismatchInRow(Locator loc, Columns columns, HashMap<String, String> sheetParameters)
			throws Exception {

		List<Integer> columnsToCheck = new ArrayList<Integer>();
		List<String> logs = new ArrayList<String>();
		if (sheetParameters.containsKey("columnsToCheck")) {

			for (String columnName : sheetParameters.get("columnsToCheck").split(";"))
				columnsToCheck.add(columns.sheet2.get(columnName));

		} else {
			columnsToCheck.addAll(columns.sheet2.values());
		}

		for (Integer columnNumber : columnsToCheck) {

			loc.cell = loc.row.getCell(columnNumber);

			logs.add(checkforMismatchInCell(loc, columns));

		}

		return logs;

	}

	private String checkforMismatchInCell(Locator loc, Columns columns) throws Exception {

		// this code to show column Names in log is not needed will effect performance

		int colNumber = loc.cell.getColumnIndex();
		String columnName = "";
		for (String key : columns.sheet2.keySet()) {
			if (columns.sheet2.get(key) == colNumber) {
				columnName = key;
				break;
			}
		}

		if (isCellAdded(loc)) {
			return columnName + ": " + getStringCellValue(loc.cell) + " is added";
		} else if (isCellRemoved(loc)) {
			return columnName + ": " + getStringCellValue(loc.cell) + " is Removed";
		} else if (isCellMisc(loc)) {
			return columnName + ": " + getStringCellValue(loc.cell) + " is Misc";
		}

		else if (loc.cell.getCellType() == CellType.STRING) {

			if (loc.cell.getStringCellValue().contains("-->")) {
				return columnName + ": " + getStringCellValue(loc.cell) + " is Modified";
			}
		}

		return columnName + ": " + getStringCellValue(loc.cell) + ": Matched";

	}

	private boolean isCellAdded(Locator loc) {
		CellStyle cs = loc.cell.getCellStyle();
		if (cs.getFillForegroundColor() == IndexedColors.GREEN.index)
			return true;
		return false;
	}

	private boolean isCellRemoved(Locator loc) {
		CellStyle cs = loc.cell.getCellStyle();
		if (cs.getFillForegroundColor() == IndexedColors.ORANGE.index)
			return true;
		return false;
	}

	private boolean isCellMisc(Locator loc) {
		CellStyle cs = loc.cell.getCellStyle();
		if (cs.getFillForegroundColor() == IndexedColors.RED.index)
			return true;
		return false;
	}

}

class ExcelHandler {

	CellStyle cellupdated;
	CellStyle cellAdded;
	CellStyle cellRemoved;
	CellStyle cellMisc;

	public void CreateReport(HashMap<String, String> dataMap) throws Exception {

		String fileName = dataMap.get("fileName");
		String excelFilePath = "";
		String logFilePath = "";
		excelFilePath =  dataMap.get("fileName") + ".xlsx";
		logFilePath = dataMap.get("fileName") + ".txt";
		System.out.println("Excel Output File is saved to " + excelFilePath);
		if (fileName.matches("^[A-Z]:.*")) {
			System.out.println("File paths are changed to folder Paths");
			excelFilePath = fileName + ".xlsx";
			logFilePath = fileName + ".txt";
		}

		File file = new File(logFilePath);
		if (!file.exists()) {
			throw new Exception("Unable to find the Log File at location " + logFilePath);
		}

		Workbook wb = new XSSFWorkbook();
		Row row;
		Cell cell;

		cellupdated = wb.createCellStyle();
		cellupdated.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellupdated.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		cellAdded = wb.createCellStyle();
		cellAdded.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellAdded.setFillForegroundColor(IndexedColors.GREEN.index);
		cellRemoved = wb.createCellStyle();
		cellRemoved.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellRemoved.setFillForegroundColor(IndexedColors.ORANGE.index);
		cellMisc = wb.createCellStyle();
		cellMisc.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellMisc.setFillForegroundColor(IndexedColors.RED.index);
		BufferedReader br = new BufferedReader(new FileReader(file));

		String st;

		while (true) {

			st = br.readLine();

			if (st == null) {
				Thread.sleep(2000);
			}

			else {
				// System.out.println(st);
				if (st.toLowerCase().contains("Starting comparision of Sheet".toLowerCase())) {
					// System.out.println("indide adding sheet");
					String SheetName = st.toLowerCase().split("starting comparision of sheet")[1];
					// System.out.println("SheetName "+SheetName.trim());
					addSheet(br, wb.createSheet(SheetName.trim()));
				} else if (st.toLowerCase().contains("end comparision of workbook")) {
					break;
				}
			}

		}

		File excelFile = new File(excelFilePath);
		if (excelFile.exists()) {
			excelFile.delete();
		}

		try {
			FileOutputStream fop = new FileOutputStream(excelFile);
			wb.write(fop);
			fop.close();
		} catch (Exception e) {
			System.out.println("Unable to Create Output Excel File");
		}

	}

	public void addSheet(BufferedReader br, Sheet sheet) throws Exception {
		int i = 0;
		while (true) {
			String str = br.readLine();
			if (str == null) {
				Thread.sleep(2000);
			} else {
				if (str.toLowerCase().contains("Starting row comparision".toLowerCase())) {
					addRow(br, sheet.createRow(i++));
				} else if (str.toLowerCase().contains("End Comparion of Sheet".toLowerCase())) {
					return;
				}
			}
		}
	}

	public void addRow(BufferedReader br, Row row) throws Exception {

		int j = 0;
		while (true) {
			String str = br.readLine();
			if (str == null) {
				Thread.sleep(2000);
			}
			if (str != null) {
				if (!str.toLowerCase().contains("Ending row comparision".toLowerCase())) {
					addCell(str, row.createCell(j++));

				} else {
					return;
				}
			}
		}
	}

	public void addCell(String str, Cell cell) {

		if (str.toLowerCase().contains("is equal to")) {
			cell.setCellValue(getcellValueFromText(str));
		} else if (str.toLowerCase().contains("is updated to")) {

			cell.setCellValue(getCellStringFromLog(str, "is updated to"));
			
			cell.setCellStyle(cellupdated);
		} else if (str.toLowerCase().contains("Cell Data-Type does not Match".toLowerCase())) {
			cell.setCellValue(getCellStringFromLog(str, "Cell Data-Type does not Match"));
			cell.setCellStyle(cellupdated);
		} else if (str.toLowerCase().contains("is removed")) {
			cell.setCellValue(getcellValueFromText(str));
			cell.setCellStyle(cellRemoved);
		} else if (str.toLowerCase().contains("is added")) {
			cell.setCellValue(getcellValueFromText(str));
			cell.setCellStyle(cellAdded);
		} else if (str.startsWith("Cell")) {
			cell.setCellValue(getcellValueFromText(str));
			cell.setCellStyle(cellMisc);
		}

	}

	private String getCellStringFromLog(String log, String delimitter) {
		String[] cellvalues = log.split(delimitter);
		String cellValue = getcellValueFromText(cellvalues[0]) + "-->" + getcellValueFromText(cellvalues[1]);
		return cellValue;
	}

	private String getcellValueFromText(String str) {
		return str.substring(str.indexOf("[") + 1, str.indexOf("]"));
	}

}
