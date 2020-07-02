package ActivityTracker.Model;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ActivityTrackerOperation {

	String leftTimeMsg = "";
	String Path = "G:\\excel\\Test\\TestExcel.xlsx";
	String TextPath = "G:\\excel\\Test\\TestDoc.txt";
	
	

	public void UpdateInOldExcel() throws InvalidFormatException, EncryptedDocumentException, IOException {
		BufferedReader reader = new BufferedReader(new FileReader(TextPath));
		FileInputStream inputStream = new FileInputStream(new File(Path));
		Workbook workbook = WorkbookFactory.create(inputStream);

		ArrayList<String> words = new ArrayList<String>();
		String line;
		String cellValue = "";

		Sheet sheet = workbook.getSheetAt(0);
		System.out.println(sheet);

		while ((line = reader.readLine()) != null) {

			words.add(line);
			// System.out.println(words);
		}
		reader.close();

		for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row != null) {
				Cell cell = row.getCell(2);
				Cell cell1 = row.getCell(3);

				if (cell != null) {
					// Found column and there is value in the cell.
					if (cell.getCellType() == CellType.STRING) {
						cellValue = cell.getStringCellValue();

					} else if (cell.getCellType() == CellType.NUMERIC) {

						double newValue = cell.getNumericCellValue();

						cellValue = String.valueOf(newValue);

					}

					for (String sLine : words) {
						if (sLine.contains(cellValue)) {
							int index = words.indexOf(sLine);
							System.out.println("");

							System.out.println(
									"Got a match of " + cellValue + " at line " + index + "with time" + cell1);

						}

//						else
//							System.out.println("Not found");

					}

				}
			}
		}

	}

}
