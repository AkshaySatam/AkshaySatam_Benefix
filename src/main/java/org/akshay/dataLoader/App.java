package org.akshay.dataLoader;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;
import com.itextpdf.kernel.pdf.canvas.parser.filter.TextRegionEventFilter;
import com.itextpdf.kernel.pdf.canvas.parser.listener.FilteredTextEventListener;
import com.itextpdf.kernel.pdf.canvas.parser.listener.ITextExtractionStrategy;
import com.itextpdf.kernel.pdf.canvas.parser.listener.LocationTextExtractionStrategy;

/*import org.apache.pdfbox.cos.COSBase;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.text.PDFTextStripper;*/
/**
 * Hello world!
 *
 */
public class App {
	// TODO Currently the dimensions are hard-coded. can be easily changed. Will do
	// this later.
	static String[][] excelMapper = new String[1000][52];
	static int excelRowNum = 1;
	static HashMap<String, String> statesUSA = new HashMap<String, String>();

	public static void main(String[] args) throws IOException {
		initializeStatesMap();

		// Get all the files in the current directory and put them in a collection
		ArrayList<String> pdfFiles = new ArrayList<String>();

		getPDFfilesInDirectory("./src/main/java/org/akshay/dataLoader/", pdfFiles);

		PdfReader pdfReader;
		PdfDocument pdfDoc;

		// Do the following for every entry in the collection
		for (String s : pdfFiles) {
			pdfReader = new PdfReader("./src/main/java/org/akshay/dataLoader/" + s);
			pdfDoc = new PdfDocument(pdfReader);
			processDocument(pdfDoc);
		}

	}

	private static void processDocument(PdfDocument pdfDoc) throws IOException {
		// do the following for each page

		String planName, startDate, endDate, states, grpRating;
		int x,y,ind ;
		// traverse through all pages of the PDF
		for (int page = 1; page <= pdfDoc.getNumberOfPages(); page++, excelRowNum++) {
			x = 119;
			
			ind =0;
			startDate = extractStartDate(pdfDoc, page);
			excelMapper[excelRowNum][0] = startDate;

			endDate = extractEndDate(pdfDoc, page);
			excelMapper[excelRowNum][1] = endDate;

			planName = extractPlanName(pdfDoc, page);
			excelMapper[excelRowNum][2] = planName;

			states = extractStates(pdfDoc, page);
			excelMapper[excelRowNum][3] = statesUSA.get(states.toLowerCase());

			grpRating = extractGrpRating(pdfDoc, page);
			excelMapper[excelRowNum][4] = extractNumber(grpRating);
			
			for(int k=0; k<3; k++, x+= 220) {
				y = 355;
				for(int l=0;l<15;l++, y-=15) {
					excelMapper[excelRowNum][5+ind++]= getAgeGroupInfo(pdfDoc,page,x,y);
					}
			}
			
		}
		// Enter the info in the 2D array, to be sent to the Excel sheet
		insertDateIntoExcelSheet(excelMapper);
		// printExcelMapper();
	}

	private static String getAgeGroupInfo(PdfDocument pdfDoc, int page, int x, int y) {
		Rectangle rect = new Rectangle(x, y, 50, 11);
		return getTextAtDesiredLocation(rect, pdfDoc, page);
	}

	private static String extractNumber(String grpRating) {
		int i = 0;
		while (i < grpRating.length() && !Character.isDigit(grpRating.charAt(i)))
			i++;
		return grpRating.substring(i);
	}

	private static String extractGrpRating(PdfDocument pdfDoc, int page) {
		Rectangle rect = new Rectangle(404, 447, 48, 11);
		return getTextAtDesiredLocation(rect, pdfDoc, page);
	}

	private static void printExcelMapper() {
		for (int i = 0; i < 50; i++) {
			for (int j = 0; j < 4; j++) {
				System.out.println(excelMapper[i][j]);
			}
		}

	}

	private static String pruneInput(String string) {
		if (string == null || string.contains("Exception")) {
			return "";
		}
		return string;
	}

	private static String extractStates(PdfDocument pdfDoc, int page) {
		Rectangle rect = new Rectangle(308, 482, 130, 20);
		return getTextAtDesiredLocation(rect, pdfDoc, page);
	}

	private static String extractEndDate(PdfDocument pdfDoc, int page) {
		Rectangle rect = new Rectangle(645, 503, 50, 11);
		return getTextAtDesiredLocation(rect, pdfDoc, page);
	}

	private static String extractStartDate(PdfDocument pdfDoc, int page) {
		Rectangle rect = new Rectangle(572, 503, 50, 11);
		return getTextAtDesiredLocation(rect, pdfDoc, page);
	}

	private static void insertDateIntoExcelSheet(String[][] data) throws IOException {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Benefit Data");
		int rowNum = 0;

		for (String[] d : data) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (String field : d) {
				Cell cell = row.createCell(colNum++);
				cell.setCellValue((String) field);
			}
		}

		FileOutputStream outputStream = new FileOutputStream(
				"./src/main/java/org/akshay/dataLoader/BeneFix Small Group Plans upload template.xlsx");
		workbook.write(outputStream);
		workbook.close();
	}

	private static String extractPlanName(PdfDocument pdfDoc, int page) {
		Rectangle rect = new Rectangle(390, 402, 167, 22);
		return getTextAtDesiredLocation(rect, pdfDoc, page);
	}

	private static String getTextAtDesiredLocation(Rectangle rect, PdfDocument pdfDoc, int page) {
		TextRegionEventFilter regionFilter = new TextRegionEventFilter(rect);
		ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(),
				regionFilter);

		String str = PdfTextExtractor.getTextFromPage(pdfDoc.getPage(page), strategy);
		return str;

	}

	private static void getPDFfilesInDirectory(String path, ArrayList<String> pdfFiles) {
		File dir = new File(path);
		File[] allFiles = dir.listFiles();
		if (allFiles != null) {
			for (File file : allFiles) {
				// System.out.println(file.getName());
				if (file.getName().endsWith(".pdf")) {
					pdfFiles.add(file.getName());
				}
			}
		} else {
			System.out.println("The given directory contains no PDF files. Program exiting");
			System.exit(1);
		}

	}

	private static void initializeStatesMap() {
		// TODO Auto-generated method stub
		statesUSA.put("alabama", "AL");
		statesUSA.put("alaska", "AK");
		statesUSA.put("arizona", "AZ");
		statesUSA.put("arkansas", "AR");
		statesUSA.put("california", "CA");
		statesUSA.put("colorado", "CO");
		statesUSA.put("conneticut", "CT");
		statesUSA.put("delaware", "DE");
		statesUSA.put("florida", "FL");
		statesUSA.put("georgia", "GA");
		statesUSA.put("hawaii", "HI");
		statesUSA.put("idaho", "ID");
		statesUSA.put("illinois", "IL");
		statesUSA.put("indiana", "IN");
		statesUSA.put("iowa", "IA");
		statesUSA.put("kansas", "KS");
		statesUSA.put("kentucky", "KY");
		statesUSA.put("louisinia", "LA");
		statesUSA.put("maine", "ME");
		statesUSA.put("maryland", "MD");
		statesUSA.put("massachusetts", "MA");
		statesUSA.put("michigan", "MI");
		statesUSA.put("minnesota", "MN");
		statesUSA.put("mississippi", "MS");
		statesUSA.put("missouri", "MO");
		statesUSA.put("montana", "MT");
		statesUSA.put("nebraska", "NE");
		statesUSA.put("nevada", "NV");
		statesUSA.put("new hampshire", "NH");
		statesUSA.put("new jersey", "NJ");
		statesUSA.put("new mexico", "NM");
		statesUSA.put("new york", "NY");
		statesUSA.put("north carolina", "NC");
		statesUSA.put("north dakota", "ND");
		statesUSA.put("ohio", "OH");
		statesUSA.put("oklahoma", "OK");
		statesUSA.put("oregon", "OR");
		statesUSA.put("pennsylvania", "PA");
		statesUSA.put("rhode island", "RI");
		statesUSA.put("south carolina", "SC");
		statesUSA.put("south dakota", "SD");
		statesUSA.put("tenessee", "TN");
		statesUSA.put("texas", "TX");
		statesUSA.put("utah", "UT");
		statesUSA.put("vermont", "VT");
		statesUSA.put("virginia", "VA");
		statesUSA.put("washington", "WA");
		statesUSA.put("west virginia", "WV");
		statesUSA.put("wisconsin", "WI");
		statesUSA.put("wyoming", "WY");
	}
}
