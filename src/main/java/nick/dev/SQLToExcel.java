package nick.dev;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Properties;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class SQLToExcel {

	private boolean debug = false;

	public Connection getSQLConnectionMSSQLServerWindowsAuth(String serverName, String databaseName) {
		StringBuffer sb = new StringBuffer();
		sb.append("jdbc:sqlserver://" + serverName);
		sb.append(";databaseName=" + databaseName);
		sb.append(";integratedSecurity=true;");
		return this.getSQLConnection(sb.toString(), "com.microsoft.sqlserver.jdbc.SQLServerDriver");
	}

	public Connection getSQLConnection(String jdbcUrl, String driver) {
		Connection connection = null;
		this.printDebug("Getting Connection:");
		try {
			Class.forName(driver);
			connection = DriverManager.getConnection(driver);
		} catch (SQLException | ClassNotFoundException e) {
			this.printDebug("Could not get a connection!");
			e.printStackTrace();
		}
		this.printDebug("Getting the Connection Succeeded!");
		return connection;
	}

	public ResultSet runQuery(Connection connection, String query) {
		if (query.endsWith(".sql")) {
			query = this.getQueryFromFile(query);
		}
		this.printDebug("Executing Query...");
		this.printDebug(query);
		PreparedStatement ps = null;
		ResultSet rs = null;
		try {
			ps = connection.prepareStatement(query, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = ps.executeQuery();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return rs;
	}

	public String getQueryFromFile(String fileName) {
		StringBuffer sb = new StringBuffer();
		Path path = Paths.get(fileName, new String[0]);
		Scanner scanner = null;
		try {
			scanner = new Scanner(path, StandardCharsets.UTF_8.name());
		} catch (IOException e) {
			e.printStackTrace();
		}
		while (scanner.hasNextLine()) {
			final String Input = scanner.nextLine();
			sb.append(String.valueOf(Input) + " ");
		}
		return sb.toString();
	}

	public void runQueryFromProperties() {
		this.runQueryFromProperties("ReportConfig.properties");
	}

	public void runQueryFromProperties(String filePath) {
		Properties properties = new Properties();
		try {
			properties.load(new FileInputStream(filePath));
			Connection connection = this.getSQLConnectionMSSQLServerWindowsAuth(properties.getProperty("SQLServer"),
					properties.getProperty("SQLDatabase"));
			ResultSet resultSet = this.runQuery(connection, properties.getProperty("SQLStatement"));
			this.createExcelFile(resultSet, properties.getProperty("reportName") + "_"
					+ new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime()) + ".xlsx",
					properties.getProperty("sheetName"));
		} catch (IOException | SQLException e) {
			e.printStackTrace();
		}
	}

	public void createExcelFile(ResultSet resultSet, String fileName, String sheetName)
			throws IOException, SQLException {
		long startTime = System.nanoTime();
		ResultSetMetaData rsmd = resultSet.getMetaData();
		int columnNumber = rsmd.getColumnCount();
		// Declare Excel Objects:
		if (!fileName.endsWith(".xlsx")) {
			this.printDebug("Please check extension. This is not a valid .xlsx file!");
			throw new RuntimeException();
		}
		Path path = Paths.get(fileName);
		Files.deleteIfExists(path);
		FileOutputStream fileOut = new FileOutputStream(fileName);
		SXSSFWorkbook workbook = new SXSSFWorkbook(10000);
		SXSSFSheet worksheet = workbook.createSheet(sheetName);
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		SXSSFRow rowhead = worksheet.createRow(0);
		SXSSFCell cell = rowhead.createCell(0);
		int index = 1;
		font.setUnderline(XSSFFont.U_SINGLE);
		style.setFont(font);
		// Create the column headers:
		for (int i = 1; i <= columnNumber; i++) {
			rsmd.getColumnName(i);
			cell = rowhead.createCell(i - 1);
			cell.setCellStyle(style);
			cell.setCellValue(rsmd.getColumnName(i) == null ? "" : rsmd.getColumnName(i));
		}
		// Get total row count:
		resultSet.last();
		int numRows = resultSet.getRow();
		int incrementPercentageForProgress = 10;
		int incrementStep = 1;
		int incrementValue = 0;
		int incrementPercentValue = numRows / incrementPercentageForProgress;
		resultSet.beforeFirst();
		// Populate the rest of the rows with data:
		this.printDebug("Generating excel file!");
		while (resultSet.next()) {
			rowhead = worksheet.createRow(index);
			for (int i = 1; i <= columnNumber; i++) {
				rowhead.createCell(i - 1).setCellValue(resultSet.getString(i));
			}
			index++;
			incrementValue++;
			if (incrementValue >= incrementPercentValue) {
				this.printDebug((incrementStep * incrementPercentageForProgress) + " % of excel file generated... (~"
						+ ((incrementValue * incrementStep) + 3) + " rows)");
				incrementValue = 0;
				incrementStep++;
			}
		}
		// Freezing top row:
		worksheet.createFreezePane(0, 1);
		workbook.write(fileOut);
		workbook.close();
		fileOut.flush();
		fileOut.close();
		resultSet.close();
		long difference = System.nanoTime() - startTime;
		this.printDebug("Total Excel generation time: " + String.format("%d min, %d sec",
				TimeUnit.NANOSECONDS.toHours(difference), TimeUnit.NANOSECONDS.toSeconds(difference)
						- TimeUnit.MINUTES.toSeconds(TimeUnit.NANOSECONDS.toMinutes(difference))));
		this.printDebug("I finished! File created here:");
		this.printDebug(path);
	}

	public void printDebug(Object o) {
		if (this.debug) {
			System.out.println(o.toString());
		}
	}

	public void setDebug(boolean debug) {
		this.debug = debug;
	}

	public static void main(String[] args) {
		SQLToExcel ste = new SQLToExcel();
		ste.setDebug(true);
		ste.runQueryFromProperties();
	}

}