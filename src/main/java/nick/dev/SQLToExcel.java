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

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SQLToExcel {

	public Connection getSQLConnectionMSSQLServerWindowsAuth(String serverName, String databaseName) {
		StringBuffer sb = new StringBuffer();
		sb.append("jdbc:sqlserver://" + serverName);
		sb.append(";databaseName=" + databaseName);
		sb.append(";integratedSecurity=true;");
		return this.getSQLConnection(sb.toString(), "com.microsoft.sqlserver.jdbc.SQLServerDriver");
	}

	public Connection getSQLConnection(String jdbcUrl, String driver) {
		Connection connection = null;
		System.out.println("Getting Connection:");
		try {
			Class.forName(driver);
			connection = DriverManager.getConnection(driver);
		} catch (SQLException | ClassNotFoundException e) {
			System.out.println("Could not get a connection!");
			e.printStackTrace();
		}
		System.out.println("Getting the Connection Succeeded!");
		return connection;
	}

	public ResultSet runQuery(Connection connection, String query, boolean showQuery) {
		if (query.endsWith(".sql")) {
			query = this.getQueryFromFile(query);
		}
		System.out.println("Executing Query...");
		if (showQuery) {
			System.out.println(query);
		}
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
			ResultSet resultSet = this.runQuery(connection, properties.getProperty("SQLStatement"),
					Boolean.parseBoolean(properties.getProperty("showQuery", "false").toLowerCase()));
			this.createExcelFile(resultSet, properties.getProperty("reportName") + "_"
					+ new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime()) + ".xlsx",
					properties.getProperty("sheetName"));
		} catch (IOException | SQLException e) {
			e.printStackTrace();
		}
	}

	public void createExcelFile(ResultSet resultSet, String fileName, String sheetName)
			throws IOException, SQLException {
		ResultSetMetaData rsmd = resultSet.getMetaData();
		int columnNumber = rsmd.getColumnCount();
		// Declare Excel Objects:
		if (!fileName.endsWith(".xlsx")) {
			System.out.println("Please check extension. This is not a valid .xlsx file!");
			throw new RuntimeException();
		}
		Path path = Paths.get(fileName);
		Files.deleteIfExists(path);
		FileOutputStream fileOut = new FileOutputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet worksheet = workbook.createSheet(sheetName);
		XSSFCellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		XSSFRow rowhead = worksheet.createRow(0);
		XSSFCell cell = rowhead.createCell(0);
		int index = 1;
		font.setUnderline(XSSFFont.U_SINGLE);
		style.setFont(font);
		// Create the column headers:
		for (int i = 1; i <= columnNumber; i++) {
			rsmd.getColumnName(i);
			cell = rowhead.createCell(i - 1);
			cell.setCellStyle(style);
			cell.setCellValue(rsmd.getColumnName(i));
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
		System.out.println("Generating excel file!");
		while (resultSet.next()) {
			rowhead = worksheet.createRow(index);
			for (int i = 1; i <= columnNumber; i++) {
				rowhead.createCell(i - 1).setCellValue(resultSet.getString(i));
			}
			index++;
			incrementValue++;
			if (incrementValue >= incrementPercentValue) {
				System.out.println((incrementStep * incrementPercentageForProgress) + " % of excel file generated... (~"
						+ ((incrementValue * incrementStep) + 3) + " rows)");
				incrementValue = 0;
				incrementStep++;
			}
		}
		for (int i = 0; i < worksheet.getRow(0).getPhysicalNumberOfCells(); i++) {
			worksheet.autoSizeColumn(i);
		}
		// Freezing top row:
		worksheet.createFreezePane(0, 1);
		workbook.write(fileOut);
		workbook.close();
		fileOut.flush();
		fileOut.close();
		resultSet.close();
		System.out.println("I finished! File created here:");
		System.out.println(path);
	}

	public static void main(String[] args) {
		SQLToExcel ste = new SQLToExcel();
		ste.runQueryFromProperties();
	}

}