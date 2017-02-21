package com.excelReader.sponsorExclusions;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excelReader.bean.ChannelBean;
import com.excelReader.bean.OfferBean;

public class ExcelReadExample {

	private static final String DB_DRIVER = "oracle.jdbc.driver.OracleDriver";
	private static final String DB_URL = "jdbc:oracle:thin:@sv2404.ca.sunlife:1560:sx3d1";
	private static final String DB_USER = "EXCOAPP";
	private static final String DB_PASSWORD = "excapp";

	public static void main(String[] args) throws Exception {

		String filename = "D:/NBA/Template/SponsorExclusionsLoadTemplate.xlsx";
		FileInputStream fis = null;

		try {

			List<OfferBean> offerList = new ArrayList<OfferBean>();
			fis = new FileInputStream(filename);

			XSSFWorkbook workbook = new XSSFWorkbook(fis);

			XSSFSheet offerSheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = offerSheet.iterator();
			while (rowIterator.hasNext()) {
				OfferBean offerBean = new OfferBean();
				Row row = rowIterator.next();
				if (row.getRowNum() == 0) {
					continue;
				}
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					int cellNum = cell.getColumnIndex();

					if (cellNum == 0) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						offerBean.setOffer(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 1) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						offerBean.setRuleType(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 2) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						offerBean.setParentValue(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 3) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						offerBean.setCriteriaValue(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 4) {
						cell.setCellType(Cell.CELL_TYPE_STRING);

						if (cell.getStringCellValue().equalsIgnoreCase("Off")) {
							offerBean.setRuleAccess("N");
						} else {
							offerBean.setRuleAccess("Y");
						}
						continue;
					}
				}
				offerList.add(offerBean);
			}

			List<ChannelBean> channelList = new ArrayList<ChannelBean>();
			XSSFSheet channelSheet = workbook.getSheetAt(1);
			Iterator<Row> rowIterator1 = channelSheet.iterator();
			while (rowIterator1.hasNext()) {
				ChannelBean channelBean = new ChannelBean();
				Row row = rowIterator1.next();
				if (row.getRowNum() == 0) {
					continue;
				}
				Iterator<Cell> cellIterator1 = row.cellIterator();

				while (cellIterator1.hasNext()) {
					Cell cell = cellIterator1.next();
					int cellNum = cell.getColumnIndex();

					if (cellNum == 0) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						channelBean.setChannel(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 1) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						channelBean.setRuleType(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 2) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						channelBean.setParentValue(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 3) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						channelBean.setCriteriaValue(cell.getStringCellValue());

						continue;
					}
					if (cellNum == 4) {
						cell.setCellType(Cell.CELL_TYPE_STRING);

						if (cell.getStringCellValue().equalsIgnoreCase("Off")) {
							channelBean.setRuleAccess("N");
						} else {
							channelBean.setRuleAccess("Y");
						}
						continue;
					}
				}
				channelList.add(channelBean);
			}
			insertSponsorExclusionData(offerList, channelList);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (fis != null) {
				fis.close();
			}
		}
	}

	private static void insertSponsorExclusionData(List<OfferBean> offerList,
			List<ChannelBean> channelList) throws SQLException,
			FileNotFoundException {
		
		String insertOfferQuery = "";
		String insertChannelQuery = "";
		String insertRuleQuery = "";
		FileWriter fwrite = null;
		List<String> insertStatements = new ArrayList<String>();

		Set<String> offerUniqueSet = new HashSet<String>();
		for(OfferBean offer : offerList){
			offerUniqueSet.add(offer.getOffer());
		}
		
		for(String offerName : offerUniqueSet){
			if (!isOfferExist(offerName)) {
				insertOfferQuery = "INSERT INTO SX3EXCO.OFFER_T(OFFER_ID,OFFER_NM,DEFAULT_ACCESS_IND,"
						+ "CREATE_TS,CREATE_USER_ID,UPDATE_TS,UPDATE_USER_ID) VALUES(SX3EXCO.OFFER_ID_SEQ.NEXTVAL,'"
						+ offerName
						+ "','Y',SYSDATE,'BATCH',SYSDATE,'BATCH');";
			}
			insertStatements.add(insertOfferQuery);
		}
		
		Set<String> channelUniqueSet = new HashSet<String>();
		for(ChannelBean channel : channelList){
			channelUniqueSet.add(channel.getChannel());
		}
		
		
		for(String channelName : channelUniqueSet){
			if (!isOfferExist(channelName)) {
				insertChannelQuery = "INSERT INTO SX3EXCO.CHANNEL_T(CHANNEL_ID,CHANNEL_NM,DEFAULT_ACCESS_IND,"
					+ "CREATE_TS,CREATE_USER_ID,UPDATE_TS,UPDATE_USER_ID) VALUES(SX3EXCO.CHANNEL_ID_SEQ.NEXTVAL,'"
					+ channelName
					+ "','Y',SYSDATE,'BATCH',SYSDATE,'BATCH');";
			}
			insertStatements.add(insertChannelQuery);
		}
		
		for (OfferBean offer : offerList) {
			String ruletype = "";
			if ("Advisor".equalsIgnoreCase(offer.getRuleType())) {
				ruletype = "1";
			} else if ("Sponsor".equalsIgnoreCase(offer.getRuleType())) {
				ruletype = "2";
			} else if ("Arrangement".equalsIgnoreCase(offer.getRuleType())) {
				ruletype = "3";
			} else {
				ruletype = "4";
			}

			if (offer.getCriteriaValue() != null && !(offer.getCriteriaValue().isEmpty())) {
				insertRuleQuery = "INSERT INTO SX3EXCO.RULE_T(RULE_ID,PARENT_RULE_ID,CHANNEL_ID,OFFER_ID,"
						+ "RULE_TYPE_ID,CRITERIA_VAL,RULE_ACCESS_IND,CREATE_TS,CREATE_USER_ID,UPDATE_TS,UPDATE_USER_ID) VALUES("
						+ "SX3EXCO.RULE_ID_SEQ.NEXTVAL,"
						+ "(SELECT PARENT_RULE_ID FROM SX3EXCO.RULE_T WHERE CRITERIA_VAL = '"
						+ offer.getParentValue()
						+ "' AND OFFER_ID = "
						+ "(SELECT OFFER_ID FROM SX3EXCO.OFFER_T WHERE OFFER_NM = '"
						+ offer.getOffer()
						+ "' AND RULE_TYPE_ID = "
						+ ruletype
						+ "))"
						+ ",null,"
						+ "(SELECT OFFER_ID FROM SX3EXCO.OFFER_T WHERE OFFER_NM = '"
						+ offer.getOffer()
						+ "'),"
						+ ruletype
						+ ",'"
						+ offer.getCriteriaValue()
						+ "','"
						+ offer.getRuleAccess()
						+ "',SYSDATE,'BATCH',SYSDATE,'BATCH');";
			} else {
				insertRuleQuery = "INSERT INTO SX3EXCO.RULE_T(RULE_ID,PARENT_RULE_ID,CHANNEL_ID,OFFER_ID,"
						+ "RULE_TYPE_ID,CRITERIA_VAL,RULE_ACCESS_IND,CREATE_TS,CREATE_USER_ID,UPDATE_TS,UPDATE_USER_ID) VALUES("
						+ "SX3EXCO.RULE_ID_SEQ.NEXTVAL,"
						+ "(SELECT PARENT_RULE_ID FROM SX3EXCO.RULE_T WHERE CRITERIA_VAL = '"
						+ offer.getParentValue()
						+ "' AND OFFER_ID ="
						+ "(SELECT OFFER_ID FROM SX3EXCO.OFFER_T WHERE OFFER_NM = '"
						+ offer.getOffer()
						+ "' AND RULE_TYPE_ID = "
						+ ruletype
						+ "))"
						+ ",null,"
						+ "(SELECT OFFER_ID FROM SX3EXCO.OFFER_T WHERE OFFER_NM = '"
						+ offer.getOffer()
						+ "'),"
						+ ruletype
						+ ",' ','"
						+ offer.getRuleAccess()
						+ "',SYSDATE,'BATCH',SYSDATE,'BATCH');";
			}
				insertStatements.add(insertRuleQuery);
				
			try {
				fwrite = new FileWriter("D:/insertSqlScripts.sql");
				for (String insertStmtFile : insertStatements) {
					fwrite.write(insertStmtFile);
					fwrite.write("\n\n");
				}
			} catch (IOException ioe) {
				ioe.printStackTrace();
			} finally {
				try {
					if (fwrite != null) {
						fwrite.close();
					}
				} catch (IOException ioe) {
					ioe.printStackTrace();
				}
			}
		}
	}

	private static boolean isOfferExist(String offerName) throws SQLException {
		
		Connection dbConnection = null;
		PreparedStatement preparedStatement = null;
		String selectOffer = "SELECT * FROM SX3EXCO.OFFER_T WHERE OFFER_NM = ?";
		boolean result = false;

		try {
			dbConnection = getDBConnection();
			preparedStatement = dbConnection.prepareStatement(selectOffer);
			preparedStatement.setString(1, offerName);
			int i = preparedStatement.executeUpdate();
			if (i > 0)
				result = true;
			else
				result = false;
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			if (preparedStatement != null) {
				preparedStatement.close();
			}
			if (dbConnection != null) {
				dbConnection.close();
			}
		}
		return result;
	}

	private static boolean isChannelExist(String channelName)
			throws SQLException {
		
		Connection dbConnection = null;
		PreparedStatement preparedStatement = null;
		boolean result = false;
		String selectChannel = "SELECT * FROM SX3EXCO.CHANNEL_T WHERE CHANNEL_NM = ?";
		
		try {
			dbConnection = getDBConnection();
			preparedStatement = dbConnection.prepareStatement(selectChannel);
			preparedStatement.setString(1, channelName);
			int i = preparedStatement.executeUpdate();
			if (i > 0)
				result = true;
			else
				result = false;
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			if (preparedStatement != null) {
				preparedStatement.close();
			}
			if (dbConnection != null) {
				dbConnection.close();
			}
		}
		return result;
	}

	private static Connection getDBConnection() {

		Connection dbConnection = null;
		try {
			Class.forName(DB_DRIVER);
		} catch (ClassNotFoundException e) {
			System.out.println(e.getMessage());
		}
		try {
			dbConnection = DriverManager.getConnection(DB_URL, DB_USER, DB_PASSWORD);
		} catch (SQLException e) {
			System.out.println(e.getMessage());
		}
		return dbConnection;
	}
}
