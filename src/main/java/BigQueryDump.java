import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class BigQueryDump {

    public static void flushToExcel(ArrayList < ExcelElements > datasets) throws FileNotFoundException {

        System.out.println("Flush " + datasets.size() + " record(s) to big_query_data.xlsx");

        String[] columns = {
                "ACCOUNT_NAME",
                "ACCOUNT_OWNER",
                "BUSINESS_PHONE",
                "COMPANY",
                "COMPANY_TYPE",
                "CREATED",
                "UPDATED",
                "DESCRIPTION",
                "DOOR_COUNT",
                "EMAIL",
                "FIRST_NAME",
                "LAST_NAME",
                "GROWTH_PLAN",
                "GROWTH_TARGET_2020",
                "LEAD_OWNER",
                "LEAD_STATUS",
                "MAILING_CITY",
                "MAILING_COUNTRY",
                "MAILING_STATE",
                "MAILING_STREET",
                "MAILING_ZIP",
                "MOBILE_PHONE",
                "PARTNER",
                "SOURCE",
                "OTHER_CITY",
                "OTHER_STATE",
                "OTHER_COUNTRY",
                "OTHER_ZIP",
                "OTHER_STREET",
                "WEBSITE_1",
                "WEBSITE_2",
                "SURVEY_CONDUCTED",
                "DEAL_NAME",
                "CLOSING_DATE",
                "CONTACT_TYPE",
                "CONTACT_NAME",
                "BDMS"
//                "PARTNER_FILE_ID"
        };

        // Create a Workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a Sheet
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Create a Row
        XSSFRow headerRow = sheet.createRow(0);

        // Create cells
        for (int i = 0; i < columns.length; i++) {
            XSSFCell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with employees data
        int rowNum = 1;
        for (ExcelElements excelElement: datasets) {
            Row row = sheet.createRow(rowNum++);

            try {
                row.createCell(0).setCellValue(excelElement.getAccountName().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {}

            try {
                row.createCell(1).setCellValue(excelElement.getAccountOwner().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(2).setCellValue(excelElement.getBusinessPhone().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(3).setCellValue(excelElement.getCompany().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(4).setCellValue(excelElement.getCompanyType().replace("\n", "").replace("\r", "").replace(".0", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(5).setCellValue(excelElement.getCreated().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(6).setCellValue(excelElement.getUpdated().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(7).setCellValue(excelElement.getDescription().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(8).setCellValue(excelElement.getDoorCount().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(9).setCellValue(excelElement.getEmail().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(10).setCellValue(excelElement.getFirstName().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(11).setCellValue(excelElement.getLastName().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }

            try {
                row.createCell(12).setCellValue(excelElement.getGrowthPlan().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(13).setCellValue(excelElement.getGrowthTarget2020().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(14).setCellValue(excelElement.getLeadOwner().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(15).setCellValue(excelElement.getLeadStatus().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(16).setCellValue(excelElement.getOtherCity().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(17).setCellValue(excelElement.getOtherState().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(18).setCellValue(excelElement.getOtherCountry().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(19).setCellValue(excelElement.getOtherZip().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(20).setCellValue(excelElement.getOtherStreet().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(21).setCellValue(excelElement.getWebsite1().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(22).setCellValue(excelElement.getWebsite2().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(23).setCellValue(excelElement.getSurveyConducted().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }

            try {
                row.createCell(24).setCellValue(excelElement.getDealName().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }

            try {
                row.createCell(25).setCellValue(excelElement.getClosingDate().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }

            try {
                row.createCell(26).setCellValue(excelElement.getContactType().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }

            try {
                row.createCell(27).setCellValue(excelElement.getContactName().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
            try {
                row.createCell(28).setCellValue(excelElement.getBdms().replace("\n", "").replace("\r", ""));
            } catch (Exception e) {

            }
//            try {
//                row.createCell(29).setCellValue(excelElement.getPartnerFileId().replace("\n", "").replace("\r", ""));
//            } catch (Exception e) {
//
//            }

        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("standard-data.xlsx");
        try {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.close();
        } catch (IOException e) {}

        // Closing the workbook
        try {
            workbook.close();
            System.out.println("Done | Write Complete.");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}