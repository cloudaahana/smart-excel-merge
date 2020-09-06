import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;

public class ExcelMerge {

    public static void main(String args[]) throws FileNotFoundException {

        ArrayList<ExcelElements> datasets = new ArrayList<>();

        //Get all the excel file except master excel file
        ArrayList<String> excelFiles = new ArrayList<>();
        File folder = new File(System.getProperty("user.dir"));

        File[] listOfFiles = folder.listFiles();
        System.out.print("Ready to merge:\n");
        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                String filename = listOfFiles[i].getName();
                if (filename.contains(".xlsx") && !filename.contains("standard-data.xlsx") && !filename.contains("~$") && !filename.contains(".~")) {
                    excelFiles.add(filename);
                    System.out.print(filename + " | ");
                }
            }
        }
        if (excelFiles.size() < 1) {
            System.out.print("ERROR! Not Enough Files to Perform Standardization/Merge Operation.\n");
            System.exit(0);
        }
        System.out.print("\n\n");
        int blank_row = 0;

        //Read file one by one
        for (String excelFile : excelFiles) {

            int accountName_pos = 0;
            int accountOwner_pos = 0;
            int businessPhone_pos = 0;
            int company_pos = 0;
            int companyType_pos = 0;
            int created_pos = 0;
            int updated_pos = 0;
            int description_pos = 0;
            int doorCount_pos = 0;
            int email_pos = 0;
            int firstName_pos = 0;
            int lastName_pos = 0;
            int growthPlan_pos = 0;
            int growthTarget2020_pos = 0;
            int leadOwner_pos = 0;
            int leadStatus_pos = 0;
            int mailingCity_pos = 0;
            int mailingCountry_pos = 0;
            int mailingState_pos = 0;
            int mailingStreet_pos = 0;
            int mailingZip_pos = 0;
            int mobilePhone_pos = 0;
            int partner_pos = 0;
            int source_pos = 0;
            int otherCity_pos = 0;
            int otherState_pos = 0;
            int otherCountry_pos = 0;
            int otherZip_pos = 0;
            int otherStreet_pos = 0;
            int website1_pos = 0;
            int website2_pos = 0;
            int surveyConducted_pos = 0;
            int dealName_pos = 0;
            int closingDate_pos = 0;
            int contactType_pos = 0;
            int contactName_pos = 0;
            int bdms_pos = 0;
            //int partnerFileId;

            InputStream ExcelFileToRead = new FileInputStream(excelFile);
            XSSFWorkbook wb = null;
            try {
                wb = new XSSFWorkbook(ExcelFileToRead);
            } catch (IOException e) {
                e.printStackTrace();
            }
            XSSFSheet sheet = wb.getSheetAt(0);

            for (int i = 0; i <= 35; i++) {
                try {
                    String value = sheet.getRow(0).getCell(i).toString().toLowerCase().trim();
                    if (value.equals("account name"))
                        accountName_pos = i + 1;
                    else if (value.equals("account owner") || value.equals("lead owner") || value.equals("contact owner"))
                        accountOwner_pos = i + 1;
                    else if (value.equals("phone") || value.equals("office phone") || value.equals("business phone"))
                        businessPhone_pos = i + 1;
                    else if (value.equals("company"))
                        company_pos = i + 1;
                    else if (value.equals("company type"))
                        companyType_pos = i + 1;
                    else if (value.equals("created time") || value.equals("created"))
                        created_pos = i + 1;
                    else if (value.equals("modified time") || value.equals("last activity time"))
                        updated_pos = i + 1;
                    else if (value.equals("description"))
                        description_pos = i + 1;
                    else if (value.equals("door count") || value.equals("right door count"))
                        doorCount_pos = i + 1;
                    else if (value.equals("email"))
                        email_pos = i + 1;
                    else if (value.equals("first name"))
                        firstName_pos = i + 1;
                    else if (value.equals("last name"))
                        lastName_pos = i + 1;
                    else if (value.equals("growth plan"))
                        growthPlan_pos = i + 1;
                    else if (value.equals("2020 growth target"))
                        growthTarget2020_pos = i + 1;
                    else if (value.equals("lead owner"))
                        leadOwner_pos = i + 1;
                    else if (value.equals("lead source"))
                        leadStatus_pos = i + 1;
                    else if (value.equals("mailing city") || value.equals("city"))
                        mailingCity_pos = i + 1;
                    else if (value.equals("mailing country"))
                        mailingCountry_pos = i + 1;
                    else if (value.equals("mailing state"))
                        mailingState_pos = i + 1;
                    else if (value.equals("mailing street") || value.equals("street 1"))
                        mailingStreet_pos = i + 1;
                    else if (value.equals("mailing zip"))
                        mailingZip_pos = i + 1;
                    else if (value.equals(value.equals("mobile") || value.equals("mobile phone")))
                        mobilePhone_pos = i + 1;
                    else if (value.equals("partner"))
                        partner_pos = i + 1;
                    else if (value.equals("other city"))
                        otherCity_pos = i + 1;
                    else if (value.equals("other state"))
                        otherState_pos = i + 1;
                    else if (value.equals("other country"))
                        otherCountry_pos = i + 1;
                    else if (value.equals("other zip"))
                        otherZip_pos = i + 1;
                    else if (value.equals("other street") || value.equals("street 2"))
                        otherStreet_pos = i + 1;
                    else if (value.equals("website"))
                        website1_pos = i + 1;
                    else if (value.equals("website 2"))
                        website2_pos = i + 1;
                    else if (value.equals("survey conducted"))
                        surveyConducted_pos = i + 1;
                    else if (value.equals("deal name"))
                        dealName_pos = i + 1;
                    else if (value.equals("closing date"))
                        closingDate_pos = i + 1;
                    else if (value.equals("contact type"))
                        contactType_pos = i + 1;
                    else if (value.equals("contact name"))
                        contactName_pos = i + 1;
                    else if (value.equals("bdms"))
                        bdms_pos = i + 1;
                } catch (Exception e) {
                }
            }

            int rowLength = sheet.getPhysicalNumberOfRows();
            int records = rowLength - 1;
            System.out.println("Scan " + records + " record(s) from " + excelFile);

            for (int j = 1; j < rowLength; j++) {

                ExcelElements data = new ExcelElements();


                if (accountName_pos != 0) {
                    try {
                        data.setAccountName(sheet.getRow(j).getCell(accountName_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (accountOwner_pos != 0) {
                    try {
                        data.setAccountOwner(sheet.getRow(j).getCell(accountOwner_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (businessPhone_pos != 0) {
                    try {
                        data.setBusinessPhone(sheet.getRow(j).getCell(businessPhone_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (company_pos != 0) {
                    try {
                        data.setCompany(sheet.getRow(j).getCell(company_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (companyType_pos != 0) {
                    try {
                        data.setCompanyType(sheet.getRow(j).getCell(companyType_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (created_pos != 0) {
                    try {
                        data.setCreated(sheet.getRow(j).getCell(created_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (updated_pos != 0) {
                    try {
                        data.setUpdated(sheet.getRow(j).getCell(updated_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (description_pos != 0) {
                    try {
                        data.setDescription(sheet.getRow(j).getCell(description_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (doorCount_pos != 0) {
                    try {
                        data.setDoorCount(sheet.getRow(j).getCell(doorCount_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (email_pos != 0) {
                    try {
                        data.setEmail(sheet.getRow(j).getCell(email_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (firstName_pos != 0) {
                    try {
                        data.setFirstName(sheet.getRow(j).getCell(firstName_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (lastName_pos != 0) {
                    try {
                        data.setLastName(sheet.getRow(j).getCell(lastName_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (growthPlan_pos != 0) {
                    try {
                        data.setGrowthPlan(sheet.getRow(j).getCell(growthPlan_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (growthTarget2020_pos != 0) {
                    try {
                        data.setGrowthTarget2020(sheet.getRow(j).getCell(growthTarget2020_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (leadOwner_pos != 0) {
                    try {
                        data.setLeadOwner(sheet.getRow(j).getCell(leadOwner_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (leadStatus_pos != 0) {
                    try {
                        data.setLeadStatus(sheet.getRow(j).getCell(leadStatus_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (mailingCity_pos != 0) {
                    try {
                        data.setMailingCity(sheet.getRow(j).getCell(mailingCity_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (mailingCountry_pos != 0) {
                    try {
                        data.setMailingCountry(sheet.getRow(j).getCell(mailingCountry_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (mailingState_pos != 0) {
                    try {
                        data.setMailingState(sheet.getRow(j).getCell(mailingState_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (mailingStreet_pos != 0) {
                    try {
                        data.setMailingStreet(sheet.getRow(j).getCell(mailingStreet_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (mailingZip_pos != 0) {
                    try {
                        data.setMailingZip(sheet.getRow(j).getCell(mailingZip_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (mobilePhone_pos != 0) {
                    try {
                        data.setMobilePhone(sheet.getRow(j).getCell(mobilePhone_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (partner_pos != 0) {
                    try {
                        data.setPartner(sheet.getRow(j).getCell(partner_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (source_pos != 0) {
                    try {
                        data.setSource(sheet.getRow(j).getCell(source_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (otherCity_pos != 0) {
                    try {
                        data.setOtherCity(sheet.getRow(j).getCell(otherCity_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (otherState_pos != 0) {
                    try {
                        data.setOtherState(sheet.getRow(j).getCell(otherState_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (otherCountry_pos != 0) {
                    try {
                        data.setOtherCountry(sheet.getRow(j).getCell(otherCountry_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (otherZip_pos != 0) {
                    try {
                        data.setOtherZip(sheet.getRow(j).getCell(otherZip_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (otherStreet_pos != 0) {
                    try {
                        data.setOtherStreet(sheet.getRow(j).getCell(otherStreet_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (website1_pos != 0) {
                    try {
                        data.setWebsite1(sheet.getRow(j).getCell(website1_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (website2_pos != 0) {
                    try {
                        data.setWebsite2(sheet.getRow(j).getCell(website2_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (surveyConducted_pos != 0) {
                    try {
                        data.setSurveyConducted(sheet.getRow(j).getCell(surveyConducted_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (dealName_pos != 0) {
                    try {
                        data.setDealName(sheet.getRow(j).getCell(dealName_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (closingDate_pos != 0) {
                    try {
                        data.setClosingDate(sheet.getRow(j).getCell(closingDate_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (contactType_pos != 0) {
                    try {
                        data.setContactType(sheet.getRow(j).getCell(contactType_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (contactName_pos != 0) {
                    try {
                        data.setContactName(sheet.getRow(j).getCell(contactName_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (bdms_pos != 0) {
                    try {
                        data.setBdms(sheet.getRow(j).getCell(bdms_pos - 1).toString());
                    } catch (Exception e) {
                    }
                }

                if (data.isEmpty() == true) {
                    blank_row++;

                } else {
                    datasets.add(data);
                }
            }

        }
        System.out.println("All Blank Row(s) will be Ignored \n" + blank_row + " Blank Row(s) Found.");
        BigQueryDump.flushToExcel(datasets);
        System.out.print("\nXLSX to CSV Conversion Started.\n");

        try {
            String excelfileName = "standard-data.xlsx";
            String csvFileName = "bigquery-data.csv";

            ConvertXLSXToCSV.excelXToCSV(excelfileName, csvFileName);
            System.out.println("OK | Conversion Completed.\nAll Done!");

        } catch (Exception e) {
            System.out.println("ERROR! Conversion Failed, Please Re-run the Merge operation.\n");
        }
        try {
            PythonHandler.executePython();
        } catch (Exception e) {
            System.out.print("Failed to Execute Python Upload Script");
        }
    }
}