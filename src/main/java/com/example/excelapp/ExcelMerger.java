package com.example.excelapp;

import java.io.*;
import java.util.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.DateTime;
import org.joda.time.chrono.EthiopicChronology;


public class ExcelMerger {

    public ExcelMerger(){

    }

    public void generateGpx(String file1, String filename) throws IOException {
        String outputFilePath = filename.substring(0, filename.length() - 4) + "gpx";
        String file1ColumnName = "Business Partner";
        Map<String, Map<String, Object>> file1Data = readExcelFile(file1, file1ColumnName, filename);
        Map<String, Map<String, Object>> file2Data = readGpsData();
        List<Map<String, Object>> mergedData = new ArrayList();
        Iterator var9 = file1Data.keySet().iterator();

        Map row;
        int count = 0;
        int noGps = 0;
        while(var9.hasNext()) {
            String commonColumnValue = (String)var9.next();
            if (file2Data.containsKey(commonColumnValue)) {
                count++;
                row = (Map)file1Data.get(commonColumnValue);
                Map<String, Object> file2Row = (Map)file2Data.get(commonColumnValue);
                row.remove("Arrears");
                row.remove("Adjustment");
                row.remove("Bill Doc No");
                row.remove("Bill Key Date");
                row.remove("Bill Month");
                row.remove("CSC Code");
                row.remove("CSC Office");
                row.remove("Column Status");
                row.remove("Cons Charge");
                row.remove("Contract");
                row.remove("Contract Account");
                row.remove("Created By");
                row.remove("Created On");
                row.remove("Demand Charge");
                row.remove("District Office");
                row.remove("Invoice No");
                row.remove("KVARH");
                row.remove("KW");
                row.remove("KWH Sold");
                row.remove("MF_KVARH");
                row.remove("MF_KW");
                row.remove("MF_KWH");
                row.remove("MIN Charge");
                row.remove("MISC Charge");
                row.remove("Meter Reading Unit");
                row.remove("No of Days");
                row.remove("OVDL PNLTY");
                row.remove("PF Charge");
                row.remove("PT Charge");
                row.remove("Portion");
                row.remove("Posting Date");
                row.remove("Present Reading");
                row.remove("Previous Reading");
                row.remove("REGION");
                row.remove("SBSDY/OWNCONS");
                row.remove("SRVC Charge");
                row.remove("Tariff");
                row.putAll(file2Row);
                mergedData.add(row);
            }
            else{
                //System.out.println(commonColumnValue);
                noGps++;
            }
        }
        System.out.println("The count is : " + count);
        System.out.println("Data with no GPS: " + noGps);
        try {
            PrintWriter writer = new PrintWriter(new File(outputFilePath));

            try {
                writer.println("<?xml version='1.0' encoding='UTF-8' standalone='yes' ?>\n<gpx version=\"1.1\" creator=\"GDAL 3.0.4\" xmlns=\"http://www.topografix.com/GPX/1/1\" xmlns:osmand=\"https://osmand.net\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd\">\n  <metadata>\n    <bounds minlat=\"9.3059081\" minlon=\"42.1351272\" maxlat=\"9.3102854\" maxlon=\"42.1381259\" />\n  </metadata>\n");
                Iterator var16 = mergedData.iterator();

                while(var16.hasNext()) {
                    row = (Map)var16.next();
                    List<String> values = new ArrayList();
                    values.add("<wpt lat=\"");
                    values.add(row.get("GPS_LAT").toString());
                    values.add("\" lon=\"");
                    values.add(row.get("GPS_LNG").toString());
                    values.add("\"><time>2021-07-29T14:09:31Z</time>\n");
                    values.add("\t<name>");
                    values.add(row.get("Customer Name").toString());
                    values.add(",");
                    values.add(row.get("Business Partner").toString());
                    values.add(",");
//                    values.add(String.valueOf(Integer.parseInt(row.get("Meter No").toString())));
                    values.add(row.get("Meter No").toString());
                    values.add(",");
                    values.add(row.get("Invoice Amount").toString().trim());
                    values.add("</name>\n");
                    values.add("\t<extensions>\n\t\t<osmand:color>#00842b</osmand:color>\n\t\t<osmand:icon>special_star</osmand:icon>\n\t\t<osmand:background>circle</osmand:background>\n\t</extensions>\n</wpt>");
                    writer.println(String.join("", values));
                }
            } catch (Throwable var14) {
                try {
                    writer.close();
                } catch (Throwable var13) {
                    var14.addSuppressed(var13);
                }

                throw var14;
            }
            writer.close();
        } catch (FileNotFoundException var15) {
            var15.printStackTrace();
        }

    }

    private Map<String, Map<String, Object>> readExcelFile(String filePath, String commonColumnName, String fileName) throws IOException {
        Map<String, Map<String, Object>> data = new HashMap();
        Workbook workbook = null;

        try {
            workbook = WorkbookFactory.create(new File(filePath));
            try {
                Sheet sheet = workbook.getSheetAt(0);
                Row headerRow = sheet.getRow(0);
                int commonColumnIndex = -1;
                int payStatus = -1;
                int billMonth = -1;
                String lastMonth = fileName.substring(0, fileName.length() - 5);
                System.out.println(lastMonth);
                int i;
                for(i = 0; i < headerRow.getLastCellNum(); ++i) {
                    Cell cell = headerRow.getCell(i);
                    if (cell.getStringCellValue().equalsIgnoreCase(commonColumnName)) {
                        commonColumnIndex = i;
                    }
                    else if(cell.getStringCellValue().equalsIgnoreCase("Column Status")){
                        payStatus = i;
                    }
                    else if(cell.getStringCellValue().equalsIgnoreCase("Bill Month")){
                        billMonth = i;
                    }
                }

                if (commonColumnIndex == -1) {
                    throw new RuntimeException("Column not found: " + commonColumnName);
                }
                int paidCount = 0;
                int rowCount = 0;
                for(i = 1; i <= sheet.getLastRowNum(); ++i) {
                    rowCount++;
                    Row row = sheet.getRow(i);
                    if (row != null && "NOT PAID".equalsIgnoreCase(row.getCell(payStatus).getStringCellValue()) && lastMonth.equalsIgnoreCase(row.getCell(billMonth).getStringCellValue())){
                        Map<String, Object> rowData = new HashMap();
                        paidCount++;
                        for(int j = 0; j < headerRow.getLastCellNum(); ++j) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                String headerValue = headerRow.getCell(j).getStringCellValue();
                                rowData.put(headerValue, getCellValue(cell));
                            }
                        }

                        Cell commonColumnCell = row.getCell(commonColumnIndex);
                        if (commonColumnCell != null) {
                            Object commonColumnValue = getCellValue(commonColumnCell);
                            data.put(commonColumnValue.toString(), rowData);
                        }
                    }
                }
                System.out.println(commonColumnName + " row count : " + rowCount);
                System.out.println(commonColumnName + " paid count: " + paidCount);
            } catch (Throwable var14) {
                if (workbook != null) {
                    try {
                        workbook.close();
                    } catch (Throwable var13) {
                        var14.addSuppressed(var13);
                    }
                }

                throw var14;
            }

            if (workbook != null) {
                workbook.close();
            }
        } catch (EncryptedDocumentException | IOException var15) {
            var15.printStackTrace();
        }

        return data;
    }

    private Object getCellValue(Cell commonColumnCell) {
        Object commonColumnValue;
        switch (commonColumnCell.getCellType()) {
            case STRING:
                commonColumnValue = commonColumnCell.getStringCellValue();
                break;
            case NUMERIC:
                DataFormatter df = new DataFormatter();
                commonColumnValue = df.formatCellValue(commonColumnCell);
                break;
            default:
                commonColumnValue = null;
        }
        return commonColumnValue;
    }

    private Map<String, Map<String, Object>> readGpsData() {
        Map<String, Map<String, Object>> data = new HashMap();
        Workbook workbook = null;
        String commonColumnName = "Bp";

        try {
            InputStream in = this.getClass().getResourceAsStream("/bb.xlsx");
            if(in != null) {
                workbook = WorkbookFactory.create(in);
            }
            try {
                Sheet sheet = workbook.getSheetAt(0);
                Row headerRow = sheet.getRow(0);
                int commonColumnIndex = -1;
                String lastMonth = getLastMonth("21.03.2023.xlsx");
                System.out.println(lastMonth + " hell");
                int i;
                for(i = 0; i < headerRow.getLastCellNum(); ++i) {
                    Cell cell = headerRow.getCell(i);
                    if (cell.getStringCellValue().equalsIgnoreCase(commonColumnName)) {
                        commonColumnIndex = i;
                        break;
                    }
                }

                if (commonColumnIndex == -1) {
                    throw new RuntimeException("Column not found: " + commonColumnName);
                }
                for(i = 1; i <= sheet.getLastRowNum(); ++i) {
                    Row row = sheet.getRow(i);
                    if (row != null){
                        Map<String, Object> rowData = new HashMap();
                        for(int j = 0; j < headerRow.getLastCellNum(); ++j) {
                            String headerValue = headerRow.getCell(j).getStringCellValue();
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                rowData.put(headerValue, getCellValue(cell));
                            }
                        }

                        Cell commonColumnCell = row.getCell(commonColumnIndex);
                        if (commonColumnCell != null) {
                            Object commonColumnValue = getCellValue(commonColumnCell);
                            data.put(commonColumnValue.toString(), rowData);
                        }
                    }
                }
                System.out.println(commonColumnName + " row count : " + i);
            } catch (Throwable var14) {
                if (workbook != null) {
                    try {
                        workbook.close();
                    } catch (Throwable var13) {
                        var14.addSuppressed(var13);
                    }
                }

                throw var14;
            }

            if (workbook != null) {
                workbook.close();
            }
        } catch (EncryptedDocumentException | IOException var15) {
            var15.printStackTrace();
        }

        return data;
    }

    public String getLastMonth(String filename){

        String[] months = {
                "MESKEREM",
                "TIKIMIT",
                "HIDAR",
                "TAHISAS",
                "TIR",
                "YEKATIT",
                "MEGABIT",
                "MIYAZIA",
                "GINBOT",
                "SENE",
                "HAMLE",
                "NEHASIE",
                "PAGUME"
        };
        String[] day = filename.split("\\.");

        org.joda.time.DateTime dtISO = new DateTime(Integer.parseInt(day[2]), Integer.parseInt(day[1]), Integer.parseInt(day[0]), 12, 0, 0, 0);

        // find out what the same instant is using the Ethiopic Chronology
        DateTime dtEthiopic = dtISO.withChronology(EthiopicChronology.getInstance());
        System.out.println(dtEthiopic.toDateTimeISO());
        System.out.println(dtEthiopic);
        dtEthiopic = dtEthiopic.minusMonths(1);
        if(dtEthiopic.getMonthOfYear() == 13){
            dtEthiopic = dtEthiopic.minusMonths(1);
        }
        String lastMonth = months[dtEthiopic.getMonthOfYear() - 1] + "-" + dtEthiopic.getYearOfEra();
        return lastMonth;
    }

}
