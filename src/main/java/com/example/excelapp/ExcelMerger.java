package com.example.excelapp;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelMerger {

    public ExcelMerger(){

    }

    public void generateGpx(String file1, String file2) {
//        String file1Path = ExcelMerger.class.getClassLoader().getResource(file1).getPath();
//        String file2Path = ExcelMerger.class.getClassLoader().getResource(file2).getPath();
        String outputFilePath = "output.gpx";
        String file1ColumnName = "Business Partner";
        String file2ColumnName = "Bp";
        Map<String, Map<String, Object>> file1Data = readExcelFile(file1, file1ColumnName);
        Map<String, Map<String, Object>> file2Data = readExcelFile(file2, file2ColumnName);
        List<Map<String, Object>> mergedData = new ArrayList();
        Iterator var9 = file1Data.keySet().iterator();

        Map row;
        while(var9.hasNext()) {
            String commonColumnValue = (String)var9.next();
            if (file2Data.containsKey(commonColumnValue)) {
                System.out.println("match");
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
        }

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
                    values.add(String.valueOf(Integer.parseInt(row.get("Meter No").toString())));
                    values.add(",");
                    values.add(row.get("Invoice Amount").toString());
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

    private static Map<String, Map<String, Object>> readExcelFile(String filePath, String commonColumnName) {
        Map<String, Map<String, Object>> data = new HashMap();

        try {
            System.out.println(filePath + " " + commonColumnName);
            Workbook workbook = WorkbookFactory.create(new File(filePath));

            try {
                Sheet sheet = workbook.getSheetAt(0);
                Row headerRow = sheet.getRow(0);
                int commonColumnIndex = -1;

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
                    if (row != null) {
                        Map<String, Object> rowData = new HashMap();

                        for(int j = 0; j < headerRow.getLastCellNum(); ++j) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                String headerValue = headerRow.getCell(j).getStringCellValue();
                                switch (cell.getCellType()) {
                                    case STRING:
                                        rowData.put(headerValue, cell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        rowData.put(headerValue, cell.getNumericCellValue());
                                        break;
                                    case BOOLEAN:
                                        rowData.put(headerValue, cell.getBooleanCellValue());
                                }
                            }
                        }

                        Cell commonColumnCell = row.getCell(commonColumnIndex);
                        if (commonColumnCell != null) {
                            Object commonColumnValue;
                            switch (commonColumnCell.getCellType()) {
                                case STRING:
                                    commonColumnValue = commonColumnCell.getStringCellValue();
                                    break;
                                case NUMERIC:
                                    commonColumnValue = commonColumnCell.getNumericCellValue();
                                    break;
                                case BOOLEAN:
                                    commonColumnValue = commonColumnCell.getBooleanCellValue();
                                    break;
                                default:
                                    commonColumnValue = null;
                            }

                            data.put(commonColumnValue.toString(), rowData);
                        }
                    }
                }
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
}
