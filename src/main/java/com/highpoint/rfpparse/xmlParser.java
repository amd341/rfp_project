package com.highpoint.rfpparse;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.entity.ContentType;
import org.apache.http.nio.entity.NStringEntity;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.elasticsearch.client.Response;
import org.elasticsearch.client.RestClient;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by alex on 6/13/17.
 */
public class xmlParser {

    //private static final String FILE_NAME = "/home/alex/documents/excels/enterprise.xlsx";

    public void xmlParse(String fileName){
        boolean isBlankRow = true;
        //int count = 0;
        try {
            System.out.println("trying...");
            FileInputStream excelFile = new FileInputStream(new File(fileName));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(1);
            Iterator<Row> iterator = datatypeSheet.iterator();
            System.out.println("tried");

            while (iterator.hasNext()) {
                //count++;
                //System.out.println("first while..." + count);
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    //count++;
                    //System.out.println("second while..." + count);
                    Cell currentCell = cellIterator.next();

                    //System.out.print("      " + currentCell.getCellTypeEnum());

                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "--");
                        isBlankRow = false;
                    }
                    else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "--");
                        isBlankRow = false;
                    }
                }
                if (isBlankRow == false){
                    System.out.println();
                    isBlankRow = true;
                }

            }
        }
        catch (FileNotFoundException e){
            e.printStackTrace();
        }
        catch (IOException e){
            e.printStackTrace();
        }
    }

    public List<String> ExcelToJSON(String fileName, String service, String company, String date){
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,Object> sectionHash = new HashMap<>();
        List<String> sectionsList = new ArrayList<>();
        boolean isBlankRow = true;

        try{
            //initiating variable for the excel file, workbook, and sheet

            FileInputStream excelFile = new FileInputStream(new File(fileName));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            StringBuilder body = new StringBuilder();
            //iterator to iterate through sheets
            Iterator<Sheet> sheetIterator = workbook.iterator();

            int count = 0;
            while(count < workbook.getNumberOfSheets()){
                datatypeSheet = workbook.getSheetAt(count);

                sectionHash.put("company", company);
                sectionHash.put("service", service);
                sectionHash.put("date", date);
                sectionHash.put("heading", workbook.getSheetName(count));

                //iterate through rows
                Iterator<Row> iterator = datatypeSheet.iterator();
                //System.out.println("tried");

                while (iterator.hasNext()) {
                    //count++;
                    //System.out.println("first while..." + count);


                    Row currentRow = iterator.next();
                    Iterator<Cell> cellIterator = currentRow.iterator();

                    while (cellIterator.hasNext()) {
                        //count++;
                        //System.out.println("second while..." + count);
                        Cell currentCell = cellIterator.next();

                        //System.out.print("      " + currentCell.getCellTypeEnum());

                        if (currentCell.getCellTypeEnum() == CellType.STRING) {
                            body.append(currentCell.getStringCellValue() + "--");
                            //System.out.print(currentCell.getStringCellValue() + "--");
                            isBlankRow = false;
                        }
                        else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                            body.append(currentCell.getNumericCellValue() + "--");
                            //System.out.print(currentCell.getNumericCellValue() + "--");
                            isBlankRow = false;
                        }
                    }
                    if (isBlankRow == false){
                        body.append("\n");
                        //System.out.println();
                        isBlankRow = true;
                    }

                }
                count++;
                sectionHash.put("body",body);
                sectionsList.add(objectMapper.writeValueAsString(sectionHash));
                sectionHash = new HashMap<>();
                body = new StringBuilder();
            }

        }
        catch (FileNotFoundException e){
            e.printStackTrace();
        }
        catch (IOException e){
            e.printStackTrace();
        }
        for(String s : sectionsList) {
            System.out.println(s);
        }
        return(sectionsList);
    }

    public String postToES(List<String> sectionsList, String hostname, int port, String scheme, String index, String type) throws IOException{
        RestClient restClient = RestClient.builder(new HttpHost(hostname, port, scheme)).build();

        //System.out.println("we here 11");

        HttpEntity entity = new NStringEntity(String.valueOf(sectionsList), ContentType.APPLICATION_JSON);

        //System.out.println("we here 222");

        String actionMetaData = String.format("{ \"index\" : { \"_index\" : \"%s\", \"_type\" : \"%s\" } }%n", index, type);

        StringBuilder prepString = new StringBuilder();

        for (String s : sectionsList){
            prepString.append(actionMetaData);
            prepString.append(s);
            if (sectionsList.indexOf(s) != sectionsList.size()-1) {
                prepString.append("\n");
            }

        }

        System.out.println();
        System.out.println(prepString);

        Response response = restClient.performRequest("POST", "/rfps3/rfp/_bulk",
                Collections.emptyMap(), entity);
        restClient.close();

        return response.toString();
    }
}