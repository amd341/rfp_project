package com.highpoint.rfpparse;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Parser main class
 *
 */
public class App
{

    public enum Choice {
        ORIGINAL, EXCELBASIC, EXCEL
    }

    //Hey!
    //to run excelparser now with intellij go to Run, Edit Configurations, type the path to the excel file
    //as the first argument, then a space, then the path to 'config2' which is in the root of this repo
    //config2 has the basic section tags you had before, but now in that separate file so it's more flexibile

    /**
     * @param args takes two arguments, first the filepath to be parsed, and second
     *             a config file that consists of key:value pairs separated by ::
     */
    public static void main(final String[] args) {

        Choice choice = Choice.ORIGINAL;

        Boolean isXLSX = true;

        try {
            String info = new String(Files.readAllBytes(Paths.get(args[1])));
            Map<String, Object> map = new HashMap<>();

            String ext = FilenameUtils.getExtension(args[0]);

            if(ext.equals("xls")){
                isXLSX = false;
            }
            else if(ext.equals("xlsx")){
                isXLSX = true;
            }

            // split on ':' and on '::'
            String[] parts = info.split("::?");

            for (int i = 0; i < parts.length; i += 2) {
                map.put(parts[i], parts[i + 1]);
            }

            if (choice == Choice.ORIGINAL) {
                DocxParser p = new DocxParser(new FileInputStream(args[0]), map);
                System.out.println(SearchClient.bulkIndex("search-elastic-test-yyco5dncwicwd2nufqhakzek2e.us-east-1.es.amazonaws.com", 443, "https", "rfps2", "rfp2", p.getHighlightedAsStrings()));

            } else if (choice == Choice.EXCELBASIC) {
                ExcelParser x = new ExcelParser(new FileInputStream(args[0]), map, isXLSX);
                List<String> jstr = x.getJsonStrings();
                System.out.println(jstr);

            } else if (choice == Choice.EXCEL) {

                ExcelParser x = new ExcelParser(new FileInputStream(args[0]), map, isXLSX);

                String resp = x.bulkIndex("search-elastic-test-yyco5dncwicwd2nufqhakzek2e.us-east-1.es.amazonaws.com", 443, "https", "rfps3", "rfp");
                System.out.println(resp);

            }

        } catch (IOException | InvalidFormatException e) {
            System.out.println("something's wrong");
            e.printStackTrace();
        }
    }
}
