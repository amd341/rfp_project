package com.highpoint.rfpparse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.highpoint.rfpparse.App.Choice.ORIGINAL;

/**
 * Hello world!
 *
 */
public class App 
{
    public enum Choice {
        ORIGINAL, EXCELBASIC, EXCEL
    }

    public static void main(final String[] args) {

        Choice choice = Choice.EXCEL;

        if (choice == Choice.ORIGINAL) {
            try {
                String info = new String(Files.readAllBytes(Paths.get(args[1])));
                Map<String, Object> map = new HashMap<>();

                // split on ':' and on '::'
                String[] parts = info.split("::?");

                for (int i = 0; i < parts.length; i += 2) {
                    map.put(parts[i], parts[i + 1]);
                }

                Parser p = new Parser(new FileInputStream(args[0]), map);
                System.out.println(p.getSubSections());
            } catch (IOException | InvalidFormatException e) {
                System.out.println("Invalid input/output file name or format");
                e.printStackTrace();
            }
        }
        else if (choice == Choice.EXCELBASIC) {
            xmlParser x = new xmlParser();
            x.xmlParse("/home/alex/Documents/excels/enterprise.xlsx");
            System.out.println("we're here");
        }
        else if (choice == Choice.EXCEL){
            xmlParser x = new xmlParser();
            List<String> listoString = x.ExcelToJSON("/home/alex/Documents/excels/enterprise.xlsx");
        }


    }
}
