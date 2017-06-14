package com.highpoint.rfpparse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * Parser main class
 *
 */
public class App 
{
    /**
     * @param args takes two arguments, first the filepath to be parsed, and second
     *             a config file that consists of key:value pairs separated by ::
     */
    public static void main(final String[] args) {
        try {
            String info = new String(Files.readAllBytes(Paths.get(args[1])));
            Map<String, Object> map = new HashMap<>();

            // split on ':' and on '::'
            String[] parts = info.split("::?");

            for (int i = 0; i < parts.length; i += 2) {
                map.put(parts[i], parts[i + 1]);
            }

            Parser p = new Parser(new FileInputStream(args[0]), map);
            //p.bulkIndexSections("search-elastic-test-yyco5dncwicwd2nufqhakzek2e.us-east-1.es.amazonaws.com",
             //       443, "https", "rfps2", "rfp2", true);
            String r = p.mapNewFields("search-elastic-test-yyco5dncwicwd2nufqhakzek2e.us-east-1.es.amazonaws.com",
                    443, "https","rfps2", "rfp2");

            System.out.println(r);

        } catch (IOException | InvalidFormatException e) {
            System.out.println("Invalid input/output file name or format");
            e.printStackTrace();
        }


    }
}
