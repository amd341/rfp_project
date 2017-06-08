package com.highpoint.rfpparse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main(final String[] args) {
        try {
            String info = new String(Files.readAllBytes(Paths.get(args[0])));
            String[] arr = info.split("\\r?\\n");

            Parser p = new Parser(new FileInputStream(arr[0]), arr[1], arr[2], arr[3], arr[4], arr[5]);
            System.out.println(p.getSubSections());
        } catch (IOException | InvalidFormatException e) {
            System.out.println("Invalid input/output file name or format");
            e.printStackTrace();
        }


    }
}
