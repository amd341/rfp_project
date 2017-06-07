package com.highpoint.rfpparse;

import org.junit.Test;

import java.text.SimpleDateFormat;
import java.util.*;

import static org.junit.Assert.*;

/**
 * Created by brenden on 6/6/17.
 */
public class ParserTest {
    @Test
    public void getSectionsShouldIdentifyHeadingsAndText() throws Exception {
        Date d = new Date();
        //heading, then paragraph, then heading, then table
        SimpleDateFormat sm = new SimpleDateFormat("yyyy-MM-dd");
        Parser p = new Parser(getClass().getResourceAsStream("/test.docx"), "test", sm.format(d), "test", "test", "test");
        List<Map<String,Object>> sections = p.getSections();

        Map<String,Object> expectedSection1 = new HashMap<>();
        Map<String,Object> expectedSection2 = new HashMap<>();
        expectedSection1.put("heading","Heading 1");
        expectedSection1.put("body", "Paragraph");
        expectedSection1.put("company", "test");
        expectedSection1.put("date", sm.format(d));
        expectedSection1.put("type", "test");
        expectedSection1.put("service", "test");
        expectedSection2.put("heading", "Heading 1");
        expectedSection2.put("body", "Table\n");
        expectedSection2.put("company", "test");
        expectedSection2.put("date", sm.format(d));
        expectedSection2.put("type", "test");
        expectedSection2.put("service", "test");
        List<Map<String,Object>> expectedSections = new ArrayList<>(Arrays.asList(expectedSection1,expectedSection2));

        assertEquals(expectedSections, sections);

        //heading, then table, then heading, then paragraph
        p = new Parser(getClass().getResourceAsStream("/test1.docx"), "test", sm.format(d), "test", "test", "test");
        sections = p.getSections();

        expectedSections = new ArrayList<>(Arrays.asList(expectedSection2,expectedSection1));

        assertEquals(expectedSections, sections);

        //only paragraph
        p = new Parser(getClass().getResourceAsStream("/test2.docx"), "test", sm.format(d), "test", "test", "test");
        sections = p.getSections();
        expectedSection1.put("heading", "No heading");
        expectedSections = new ArrayList<>(Arrays.asList(expectedSection1));

        assertEquals(expectedSections, sections);

        //only table
        p = new Parser(getClass().getResourceAsStream("/test3.docx"), "test", sm.format(d), "test", "test", "test");
        sections = p.getSections();

        expectedSection2.put("heading", "No heading");
        expectedSections = new ArrayList<>(Arrays.asList(expectedSection2));

        assertEquals(expectedSections, sections);

    }


}