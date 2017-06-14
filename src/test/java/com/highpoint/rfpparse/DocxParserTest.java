package com.highpoint.rfpparse;

import org.junit.Test;

import java.util.*;

import static org.junit.Assert.*;

/**
 * Created by brenden on 6/6/17.
 */
public class DocxParserTest {
    @Test
    public void getSectionsShouldIdentifyHeadingsAndText() throws Exception {
        //heading, then paragraph, then heading, then table
        Map<String,Object> map = new HashMap<>();
        map.put("test", "test");
        DocxParser p = new DocxParser(getClass().getResourceAsStream("/test.docx"), map);
        List<Map<String,Object>> sections = p.getSections();

        Map<String,Object> expectedSection1 = new HashMap<>();
        Map<String,Object> expectedSection2 = new HashMap<>();
        expectedSection1.put("heading","Heading 1");
        expectedSection1.put("body", "Paragraph");
        expectedSection1.put("test", "test");
        expectedSection2.put("heading", "Heading 1");
        expectedSection2.put("body", "Table\n");
        expectedSection2.put("test","test");
        List<Map<String,Object>> expectedSections = new ArrayList<>(Arrays.asList(expectedSection1,expectedSection2));

        assertEquals(expectedSections, sections);

        //heading, then table, then heading, then paragraph
        p = new DocxParser(getClass().getResourceAsStream("/test1.docx"), map);
        sections = p.getSections();

        expectedSections = new ArrayList<>(Arrays.asList(expectedSection2,expectedSection1));

        assertEquals(expectedSections, sections);

        //only paragraph
        p = new DocxParser(getClass().getResourceAsStream("/test2.docx"), map);
        sections = p.getSections();
        expectedSection1.put("heading", "No heading");
        expectedSections = new ArrayList<>(Arrays.asList(expectedSection1));

        assertEquals(expectedSections, sections);

        //only table
        p = new DocxParser(getClass().getResourceAsStream("/test3.docx"), map);
        sections = p.getSections();

        expectedSection2.put("heading", "No heading");
        expectedSections = new ArrayList<>(Arrays.asList(expectedSection2));

        assertEquals(expectedSections, sections);

    }


}