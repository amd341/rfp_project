package com.highpoint.rfpparse;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.entity.ContentType;
import org.apache.http.nio.entity.NStringEntity;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.elasticsearch.client.Response;
import org.elasticsearch.client.RestClient;

import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by Brenden Sosnader on 6/6/17.
 * Parser class for parsing files in Microsoft Word format into sections based on Heading 1's and Heading 2's
 * as well as uploading parsed sections to an elasticsearch instance 
 */
public class Parser {

    private XWPFDocument xdoc;
    private String company;
    private String date;
    private String type;
    private String service;
    private String output;



    /**
     * @param input the path to a rfp docx file to be parsed
     * @param company the company the rfp is for
     * @param date the date the rfp was submitted
     * @param type the type of company
     * @param service the service/services provided to the company
     * @throws IOException if file paths are incorrect
     * @throws InvalidFormatException if it is not a docx file
     */

    public Parser(InputStream input, String company, String date, String type, String service, String output) throws IOException, InvalidFormatException {
        xdoc = new XWPFDocument(OPCPackage.open(input));
        this.company = company;
        this.date = date;
        this.type = type;
        this.service = service;
        this.output = output;

    }
    public List<Map<String,Object>> getSections()
    {
        Map<String,Object> section = new HashMap<>();
        List<Map<String,Object>> sections = new ArrayList<>();
        List<IBodyElement> elements = xdoc.getBodyElements();
        String heading = "No heading";
        StringBuilder body = new StringBuilder();

        for (IBodyElement element : elements) {
            if (element.getElementType().toString().equals("PARAGRAPH")) {
                XWPFParagraph paragraph = (XWPFParagraph) element;

                if (paragraph.getStyleID() != null) {
                    if (paragraph.getStyleID().equals("Heading1")) {
                        if (body.length() > 0) {
                            section.put("body", body.toString());
                            section.put("heading", heading);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(section);
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        heading = paragraph.getText();
                    } else {
                        body.append(paragraph.getText()).append("\n");
                    }
                } else {
                    body.append(paragraph.getText()).append("\n");
                }

            } else if (element.getElementType().toString().equals("TABLE")) {
                XWPFTable table = (XWPFTable) element;

                if (table.getStyleID() != null) {
                    if (table.getStyleID().equals("Heading1")) {
                        if (body.length() > 0) {
                            section.put("body", body.toString());
                            section.put("heading", heading);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(section);
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        heading = table.getText();
                    } else {
                        body.append(table.getText()).append("\n");
                    }
                } else {
                    body.append(table.getText()).append("\n");
                }
            }
        }
        section.put("heading", heading);
        section.put("body", body.toString());
        section.put("company", company);
        section.put("date", date);
        section.put("type", type);
        section.put("service", service);
        sections.add(section);
        return sections;
    }

    public List<Map<String,Object>> getSubSections()
    {
        Map<String,Object> section = new HashMap<>();
        List<Map<String,Object>> sections = new ArrayList<>();
        List<IBodyElement> elements = xdoc.getBodyElements();
        String headingOne = "No heading";
        String headingTwo = "No heading";
        String headingOnePre = "No heading";
        StringBuilder body = new StringBuilder();

        for(IBodyElement element : elements)
        {
            if(element.getElementType().toString().equals("PARAGRAPH"))
            {
                XWPFParagraph paragraph = (XWPFParagraph) element;

                if(paragraph.getStyleID() != null)
                {
                    if(paragraph.getStyleID().equals("Heading1"))
                    {
                        headingOnePre = paragraph.getText();
                    }
                    else if(paragraph.getStyleID().equals("Heading2"))
                    {
                        if(body.length() > 0)
                        {
                            section.put("body", body.toString());
                            section.put("headingOne", headingOne);
                            section.put("headingTwo", headingTwo);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(section);
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        headingTwo = paragraph.getText();
                        headingOne = headingOnePre;
                    }
                    else
                    {
                        body.append(paragraph.getText()).append("\n");
                    }
                }
                else
                {
                    body.append(paragraph.getText()).append("\n");
                }
            }
            else if(element.getElementType().toString().equals("TABLE"))
            {
                XWPFTable table = (XWPFTable) element;

                if(table.getStyleID() != null)
                {
                    if(table.getStyleID().equals("Heading1"))
                    {
                        headingOne = table.getText();
                    }
                    else if(table.getStyleID().equals("Heading2"))
                    {
                        System.out.print("hmmm");
                        if(body.length() > 0)
                        {
                            section.put("body", body.toString());
                            section.put("headingOne", headingOne);
                            section.put("headingTwo", headingTwo);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(section);
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        headingTwo = table.getText();
                    }
                    else
                    {
                        body.append(table.getText()).append("\n");
                    }
                }
                else
                {
                    body.append(table.getText()).append("\n");
                }
            }
        }
        section.put("body", body.toString());
        section.put("headingOne", headingOne);
        section.put("headingTwo", headingTwo);
        section.put("company", company);
        section.put("date", date);
        section.put("type", type);
        section.put("service", service);
        sections.add(section);

        return sections;
    }


    public String getJson() throws JsonProcessingException
    {
        ObjectMapper objectMapper = new ObjectMapper();

        return objectMapper.writeValueAsString(getSections());
    }

    public void writeToFile() throws IOException
    {
        FileWriter fileWriter = new FileWriter(output);
        fileWriter.write(getJson());
        fileWriter.close();
    }
    public String bulkIndexSections(String hostname, int port, String scheme, String index, String type) throws IOException {
        RestClient restClient = RestClient.builder(
                new HttpHost(hostname, port, scheme)).build();

        String actionMetaData = String.format("{ \"index\" : { \"_index\" : \"%s\", \"_type\" : \"%s\" } }%n", index, type);

        List<String> bulkData = getSectionsAsStrings();
        StringBuilder bulkRequestBody = new StringBuilder();
        for (String bulkItem : bulkData)
        {
            bulkRequestBody.append(actionMetaData);
            bulkRequestBody.append(bulkItem);
            bulkRequestBody.append("\n");
        }
        HttpEntity entity = new NStringEntity(bulkRequestBody.toString(), ContentType.APPLICATION_JSON);

        Response response = restClient.performRequest("POST", "/rfps/rfp/_bulk",
                Collections.emptyMap(), entity);
        restClient.close();
        return response.toString();
    }

    public List<String> getSectionsAsStrings() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,Object> section = new HashMap<>();
        List<String> sections = new ArrayList<>();
        List<IBodyElement> elements = xdoc.getBodyElements();
        String heading = "No heading";
        StringBuilder body = new StringBuilder();

        for (IBodyElement element : elements) {
            if (element.getElementType().toString().equals("PARAGRAPH")) {
                XWPFParagraph paragraph = (XWPFParagraph) element;

                if (paragraph.getStyleID() != null) {
                    if (paragraph.getStyleID().equals("Heading1")) {
                        if (body.length() > 0) {
                            section.put("body", body.toString());
                            section.put("heading", heading);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(objectMapper.writeValueAsString(section));
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        heading = paragraph.getText();
                    } else {
                        body.append(paragraph.getText()).append("\n");
                    }
                } else {
                    body.append(paragraph.getText()).append("\n");
                }

            } else if (element.getElementType().toString().equals("TABLE")) {
                XWPFTable table = (XWPFTable) element;

                if (table.getStyleID() != null) {
                    if (table.getStyleID().equals("Heading1")) {
                        if (body.length() > 0) {
                            section.put("body", body.toString());
                            section.put("heading", heading);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(objectMapper.writeValueAsString(section));
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        heading = table.getText();
                    } else {
                        body.append(table.getText()).append("\n");
                    }
                } else {
                    body.append(table.getText()).append("\n");
                }
            }
        }
        section.put("heading", heading);
        section.put("body", body.toString());
        section.put("company", company);
        section.put("date", date);
        section.put("type", type);
        section.put("service", service);
        sections.add(objectMapper.writeValueAsString(section));
        return sections;
    }

    public List<String> getSubSectionsAsStrings() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,Object> section = new HashMap<>();
        List<String> sections = new ArrayList<>();
        List<IBodyElement> elements = xdoc.getBodyElements();
        String headingOne = "No heading";
        String headingTwo = "No heading";
        String headingOnePre = "No heading";
        StringBuilder body = new StringBuilder();

        for(IBodyElement element : elements)
        {
            if(element.getElementType().toString().equals("PARAGRAPH"))
            {
                XWPFParagraph paragraph = (XWPFParagraph) element;

                if(paragraph.getStyleID() != null)
                {
                    if(paragraph.getStyleID().equals("Heading1"))
                    {
                        headingOnePre = paragraph.getText();
                    }
                    else if(paragraph.getStyleID().equals("Heading2"))
                    {
                        if(body.length() > 0)
                        {
                            section.put("body", body.toString());
                            section.put("headingOne", headingOne);
                            section.put("headingTwo", headingTwo);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(objectMapper.writeValueAsString(section));
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        headingTwo = paragraph.getText();
                        headingOne = headingOnePre;
                    }
                    else
                    {
                        body.append(paragraph.getText()).append("\n");
                    }
                }
                else
                {
                    body.append(paragraph.getText()).append("\n");
                }
            }
            else if(element.getElementType().toString().equals("TABLE"))
            {
                XWPFTable table = (XWPFTable) element;

                if(table.getStyleID() != null)
                {
                    if(table.getStyleID().equals("Heading1"))
                    {
                        headingOne = table.getText();
                    }
                    else if(table.getStyleID().equals("Heading2"))
                    {
                        System.out.print("hmmm");
                        if(body.length() > 0)
                        {
                            section.put("body", body.toString());
                            section.put("headingOne", headingOne);
                            section.put("headingTwo", headingTwo);
                            section.put("company", company);
                            section.put("date", date);
                            section.put("type", type);
                            section.put("service", service);
                            sections.add(objectMapper.writeValueAsString(section));
                            section = new HashMap<>();
                            body = new StringBuilder();
                        }
                        headingTwo = table.getText();
                    }
                    else
                    {
                        body.append(table.getText()).append("\n");
                    }
                }
                else
                {
                    body.append(table.getText()).append("\n");
                }
            }
        }
        section.put("body", body.toString());
        section.put("headingOne", headingOne);
        section.put("headingTwo", headingTwo);
        section.put("company", company);
        section.put("date", date);
        section.put("type", type);
        section.put("service", service);
        sections.add(objectMapper.writeValueAsString(section));

        return sections;
    }

    public XWPFDocument getXdoc()
    {
        return xdoc;
    }



}
