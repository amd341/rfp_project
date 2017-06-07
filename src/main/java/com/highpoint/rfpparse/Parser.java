package com.highpoint.rfpparse;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.HttpStatus;
import org.apache.http.entity.ContentType;
import org.apache.http.nio.entity.NStringEntity;
import org.apache.http.util.EntityUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.elasticsearch.client.Response;
import org.elasticsearch.client.RestClient;

import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by brenden on 6/6/17.
 */
public class Parser {

    private XWPFDocument xdoc;
    private String company;
    private String date;
    private String type;
    private String service;
    private String output;
    private SimpleDateFormat sm;



    /**
     * @param input
     * @param company
     * @param date
     * @param type
     * @param service
     * @throws IOException
     * @throws InvalidFormatException
     */

    public Parser(InputStream input, String company, String date, String type, String service, String output) throws IOException, InvalidFormatException {
        xdoc = new XWPFDocument(OPCPackage.open(input));
        sm = new SimpleDateFormat("yyyy-MM-dd");
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
                        body.append(paragraph.getText());
                    }
                } else {
                    body.append(paragraph.getText());
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
                        body.append(table.getText());
                    }
                } else {
                    body.append(table.getText());
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

    public String getJson() throws JsonProcessingException
    {
        ObjectMapper objectMapper = new ObjectMapper();
        String json = objectMapper.writeValueAsString(getSections());

        return json;
    }

    public void writeToFile() throws IOException
    {
        FileWriter fileWriter = new FileWriter(output);
        fileWriter.write(getJson());
        fileWriter.close();
    }
    public String index() throws IOException {
        RestClient restClient = RestClient.builder(
                new HttpHost("search-elastic-test-yyco5dncwicwd2nufqhakzek2e.us-east-1.es.amazonaws.com", 443, "https")).build();

        String actionMetaData = String.format("{ \"index\" : { \"_index\" : \"%s\", \"_type\" : \"%s\" } }%n", "rfps", "rfp");

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
                        body.append(paragraph.getText() + " ");
                    }
                } else {
                    body.append(paragraph.getText() + " ");
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
                        body.append(table.getText() + " ");
                    }
                } else {
                    body.append(table.getText() + " ");
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

    public XWPFDocument getXdoc()
    {
        return xdoc;
    }



}
