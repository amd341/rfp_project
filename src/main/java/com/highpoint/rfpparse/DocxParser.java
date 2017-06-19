package com.highpoint.rfpparse;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.entity.ContentType;
import org.apache.http.nio.entity.NStringEntity;
import org.apache.http.util.EntityUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.elasticsearch.client.Response;
import org.elasticsearch.client.RestClient;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by Brenden Sosnader on 6/6/17.
 * DocxParser class for parsing files in Microsoft Word format into sections based on Heading 1's and Heading 2's
 * as well as uploading parsed sections to an elasticsearch instance
 */
public class DocxParser {

    private XWPFDocument xdoc;
    private Map<String,Object> entries;


    /**
     * @param input path to docx file to be parsed
     * @param entries optional key/value pairs to be added to index for greater classification
     * @throws IOException if input filepath is wrong
     * @throws InvalidFormatException if file is not a docx
     */
    public DocxParser(InputStream input, Map<String,Object> entries) throws IOException, InvalidFormatException {
        xdoc = new XWPFDocument(OPCPackage.open(input));
        this.entries = entries;

    }

    /**
     * bulk indexes sections from word document
     * @param hostname hostname of elasticsearch instance
     * @param port port for accessing
     * @param scheme scheme of access
     * @param index name of index to add data to
     * @param type name of type to add data to
     * @return String response from elasticsearch
     * @throws IOException if hostname is incorrect
     */
    public String bulkIndex(String hostname, int port, String scheme, String index, String type, List<String> bulkData) throws IOException {
        RestClient restClient = RestClient.builder(
                new HttpHost(hostname, port, scheme)).build();

        String actionMetaData = String.format("{ \"index\" : { \"_index\" : \"%s\", \"_type\" : \"%s\" } }%n", index, type);

        StringBuilder bulkRequestBody = new StringBuilder();
        for (String bulkItem : bulkData) {
            bulkRequestBody.append(actionMetaData);
            bulkRequestBody.append(bulkItem);
            bulkRequestBody.append("\n");
        }
        HttpEntity entity = new NStringEntity(bulkRequestBody.toString(), ContentType.APPLICATION_JSON);

        Response response = restClient.performRequest("POST", "/"+index+"/"+type+"/_bulk",
                Collections.emptyMap(), entity);
        restClient.close();
        return response.toString();
    }

    /**
     * adds any new keys as keywords to elasticsearch. this is assuming there is already a default
     * mapping which should have basic, essential fields like body or date in whichever datatype you prefer
     * @param hostname hostname of elasticsearch instance
     * @param port port for accessing
     * @param scheme scheme of access
     * @param index name of index to add data to
     * @param type name of type to add data to
     * @return String response from server
     * @throws IOException if connection or request is bad
     */
    public String mapNewFields(String hostname, int port, String scheme, String index, String type) throws IOException {
        RestClient restClient = RestClient.builder(
                new HttpHost(hostname, port, scheme)).build();

        Response response = restClient.performRequest("GET", "/"+index+"/"+type+"/_mapping");
        ObjectMapper objectMapper = new ObjectMapper();

        JsonNode root = objectMapper.readTree(EntityUtils.toString(response.getEntity()));

        for (String key : entries.keySet())
        {
            if(root.findPath(key).isMissingNode())
            {
                String newMapping = String.format("{ \"properties\": { \"%s\": { \"type\": \"keyword\" } } } ", key);
                HttpEntity entity = new NStringEntity(newMapping, ContentType.APPLICATION_JSON);
                response = restClient.performRequest("PUT", "/"+index+"/"+type+"/_mapping", Collections.emptyMap(), entity);
            }
        }
        restClient.close();
        return EntityUtils.toString(response.getEntity());

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
                            section.putAll(entries);
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
                            section.putAll(entries);
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
        section.putAll(entries);
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
                            section.putAll(entries);
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
                        if(body.length() > 0)
                        {
                            section.put("body", body.toString());
                            section.put("headingOne", headingOne);
                            section.put("headingTwo", headingTwo);
                            section.putAll(entries);
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
        section.putAll(entries);
        sections.add(section);

        return sections;
    }

    private List<String> getSectionsAsStrings() throws JsonProcessingException {
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
                            section.putAll(entries);
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
                            section.putAll(entries);
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
        section.putAll(entries);
        sections.add(objectMapper.writeValueAsString(section));
        return sections;
    }

    private List<String> getSubSectionsAsStrings() throws JsonProcessingException {
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
                            section.putAll(entries);
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
                            section.putAll(entries);
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
        section.putAll(entries);
        sections.add(objectMapper.writeValueAsString(section));

        return sections;
    }

    public List<Map<String,Object>> getHighlighted() {
        Map<String,Object> section = new HashMap<>();
        List<Map<String,Object>> sections = new ArrayList<>();
        List<IBodyElement> elements = xdoc.getBodyElements();
        boolean gotQuestion = false;
        String body;
        String question = "no question";

        for(IBodyElement element : elements)
        {
            if(element.getElementType().toString().equals("PARAGRAPH"))
            {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                for (XWPFRun run : paragraph.getRuns()) {
                    if (run.isHighlighted()) {
                        if (run.getColor() != null) {
                            if (run.getColor().equals("FF0000") && !gotQuestion) {
                                question = paragraph.getText();
                                gotQuestion = true;

                            } else if (run.getColor().equals("0070C0") && gotQuestion) {
                                body = paragraph.getText();
                                section.put("body", body);
                                section.put("question", question);
                                section.putAll(entries);
                                sections.add(section);
                                section = new HashMap<>();
                                gotQuestion = false;
                            }
                        }
                    }
                }

            }
            else if(element.getElementType().toString().equals("TABLE"))
            {
                XWPFTable table = (XWPFTable) element;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for(XWPFRun run : paragraph.getRuns()) {
                                if (run.isHighlighted()) {
                                    if (run.getColor() != null) {
                                        if (run.getColor().equals("FF0000") && !gotQuestion) {
                                            question = paragraph.getText();
                                            gotQuestion = true;

                                        } else if (run.getColor().equals("0070C0") && gotQuestion) {
                                            body = cell.getText();
                                            section.put("body", body);
                                            section.put("question", question);
                                            section.putAll(entries);
                                            sections.add(section);
                                            section = new HashMap<>();
                                            gotQuestion = false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        return sections;
    }

    public List<String> getHighlightedAsStrings() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,Object> section = new HashMap<>();
        List<String> sections = new ArrayList<>();
        List<IBodyElement> elements = xdoc.getBodyElements();
        boolean gotQuestion = false;
        String body;
        String question = "no question";

        for(IBodyElement element : elements)
        {
            if(element.getElementType().toString().equals("PARAGRAPH"))
            {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                for (XWPFRun run : paragraph.getRuns()) {
                    if (run.isHighlighted()) {
                        if (run.getColor() != null) {
                            if (run.getColor().equals("FF0000") && !gotQuestion) {
                                question = paragraph.getText();
                                gotQuestion = true;

                            } else if (run.getColor().equals("0070C0") && gotQuestion) {
                                body = paragraph.getText();
                                section.put("body", body);
                                section.put("question", question);
                                section.putAll(entries);
                                sections.add(objectMapper.writeValueAsString(section));
                                section = new HashMap<>();
                                gotQuestion = false;
                            }
                        }
                    }
                }

            }
            else if(element.getElementType().toString().equals("TABLE"))
            {
                XWPFTable table = (XWPFTable) element;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for(XWPFRun run : paragraph.getRuns()) {
                                if (run.isHighlighted()) {
                                    if (run.getColor() != null) {
                                        if (run.getColor().equals("FF0000") && !gotQuestion) {
                                            question = paragraph.getText();
                                            gotQuestion = true;

                                        } else if (run.getColor().equals("0070C0") && gotQuestion) {
                                            body = cell.getText();
                                            section.put("body", body);
                                            section.put("question", question);
                                            section.putAll(entries);
                                            sections.add(objectMapper.writeValueAsString(section));
                                            section = new HashMap<>();
                                            gotQuestion = false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        return sections;
    }
}
