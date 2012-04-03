package com.trebogeer.xlsx.stream;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.io.IOException;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static com.trebogeer.xlsx.stream.XLSXConstants.SHEET_LOCATION;

/**
 * @author dimav
 *         Date: 6/21/11
 *         Time: 2:43 PM
 */
class SpreadSheetWriter {

    private final XMLStreamWriter out;
    private int rownum = -1;
    private List<? extends ColumnHeader> columnHeaders = null;
    private final ZipOutputStream zipOutputStream;

    SpreadSheetWriter(ZipOutputStream out, int index) {
        try {
            this.zipOutputStream = out;
            this.out = XMLOutputFactory.newInstance().createXMLStreamWriter(out);
            zipOutputStream.putNextEntry(new ZipEntry(SHEET_LOCATION + "sheet" + index + ".xml"));
        } catch (Exception e) {
            throw new RuntimeException("Failed to create writer.", e);
        }
    }

    SpreadSheetWriter(ZipOutputStream out, int index, List<? extends ColumnHeader> columnHeaders) {
        this(out, index);
        this.columnHeaders = columnHeaders;
    }

    public void writeHeaders() throws XMLStreamException {
        if (columnHeaders != null && !columnHeaders.isEmpty()) {
            insertNextRow();
            for (int i = 0; i < columnHeaders.size(); i++) {
                ColumnHeader header = columnHeaders.get(i);
                createCell(i, header.getName());
            }
            endRow();
        }
    }

    public void beginSheet() throws XMLStreamException {
        //<?xml version="1.0" encoding="UTF-8"?>
        out.writeStartDocument("1.0");
        out.writeStartElement("worksheet");
        out.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writeColumnDefinition();
        out.writeStartElement("sheetData");
        writeHeaders();
    }

    public void endSheet() throws XMLStreamException, IOException {
        out.writeEndElement();
        writeValidations();
        out.writeEndElement();
        out.writeEndDocument();
        out.close();
        zipOutputStream.closeEntry();
    }

    private void writeValidations() throws XMLStreamException {
        // out.write();
    }

    private void writeColumnDefinition() throws XMLStreamException {
         if (columnHeaders != null && !columnHeaders.isEmpty()) {
            out.writeStartElement("cols");
            for (int i = 0; i < columnHeaders.size(); i++) {
                out.writeEmptyElement("col");
                out.writeAttribute("collapsed", "false");
                out.writeAttribute("hidden", "false");
                out.writeAttribute("min", String.valueOf(i + 1));
                out.writeAttribute("max", String.valueOf(i + 1));
                //out.writeAttribute("style", "1");
                out.writeAttribute("width", "22.5098039215686");
            }
            out.writeEndElement();
        }
    }

    public void createCell(int columnIndex, String value, int styleIndex) throws XMLStreamException {
        String ref = ExcelUtils.getCellReference(columnIndex, rownum);
        out.writeStartElement("c");
        out.writeAttribute("r", ref);
        out.writeAttribute("t", "inlineStr");
        if (styleIndex != -1) out.writeAttribute("s", String.valueOf(styleIndex));
        out.writeStartElement("is");
        out.writeStartElement("t");
        out.writeCharacters(value);
        out.writeEndElement();
        out.writeEndElement();
        out.writeEndElement();
    }

    public void createCell(int columnIndex, String value) throws XMLStreamException {
        createCell(columnIndex, value, -1);
    }

    public void createCell(int columnIndex, double value, int styleIndex) throws XMLStreamException {
        String ref = ExcelUtils.getCellReference(columnIndex, rownum);
        out.writeStartElement("c");
        out.writeAttribute("r", ref);
        out.writeAttribute("t", "n");
        if (styleIndex != -1) out.writeAttribute("s", String.valueOf(styleIndex));
        out.writeStartElement("v");
        out.writeCharacters(String.valueOf(value));
        out.writeEndElement();
        out.writeEndElement();
    }

    public void createCell(int columnIndex, double value) throws XMLStreamException {
        createCell(columnIndex, value, -1);
    }

    public void insertNextRow() throws XMLStreamException {
        rownum++;
        out.writeStartElement("row");
        out.writeAttribute("r", String.valueOf(rownum + 1));
    }

    /**
     * Insert row end marker
     *
     * @throws javax.xml.stream.XMLStreamException xml writing exception
     */
    public void endRow() throws XMLStreamException {
        out.writeEndElement();
    }
}
