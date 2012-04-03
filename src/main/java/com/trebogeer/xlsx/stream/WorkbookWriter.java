package com.trebogeer.xlsx.stream;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import static com.trebogeer.xlsx.stream.XLSXConstants.RID;
import static com.trebogeer.xlsx.stream.XLSXConstants.WORKBOOK_LOCATION;
import static com.trebogeer.xlsx.stream.XLSXConstants.WORKBOOK_REL_LOCATION;
import static com.trebogeer.xlsx.stream.XLSXConstants.WORKSHEET_NAMESPACE;

/**
 * @author dimav
 *         Date: 6/21/11
 *         Time: 3:15 PM
 */
public class WorkbookWriter {

    private final ZipFile template;
    private final ZipOutputStream zos;
    private Document workbook;
    private Document workbookRelations;
    private Transformer transformer;

    public static WorkbookWriter getWorkbookWriter(final ZipFile template, final ZipOutputStream os) throws IOException, SAXException, ParserConfigurationException, TransformerConfigurationException {
        assert template != null;
        assert os != null;
        WorkbookWriter wbw = new WorkbookWriter(template, os);
        wbw.init();
        return wbw;
    }

    private WorkbookWriter(ZipFile template, ZipOutputStream os) {
        assert template != null;
        assert os != null;
        this.template = template;
        this.zos = os;
    }

    private void init() throws IOException, SAXException, ParserConfigurationException, TransformerConfigurationException {
        transformer = TransformerFactory.newInstance().newTransformer();
        Enumeration<? extends ZipEntry> templateEntries = template.entries();
        while (templateEntries.hasMoreElements()) {

            ZipEntry entry = templateEntries.nextElement();
            if (WORKBOOK_LOCATION.equals(entry.getName())) {
                workbook = readXmlZipEntryToDOM(entry);
            } else if (WORKBOOK_REL_LOCATION.equals(entry.getName())) {
                workbookRelations = readXmlZipEntryToDOM(entry);
            } else {
                ZipEntry newEntry = new ZipEntry(entry.getName());

                zos.putNextEntry(newEntry);

                BufferedInputStream bis = new BufferedInputStream(template.getInputStream(entry));

                while (bis.available() > 0) {
                    zos.write(bis.read());
                }
                zos.closeEntry();
                bis.close();
            }
        }
    }

    private Document readXmlZipEntryToDOM(ZipEntry entry) throws ParserConfigurationException, IOException, SAXException {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        DocumentBuilder documentBuilder = dbf.newDocumentBuilder();
        return documentBuilder.parse(template.getInputStream(entry));
    }

    /*
     <sheet name="Data Definition" sheetId="1" state="visible" r:id="rId2"/>
    */
    public SpreadSheetWriter addNextSheet(String name, boolean isHidden) throws IOException {
        int index = addSheet(name, false);
        return new SpreadSheetWriter(zos, index);
    }


    public SpreadSheetWriter addNextSheet(String name, boolean isHidden, List<? extends ColumnHeader> columnHeaders) throws IOException {
        int index = addSheet(name, isHidden);
        return new SpreadSheetWriter(zos, index, columnHeaders);
    }

    private int addSheet(String name, boolean isHidden) throws IOException {
        NodeList nl = workbook.getElementsByTagName("sheets");
        if (nl != null && nl.getLength() > 0) {
            Node sheets = nl.item(0);
            if (sheets != null) {
                int sheetId = sheets.getChildNodes().getLength() + 1;
                Element newSheet = workbook.createElement("sheet");
                newSheet.setAttribute("name", name);
                newSheet.setAttribute("sheetId", String.valueOf(sheetId));
                String state = isHidden ? SheetState.hidden.name() : SheetState.visible.name();
                newSheet.setAttribute("state", state);
                String rid = RID + getNextRID();
                newSheet.setAttribute("r:id", rid);
                sheets.appendChild(newSheet);

                Element newRelation = workbookRelations.createElement("Relationship");
                newRelation.setAttribute("Type", WORKSHEET_NAMESPACE);
                newRelation.setAttribute("Target", "worksheets/sheet" + sheetId + ".xml");
                newRelation.setAttribute("Id", rid);
                NodeList relations = workbookRelations.getElementsByTagName("Relationships");
                if (relations != null && relations.getLength() > 0) {
                    Node relation = relations.item(0);
                    if (relation != null) {
                        relation.appendChild(newRelation);
                    }
                }
                return sheetId;
            }
        }
        return -1;
    }

    private int getNextRID() {
        NodeList relations = workbookRelations.getElementsByTagName("Relationship");
        if (relations != null) {
            return relations.getLength() + 1;
        }
        return 1;
    }

    private enum SheetState {
        visible,
        hidden
    }

    public void flush() throws IOException, TransformerException {
        flushEntryToZip(workbook, WORKBOOK_LOCATION);
        flushEntryToZip(workbookRelations, WORKBOOK_REL_LOCATION);
    }

    private void flushEntryToZip(Document doc, String location) throws IOException, TransformerException {
        zos.putNextEntry(new ZipEntry(location));
        //create string from xml tree
        OutputStreamWriter sw = new OutputStreamWriter(zos);
        StreamResult result = new StreamResult(sw);
        DOMSource source = new DOMSource(doc);
        transformer.transform(source, result);
        //sw.flush();
        // sw.close();
        zos.closeEntry();
    }

}
