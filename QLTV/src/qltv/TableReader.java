/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package qltv;

import java.io.*;
import java.util.Iterator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
//import org.apache.poi.hwpf.HWPFDocument;
//import org.apache.poi.hwpf.usermodel.*;
//import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class TableReader {

    public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream("C:\\Users\\Pham Ngoc Minh\\Desktop\\TestWord.docx");
            XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
            Iterator bodyElementIterator = xdoc.getBodyElementsIterator();
            while (bodyElementIterator.hasNext()) {
                IBodyElement element = (IBodyElement) bodyElementIterator.next();

                if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                    java.util.List<XWPFTable> tableList = element.getBody().getTables();
                    for (XWPFTable table : tableList) {
                        System.out.println("Total Number of Rows of Table:" + table.getNumberOfRows());
                        for (int i = 0; i < table.getRows().size(); i++) {
                            for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
                                removeParagraphs(table.getRow(i).getCell(j));
                                XWPFParagraph paragraph = table.getRow(i).getCell(j).addParagraph();
                                paragraph.createRun().setText("Phạm Ngọc Minh");
                            }
                        }
                    }
                }
            }
            OutputStream out = new FileOutputStream("C:\\Users\\Pham Ngoc Minh\\Desktop\\TestWord.docx");
            xdoc.write(out);
            out.close();
        } catch (IOException | InvalidFormatException ex) {
        }

    }

    private static void removeParagraphs(XWPFTableCell tableCell) {
        int count = tableCell.getParagraphs().size();
        for (int i = 0; i < count; i++) {
            tableCell.removeParagraph(i);
        }
    }
}
