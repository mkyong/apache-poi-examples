package com.mkyong.poi.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class ReadParseDocument {

    public static void main(String[] args) throws IOException {

        String fileName = "c:\\test\\simple.docx";

        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get(fileName)))) {

            /*
            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(doc);
            String docText = xwpfWordExtractor.getText();
            System.out.println(docText);

            // find number of words in the document
            long count = Arrays.stream(docText.split("\\s+")).count();
            System.out.println("Total words: " + count);
            */

            List<XWPFParagraph> list = doc.getParagraphs();
            for (XWPFParagraph paragraph : list) {
                System.out.println(paragraph.getText());
            }

        }

    }

}
