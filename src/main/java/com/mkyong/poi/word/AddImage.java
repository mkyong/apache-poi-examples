package com.mkyong.poi.word;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class AddImage {

    public static void main(String[] args) throws IOException, InvalidFormatException {

        String imgFile = "c:\\test\\google.png";

        try (XWPFDocument doc = new XWPFDocument()) {

            XWPFParagraph p = doc.createParagraph();
            XWPFRun r = p.createRun();
            r.setText(imgFile);
            r.addBreak();

            // add png image
            try (FileInputStream is = new FileInputStream(imgFile)) {
                r.addPicture(is,
                        Document.PICTURE_TYPE_PNG,         // png file
                        imgFile,
                        Units.toEMU(400),
                        Units.toEMU(200));          // 400x200 pixels
            }

            try (FileOutputStream out = new FileOutputStream("c:\\test\\images.docx")) {
                doc.write(out);
            }
        }

    }
}
