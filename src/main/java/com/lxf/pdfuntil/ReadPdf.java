package com.lxf.pdfuntil;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;


public class ReadPdf {
    public static void main(String[] args) {
        String inputPath = "D:/LXF/PdfTest/";
        String inputFileName = "pdf_Demo.pdf";
        String inputPath_fileName = inputPath + inputFileName;

        String outputPath = "D:/LXF/";
        String outFileName = "demo.txt";
        String outputPath_fileName = outputPath + outFileName;

        readPdf(inputPath_fileName, outputPath_fileName);
    }

    // PDF输入路径，PDF输出路径
    public static void readPdf(String inputPath, String outputPath) {
        if (!"".equals(inputPath) && !"".equals(outputPath)) {
            try (PDDocument document = PDDocument.load(new File(inputPath))) {
                if (!document.isEncrypted()) {
                    PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                    stripper.setSortByPosition(true);
                    PDFTextStripper tStripper = new PDFTextStripper();
                    String pdfFileInText = tStripper.getText(document);
                    String[] lines = pdfFileInText.split("\\r?\\n");
                    FileWriter fw = new FileWriter(outputPath);
                    for (String line : lines) {
                        System.out.println(line);
                        fw.write(line + "\n");
                    }
                    fw.close();
                    document.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("文件路径非法");
            return;
        }
    }
}