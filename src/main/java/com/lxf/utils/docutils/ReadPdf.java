package com.lxf.utils.docutils;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.io.File;
import java.io.IOException;


public class ReadPdf {
    // 测试
/*    public static void main(String[] args) {
        String inputPath = "D:/LXF/PdfTest/";
        String inputFileName = "pdf_Demo.pdf";
        String inputPath_fileName = inputPath + inputFileName;

        System.out.println(getPdfStr(inputPath_fileName));
    }*/

    /**
     * 根据PDF文件路径 获取PDF内文字信息
     *
     * @param filePath 文件路径
     * @return String pdfStr
     */
    public static String getPdfStr(String filePath) {
        File file = new File(filePath);
        StringBuilder pdfFileInText = new StringBuilder();
        if (file.exists()) {
            try (PDDocument document = PDDocument.load(file)) {
                if (!document.isEncrypted()) {
                    PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                    stripper.setSortByPosition(true);
                    PDFTextStripper tStripper = new PDFTextStripper();
                    pdfFileInText.append(tStripper.getText(document));

/*                  String[] lines = pdfFileInText.split("\\r?\\n");
                    FileWriter fw = new FileWriter(outputPath);
                    for (String line : lines) {
                        System.out.println(line);
                        fw.write(line + "\n");
                    }
                    fw.close();
*/
                    document.close();
//                  System.out.println("PDF 文字信息:\n"+pdfFileInText);
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                System.gc();
            }
        } else {
            System.out.println("PDF文件不存在,检查文件路径");
        }
        return pdfFileInText.toString();
    }
}