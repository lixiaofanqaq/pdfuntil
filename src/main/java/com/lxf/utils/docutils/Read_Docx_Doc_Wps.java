package com.lxf.utils.docutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import com.lxf.utils.documentInspect.FileTypeInspect;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 用poi读取doc,docx文档内容检测关键字
 */
public class Read_Docx_Doc_Wps {
    // 测试
    public static void main(String[] args) throws IOException {
//        String filePath = "D:\\LXF\\Test\\doctest.doc";
//        String docStr = getDocStr(filePath);
//        System.out.println(FileTypeInspect.KeywordInspect(docStr, "百年"));
//
//        String filePath2 = "D:\\LXF\\Test\\doctest.docx";
//        String docxStr = getDocxStr(filePath2);
//        System.out.println(FileTypeInspect.KeywordInspect(docStr, "奋斗"));

        String filePath3 = "D:\\LXF\\Test\\wpstest.wps";
        String wpsStr = getWpsStr(filePath3);
        System.out.println(FileTypeInspect.KeywordInspect(wpsStr, "奋斗"));
    }

    /**
     * 读取doc文件内容，将文档内容转换为string字符串输出
     *
     * @param filePath doc文件路径
     * @return String docStr
     * @throws IOException IO异常
     */
    public static String getDocStr(String filePath) throws IOException {
        File file = new File(filePath);
        StringBuilder docStr = new StringBuilder();
        if (file.exists()) {
            FileInputStream fis = new FileInputStream(file);
            WordExtractor extractor = new WordExtractor(fis);

            docStr.append(extractor.getText());
            fis.close();
            extractor.close();
//          System.out.println("Doc_Wps_Result:\n" + docStr);
        } else {
            System.out.println("DOC/WPS 文件不存在");
        }
        return docStr.toString();
    }

    /**
     * 读取docx文件内容,将文档内容转换为string字符串输出
     *
     * @param filePath docx文件路径
     * @return String docxStr
     * @throws IOException IO异常
     */
    public static String getDocxStr(String filePath) throws IOException {
        File file = new File(filePath);
        StringBuilder docxStr = new StringBuilder();
        if (file.exists()) {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument xdoc = new XWPFDocument(new FileInputStream(file));
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            docxStr.append(extractor.getText());

            fis.close();
            xdoc.close();
            extractor.close();
//          System.out.println("Docx_Result:\n" + docxStr);
        } else {
            System.out.println("DOCX 文件不存在");
        }
        return docxStr.toString();
    }

    /**
     * 读取wps文件内容，将文档内容转换为string字符串输出
     * wps 与doc本质一致，用一个方法即可读取
     *
     * @param filePath 文件路径
     * @return String wpsStr
     * @throws IOException IO异常
     */
    public static String getWpsStr(String filePath) throws IOException {
        return getDocStr(filePath);
    }
}
