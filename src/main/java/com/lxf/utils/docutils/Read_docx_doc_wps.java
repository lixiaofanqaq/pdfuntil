package com.lxf.utils.docutils;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 用poi读取doc,docx文档内容检测关键字
 */
public class Read_docx_doc_wps {
    // 测试
    public static void main(String[] args) throws IOException {
        String filePath = "D:\\LXF\\Test\\doctest.doc";
        String filePath2 = "D:\\LXF\\Test\\docxtest.docx";
        String filePath3 = "D:\\LXF\\Test\\wpstest.wps";

//        String docStr = getDocStr(filePath);
//        String docxStr = getDocxStr(filePath2);
//        String wpsStr = getWpsStr(filePath3);

//        System.out.println(FileTypeInspect.KeywordInspect(docStr, "百年"));
//        System.out.println(FileTypeInspect.KeywordInspect(docxStr, "奋斗"));
//        System.out.println(FileTypeInspect.KeywordInspect(wpsStr, "奋斗"));

//        for (String s : getDocxImages(filePath2)) {
//            System.out.println(s);
//        }
        for (String s : getDocxImages(filePath2)) {
            System.out.println(s);
        }
    }

    /**
     * 读取doc文件内容，将文档内容转换为string字符串输出
     *
     * @param filePath doc文件路径
     * @return String docStr
     * @throws IllegalArgumentException 非DOC文件异常
     */
    public static String getDocStr(String filePath) {
        File file = new File(filePath);
        StringBuilder docStr = new StringBuilder();
        try {
            if (file.exists() && file.isFile()) {
                FileInputStream fis = new FileInputStream(file);
                WordExtractor extractor = new WordExtractor(fis);

                docStr.append(extractor.getText());
                fis.close();
                extractor.close();
                System.out.println("Doc_Wps_Result:\n" + docStr);
            } else {
                System.out.println("DOC/WPS 文件不存在");
            }
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
            return "noDoc";
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return docStr.toString();
    }

    /**
     * 读取docx文件内容,将文档内容转换为string字符串输出
     *
     * @param filePath docx文件路径
     * @return String docxStr
     */
    public static String getDocxStr(String filePath) {
        File file = new File(filePath);
        StringBuilder docxStr = new StringBuilder();
        try {
            if (file.exists() && file.isFile()) {
                FileInputStream fis = new FileInputStream(file);
                XWPFDocument xdoc = new XWPFDocument(new FileInputStream(file));
                XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);

                docxStr.append(extractor.getText());

                fis.close();
                xdoc.close();
                extractor.close();
                System.out.println("Docx_Result:\n" + docxStr);
            } else {
                System.out.println("DOCX 文件不存在");
            }
        } catch (POIXMLException e) {
            e.printStackTrace();
            return "noDocx";
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
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

    /**
     * 读取docx 文件中的图片,并输出其路径。
     *
     * @param filePath docx文件路径
     * @return list
     */
    public static List<String> getDocxImages(String filePath) {
        List<String> imagesList = new ArrayList<>();
        File file = new File(filePath);
        try {
            if (file.exists() && file.isFile()) {
                // 用XWPFWordExtractor来获取文字
                FileInputStream fis = new FileInputStream(file);
                XWPFDocument docx = new XWPFDocument(fis);
                //获取去除后缀的文件路径
                String fileDirectory = filePath.substring(0, filePath.lastIndexOf("."));
                File imagesFile = new File(fileDirectory);
                // 用XWPFDocument的getAllPictures来获取所有的图片
                List<XWPFPictureData> picList = docx.getAllPictures();
                for (XWPFPictureData pic : picList) {
                    // System.out.println(pic.getPictureType() + File.separator + pic.suggestFileExtension() + File.separator + pic.getFileName());
                    byte[] imageByte = pic.getData();
//                System.out.println("imageByte:" + imageByte.length);
                    //如果文件夹不存在
                    if (!imagesFile.exists()) {
                        //创建文件夹
                        imagesFile.mkdir();
                    }
                    // 大于 xx bites的图片我们才弄下来，消除word中莫名的小图片的影响
                    if (imageByte.length > 100) {
                        FileOutputStream fos = new FileOutputStream(fileDirectory + "\\" + pic.getFileName());
                        fos.write(imageByte);
                        imagesList.add(fileDirectory + "\\" + pic.getFileName());
                    }
                }
                fis.close();
                docx.close();
            } else {
                System.out.println("DOCX 文件不存在");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return imagesList;
    }

    /**
     * 读取doc文件中的图片，并输出其路径
     *
     * @param filePath doc文件路径
     * @return List
     */
    public static List<String> getDocImages(String filePath) {
        List<String> imagesList = new ArrayList<>();
        File file = new File(filePath);
        try {
            if (file.exists() && file.isFile()) {
                InputStream is = new FileInputStream(file);
                HWPFDocument doc = new HWPFDocument(is);

                //文档图片内容
                PicturesTable picturesTable = doc.getPicturesTable();
                List<Picture> allPictures = picturesTable.getAllPictures();
                //获取去除后缀的文件路径
                String fileDirectory = filePath.substring(0, filePath.lastIndexOf("."));
                File imagesFile = new File(fileDirectory);
                // System.out.println(fileDirectory);
                int imageCount = 0;
                String imagePathName = "";
                for (Picture picture : allPictures) {
                    //如果文件夹不存在
                    if (!imagesFile.exists()) {
                        //创建文件夹
                        imagesFile.mkdir();
                    }
                    imageCount++;
                    // 输出图片到磁盘
                    imagePathName = fileDirectory + "\\image" + imageCount + "." + picture.suggestFileExtension();
                    FileOutputStream fos = new FileOutputStream(new File(imagePathName));
                    picture.writeImageContent(fos);
                    imagesList.add(imagePathName);
                    fos.close();
                }
            } else {
                System.out.println("doc/wps 文件不存在");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return imagesList;
    }

    /**
     * 读取wps文件中的图片，并输出其路径
     *
     * @param filePath wps文件路径
     * @return List
     * @throws IOException IO异常
     */
    public static List<String> getWpsImages(String filePath) throws IOException {
        return getDocImages(filePath);
    }

}
