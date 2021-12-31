package com.lxf.utils.docutils;

import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Read_ppt_pptx_dps {
/*    public static void main(String[] args) throws IOException {
        String filePath = "D:\\LXF\\Test\\111p.ppt";
        String filePath2 = "D:\\LXF\\Test\\111x.pptx";
        String filePath3 = "D:\\LXF\\Test\\111d.dps";
        String filePath4 = "D:\\LXF\\Test\\111.txt";
//        System.out.println(getPptStr(filePath));
//        System.out.println(getPptxStr(filePath2));
//        System.out.println(getDpsStr(filePath3));
        for (String image : getAllImages(filePath4)) {
            System.out.println(image);
        }
    }*/

    /**
     * 识别PPT中文字与表格，图片暂时识别不到，会跳过。
     *
     * @param filePath 文件路径
     * @return String pptStr
     */
    public static String getPptStr(String filePath) {
        File file = new File(filePath);
        StringBuilder pptStr = new StringBuilder();
        try {
            if (file.exists() && file.isFile()) {
                FileInputStream fis = new FileInputStream(file);
                HSLFSlideShow hslfSlideShow = new HSLFSlideShow(fis);
                SlideShowExtractor<HSLFShape, HSLFTextParagraph> slideShowExtractor = new SlideShowExtractor<>(hslfSlideShow);
                List<HSLFSlide> slides = hslfSlideShow.getSlides();

                for (HSLFSlide slide : slides) {
                    pptStr.append(slideShowExtractor.getText(slide));
                }
                hslfSlideShow.close();
                slideShowExtractor.close();
                fis.close();
//            System.out.println(pptStr.toString());
            } else {
                System.out.println("PPT 文件内容为空");
            }
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
            return "noPpt";
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return pptStr.toString();
    }

    /**
     * 识别PPTX中文字与表格，图片暂时识别不到，会跳过。
     *
     * @param filePath 文件路径
     * @return String字符串
     */
    public static String getPptxStr(String filePath) {
        File file = new File(filePath);
        StringBuilder pptxStr = new StringBuilder();
        try {
            if (file.exists() && file.isFile()) {
                FileInputStream fis = new FileInputStream(file);
                XMLSlideShow xmlSlideShow = new XMLSlideShow(fis);
                // 获得每一张幻灯片
                List<XSLFSlide> slides = xmlSlideShow.getSlides();

                for (XSLFSlide slide : slides) {
                    CTSlide rawSlide = slide.getXmlObject();
                    CTGroupShape spTree = rawSlide.getCSld().getSpTree();
                    List<CTShape> spList = spTree.getSpList();

                    for (CTShape shape : spList) {
                        CTTextBody txBody = shape.getTxBody();

                        if (txBody == null) {
                            continue;
                        }
                        List<CTTextParagraph> pList = txBody.getPList();

                        for (CTTextParagraph textParagraph : pList) {
                            List<CTRegularTextRun> textRuns = textParagraph.getRList();

                            for (CTRegularTextRun textRun : textRuns) {
                                pptxStr.append(textRun.getT());
                            }
                        }
                    }
                }
                fis.close();
                xmlSlideShow.close();
//            System.out.println(pptxStr.toString());
            } else {
                System.out.println("PPTX 文件内容为空");
            }
        } catch (POIXMLException e) {
            e.printStackTrace();
            return "noPptx";
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return pptxStr.toString();
    }

    /**
     * 识别DPS中文字与表格，图片暂时识别不到，会跳过。
     * DPS 与 PPT  文件本质一样。
     *
     * @param filePath 文件路径
     * @return String dpsStr
     */
    public static String getDpsStr(String filePath) {
        return getPptStr(filePath);
    }

    /**
     * 识别 ppt,pptx,dps文件中的图片并将路径存储到List集合里
     *
     * @param filePath 文件路径
     * @return list
     */
    public static List<String> getAllImages(String filePath) throws IOException {
        List<String> imageList = new ArrayList<>();
        File file = new File(filePath);
        InputStream is = new FileInputStream(file);
        SlideShow slideShow = null;
        try {
            if (file.exists() && file.isFile()) {
                if (filePath.endsWith(".ppt") || filePath.endsWith(".dps")) {
                    slideShow = new HSLFSlideShow(is);
                } else if (filePath.endsWith(".pptx")) {
                    slideShow = new XMLSlideShow(is);
                }
                if (slideShow != null) {
                    // 图片内容
                    List<?> pictures = slideShow.getPictureData();
                    //获取去除后缀的文件路径
                    getImagePath(filePath, imageList, pictures);
                } else {
                    System.out.println("null,文件格式不支持");
                    return null;
                }
            } else {
                System.out.println("ppt/pptx/dps 文件不存在");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (slideShow != null) {
                slideShow.close();
            }
            is.close();
            System.gc();
        }
        return imageList;
    }

    /**
     * 复用通用代码
     *
     * @param filePath  文件路径
     * @param imageList 图片路径集合
     * @param pictures  图片
     */
    public static void getImagePath(String filePath, List<String> imageList, List<?> pictures) {
        String fileDirectory = filePath.substring(0, filePath.lastIndexOf("."));
        File imagesFile = new File(fileDirectory);
        int imageCount = 0;
        String imagePathName = "";
        try {
            if (pictures != null && !pictures.isEmpty()) {
                if (!imagesFile.exists()) {
                    //创建文件夹
                    imagesFile.mkdir();
                }
                for (Object object : pictures) {
                    imageCount++;
                    PictureData picture = (PictureData) object;
                    byte[] data = picture.getData();
                    imagePathName = fileDirectory + "\\image" + imageCount + ".png";
                    FileOutputStream fos = new FileOutputStream(new File(imagePathName));
                    fos.write(data);
                    fos.close();
                    imageList.add(imagePathName);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
    }

}
