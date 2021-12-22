package com.lxf.utils.docutils;

import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class Read_ppt_pptx_dps {
    public static void main(String[] args) {
        String filePath = "D:\\LXF\\Test\\111.ppt";
        String filePath2 = "D:\\LXF\\Test\\111.pptx";
        String filePath3 = "D:\\LXF\\Test\\111.dps";
        System.out.println(getPptStr(filePath));
        System.out.println(getPptxStr(filePath2));
        System.out.println(getDpsStr(filePath3));
    }

    /**
     * 识别PPT中文字与表格，图片暂时识别不到，会跳过。
     *
     * @param filePath 文件路径
     * @return String pptStr
     */
    public static String getPptStr(String filePath) {
        File file = new File(filePath);
        StringBuilder pptStr = new StringBuilder();

        if (file.exists()) {
            try {
                FileInputStream fis = new FileInputStream(file);
                HSLFSlideShow hslfSlideShow = new HSLFSlideShow(fis);
                SlideShowExtractor<HSLFShape, HSLFTextParagraph> slideShowExtractor = new SlideShowExtractor<>(hslfSlideShow);
                List<HSLFSlide> slides = hslfSlideShow.getSlides();

                for (HSLFSlide slide : slides) {
                    pptStr.append(slideShowExtractor.getText(slide));
                }
                fis.close();
                hslfSlideShow.close();
                slideShowExtractor.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
//            System.out.println(pptStr.toString());
        } else {
            System.out.println("PPT 文件内容为空");
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

        if (file.exists()) {
            try {
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
            } catch (Exception e) {
                e.printStackTrace();
            }
//            System.out.println(pptxStr.toString());
        } else {
            System.out.println("PPTX 文件内容为空");
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

}
