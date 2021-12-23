package com.lxf.utils.pdfutils;

import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TransitionPdf {

    public static void main(String[] args) throws TesseractException {
//        路径示范:String filePath = "D://Test/xxx.pdf";
        String filePath = "D:/LXF/PdfTest/Test2.pdf";
        List<String> imageList = pdfToImage(filePath);
        for (String s : imageList) {
            System.out.println(s);
        }
//        List<String> wordlist = imageToText(imageList);
//        for (String s : wordlist) {
//            System.out.println(s);
//        }

//        boolean isPass = SelectKeyword(TextList, "理论");
//        System.out.println(isPass);

//        String result = FindOCR("D:/LXF/PdfTest/Test2/page_3.png");
//        System.out.println(result);

//        List<String> list = new ArrayList<>();
//        list.add("理论是行动的先导，思想是前进的旗帜。");
//        list.add("党的十九届六中全会审议通过");
//        String keyWord = "理论";
//        boolean isPass = SelectKeyword(list, keyWord);
//        System.out.println(isPass);
    }

    /**
     * PDF文件转换为后缀为PNG或JPG的图片,按页转换.
     *
     * @param filePath PDF路径
     * @return List<String> 图片路径集合
     */
    public static List<String> pdfToImage(String filePath) {
        String imagePath;
        List<String> list = new ArrayList<>();
        if (!"".equals(filePath)) {
            String fileDirectory = filePath.substring(0, filePath.lastIndexOf("."));//获取去除后缀的文件路径

            File file = new File(filePath);
            try {
                File f = new File(fileDirectory);
                if (!f.exists()) {
                    f.mkdir();
                }
                PDDocument doc = PDDocument.load(file);
                PDFRenderer renderer = new PDFRenderer(doc);
                int pageCount = doc.getNumberOfPages();
                for (int i = 0; i < pageCount; i++) {
                    // 方式1,第二个参数是设置缩放比(即像素)
                    // BufferedImage image = renderer.renderImageWithDPI(i, 296);
                    // 方式2,第二个参数是设置缩放比(即像素)
                    BufferedImage image = renderer.renderImage(i, 3.5f);  //第二个参数越大生成图片分辨率越高，转换时间也就越长
                    imagePath = fileDirectory + "/page_" + (i + 1) + ".png";
                    ImageIO.write(image, "PNG", new File(imagePath));
                    list.add(imagePath);
                }
                doc.close();
                System.out.println("PDF转换图片成功!");
                System.out.println("图片存放路径:" + fileDirectory);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("文件路径不存在!!");
        }
        return list;
    }

    /**
     * 识别图片里的文字并存到list集合里  识别率不高
     *
     * @param imagePathList 图片地址集合
     * @return list<String>
     */
    public static List<String> imageToText(List<String> imagePathList) throws TesseractException {

        List<String> list = new ArrayList<>();

        //将图片地址里的图片写入到List<String>中,使用OCR识别
        ITesseract instance = new Tesseract();

        //如果未将tessdata放在根目录下需要指定绝对路径
        // instance.setDatapath("C:/LXF/Idea/IdeaProjects/pdfuntil/tessdata");
        //如果需要识别英文之外的语种，需要指定识别语种，并且需要将对应的语言包放进项目中
        instance.setLanguage("chi_sim");

        // 指定识别图片
        for (String imagePath : imagePathList) {
            File imgDir = new File(imagePath);
            long startTime = System.currentTimeMillis();
            if (!imgDir.exists()) {
                System.out.println("图片不存在");
            }
            String ocrResult = instance.doOCR(imgDir);
            // 输出识别结果
            System.out.println("OCR Result: \n" + ocrResult + "\n 耗时：" + (System.currentTimeMillis() - startTime) / 1000 + "s");
            list.add(ocrResult);
        }
        return list;
    }

    /**
     * 识别图片信息 使用中文训练裤
     * @param srImage 图片
     * @return 文字信息
     */
    public static String FindOCR(String srImage) {
        try {
            System.out.println("start");
            double start = System.currentTimeMillis();
            File imageFile = new File(srImage);
            if (!imageFile.exists()) {
                return "图片不存在";
            }
            BufferedImage textImage = ImageIO.read(imageFile);
            Tesseract instance = Tesseract.getInstance();
            instance.setDatapath("C:/LXF/Idea/IdeaProjects/pdfuntil/tessdata");//设置训练库
            instance.setLanguage("chi_sim");//中文识别
            String result = null;
            result = instance.doOCR(textImage);
            double end = System.currentTimeMillis();
            System.out.println("耗时" + (end - start) / 1000 + " s");
            return result;
        } catch (Exception e) {
            e.printStackTrace();
            return "发生未知错误";
        }
    }

    /**
     * 判断 List<String> 集合里的元素是否包含某个关键字
     *
     * @param list String集合
     * @param keyword 关键字
     * @return true false
     */
    public static boolean SelectKeyword(List<String> list, String keyword) {
        int keywordCount = 0;
        for (String s : list) {
            if (s.contains(keyword)) {
                keywordCount++;
                System.out.println("包含关键字:" + keyword + "  count:" + keywordCount);
            }
        }
        return keywordCount == 0;
    }
}


