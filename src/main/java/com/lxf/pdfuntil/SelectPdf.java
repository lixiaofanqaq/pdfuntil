package com.lxf.pdfuntil;

import com.itextpdf.awt.geom.Rectangle2D.Float;
import com.itextpdf.text.pdf.PdfDictionary;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 输入PDF地址以及关键字
 * 可以查找到关键字坐标以及关键字出现次数,返回状态码
 */
@SuppressWarnings("all")
public class SelectPdf {
    public static void main(String[] args) throws IOException {
        String inputPath = "D:/LXF/PdfTest/";
        String inputFileName = "pdf_Demo.pdf";
        String inputPath_fileName = inputPath + inputFileName;
        String keyword = "党";
        //状态码： 合格:200 不合格:403
        int stateCode = isPass(inputPath_fileName, keyword);
        System.out.println("Code:" + stateCode);
    }

    public static int isPass(String path, String keyword) throws IOException {
        if (!"".equals(path) && !"".equals(keyword)) {
            if (getKeyword(path, keyword)) {
                System.out.println("successfully pass!");
                return 200;
            } else {
                System.out.println("invalid characters exist!");
                return 403;
            }
        }
        return 403;
    }

    public static boolean getKeyword(String path, String keyword) throws IOException {
        //1.给定文件
        File pdfFile = new File(path);

        //2.定义一个byte数组，长度为文件的长度
        byte[] pdfData = new byte[(int) pdfFile.length()];

        //3.IO流读取文件内容到byte数组
        try (FileInputStream inputStream = new FileInputStream(pdfFile)) {
            inputStream.read(pdfData);
        } catch (IOException e) {
            throw e;
        }

        //4.调用方法，给定关键字和文件
        List<float[]> positions = findKeywordPostions(pdfData, keyword);

        //5.返回值类型是  List<float[]> 每个list元素代表一个匹配的位置，分别为 float[0]所在页码  float[1]所在x轴 float[2]所在y轴
        System.out.println("total:" + positions.size());
        if (positions.size() > 0) {
            for (float[] position : positions) {
                System.out.print("pageNum: " + (int) position[0]);
                System.out.print("\tx: " + position[1]);
                System.out.println("\ty: " + position[2]);
            }
        }
        //非法字符数量大于1 非法，
        return positions.size() <= 0;
    }

    /**
     * findKeywordPostions
     *
     * @param pdfData 通过IO流 PDF文件转化的byte数组
     * @param keyword 关键字
     * @return List<float [ ]> : float[0]:pageNum float[1]:x float[2]:y
     */
    public static List<float[]> findKeywordPostions(byte[] pdfData, String keyword) throws IOException {
        List<float[]> result = new ArrayList<>();
        List<PdfPageContentPositions> pdfPageContentPositions = getPdfContentPostionsList(pdfData);


        for (PdfPageContentPositions pdfPageContentPosition : pdfPageContentPositions) {
            List<float[]> charPositions = findPositions(keyword, pdfPageContentPosition);
            if (charPositions.size() < 1) {
                continue;
            }
            result.addAll(charPositions);
        }
        return result;
    }

    private static List<PdfPageContentPositions> getPdfContentPostionsList(byte[] pdfData) throws IOException {
        PdfReader reader = new PdfReader(pdfData);

        List<PdfPageContentPositions> result = new ArrayList<>();

        int pages = reader.getNumberOfPages();
        for (int pageNum = 1; pageNum <= pages; pageNum++) {
            float width = reader.getPageSize(pageNum).getWidth();
            float height = reader.getPageSize(pageNum).getHeight();

            PdfRenderListener pdfRenderListener = new PdfRenderListener(pageNum, width, height);

            //解析pdf，定位位置
            PdfContentStreamProcessor processor = new PdfContentStreamProcessor(pdfRenderListener);
            PdfDictionary pageDic = reader.getPageN(pageNum);
            PdfDictionary resourcesDic = pageDic.getAsDict(PdfName.RESOURCES);
            try {
                processor.processContent(ContentByteUtils.getContentBytesForPage(reader, pageNum), resourcesDic);
            } catch (IOException e) {
                reader.close();
                throw e;
            }

            String content = pdfRenderListener.getContent();
            List<CharPosition> charPositions = pdfRenderListener.getcharPositions();

            List<float[]> positionsList = new ArrayList<>();
            for (CharPosition charPosition : charPositions) {
                float[] positions = new float[]{charPosition.getPageNum(), charPosition.getX(), charPosition.getY()};
                positionsList.add(positions);
            }

            PdfPageContentPositions pdfPageContentPositions = new PdfPageContentPositions();
            pdfPageContentPositions.setContent(content);
            pdfPageContentPositions.setPostions(positionsList);

            result.add(pdfPageContentPositions);
        }
        reader.close();
        return result;
    }

    private static List<float[]> findPositions(String keyword, PdfPageContentPositions pdfPageContentPositions) {

        List<float[]> result = new ArrayList<>();

        String content = pdfPageContentPositions.getContent();
        List<float[]> charPositions = pdfPageContentPositions.getPositions();

        for (int pos = 0; pos < content.length(); ) {
            int positionIndex = content.indexOf(keyword, pos);
            if (positionIndex == -1) {
                break;
            }
            float[] postions = charPositions.get(positionIndex);
            result.add(postions);
            pos = positionIndex + 1;
        }
        return result;
    }

    private static class PdfPageContentPositions {
        private String content;
        private List<float[]> positions;

        public String getContent() {
            return content;
        }

        public void setContent(String content) {
            this.content = content;
        }

        public List<float[]> getPositions() {
            return positions;
        }

        public void setPostions(List<float[]> positions) {
            this.positions = positions;
        }
    }

    private static class PdfRenderListener implements RenderListener {
        private final int pageNum;
        private final float pageWidth;
        private final float pageHeight;
        private final StringBuilder contentBuilder = new StringBuilder();
        private final List<CharPosition> charPositions = new ArrayList<>();

        public PdfRenderListener(int pageNum, float pageWidth, float pageHeight) {
            this.pageNum = pageNum;
            this.pageWidth = pageWidth;
            this.pageHeight = pageHeight;
        }

        @Override
        public void beginTextBlock() {
        }

        public void renderText(TextRenderInfo renderInfo) {
            List<TextRenderInfo> characterRenderInfos = renderInfo.getCharacterRenderInfos();
            for (TextRenderInfo textRenderInfo : characterRenderInfos) {
                String word = textRenderInfo.getText();
                if (word.length() > 1) {
                    word = word.substring(word.length() - 1, word.length());
                }
                Float rectangle = textRenderInfo.getAscentLine().getBoundingRectange();

                float x = (float) rectangle.getX();
                float y = (float) rectangle.getY();

                //这两个是关键字在所在页面的XY轴的百分比
                float xPercent = Math.round(x / pageWidth * 10000) / 10000f;
                float yPercent = Math.round((1 - y / pageHeight) * 10000) / 10000f;


                CharPosition charPosition = new CharPosition(pageNum, (float) x, (float) y);
                charPositions.add(charPosition);
                contentBuilder.append(word);
            }
        }

        public void endTextBlock() {
        }

        public void renderImage(ImageRenderInfo renderInfo) {
        }

        public String getContent() {
            return contentBuilder.toString();
        }

        public List<CharPosition> getcharPositions() {
            return charPositions;
        }
    }

    private static class CharPosition {
        private int pageNum = 0;
        private float x = 0;
        private float y = 0;

        public CharPosition(int pageNum, float x, float y) {
            this.pageNum = pageNum;
            this.x = x;
            this.y = y;
        }

        public int getPageNum() {
            return pageNum;
        }

        public float getX() {
            return x;
        }

        public float getY() {
            return y;
        }

        @Override
        public String toString() {
            return "[pageNum=" + this.pageNum + ",x=" + this.x + ",y=" + this.y + "]";
        }
    }
}
