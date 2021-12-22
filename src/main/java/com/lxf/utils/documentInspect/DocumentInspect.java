package com.lxf.utils.documentInspect;

import com.lxf.utils.docutils.Read_Docx_Doc_Wps;
import com.lxf.utils.docutils.Read_ppt_pptx_dps;
import com.lxf.utils.docutils.Read_xls_xlsx_et;
import com.lxf.domain.ResultDomain;
import com.lxf.utils.docutils.ReadPdf;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DocumentInspect {

    private static final int RIGHTFUL_CODE = 200; //合格
    private static final int ILLEGAL_CODE = 403;  //不合格
    private static final int NO_SUPPORT_CODE = 0; //文件格式不支持

    private static final String RIGHTFUL_MSG = "文件检测合格";
    private static final String ILLEGAL_MSG = "文件中包含关键字!";
    private static final String UN_TRUE_MSG = "文件未通过真实性检测";
    private static final String NONE_MSG = "文件未被检测,请检查文件格式是否支持";

    public static void main(String[] args) throws IOException {
        List<String> keyWordList = new ArrayList<String>();
        keyWordList.add("电饭锅");
        keyWordList.add("地方");
        keyWordList.add("546");
        keyWordList.add("sdf");
        System.out.println(documentInspect(keyWordList).toString());// new Date()为获取当前系统时间
    }

    /**
     * 文件格式及关键字检测,根据相应的格式调取 接口检测文件是否合格
     *
     * @param keyWordList 关键字集合
     * @return ResultDomain: code-状态码 | msg-消息
     * @throws IOException IO异常
     */
    public static ResultDomain documentInspect(List<String> keyWordList) throws IOException {
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        ResultDomain resultDomain = new ResultDomain(NO_SUPPORT_CODE, NONE_MSG);
//        String filePath = "D:\\LXF\\Test\\测试112.et";
//        String filePath = "D:\\LXF\\Test\\111.ppt";
//        String filePath = "D:\\LXF\\Test\\111.pptx";
//        String filePath = "D:\\LXF\\Test\\111.pptx";
        String filePath = "D:\\LXF\\Test\\wpstest.wps";
//        String filePath = "D:\\LXF\\Test\\pdf_Demo.pdf";

        System.out.println("****************************文件检测开始****************************");
        // 先检测文件的真实性
        if (FileTypeInspect.isPass(filePath)) {
            // 用来统计关键字数量
            int keyCount = 0;
            // 根据后缀名判断文件格式
            String fileName = FileTypeInspect.getFileName(filePath);
            String fileType = FileTypeInspect.getFileSuffix(fileName);
            System.out.println("文件格式为==>" + fileType);
            // 如果文件格式为以下类型 支持进行检测
            switch (fileType) {
                case "pdf":
                    isPdf(filePath, keyWordList, keyCount, resultDomain);
                    break;
                case "doc":
                    isDoc(filePath, keyWordList, keyCount, resultDomain);
                    break;
                case "wps":
                    isWps(filePath, keyWordList, keyCount, resultDomain);
                    break;
                case "docx":
                    isDocx(filePath, keyWordList, keyCount, resultDomain);
                    break;
                case "xls":
                case "et":
                case "xlsx":
                    is_xls_xlsx_et(filePath, keyWordList, keyCount, resultDomain);
                    break;
                case "ppt":
                    isPpt(filePath, keyWordList, keyCount, resultDomain);
                    break;
                case "dps":
                    isDps(filePath, keyWordList, keyCount, resultDomain);
                case "pptx":
                    isPptx(filePath, keyWordList, keyCount, resultDomain);
                    break;
                default:
                    System.out.println(NONE_MSG);
                    break;
            }
        } else {
            resultDomain.setCode(ILLEGAL_CODE);
            resultDomain.setMsg(UN_TRUE_MSG);
        }
        System.out.println("检测时间:" + df.format(new Date()));
        System.out.println("****************************文件检测结束****************************");
        return resultDomain;
    }

    /**
     * Doc 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void isDoc(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String docStr = Read_Docx_Doc_Wps.getDocStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, docStr, msg);
    }

    /**
     * docx 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void isDocx(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String docxStr = Read_Docx_Doc_Wps.getDocxStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, docxStr, msg);

    }

    /**
     * xls,xlsx,et 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void is_xls_xlsx_et(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String xls_xlsx_et_Str = Read_xls_xlsx_et.getAllExcelStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, xls_xlsx_et_Str, msg);
    }

    /**
     * PPT 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void isPpt(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String ppt_Str = Read_ppt_pptx_dps.getPptStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, ppt_Str, msg);
    }

    /**
     * PPTX 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void isPptx(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String pptx_Str = Read_ppt_pptx_dps.getPptxStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, pptx_Str, msg);
    }

    /**
     * WPS 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void isWps(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String wps_Str = Read_Docx_Doc_Wps.getWpsStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, wps_Str, msg);
    }

    /**
     * DPS 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void isDps(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String dps_Str = Read_ppt_pptx_dps.getDpsStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, dps_Str, msg);
    }

    /**
     * PDF 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @throws IOException IO异常
     */
    public static void isPdf(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) throws IOException {
        String pdf_Str = ReadPdf.getPdfStr(filePath);
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(filePath, keyWordList, keyCount, resultDomain, pdf_Str, msg);
    }

    /**
     * 检测通用代码  复用
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @param xxxStr       具体的文件字符串
     * @param msg          输出的消息
     */
    public static void isXxx(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain, String xxxStr, String msg) {
        if (!xxxStr.equals("")) {
            is_Document_inspect(keyWordList, keyCount, resultDomain, xxxStr);
        } else {
            System.out.println(msg);
        }
    }

    /**
     * 检测通用代码  复用
     *
     * @param keyWordList  关键字集合
     * @param keyCount     关键字数量 >1 即文件不合格
     * @param resultDomain 对象
     * @param xx_Str       某文档字符串
     */
    public static void is_Document_inspect(List<String> keyWordList, int keyCount, ResultDomain resultDomain, String xx_Str) {
        for (String keyWord : keyWordList) {
            if (FileTypeInspect.KeywordInspect(xx_Str, keyWord)) {
                System.out.println("检测到关键字==>" + keyWord);
                keyCount++;
            }
        }
        if (keyCount == 0) {
            resultDomain.setCode(RIGHTFUL_CODE);
            resultDomain.setMsg(RIGHTFUL_MSG);
        } else {
            resultDomain.setCode(ILLEGAL_CODE);
            resultDomain.setMsg(ILLEGAL_MSG);
        }
    }

}
