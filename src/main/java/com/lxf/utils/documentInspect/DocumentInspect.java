package com.lxf.utils.documentInspect;

import com.lxf.domain.ResultDomain;
import com.lxf.utils.docutils.*;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DocumentInspect {

    private static final int RIGHTFUL_CODE = 200; //合格
    private static final int ILLEGAL_CODE = 403;  //不合格

    private static final String RIGHTFUL_MSG = "文件检测合格";

    private static final String UN_TRUE_MSG = "文件未通过真实性检测";
    private static final String NO_EXISTS_MSG = "文件不存在";
    private static final String NO_SUPPORT_MSG = "文件格式不支持";
    private static final String ILLEGAL_MSG = "文件中包含关键字,检测不合格";
    private static final String NAME_ILLEGAL_MSG = "文件名中包含关键字,检测不合格";

    public static void main(String[] args) throws IOException {
        List<String> keyWordList = new ArrayList<>();
        keyWordList.add("电饭锅");
        keyWordList.add("地方");
//        keyWordList.add("n");
//        keyWordList.add("new");
//        String filePath = "D:\\LXF\\Test\\kong.ppt";
        String filePath = "D:\\LXF\\Test\\To\\new.pptx";
//        String filePath = "D:\\LXF\\Test\\测试113.xls";
//        String filePath = "D:\\LXF\\Test\\测试114.xlsx";

//        String filePath = "D:\\LXF\\Test\\111p.ppt";
//        String filePath = "D:\\LXF\\Test\\111x.pptx";
//        String filePath = "D:\\LXF\\Test\\111d.dps";

//        String filePath = "D:\\LXF\\Test\\pdf_Demo.pdf";

//        String filePath = "D:\\LXF\\Test\\doctest.doc";
//        String filePath = "D:\\LXF\\Test\\111.docx";
//        String filePath = "D:\\LXF\\Test\\测试111.wps"
//        String filePath = "D:\\LXF\\Test\\wpstest.wps";
//        String filePath = "D:\\LXF\\Test\\111.txt";
        System.out.println(documentInspect(keyWordList, filePath));// new Date()为获取当前系统时间
    }

    /**
     * 文件格式及关键字检测,根据相应的格式调取 接口检测文件是否合格
     *
     * @param keyWordList 关键字集合
     * @return ResultDomain: code-状态码 | msg-消息
     * @throws IOException IO异常
     */
    public static ResultDomain documentInspect(List<String> keyWordList, String filePath) throws IOException {
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        ResultDomain resultDomain = new ResultDomain(ILLEGAL_CODE, NO_SUPPORT_MSG);
        System.out.println("****************************文件检测开始****************************");
        //1.先检测文件是否存在，且是否为一个文件而非文件夹
        File f = new File(filePath);
        if (f.exists() && f.isFile()) {
            // 2. 再检测文件的真实性
            if (FileTypeInspect.isPass(filePath)) {
                // 用来统计关键字数量
                int keyCount = 0;
                // 根据后缀名判断文件格式
                String fileName = FileTypeInspect.getFileName(filePath);
                String fileType = FileTypeInspect.getFileSuffix(fileName);
                String name = FileTypeInspect.getName(filePath);
                System.out.println("[文件格式]:" + fileType);
                // 检测文件名是否合格，文件名合格后再往下进行

                if (nameIsPass(keyWordList, name)) {
                    // 3.如果文件格式为以下类型 支持进行检测
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
                            break;
                        case "pptx":
                            isPptx(filePath, keyWordList, keyCount, resultDomain);
                            break;
                        case "txt":

                            break;
                        default:
                            resultDomain.setCode(ILLEGAL_CODE);
                            resultDomain.setMsg(NO_SUPPORT_MSG);
                            // System.out.println(NONE_MSG);
                            break;
                    }
                } else {
                    resultDomain.setCode(ILLEGAL_CODE);
                    resultDomain.setMsg(NAME_ILLEGAL_MSG);
                }
            } else {
                resultDomain.setCode(ILLEGAL_CODE);
                resultDomain.setMsg(UN_TRUE_MSG);
            }
        } else {
            resultDomain.setCode(ILLEGAL_CODE);
            resultDomain.setMsg(NO_EXISTS_MSG);
        }
        System.out.println("[检测时间]:" + df.format(new Date()));
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
     */
    public static void isDoc(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String docStr = Read_docx_doc_wps.getDocStr(filePath);
        String msg = null;
        if (docStr.equals("noDoc")) {
            msg = "该文件经过检测非 DOC/WPS 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, docStr, msg);
        } else {
            msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
            isXxx(keyWordList, keyCount, resultDomain, docStr, msg);
        }
    }

    /**
     * docx 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void isDocx(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String docxStr = Read_docx_doc_wps.getDocxStr(filePath);
        String msg = null;
        if (docxStr.equals("noDocx")) {
            msg = "该文件经过检测非 DOCX 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println("[提示]" + msg);
            isXxx(keyWordList, keyCount, resultDomain, docxStr, msg);
        } else if (docxStr.equals("emptyDocx")) {
            msg = "检测合格(该文件是一个空的DOCX文件)";
            resultDomain.setCode(200);
            resultDomain.setMsg(msg);
            isXxx(keyWordList, keyCount, resultDomain, docxStr, msg);
        } else {
            msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
            isXxx(keyWordList, keyCount, resultDomain, docxStr, msg);
        }

    }

    /**
     * xls,xlsx,et 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void is_xls_xlsx_et(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String xls_xlsx_et_Str = Read_xls_xlsx_et.getAllExcelStr(filePath);
        if (xls_xlsx_et_Str.equals("noXls")) {
            String msg = "该文件经过检测非 xls/xlsx/et 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, xls_xlsx_et_Str, msg);
        }
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(keyWordList, keyCount, resultDomain, xls_xlsx_et_Str, msg);
    }

    /**
     * PPT 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void isPpt(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String ppt_Str = Read_ppt_pptx_dps.getPptStr(filePath);
        if (ppt_Str.equals("noPpt")) {
            String msg = "该文件经过检测非 PPT/DPS 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, ppt_Str, msg);
        }
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(keyWordList, keyCount, resultDomain, ppt_Str, msg);
    }

    /**
     * PPTX 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void isPptx(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String pptx_Str = Read_ppt_pptx_dps.getPptxStr(filePath);
        String msg = null;
        if (pptx_Str.equals("noPptx")) {
            msg = "该文件经过检测非 PPTX 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, pptx_Str, msg);
        } else if (pptx_Str.equals("emptyPptx")) {
            msg = "检测合格(该文件是一个空的PPTX文件)";
            resultDomain.setCode(200);
            resultDomain.setMsg(msg);
            isXxx(keyWordList, keyCount, resultDomain, pptx_Str, msg);
        } else {
            msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
            isXxx(keyWordList, keyCount, resultDomain, pptx_Str, msg);
        }
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
        String wps_Str = Read_docx_doc_wps.getWpsStr(filePath);
        if (wps_Str.equals("noWps")) {
            String msg = "该文件经过检测非 WPS 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, wps_Str, msg);
        }
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(keyWordList, keyCount, resultDomain, wps_Str, msg);
    }

    /**
     * DPS 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void isDps(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String dps_Str = Read_ppt_pptx_dps.getDpsStr(filePath);
        if (dps_Str.equals("noDps")) {
            String msg = "该文件经过检测非 DPS 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, dps_Str, msg);
        }
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(keyWordList, keyCount, resultDomain, dps_Str, msg);
    }

    /**
     * PDF 文件关键字检测
     *
     * @param filePath     文件路径
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     */
    public static void isPdf(String filePath, List<String> keyWordList, int keyCount, ResultDomain resultDomain) {
        String pdf_Str = ReadPdf.getPdfStr(filePath);
        if (pdf_Str.equals("noPdf")) {
            String msg = "该文件经过检测非 PDF 文件";
            resultDomain.setCode(403);
            resultDomain.setMsg(msg);
            System.out.println(msg);
            isXxx(keyWordList, keyCount, resultDomain, pdf_Str, msg);
        }
        String msg = FileTypeInspect.getFileName(filePath) + " 内容为空";
        isXxx(keyWordList, keyCount, resultDomain, pdf_Str, msg);
    }

    /**
     * 文件检测通用代码  复用
     *
     * @param keyWordList  关键字集合
     * @param keyCount     关键字统计
     * @param resultDomain 返回结果
     * @param xxxStr       具体的文件字符串
     * @param msg          输出的消息
     */
    public static void isXxx(List<String> keyWordList, int keyCount, ResultDomain resultDomain, String xxxStr, String msg) {
        if (!xxxStr.equals("")) {
            is_Document_inspect(keyWordList, keyCount, resultDomain, xxxStr);
        } else {
            System.out.println(msg);
            is_Document_inspect(keyWordList, keyCount, resultDomain, xxxStr);
        }
    }

    /**
     * 关键字检测通用代码  复用
     *
     * @param keyWordList  关键字集合
     * @param keyCount     关键字数量 >1 即文件不合格
     * @param resultDomain 对象
     * @param xx_Str       某文档字符串
     */
    public static void is_Document_inspect(List<String> keyWordList, int keyCount, ResultDomain resultDomain, String xx_Str) {
        switch (xx_Str) {
            case "noDoc":
            case "noDocx":
            case "noXls":
            case "noXlsx":
            case "noPpt":
            case "noPptx":
            case "emptyDocx":
            case "emptyPptx":
                return;
            default:
                break;
        }
        for (String keyWord : keyWordList) {
            if (FileTypeInspect.KeywordInspect(xx_Str, keyWord)) {
                System.out.println("检测到关键字==>" + keyWord);
                keyCount++;
            }
        }
        if (keyCount == 0) {
            System.out.println(RIGHTFUL_MSG);
            resultDomain.setCode(RIGHTFUL_CODE);
            resultDomain.setMsg(RIGHTFUL_MSG);
        } else {
            resultDomain.setCode(ILLEGAL_CODE);
            resultDomain.setMsg(ILLEGAL_MSG);
        }
    }

    /**
     * 文件名关键字检测
     */
    public static boolean nameIsPass(List<String> keyWordList, String name) {
        int keywordCount = 0;
        System.out.println("[文件名]:" + name);
        for (String keyword : keyWordList) {
            if (name.contains(keyword)) {
                keywordCount++;
                System.out.println("[包含的关键字" + keywordCount + "]:" + keyword);
            }
        }
        return keywordCount == 0;
    }
}
