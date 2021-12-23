package com.lxf.utils.documentInspect;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;
import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;

/**
 * @author xiaoli
 * <p>
 * 文件格式真实性 通用检测类
 * 用获取文件头的方式,直接读取文件的前几个字节,来判断上传文件是否符合格式。
 */
public class FileTypeInspect {

    public static void main(String[] args) throws Exception {
        String path = "D:\\LXF\\documentTest\\zzz.ppt";
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        System.out.println(isPass(path));

        System.out.println("================" + df.format(new Date()) + "================");
    }

    public final static Map<String, String> FILE_TYPE_MAP = new HashMap<String, String>();

    private FileTypeInspect() {
    }

    static {
        // 初始化文件类型信息
        getAllFileType();
    }

    /**
     * getAllFileType  常见文件头信息
     */
    public static void getAllFileType() {
        // 常用的
        FILE_TYPE_MAP.put("pdf", "255044462D312E");  //Adobe Acrobat (pdf)
        FILE_TYPE_MAP.put("jpg", "FFD8FF"); //JPEG (jpg)
        FILE_TYPE_MAP.put("png", "89504E47");  //PNG (png)
        FILE_TYPE_MAP.put("tif", "49492A00");  //TIFF (tif)  图片格式

        FILE_TYPE_MAP.put("xls", "D0CF11E0");  //MS Word  D0 CF 11 E0 A1 B1 1A E1
        FILE_TYPE_MAP.put("doc", "D0CF11E0");  //word   D0 CF 11 E0 A1 B1 1A E1
        FILE_TYPE_MAP.put("et", "D0CF11E0");  //et  D0 CF 11 E0 A1 B1 1A E1
        FILE_TYPE_MAP.put("wps", "D0CF11E0");  //wps D0 CF 11 E0 A1 B1 1A E1
        FILE_TYPE_MAP.put("dps", "D0CF11E0");  //dps   WPS文件格式  D0 CF 11 E0 A1 B1 1A E1
        FILE_TYPE_MAP.put("ppt", " D0CF11E0");  //pptx  PPT的另外一种格式   D0 CF 11 E0 A1 B1 1A E1

        FILE_TYPE_MAP.put("xlsx", "504B0304");  //zip xlsx  50 4B 03 04 14 00 06
        FILE_TYPE_MAP.put("docx", "504B0304");  //docx  docx word文档格式  50 4B 03 04 14 00 06
        FILE_TYPE_MAP.put("pptx", "504B0304");  //pptx  PPT的另外一种格式  50 4B 03 04 0A 00 00

        // 基本用不上
        FILE_TYPE_MAP.put("qdf", "AC9EBD8F");  //Quicken (qdf)
        FILE_TYPE_MAP.put("gif", "47494638");  //GIF (gif)
        FILE_TYPE_MAP.put("bmp", "424D"); //Windows Bitmap (bmp)
        FILE_TYPE_MAP.put("dwg", "41433130"); //CAD (dwg)
        FILE_TYPE_MAP.put("html", "68746D6C3E");  //HTML (html)
        FILE_TYPE_MAP.put("rtf", "7B5C727466");  //Rich Text Format (rtf)
        FILE_TYPE_MAP.put("xml", "3C3F786D6C"); //xml
        FILE_TYPE_MAP.put("rar", "52617221");  //rar
        FILE_TYPE_MAP.put("psd", "38425053");  //Photoshop (psd)
        FILE_TYPE_MAP.put("eml", "44656C69766572792D646174653A");  //Email [thorough only] (eml)
        FILE_TYPE_MAP.put("dbx", "CFAD12FEC5FD746F");  //Outlook Express (dbx)
        FILE_TYPE_MAP.put("pst", "2142444E");  //Outlook (pst)
        FILE_TYPE_MAP.put("mdb", "5374616E64617264204A");  //MS Access (mdb)
        FILE_TYPE_MAP.put("wpd", "FF575043"); //WordPerfect (wpd)
        FILE_TYPE_MAP.put("eps", "252150532D41646F6265");
        FILE_TYPE_MAP.put("ps", "252150532D41646F6265");
        FILE_TYPE_MAP.put("pwl", "E3828596");  //Windows Password (pwl)
        FILE_TYPE_MAP.put("wav", "57415645");  //Wave (wav)
        FILE_TYPE_MAP.put("avi", "41564920");  //avi
        FILE_TYPE_MAP.put("ram", "2E7261FD");  //Real Audio (ram)
        FILE_TYPE_MAP.put("rm", "2E524D46");  //Real Media (rm)
        FILE_TYPE_MAP.put("mpg", "000001BA");  // 视频文件格式
        FILE_TYPE_MAP.put("mov", "6D6F6F76");  //Quicktime (mov)
        FILE_TYPE_MAP.put("asf", "3026B2758E66CF11"); //Windows Media (asf)
        FILE_TYPE_MAP.put("mid", "4D546864");  //MIDI (mid)
        FILE_TYPE_MAP.put("mp4", "00000020667479706d70");//视频格式
        FILE_TYPE_MAP.put("mp3", "49443303000000002176");
    }

    /**
     * 判断文件类型是否真实
     *
     * @param filePath 文件路径
     * @return true:真实  false:不真实
     */
    public static boolean isPass(String filePath) {
        boolean result = false;
        String fileType = getFileSuffix(getFileName(filePath));
        File f = new File(filePath);
        if (f.exists()) {
            switch (fileType) {
                case "docx":
                case "xlsx":
                case "pptx":
                    result = docx_xlsx_pptx_isPass(filePath);
                    System.out.println("文件真实性检测==>" + (result ? "合格" : "不合格"));
                    break;
                case "doc":
                case "xls":
                case "et":
                case "wps":
                case "dps":
                case "ppt":
                    result = doc_xls_et_wps_dps_ppt_isPass(filePath);
                    System.out.println("文件真实性检测==>" + (result ? "合格" : "不合格"));
                    break;
                case "pdf":
                case "jpg":
                case "png":
                case "tif":
                    result = other_isPass(filePath);
                    System.out.println("文件真实性检测==>" + (result ? "合格" : "不合格"));
                    break;
                default:
                    result = true;
                    System.out.println("当前文件格式不支持检测");
                    break;
            }
        } else {
            System.out.println("文件不存在,请检查路径是否正确!");
        }
        return result;
    }

    /**
     * 判断 pdf jpg  png  tif 文件类型是否真实
     *
     * @param filePath 文件路径
     * @return true:真实  false:不真实
     */
    public static boolean other_isPass(String filePath) {
        File f = new File(filePath);
        return getFileSuffix(getFileName(filePath)).equals(getTypeByFile(f));
    }

    /**
     * 判断 docx,xlsx,pptx 文件类型是否真实
     *
     * @param filePath 文件路径
     * @return true:真实  false:不真实
     */
    public static boolean docx_xlsx_pptx_isPass(String filePath) {
        File file = new File(filePath);
        byte[] b = new byte[50];
        try {
            InputStream is = new FileInputStream(file);
            is.read(b);
            //如果文件头和获取的 docx,xlsx,pptx这些文件头一样,就为真实文件
            if ("504B0304".equals(getFileTypeValue(b))) {
                return true;
            }
            is.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 判断  doc,xls,et,wps,dps,ppt文件是否真实
     *
     * @param filePath 文件路径
     * @return true:真实  false:不真实
     */
    public static boolean doc_xls_et_wps_dps_ppt_isPass(String filePath) {
        File file = new File(filePath);
        byte[] b = new byte[50];
        try {
            InputStream is = new FileInputStream(file);
            is.read(b);
            //如果文件头和获取的 docx,xlsx,pptx这些文件头一样,就为真实文件
            if ("D0CF11E0".equals(getFileTypeValue(b))) {
                return true;
            }
            is.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 根据文件全路径获得不带后缀的文件名
     *
     * @param path 文件路径
     * @return name 文件名
     */
    public static String getName(String path) {
        String fileName = path.substring(path.lastIndexOf("\\") + 1);
        return fileName.substring(0, fileName.lastIndexOf("."));
    }

    /**
     * 根据文件全路径取得文件名以及后缀
     *
     * @param filePath 文件全路径
     * @return filenameNow 文件名及后缀
     */
    public static String getFileName(String filePath) {
        String[] temp = filePath.split("\\\\");
        String fileNameNow = temp[temp.length - 1];
//        System.out.println("文件名及格式==>" + fileNameNow);
        return fileNameNow;
    }

    /**
     * 根据getFileName方法获得文件名及后缀,根据文件名及后缀获得文件后缀名。
     *
     * @param fileName 文件名
     * @return 文件后缀
     */
    public static String getFileSuffix(String fileName) {
        String[] strArray = fileName.split("\\.");
        int suffixIndex = strArray.length - 1;
//        System.out.println("文件类型==>" + strArray[suffixIndex]);
        return strArray[suffixIndex];
    }

    /**
     * 获取图片文件实际类型,若不是图片则返回null
     *
     * @param file 文件
     * @return fileType
     */
    public static String getImageFileType(File file) {
        if (isImage(file)) {
            try {
                ImageInputStream iis = ImageIO.createImageInputStream(file);
                Iterator<ImageReader> iter = ImageIO.getImageReaders(iis);
                if (!iter.hasNext()) {
                    return null;
                }
                ImageReader reader = iter.next();
                iis.close();
                return reader.getFormatName();
            } catch (IOException e) {
                return null;
            } catch (Exception e) {
                return null;
            }
        }
        return null;
    }

    /**
     * 判断文件类型是否为图片类型
     *
     * @param file 文件
     * @return true 是 | false 否
     */
    public static boolean isImage(File file) {
        boolean flag = false;
        try {
            BufferedImage buffReader = ImageIO.read(file);
            int width = buffReader.getWidth();
            int height = buffReader.getHeight();
            flag = width != 0 && height != 0;
        } catch (IOException e) {
            flag = false;
        } catch (Exception e) {
            flag = false;
        }
        return flag;
    }

    /**
     * 获取文件类型,包括图片,若格式不是已配置的,则返回null
     *
     * @param file 文件
     * @return fileType
     */
    public static String getTypeByFile(File file) {
        String filetype = null;
        byte[] b = new byte[50];
        try {
            InputStream is = new FileInputStream(file);
            is.read(b);
            filetype = getFileTypeByStream(b);
            is.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return filetype;
    }

    /**
     * 获取文件类型的字符流
     *
     * @param b 文件头信息byte数组
     * @return fileType
     */
    public static String getFileTypeByStream(byte[] b) {
        String filetypeHex = String.valueOf(getFileHexString(b));
        for (Entry<String, String> entry : FILE_TYPE_MAP.entrySet()) {
            String fileTypeHexValue = entry.getValue();
            if (filetypeHex.toUpperCase().startsWith(fileTypeHexValue)) {
                System.out.println("16进制码==>" + entry.getValue());
                System.out.println("文件真实类型==>" + entry.getKey());
                return entry.getKey();
            }
        }
        return null;
    }

    /**
     * 获取文件类型的16进制编码
     *
     * @return codeValue
     */
    public static String getFileTypeValue(byte[] b) {
        String filetypeHex = String.valueOf(getFileHexString(b));
//        System.out.println("=============" + filetypeHex + "==============");
        for (Entry<String, String> entry : FILE_TYPE_MAP.entrySet()) {
            String fileTypeHexValue = entry.getValue();
            if (filetypeHex.toUpperCase().startsWith(fileTypeHexValue)) {
                System.out.println("文件头16进制码==>" + entry.getValue());
                return entry.getValue();
            }
        }
        return null;
    }

    /**
     * 将要读取文件头信息的文件的byte数组转换成string类型表示
     *
     * @param b []
     * @return fileTypeHex
     */
    public static String getFileHexString(byte[] b) {
        StringBuilder fileTypeHexStr = new StringBuilder();
        if (b == null || b.length <= 0) {
            return null;
        }
        for (byte value : b) {
            int v = value & 0xFF;
            String hv = Integer.toHexString(v);
            if (hv.length() < 2) {
                fileTypeHexStr.append(0);
            }
            fileTypeHexStr.append(hv);
        }
        return fileTypeHexStr.toString();
    }

    /**
     * 检测String 字符串中是否包含有某些关键字
     *
     * @param docStr  被检测的字符串
     * @param keyword 关键字
     * @return true:包含  false:不包含
     */
    public static boolean KeywordInspect(String docStr, String keyword) {
        return docStr.contains(keyword);
    }
}
