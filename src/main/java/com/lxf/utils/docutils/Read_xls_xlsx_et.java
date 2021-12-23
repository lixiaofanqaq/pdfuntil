package com.lxf.utils.docutils;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Read_xls_xlsx_et {
    public static void main(String[] args) throws IOException {
        String path = "D:\\LXF\\Test\\测试114.xlsx";
        String path2 = "D:\\LXF\\Test\\测试113.xls";
        String path3 = "D:\\LXF\\Test\\测试112.et";

//        System.out.println(getAllExcelStr(path));

        for (String Image : getAllExcelImages(path3)) {
            System.out.println(Image);
        }

    }

    /**
     * 获取xls,xlsx,et文件（表格样式）中的内容,并转换为字符串
     *
     * @param filePath 文件路径
     * @return String excelStr
     */
    public static String getAllExcelStr(String filePath) {
        StringBuilder excelStr = new StringBuilder();

        if (!filePath.equals("")) {
            ExcelReader reader = ExcelUtil.getReader(filePath);

            List<List<Object>> readAll = reader.read();
            for (List list : readAll) {
                for (Object o : list) {
                    excelStr.append(o);
//                    System.out.println(o);
                }
            }
            reader.close();
        } else {
            System.out.println("路径不正确，找不到Excel文件");
        }
//        System.out.println(excelStr.toString());
        return excelStr.toString();
    }

    /**
     * 获取xls,xlsx,et文件（表格样式）中的图片，并存到相应的文件夹中
     *
     * @param filePath 文件路径
     * @return List 图片路径集合
     * @throws FileNotFoundException 找不到文件异常
     * @throws IOException           IO异常
     */
    public static List<String> getAllExcelImages(String filePath) throws IOException {
        List<String> imagesList = new ArrayList<>();
        File file = new File(filePath);
        if (file.exists() && file.isFile()) {
            Workbook workbook = null;
            InputStream is = new FileInputStream(file);
            if (filePath.endsWith(".xls") || filePath.endsWith(".et")) {
                workbook = new HSSFWorkbook(is);
            } else if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(is);
            }
            if (workbook != null) {
                // 图片内容
                List<?> pictures = workbook.getAllPictures();
                //获取去除后缀的文件路径
                Read_ppt_pptx_dps.getImagePath(filePath, imagesList, pictures);
            } else {
                System.out.println("Workbook==null,文件格式不正确");
                return null;
            }
        } else {
            System.out.println("xls/xlsx/et 文件不存在");
        }
        return imagesList;
    }
}

