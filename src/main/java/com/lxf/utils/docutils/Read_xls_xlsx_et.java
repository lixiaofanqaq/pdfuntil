package com.lxf.utils.docutils;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;

import java.util.List;

public class Read_xls_xlsx_et {
    public static void main(String[] args) {
        String path = "D:\\LXF\\Test\\测试111.xlsx";
        String path2 = "D:\\LXF\\Test\\测试112.xls";
        String path3 = "D:\\LXF\\Test\\测试112.et";
//        String str = getExcelStr(path);
//        System.out.println(str);
        System.out.println(getAllExcelStr(path3));
    }

    /**
     * 获取xls,xlsx,et文件（表格样式）中的内容,并转换为字符串
     *
     * @param filePath 文件完整路径
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

}

