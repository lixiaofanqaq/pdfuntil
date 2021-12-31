package com.lxf.utils.documentInspect;

import net.sf.jmimemagic.Magic;
import net.sf.jmimemagic.MagicMatch;
import org.apache.tika.mime.MimeType;
import org.apache.tika.mime.MimeTypes;

import java.io.File;


/**
 * 用获取mime码的方式验证文件格式是否真实
 */
public class FileMimeUtil {

    public static void main(String[] args) {
        try {
//            getMimeType1();
            getMimeType1("D:\\LXF\\documentTest\\zzz.dps");
            getMimeType1("D:\\LXF\\Test\\docxToxlsx.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * mime码校验示例
     * @param url
     * @throws Exception
     */
    public static void getMimeType1(String url) throws Exception {
        MimeTypes allTypes = MimeTypes.getDefaultMimeTypes();
        File file = new File(url);
        MagicMatch match = Magic.getMagicMatch(file, true, false);
        String mimeType = match.getMimeType();
        System.out.println("MIME：" + mimeType);
        MimeType mimeTypeMatch = allTypes.forName(mimeType);
        System.out.println("suffix:" + mimeTypeMatch.getExtension());
    }
}