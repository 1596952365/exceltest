package com.excel.exceltest.util;

/**
 * Created by Administrator on 2014/10/30.
 */
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;

public class FileUtil {
    private static final Logger Logger = LoggerFactory.getLogger(FileUtil.class);

    // 重命名文件；
    public static File[] renameToFiles(String[] fileNames, File[] files) {
        File[] retFiles = null;
        if (fileNames != null && fileNames.length > 0) {
            retFiles = new File[fileNames.length];

            for (int i = 0, n = fileNames.length; i < n; i++) {
                File dist = new File(fileNames[i]);
                files[i].renameTo(dist);
                retFiles[i] = dist;
            }
        }
        return retFiles;
    }

    // save文件
    public static long saveFile(File file, String fileName, String filePath)
            throws Exception {
        if (file == null) {
            return 0;
        }

        File filepath = new File(filePath);
        if (!filepath.isDirectory()) {
            filepath.mkdirs();
        }
        File filedesc = new File(filepath, fileName);

        return copyFile(file, filedesc);
    }

    // copy文件
    public static long copyFile(File fromFile, File toFile) {
        long len = 0;

        InputStream in = null;
        OutputStream out = null;
        try {
            in = new FileInputStream(fromFile);
            out = new FileOutputStream(toFile);
            byte[] t = new byte[1024];
            int ii = 0;
            while ((ii = in.read(t)) > 0) {
                out.write(t, 0, ii);
                len += ii;
            }

        } catch (IOException e) {
            Logger.error(e.getMessage(),e);

        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (Exception e) {
                    Logger.error(e.getMessage(),e);
                }
            }

            if (out != null) {
                try {
                    out.close();
                } catch (Exception e) {
                    Logger.error(e.getMessage(),e);
                }
            }

        }

        return len;
    }

    // 验证文件正确;
    public static boolean verifyFile(String fileFix, String[] exts) {
        boolean flag = false;
        for(String ext : exts){
            if(fileFix.equals(ext)){
                flag = true;
                break;
            }
        }
        return flag;
    }

    // 取得文件扩展;
    public static String getExtension(String fileName) {

        int newEnd = fileName.length();
        int i = fileName.lastIndexOf('.', newEnd);
        if (i != -1) {
            return fileName.substring(i + 1, newEnd).toLowerCase();
        } else {
            return null;
        }
    }


    /**
     * 传文件名称，返回一个输出流
     * @param fileName 需要上传的文件目录
     */
    public static PrintWriter uploadBossCountFile(String fileName){
        File file = null;
        file = new File(fileName);
        if(file.exists()==false){//如果文件不存在，则创建文件
            try{
                file.createNewFile();
            }catch(IOException e){
                Logger.error(e.getMessage(),e);
            }
        }
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(fileName);
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        PrintWriter writer = null;
        writer = new PrintWriter(new OutputStreamWriter(fos));
        return writer;
    }

    public static boolean deleteFile(String filePath){
        boolean flag = false;
        File file = new File(filePath);
        // 路径为文件且不为空则进行删除
        if (file.isFile() && file.exists()) {
            file.delete();
            flag = true;
        }
        return flag;
    }

    //资源预览类型：1. 视频 2. 音频 3. 图片 4. 文档 5. 其它 6. 动画
    public static int checkResourceType(String extension) {
        try {
            extension =  extension.toLowerCase();
        }catch (Exception e){
            Logger.error(e.getMessage(),e);
        }
        int type = 0;
        if (Constants.DOCTYPE.contains(extension)) {
            //文档资源
            type = 4;
        } else if (Constants.JPGTYPE.contains(extension)) {
            //图片资源
            type = 3;
        } else if (Constants.MP3TYPE.contains(extension)) {
            //音频资源
            type = 2;
        } else if (Constants.FLVTYPE.contains(extension)) {
            //视频资源
            type = 1;
        } else if (Constants.SWFTYPE.contains(extension)) {
            //动画
            type = 6;
        } else {
            type = 5;
        }
        return type;
    }

    public static void main(String[] args){
        
    }
}
