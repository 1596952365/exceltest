package com.excel.exceltest.util;

import javax.servlet.http.HttpSession;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by stc on 2016/9/8.
 */
public class Constants {
    public static Map<String, HttpSession> map = new HashMap<String, HttpSession>();

    public static final String DOCTYPE = "doc,ppt,pdf,docx,pptx,txt,xlsx,xls,pdf";

    public static final String JPGTYPE = "gif,jpg,png,jpeg,bmp";

    public static final String MP3TYPE = "mp3,wav,wma";

    public static final String FLVTYPE = "rmvb,avi,wmv,asf,mpeg,mpeg2,3gp,flv,mp4,f4v,f5v,f6v,rm";

    public static final String SWFTYPE = "swf";
}
