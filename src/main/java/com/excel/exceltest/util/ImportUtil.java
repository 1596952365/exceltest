package com.excel.exceltest.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;


public class ImportUtil {
	static Logger logger = LoggerFactory.getLogger(ImportUtil.class);
	public void downFile(String downloadFileName,String filePath,HttpServletResponse response) throws IOException{
		HttpHeaders responseHeaders = new HttpHeaders();
		responseHeaders.set("Content-Type", "text/plain;charset=utf-8");
		if("".equals(downloadFileName) || downloadFileName==null){
			logger.error("下载文件异常...", "未找到文件名字");
		}
		
		String dName = new String(downloadFileName.getBytes("UTF-8"), "iso-8859-1");
		File file = new File(filePath+downloadFileName);
		response.reset();
		response.addHeader("Content-Disposition", "attachment;filename=" + dName);
		response.addHeader("Content-Length", "" + file.length());
		response.setContentType("application/msexcel");

		InputStream inStream = null;
		byte[] b = new byte[100];
		int len;
		try {
			inStream = new FileInputStream(file);// 文件的存放路径
			while ((len = inStream.read(b)) > 0) {
				response.getOutputStream().write(b, 0, len);
			}
			inStream.close();
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("下载文件异常...", e.getMessage());
		} finally {
			if (null != inStream) {
				inStream.close();
			}
		}
	}
	
	public Workbook importFile(HttpServletRequest request){
		HttpHeaders responseHeaders = new HttpHeaders();
		logger.info("请求importFile 批量导入上传......");
		responseHeaders.set("Content-Type", "text/plain;charset=utf-8");
		MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
		MultipartFile mfile = multipartRequest.getFile("file");

		return getWorkbook(mfile);
	}
	
	public Collection<MultipartFile> importFile2(HttpServletRequest request){
		HttpHeaders responseHeaders = new HttpHeaders();
		logger.info("请求importFile2 批量导入上传......");
		responseHeaders.set("Content-Type", "text/plain;charset=utf-8");
		MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
		
		Map<String, MultipartFile> fileMap = multipartRequest.getFileMap();
		Collection<MultipartFile> files = fileMap.values();
		/*for(MultipartFile mfile:files){
			Workbook wb = getWorkbook(mfile);
		}*/

		return files;
	}
	
	public Workbook getWorkbook(MultipartFile mfile){
		if (mfile == null || mfile.getSize() == 0) {
			logger.info("请导入内容不为空的文件!");
			return null;
		}
		File file = toFile(mfile);
		
		if (file == null) {
			logger.info("请导入内容不为空的文件!");
			return null;
		}
		String fileName = file.getName();
		int ty1 = fileName.lastIndexOf(".xls") + 4;
		int ty2 = fileName.lastIndexOf(".xlsx") + 5;
		if (ty1 != fileName.length() && ty2 != fileName.length()) {
			logger.info("文件格式不正确！");
			return null;
		}
		String strErr = "";
		
		Workbook wb = versionExcel(file, strErr);
		if (!strErr.equals("")) {
			logger.info("列内容有问题");
			return null;
		}
		if (wb == null) {
			logger.info("文件为空或内容有问题！");
			return null;
		}
		Sheet sheet = wb.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		logger.info("导入行数：" + rows);
		file.delete();

		logger.info("请求importFile 批量导入完成");
		return wb;
	}
	
	// 将上传的MultipartFile类型文件转换为一个普通的file类型
	public static File toFile(MultipartFile mfile) {
		String path = PropertiesUtil.getPro("excel.import.path");
		String file_url = path + "/";

		File dir = new File(file_url);
		if (!dir.exists()) {
			dir.mkdirs();
		}
		String filename = mfile.getOriginalFilename();
		filename = filename.substring(filename.lastIndexOf("."), filename.length());
		File file = new File(dir + "/" + System.currentTimeMillis() + filename);
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			fos.write(mfile.getBytes());
		} catch (IOException e) {
			System.out.println(e.getMessage());
			return null;
		} finally {
			if (fos != null)
				try {
					fos.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
		return file;
	}
	
	public static Workbook versionExcel(File file, String strErr) {
		String fileName = file.getName();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
		FileInputStream fis;
		Workbook wb = null;
		try {
			fis = new FileInputStream(file);
			if (fileType != null && fileType.equals("xls")) {
				wb = new HSSFWorkbook(fis);
			} else if (fileType != null && fileType.equals("xlsx")) {
				wb = new XSSFWorkbook(fis);
			}
		} catch (Exception e) {
			// 内容中有超级链接,并链接为空为报错The hyperlink for cell BJ3 references relation
			// rId1, but that didn't exist!
			String message = e.getMessage();
			if (message.indexOf("The hyperlink for cell") != -1 && message.indexOf("references relation rId1") != -1) {
				strErr = message.substring(message.indexOf("cell") + 4, message.indexOf("references"));
				strErr = strErr.replaceAll("\\d+", "");
			}
			logger.error(e.getMessage());
			return null;
		}
		return wb;
	}
	
	/**
	 * 判断是否有空行
	 * 
	 * @param row
	 * @return
	 */
	public static boolean isBlankRow(Row row) {
		if (row == null)
			return true;
		boolean result = true;
		for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			String value = "";
			if (cell != null) {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					value = cell.getStringCellValue();
					break;
				case Cell.CELL_TYPE_NUMERIC:
					value = String.valueOf((int) cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					value = String.valueOf(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_FORMULA:
					value = String.valueOf(cell.getCellFormula());
					break;
				// case Cell.CELL_TYPE_BLANK:
				// break;
				default:
					break;
				}

				if (!value.trim().equals("")) {
					result = false;
					break;
				}
			}
		}

		return result;
	}
}
