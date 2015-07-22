package com.shenj.teworks.service.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.ConnectException;
import java.util.Properties;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

/**
 * 将Office文档转换为PDF文档
 * 
 * @author 王文路
 * @date 2015-7-22
 */
public class Office2PDFConverter {
	
	public static void main(String[] args){
		office2PDF("G:\\234.docx" , "G:\\123.pdf");
	}

	/**
	 * 环境变量下面的url.properties的绝对路径
	 */
	private static final String RUL_PATH = Thread.currentThread()
			.getContextClassLoader().getResource("").getPath()
			.replace("%20", " ")
			+ "url.properties";

	/**
	 * 将Office文档转换为PDF. 运行该函数需要用到OpenOffice, OpenOffice版本3.x
	 * 
	 * <pre>
	 * 方法示例:
	 * String sourcePath = "F:\\office\\source.doc";
	 * String destFile = "F:\\pdf\\dest.pdf";
	 * Office2PDFConverter.office2PDF(sourcePath, destFile);
	 * </pre>
	 * 
	 * @author 王文路
	 * @date 2015-7-22
	 * @param sourceFile
	 * 			源文件, 绝对路径. 可以是Office2003-2007全部格式的文档, Office2010的没测试. 包括.doc,
	 *            .docx, .xls, .xlsx, .ppt, .pptx等. 示例: F:\\office\\source.doc
	 * @param destFile
	 * 			 目标文件. 绝对路径. 示例: F:\\pdf\\dest.pdf
	 * @return
	 */
	public static boolean office2PDF(String sourceFile, String destFile) {
		
		try {
			
			// 1  找不到源文件, 则返回false
			File inputFile = new File(sourceFile);
			if (!inputFile.exists()) {
				return false;
			}

			File outputFile = new File(destFile);
			
			// 2 已经存在pdf文件，返回成功
			if (outputFile.exists())
				return true;

			// 3  如果目标路径不存在, 则新建该路径
			if (!outputFile.getParentFile().exists()) {
				outputFile.getParentFile().mkdirs();
			}
			
			// 4 从url.properties文件中读取OpenOffice的安装根目录, OpenOffice_HOME对应的键值.
			Properties prop = new Properties();
			FileInputStream fis = null;
			fis = new FileInputStream(RUL_PATH);	// 属性文件输入流
			prop.load(fis);		// 将属性文件流装载到Properties对象中
			fis.close();		// 关闭流

			String OpenOffice_HOME = prop.getProperty("OpenOffice_HOME");
			
			// 5 如果没有配置openoffice Home  ，返回错误
			if (OpenOffice_HOME == null)
				return false;
			
			// 6 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
			if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
				OpenOffice_HOME += "\\";
			}
			
			// 7 启动OpenOffice的服务 , 注意端口不要冲突
			String command = OpenOffice_HOME
					+ "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8300;urp;\" -nofirststartwizard";
			Process pro = Runtime.getRuntime().exec(command);
			
			// 8 连接到OpenOffice ，注意端口要与上面一致
			OpenOfficeConnection connection = new SocketOpenOfficeConnection(
					"127.0.0.1", 8300);
			connection.connect();

			// 8 转换pdf
			DocumentConverter converter = new OpenOfficeDocumentConverter(
					connection);
			converter.convert(inputFile, outputFile);

			// 9 关闭连接
			connection.disconnect();
			
			// 10  关闭OpenOffice服务的进程 ， 避免占用CPU
			pro.destroy();

			return true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (ConnectException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return false;
	}
}
