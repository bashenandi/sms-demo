package com.debugerman.sms.demo;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Servlet implementation class ReadExcel
 */
public class ReadExcel extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public ReadExcel() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.setContentType("text/html;charset=utf-8");
		response.setCharacterEncoding("UTF-8");
		String filename = request.getParameter("file");
		String filepath = request.getRealPath("upload/" + filename);
		try {
			OPCPackage pkg = OPCPackage.open(new File(filepath));
			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			DataFormatter formatter = new DataFormatter();
			Sheet sheet1 = (Sheet) wb.getSheetAt(0);
			for (Row row : sheet1) {
				Cell cell1 = row.getCell(0);
				Cell cell2 = row.getCell(1);
				String mobile = cell1.getStringCellValue();
				String content = cell2.getStringCellValue().toString();
				response.getWriter().append("<div>").append("<i>" + mobile + "</i>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;").append(content).append("<u>"+ this.sendSms(mobile, content) +"</u>").append("</div>");
			}
			pkg.close();
		} catch (EncryptedDocumentException | InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//response.getWriter().append("Served at: ").append(request.getContextPath());
	}
	
	private String sendSms (String mobile, String content) throws MalformedURLException{
		String result = "no content";
		BufferedReader in = null;// 读取响应输入流
		
		/**
        try {  
        	@SuppressWarnings("deprecation")
			URL connURL = new URL(String.format("http://sms.zbwin.mobi/ws/sendsms.ashx?uid=%s&pass=%s&mobile=%s&content=%s"
        			, "", "", mobile, java.net.URLEncoder.encode(content) ) );  
            // 打开URL连接
            java.net.HttpURLConnection httpConn = (java.net.HttpURLConnection) connURL  
                    .openConnection();  
            // 设置通用属性  
            
            httpConn.setRequestProperty("Connection", "Keep-Alive");  
            httpConn.setRequestProperty("User-Agent",  
                    "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1)");  
            // 建立实际的连接  
            httpConn.connect(); 
            // 定义BufferedReader输入流来读取URL的响应,并设置编码方式  
            in = new BufferedReader(new InputStreamReader(httpConn  
                    .getInputStream(), "UTF-8"));  
            String line;  
            // 读取返回的内容  
            while ((line = in.readLine()) != null) {  
                result += line;  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        } finally {  
            try {  
                if (in != null) {  
                    in.close();  
                }  
            } catch (IOException ex) {  
                ex.printStackTrace();  
            }  
        }  
        **/
        return result;
	}

}
