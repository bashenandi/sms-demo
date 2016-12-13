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

import com.microsoft.schemas.office.visio.x2012.main.CellType;

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
			StringBuilder repo = new StringBuilder("<table style='border:solid 1px #333333;'><tbody>");
			Sheet sheet1 = (Sheet) wb.getSheetAt(0);
			String mobile = "m", content = "c", result = "no";
			for (Row row : sheet1) {
				Cell cellMobile = row.getCell(0);
				Cell cellContent = row.getCell(1);
				if(cellMobile != null && cellContent != null){
					mobile = cellMobile.getStringCellValue();
					content = cellContent.getStringCellValue();
					Thread.sleep(1000);
					result = this.sendSms(mobile, content);
					repo.append(String.format("<tr><td style='border:solid 1px #333333;'>%s</td><td style='border:solid 1px #333333;'>%s</td><td style='border:solid 1px #333333;'>%s</td></tr>"
							, mobile, content, result));
				}
			}
			repo.append("</tbody></table>");
			response.getWriter().append(repo.toString());
			pkg.close();
		} catch (EncryptedDocumentException | InvalidFormatException e) {
			// TODO Auto-generated catch block
			response.getWriter().append("Base exception is:" + e.getMessage());
		} catch (Exception e){
			response.getWriter().append("Exception is:" + e.getStackTrace());
		}
	}
	
	private String sendSms (String mobile, String content) throws MalformedURLException{
		String result = "";
		BufferedReader in = null;// 读取响应输入流
		
        try {  
        	@SuppressWarnings("deprecation")
			URL connURL = new URL(String.format("http://sms.zbwin.mobi/ws/sendsms.ashx?uid=%s&pass=%s&mobile=%s&content=%s"
        			, "omp", "21f11b316d6c6defaae08e83b1c2faac", mobile, java.net.URLEncoder.encode(content) ) );  
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
        return result;
	}

}
