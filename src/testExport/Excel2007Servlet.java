package testExport;

import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2007Servlet extends HttpServlet {
	public static final String FILE_SEPARATOR = System.getProperties()
			.getProperty("file.separator");

	@Override
	protected void doPost(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {
		doGet(request, response);
	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		String fileName = "export2007中文_" + System.currentTimeMillis() + ".xlsx";
		try {
			// 工作区
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("test");
			for (int i = 0; i < 10; i++) {
				// 创建第一个sheet
				// 生成第一行
				XSSFRow row = sheet.createRow(i);
				// 给这一行的第一列赋值
				row.createCell(0).setCellValue("column1");
				// 给这一行的第一列赋值
				row.createCell(1).setCellValue("column2");
				System.out.println(i);
			}
			 // 清空response  
            response.reset();  
            response.setContentType("application/msexcel");//设置生成的文件类型  
            response.setCharacterEncoding("UTF-8");//设置文件头编码方式和文件名  
            response.setHeader("Content-Disposition", "attachment; filename=" + fileName);  
            OutputStream os=response.getOutputStream();  
            wb.write(os);  
            os.flush();  
            os.close();  
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	
}