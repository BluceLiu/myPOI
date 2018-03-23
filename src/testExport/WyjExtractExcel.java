package testExport;

import java.awt.Robot;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import OThinker.Common.DotNetToJavaStringHelper;
import OThinker.Common.Data.Database.Parameter;
import OThinker.H3.Controller.ControllerBase;
import OThinker.H3.DataAccessLib.DB.SqlDbHelper;
import net.sf.json.JSONArray;
import oracle.xml.parser.v2.oraxml;
@Controller
@RequestMapping(value = "/Portal/Bill")
public class WyjExtractExcel extends ControllerBase {
	
	public static final String FILE_SEPARATOR = System.getProperties()
			.getProperty("file.separator");
	
	
	//表单提交请求，可以传递参数，在用
	@SuppressWarnings("serial")
	@ResponseBody
	@RequestMapping(value = "/Wyj2Extract",method = {RequestMethod.POST,RequestMethod.GET},  produces = "application/json; charset=UTF-8")
	public String Extract2Excel(HttpServletRequest request, HttpServletResponse response) {
		// 定义查询结果返回行数
		String contractCode = request.getParameter("contractCode");
		String customerID = request.getParameter("customerID");
		String mobile = request.getParameter("mobile");
		String companyID = request.getParameter("companyID");
		SqlDbHelper helper = new SqlDbHelper();
		// 定义sql语句
		String sqlCommand = "select ROWNUM,a.ObjectID,a.ContractCode,a.ContractName,a.ProjectName,a.CustomerName,a.ExpenditureName,"
				+"to_char(a.ReceivableDate,'yyyy-MM-dd') as ReceivableDate,to_char(a.CostDate,'yyyy-MM-dd') as CostDate,"
				+"a.ReceivableAccount,a.HtwyjjmAccount,a.HtwyjAccount,a.YsAccount,a.ReceivedAccount,a.htwyjYqDays,"
				+"b.PlanID,b.JmType,b.State,to_char(b.JmDate,'yyyy-MM-dd') as JmDate,b.JmBl,b.JmGdAccount,b.JmBz "
				+" from SYS_CON_RECEIVABLEPLAN a LEFT JOIN Sys_Con_Plan_WYJJM b ON TRIM(a.OBJECTID) = TRIM(b.PLANID) "
				+" LEFT JOIN sys_con_contractinfocustomer d ON a.cid = d.CID "
				+" WHERE 1 = 1 AND b.OBJECTID IS NOT NULL";
		// 定义数据库执行语句参数集合
		List<Parameter> parList = new ArrayList<Parameter>();
		// 如果合同编号不为空，则作为查询条件
		if (!DotNetToJavaStringHelper.isNullOrEmpty(contractCode)) {
			sqlCommand += " and a.contractCode like '%"  + contractCode + "%'";
		}
		// 如果客户姓名不为空，则作为查询条件
		if (!DotNetToJavaStringHelper.isNullOrEmpty(customerID)) {
			sqlCommand += " and a.customerID like '%" + customerID + "%'";
		}
//		// 如果客户电话不为空，则作为查询条件
//		if (!DotNetToJavaStringHelper.isNullOrEmpty(mobile)) {
//			sqlCommand += " and c.MOBILE like '%" + mobile + "%'";
//		}
		// 如果公司id不为空，则作为查询条件
		if (!DotNetToJavaStringHelper.isNullOrEmpty(companyID)) {
			sqlCommand += " and a.companyID like '%" + companyID + "%'";
		}
		//排序
		sqlCommand += " ORDER BY a.customerid,a.projectid,a.expenditureid,a.receivableaccount";
		//List<Map<String, String>> getPagerListForMap
		List<Map<String, String>> list = new ArrayList<Map<String,String>>() {
		};
		// 转换参数集合
		Parameter[] parArr = parList.toArray(new Parameter[] {});
		try {
			list = helper.getPagerListForMap(1000,1,sqlCommand, parArr);
		} catch (Exception e) {
			e.printStackTrace();
		}
//		String docsPath = request.getSession().getServletContext()
//				.getRealPath("docs");
		String fileName = "export2007中文_" + System.currentTimeMillis() + ".xlsx";
		
		//---------------------------------------
		try {
			// 输出流
			// 工作区
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("export");
			//头标题样式
	        XSSFCellStyle headStyle = createCellStyle(wb,(short)13);
	        //设置默认列宽
            sheet.setDefaultColumnWidth(15);
			//填充第一行--------------------------
			//生成第一行
			XSSFRow row = sheet.createRow(0);
			String[] titles = {"合同编号","合同名称","客户","项目","费项","费用日期","应收日期",
					"应收金额/元","状态","减免方式","计算截止日期","减免比例","固定减免违约金金额/元",
					"合同违约金/元","合同违约金减免/元","减免后合同违约金/元"};
			for(int i=0;i<titles.length;i++){
                XSSFCell cell = row.createCell(i);
                cell.setCellStyle(headStyle);
                cell.setCellValue(titles[i]);
            }
			//-----填充第一行END--------------------------
			//填充数据行--------------------------
			if(list != null){
				System.out.println(list.size());
                for(int j=0;j<list.size();j++){
                    //创建数据行,前面有一行,列标题行
                    XSSFRow row3 = sheet.createRow(j+1);
                    row3.createCell(0).setCellValue(list.get(j).get("CONTRACTCODE"));
                    row3.createCell(1).setCellValue(list.get(j).get("CONTRACTNAME"));
                    row3.createCell(2).setCellValue(list.get(j).get("CUSTOMERNAME"));
                    row3.createCell(3).setCellValue(list.get(j).get("PROJECTNAME"));
                    row3.createCell(4).setCellValue(list.get(j).get("EXPENDITURENAME"));
                    row3.createCell(5).setCellValue(list.get(j).get("COSTDATE"));
                    row3.createCell(6).setCellValue(list.get(j).get("RECEIVABLEDATE"));
                    row3.createCell(7).setCellValue(list.get(j).get("RECEIVABLEACCOUNT"));
                    String stateName = getStateName(list.get(j).get("STATE"));
                    row3.createCell(8).setCellValue(stateName);
                    String jmtypeName = getJmtypeName(list.get(j).get("JMTYPE"));
                    row3.createCell(9).setCellValue(jmtypeName);
                    row3.createCell(10).setCellValue(list.get(j).get("JMDATE"));
                    String jmbl = getJmbl(list.get(j).get("JMBL"));
                    row3.createCell(11).setCellValue(jmbl);
                    String jmgdaccount = list.get(j).get("JMGDACCOUNT");
                    jmgdaccount = updateZero(jmgdaccount);
                    row3.createCell(12).setCellValue(jmgdaccount);
                    String wyjTotal = getWyjTotal(list.get(j).get("HTWYJJMACCOUNT"),list.get(j).get("HTWYJACCOUNT"));
                    row3.createCell(13).setCellValue(wyjTotal);
                    row3.createCell(14).setCellValue(updateZero(list.get(j).get("HTWYJJMACCOUNT")));
                    row3.createCell(15).setCellValue(updateZero(list.get(j).get("HTWYJACCOUNT")));
                    
                }
            }
			
			//填充数据行END--------------------------
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
		//返回结果
		Map<String, Object> map = new HashMap<String, Object>();
		// 把集合转换成json对象集合字符串输出到前台
		String result = JSONArray.fromObject(map).toString();
		return result;
	}
	
	/**
	 * 判断减免类型
	 * @param jmtype
	 * @return
	 */
	public String getJmtypeName(String jmtype){
		String result = "";
		switch (jmtype) {
		case "1":
			result = "计算截至日期";
			break;
		case "2":
			result = "减免比例";
			break;
		case "3":
			result = "固定违约金金额";
			break;
		default:
			break;
		}
		return result;
	}
	/**
	 * 判断状态
	 * @param state
	 * @return
	 */
	public String getStateName(String state){
		String result = "";
		switch (state) {
		case "1":
			result = "生效";
			break;
		case "2":
			result = "已撤销";
			break;
		default:
			break;
		}
		return result;
	}
	/**
	 * 计算总金额
	 * 合同违约金=合同违约金减免+减免后合同违约金
	 * @param a
	 * @param b
	 * @return
	 */
	public String getWyjTotal(String a,String b){
		String result = "";
		double result1 = 0;
		double result2 = 0;
		result1 = (a=="")?0:Double.parseDouble(a);
		result2 = (b=="")?0:Double.parseDouble(b);
		result = String.valueOf(result1+result2);
		if ("0.0".equals(result)) {
			result = "";
		}
		return result;
	}
	//减免比例百分比显示
	public String getJmbl(String c){
		String result = "";
		if (c == "") {
		}else{
			DecimalFormat df = new DecimalFormat("0.0%");
			result = String.valueOf(df.format(Double.parseDouble(c)));
		}
		if ("0.0%".equals(result)) {
			result="";
		}
		return result;
	}	
	//如果值是零，就不显示
	public String updateZero(String a){
		String result = "";
		Double b = 0.0;
		if (a == "") {
		}else{
			b = Double.parseDouble(a);
		}
		if (0 - b == 0) {
			result = "";
		}else {
			result =String.valueOf(b);
		}
		return result;
	}	
	 /**
     * 表格样式设置
     * @param workbook
     * @param fontsize
     * @return
     */
    private static XSSFCellStyle createCellStyle(XSSFWorkbook workbook, short fontsize) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);//垂直居中
        //创建字体
        XSSFFont font = workbook.createFont();
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(fontsize);
        //加载字体
        style.setFont(font);
        return style;
    }
	@Override
	public String getFunctionCode() {
		// TODO Auto-generated method stub
		return null;
	}
	//ajax请求，无法选择路径，废弃不用
	/*@ResponseBody
		@RequestMapping(value = "/WyjExtract",method = RequestMethod.POST,  produces = "application/json; charset=UTF-8")
		public String ExtractExcel(HttpServletRequest request, HttpServletResponse response) {
			// 定义返回结果
			List<Object> rowList = null;
			// 定义查询结果返回行数
			int count = 0;
			String contractCode = request.getParameter("contractCode");
			String customerName = request.getParameter("customerName");
			String mobile = request.getParameter("mobile");
			String companyName = request.getParameter("companyName");
			SqlDbHelper helper = new SqlDbHelper();			
			// 定义sql语句
			String sqlCommand = "select ROWNUM,a.ObjectID,a.ContractCode,a.ContractName,a.ProjectName,a.CustomerName,a.ExpenditureName,"
					+"to_char(a.ReceivableDate,'yyyy-MM-dd') as ReceivableDate,to_char(a.CostDate,'yyyy-MM-dd') as CostDate,"
					+"a.ReceivableAccount,a.HtwyjjmAccount,a.HtwyjAccount,a.YsAccount,a.ReceivedAccount,a.htwyjYqDays,"
					+"b.PlanID,b.JmType,b.State,to_char(b.JmDate,'yyyy-MM-dd') as JmDate,b.JmBl,b.JmGdAccount,b.JmBz "
					+" from SYS_CON_RECEIVABLEPLAN a LEFT JOIN Sys_Con_Plan_WYJJM b ON TRIM(a.OBJECTID) = TRIM(b.PLANID) "
					+" LEFT JOIN I_Con_Customer c ON c.OBJECTID = a.CUSTOMERID "
					+" WHERE 1 = 1 AND (b.state is null or b.state='1')";
			// 定义数据库执行语句参数集合
			List<Parameter> parList = new ArrayList<Parameter>();
			// 如果合同编号不为空，则作为查询条件
			if (!DotNetToJavaStringHelper.isNullOrEmpty(contractCode)) {
				sqlCommand += " and a.contractCode like '%" + contractCode + "%'";
			}
			// 如果客户姓名不为空，则作为查询条件
			if (!DotNetToJavaStringHelper.isNullOrEmpty(customerName)) {
				sqlCommand += " and a.customerName like '%" + customerName + "%'";
			}
			// 如果客户电话不为空，则作为查询条件
			if (!DotNetToJavaStringHelper.isNullOrEmpty(mobile)) {
				sqlCommand += " and c.MOBILE like '%" + mobile + "%'";
			}
			// 如果公司id不为空，则作为查询条件
			if (!DotNetToJavaStringHelper.isNullOrEmpty(companyName)) {
				sqlCommand += " and a.companyName like '%" + companyName + "%'";
			}
			//List<Map<String, String>> getPagerListForMap
			List<Map<String, String>> list = new ArrayList<Map<String,String>>() {
			};
			// 转换参数集合
			Parameter[] parArr = parList.toArray(new Parameter[] {});
			try {
				list = helper.getPagerListForMap(1000,1,sqlCommand, parArr);
			} catch (Exception e) {
				e.printStackTrace();
			}
//			String docsPath = request.getSession().getServletContext()
//					.getRealPath("docs");
			String docsPath = "E:\\";
			String fileName = "export2007中文_" + System.currentTimeMillis() + ".xlsx";
			String filePath = docsPath + FILE_SEPARATOR + fileName;
			
			
			//---------------------------------------
			try {
				// 输出流
				OutputStream os = new FileOutputStream(filePath);
				// 工作区
				XSSFWorkbook wb = new XSSFWorkbook();
				XSSFSheet sheet = wb.createSheet("export");
				//头标题样式
		        XSSFCellStyle headStyle = createCellStyle(wb,(short)13);
		        //设置默认列宽
	            sheet.setDefaultColumnWidth(15);
				//填充第一行--------------------------
				//生成第一行
				XSSFRow row = sheet.createRow(0);
				String[] titles = {"合同编号","合同名称","客户","项目","费项","费用日期","应收日期",
						"应收金额/元","状态","减免方式","计算截止日期","减免比例","固定减免违约金金额/元",
						"合同违约金/元","合同违约金减免/元","减免后合同违约金/元"};
				for(int i=0;i<titles.length;i++){
	                XSSFCell cell = row.createCell(i);
	                cell.setCellStyle(headStyle);
	                cell.setCellValue(titles[i]);
	            }
				//-----填充第一行END--------------------------
				//填充数据行--------------------------
				if(list != null){
					System.out.println(list.size());
	                for(int j=0;j<list.size();j++){
	                    //创建数据行,前面有一行,列标题行
	                    XSSFRow row3 = sheet.createRow(j+1);
	                    row3.createCell(0).setCellValue(list.get(j).get("CONTRACTCODE"));
	                    row3.createCell(1).setCellValue(list.get(j).get("CONTRACTNAME"));
	                    row3.createCell(2).setCellValue(list.get(j).get("CUSTOMERNAME"));
	                    row3.createCell(3).setCellValue(list.get(j).get("PROJECTNAME"));
	                    row3.createCell(4).setCellValue(list.get(j).get("EXPENDITURENAME"));
	                    row3.createCell(5).setCellValue(list.get(j).get("COSTDATE"));
	                    row3.createCell(6).setCellValue(list.get(j).get("RECEIVABLEDATE"));
	                    row3.createCell(7).setCellValue(list.get(j).get("RECEIVABLEACCOUNT"));
	                    String stateName = getJmtypeName(list.get(j).get("STATE"));
	                    row3.createCell(8).setCellValue(stateName);
	                    String jmtypeName = getJmtypeName(list.get(j).get("JMTYPE"));
	                    row3.createCell(9).setCellValue(jmtypeName);;
	                    row3.createCell(10).setCellValue(list.get(j).get("JMDATE"));;
	                    row3.createCell(11).setCellValue(list.get(j).get("JMBL"));;
	                    row3.createCell(12).setCellValue(list.get(j).get("JMGDACCOUNT"));;
	                    row3.createCell(13).setCellValue(list.get(j).get("HTWYJJMACCOUNT")==""?0:Double.parseDouble(list.get(j).get("HTWYJJMACCOUNT"))+
	                    		list.get(j).get("HTWYJACCOUNT")==""?0:Double.parseDouble(list.get(j).get("HTWYJACCOUNT")));;
	                    row3.createCell(14).setCellValue(list.get(j).get("HTWYJJMACCOUNT"));;
	                    row3.createCell(15).setCellValue(list.get(j).get("HTWYJACCOUNT"));;
	                    
	                }
	            }
				
				//填充数据行END--------------------------
				
				wb.write(os);
				// 关闭输出流
				os.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
			download(filePath,response);
			//返回结果
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("path", docsPath);
			// 把集合转换成json对象集合字符串输出到前台
			String result = JSONArray.fromObject(map).toString();
			return result;
		}*/
	
}
