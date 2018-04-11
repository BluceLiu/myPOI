package OThinker.H3.Controller.BizSys.EnergyBuildManager;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

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
import org.springframework.web.multipart.MultipartFile;

import OThinker.Common.DotNetToJavaStringHelper;
import OThinker.Common.Data.DbType;
import OThinker.Common.Data.Database.Parameter;
import OThinker.H3.Controller.ControllerBase;
import OThinker.H3.Controller.BizSys.Common.ToolHelper;
import OThinker.H3.DataAccessLib.DB.SqlDbHelper;
import net.sf.json.JSONArray;
@SuppressWarnings("serial")
@Controller
@RequestMapping(value = "/Portal/EnergyBuild")
public class EnergyBuildExport extends ControllerBase {
	
	public static final String FILE_SEPARATOR = System.getProperties()
			.getProperty("file.separator");
	SqlDbHelper sqlDbHelper = new SqlDbHelper();
	
	/**
	 * 能源搭接导入
	 * @param request
	 * @param response
	 * @return
	 */
	@ResponseBody
	@RequestMapping(value = "/ImportEnergyBuild",method = {RequestMethod.POST,RequestMethod.GET},  produces = "application/json; charset=UTF-8")
	public String ImportEnergyBuild(@RequestParam(value = "file" , required = true) MultipartFile myfile,
			HttpServletRequest request, HttpServletResponse response) {
		Map<String, String> map = new HashMap<>();
		String insertPlanResult;
		String insertEnergyResult;
		InputStream inputStream = null ;
		XSSFWorkbook wb =null;
		XSSFSheet sheet = null;
		String checkResult = null;
		try {
			try {
				inputStream = myfile.getInputStream();
			} catch (IOException e1) {
				e1.printStackTrace();
			}//获取文件流
			try {
				wb = new XSSFWorkbook(inputStream);
			} catch (IOException e1) {
				e1.printStackTrace();
			}//得到Excel工作簿的对象
			if (wb == null) {
				map.put("result", "上传文件无内容！");
				return JSONArray.fromObject(map).toString();
			}
			sheet = wb.getSheetAt(0);//得到Excel工作表对象
			checkResult = checkSheet(sheet);//--------调用函数
			if (!"checked".equals(checkResult)) {
				map.put("result", checkResult);
				return JSONArray.fromObject(map).toString();
			}
			List<String> planidlist = new ArrayList<>();//获取账单表id和能源表planid
			sqlDbHelper.beginTranslate();//使用全局连接
			insertPlanResult = energyPlanGetParams(sheet,planidlist);//---------调用函数
			insertEnergyResult = energyGetParams(sheet,planidlist);//----------调用函数
			if (insertPlanResult == "done" && insertEnergyResult == "done") {
				sqlDbHelper.commit();
				map.put("result", "导入成功！");
				return JSONArray.fromObject(map).toString();
			}else {
				sqlDbHelper.rollBack();
				map.put("result", insertEnergyResult);
				return JSONArray.fromObject(map).toString();
			}
		} catch (Exception e) {
			e.printStackTrace();
			map.put("result", "导入不成功,格式错误！");
			return JSONArray.fromObject(map).toString();
		}
	}
	/**
	 * 检查表格是否有问题
	 * @param sheet
	 * @return
	 */
	private String checkSheet(XSSFSheet sheet) {
		//1.检查sheet是否存在
		if (sheet == null) {
			return "上传文件第一个sheet不存在，无法读取！";
		}
		//2.检查标题行是否完整
		String checkFirstRowResult = checkFirstRow(sheet);
		if (!"checked".equals(checkFirstRowResult)) {
			return checkFirstRowResult;
		}
		//3.检查是否有空行
		String checkBlankRowResult = checkBlankRow(sheet);
		if (!"checked".equals(checkBlankRowResult)) {
			return checkBlankRowResult;
		}
		//4.检查各列的格式
		String checkColumnResult = checkColumn(sheet);
		if (!"checked".equals(checkColumnResult)) {
			return checkColumnResult;
		}
		return "checked";
	}
	/*
	 * 检查空行
	 */
	private String checkBlankRow(XSSFSheet sheet) {
		String result = "";
		//查出哪一行是空行
		XSSFRow row = null;
		XSSFCell cell = null;
		boolean flag = false;
		for (int i = 0; i < sheet.getLastRowNum()+1; i++) {
			flag = false;
			row = sheet.getRow(i);
			if (row == null || i == 0) {
				continue;
			}
			for (int j = 0; j < row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (cell == null) {
					continue;
				}
				if (cell.getCellType() != XSSFCell.CELL_TYPE_BLANK && cell.getCellType() !=3) {
					flag = true;//该行有值
					break;
				}
			}
			if(flag){  
                continue;
            }else {
            	result = "第"+(i+1)+"行是空行,请检查后重新上传！";
				return result;
			}
		}
		return "checked";
	}
	/*
	 * 检查列格式
	 */
	private String checkColumn(XSSFSheet sheet) {
		String result = "";
		XSSFRow row = null;
		XSSFCell cell = null;
		for (int i = 0; i < sheet.getLastRowNum()+1; i++) {
			row = sheet.getRow(i);
			if (row == null || i == 0) {
				continue;
			}
			for (int j = 0; j < row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (cell == null || j==0) {
					continue;
				}
				switch (j) {
				case 1:case 2:case 3:case 4:case 5:case 7:case 16:case 17:case 18:case 19:case 20:
					if (cell.getCellType()!=XSSFCell.CELL_TYPE_STRING) {
						result = "第" + (i+1) + "行第"+(j+1)+"列格式应为字符串！";
						return result;
					}
					break;
				case 6:case 12:
					if (cell.getCellType()!=XSSFCell.CELL_TYPE_NUMERIC) {
						result = "第" + (i+1) + "行第"+(j+1)+"列格式应为数字！";
						return result;
					}
					break;
				case 13:case 14:case 15:
					if (cell.getCellType()!=XSSFCell.CELL_TYPE_NUMERIC) {
						result = "第" + (i+1) + "行第"+(j+1)+"列格式应为日期！";
						return result;
					}
					break;
				default:
					break;
				}
			}
		}
		return "checked";
	}
	/*
	 * 检查标题行
	 */
	private String checkFirstRow(XSSFSheet sheet) {
		XSSFRow row = null;
		XSSFCell cell = null;
		List<String> titleList = new ArrayList<>();
		List<String> titleListOrigin = new ArrayList<>(
				Arrays.asList("序号","公司","项目","合同编号","客户","费项","单价(元)",
						"表号","倍率","起数","止数","用量","金额(元)","应收日期","开始日期","结束日期",
						"公司ID","项目ID","合同ID","客户ID","费项ID"));
		row = sheet.getRow(0);
		if (row == null) {
			return "第一行应为标题行！";
		}
		for (int j = 0; j < row.getLastCellNum(); j++) {
			cell = row.getCell(j);
			if (cell == null) {
				continue;
			}
			titleList.add(cell.getCellType()==3?null:cell.getStringCellValue());
		}
		if (titleList.containsAll(titleListOrigin) && titleListOrigin.containsAll(titleList)) {
			return "checked";
		}
		return "标题行请使用模板标题！";
	}
	/**
	 * 功能：遍历sheet，获取插入【账单表】需要的参数
	 * 功能：循环调用插入账单表方法
	 * @param sheet
	 * @return
	 */
	private String energyPlanGetParams(XSSFSheet sheet,List<String> planidList) {
		String result = "";
		List<Parameter> paraInfo = new ArrayList<Parameter>();//插入能源表所有的参数
		XSSFRow row = null;
		XSSFCell cell = null;
		int count = 0;
		int countOne = 0;
		String planid = "";
		int i = 0;
		int j = 0;
		String v_sonid = UUID.randomUUID().toString();
		try {
			for (i = 0; i < sheet.getLastRowNum()+1; i++) {
				paraInfo.clear();//参数清空
				planid = UUID.randomUUID().toString();//生成主键备用
				planidList.add(planid);//该主键和能源表的planid一直
				//进行内容处理
				row = sheet.getRow(i);
				if (row == null || i == 0) {
					continue;
				}
				for (j = 0; j < row.getLastCellNum(); j++) {
					cell = row.getCell(j);
					if (cell == null) {
						continue;
					}
					switch (j) {
					case 16:paraInfo.add(new Parameter("v_companyid", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 1:paraInfo.add(new Parameter("v_companyname", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 18:paraInfo.add(new Parameter("v_cid", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 3:paraInfo.add(new Parameter("v_contractcode", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 17:paraInfo.add(new Parameter("v_projectid", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 2:paraInfo.add(new Parameter("v_projectname", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 19:paraInfo.add(new Parameter("v_customerid", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 4:paraInfo.add(new Parameter("v_customername", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 20:paraInfo.add(new Parameter("v_expenditureid", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 5:paraInfo.add(new Parameter("v_expenditurename", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 13:paraInfo.add(new Parameter("v_receivabledate", DbType.Date, cell.getCellType()==3?null:cell.getDateCellValue()));
							paraInfo.add(new Parameter("v_costdate", DbType.Date, cell.getCellType()==3?null:getFirstDayOfMonth(cell.getDateCellValue())));break;
					case 14:paraInfo.add(new Parameter("v_startdate", DbType.Date, cell.getCellType()==3?null:cell.getDateCellValue()));break;
					case 15:paraInfo.add(new Parameter("v_enddate", DbType.Date, cell.getCellType()==3?null:cell.getDateCellValue()));break;
					case 7:paraInfo.add(new Parameter("v_remark", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));break;
					case 6:paraInfo.add(new Parameter("v_unitprice", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));break;
					case 11:paraInfo.add(new Parameter("v_numbers", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));break;
					case 12:paraInfo.add(new Parameter("v_receivableaccount", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));break;
					default:break;
					}
				}
				paraInfo.add(new Parameter("v_objectid", DbType.String, planid));
				paraInfo.add(new Parameter("v_code", DbType.String, ToolHelper.getBussinessSerialNum("YS-ZC","",1)));
				paraInfo.add(new Parameter("v_sonid", DbType.String, v_sonid));
				paraInfo.add(new Parameter("v_yjexpenditureid", DbType.String, null));
				paraInfo.add(new Parameter("v_yjexpenditurename", DbType.String, null));
				paraInfo.add(new Parameter("v_costtype", DbType.String, "4"));
				paraInfo.add(new Parameter("v_receivablecycle", DbType.String, "5"));
				paraInfo.add(new Parameter("v_tctype", DbType.String, null));
				paraInfo.add(new Parameter("v_tcnumber", DbType.Double, null));
				paraInfo.add(new Parameter("v_ccnumber", DbType.Double, null));
				paraInfo.add(new Parameter("v_bdzq", DbType.String, null));
				paraInfo.add(new Parameter("v_bdaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_htwyjyqdays", DbType.String, "0"));
				paraInfo.add(new Parameter("v_htwyjbl", DbType.Double, 0));
				paraInfo.add(new Parameter("v_receivedaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_htwyjaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_cxdmaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_htwyjjmaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_yscdaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_refundaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_hxaccount", DbType.Double, 0));//找不到
				paraInfo.add(new Parameter("v_yjhkdate", DbType.Date, null));
				paraInfo.add(new Parameter("v_tjcdtkaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_htwytkaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_zfaccount", DbType.Double, 0));//找不到
				paraInfo.add(new Parameter("v_skstatus", DbType.String, "0"));
				paraInfo.add(new Parameter("v_sfscpz", DbType.Int16, 0));//是否生成凭证
				paraInfo.add(new Parameter("v_spstate", DbType.Int16, 1));//审批状态
				paraInfo.add(new Parameter("v_contractname", DbType.String, null));//合同名称找不到
				paraInfo.add(new Parameter("v_costitemname", DbType.String, null));//费项名称不落地
				paraInfo.add(new Parameter("v_costitemid", DbType.String, null));//费项ID不落地
				paraInfo.add(new Parameter("v_operatordepartid", DbType.String, this.getUserValidator()==null?null:this.getUserValidator().getDepartment()));//部门*******
				paraInfo.add(new Parameter("v_instanceid", DbType.String, "能源搭接无流程"));
				paraInfo.add(new Parameter("v_projectcode", DbType.String,null));//项目编码不落地
				paraInfo.add(new Parameter("v_ysaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_cxdmingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_yjaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_refundingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_tjcdtkingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_htwytkingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_yjtkkingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_prerefundaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_pretjcdtkingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_prehtwytkingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_preyjtkkingaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_yjtkaccount", DbType.Double, 0));
				paraInfo.add(new Parameter("v_createdby", DbType.String, this.getUserValidator()==null?null:this.getUserValidator().getUserID()));
				paraInfo.add(new Parameter("v_createdtime", DbType.Date, new Date()));
				paraInfo.add(new Parameter("v_operatordepartname", DbType.String, this.getUserValidator()==null?null:this.getUserValidator().getDepartmentName()));
				paraInfo.add(new Parameter("v_mobiles", DbType.String, null));
				countOne = energyPlanInsert(paraInfo.toArray(new Parameter[]{}));
				count += countOne;
			}
		} catch (Exception e) {
			e.printStackTrace();
			result = "第" + (i+1) +"行第"+(j+1)+"列格式不对！";
			return result;
		}
		if (count == sheet.getLastRowNum()) {
			result = "done";
			return result;
		}else {
			result = "第" + (i+1) +"行第"+(j+1)+"列格式不对！";
			return result;
		}
	}
	/**
	 * 获取当月第一天
	 */
	private Date getFirstDayOfMonth(Date date) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cal.set(Calendar.DAY_OF_MONTH, 1);
		return cal.getTime();
	}
	/**
	 * 功能：插入一条数据到【账单表】
	 * @param parameters
	 * @return
	 */
	private int energyPlanInsert(Parameter[] parameters) {
		String strSql = "INSERT INTO sys_con_receivableplan " +
		"  (objectid, code, sonid, companyid, companyname, contractcode, projectid, projectname, customerid, customername, "
		+ "expenditureid, expenditurename, yjexpenditureid, yjexpenditurename, costtype, receivablecycle, receivabledate, "
		+ "costdate, startdate, enddate, unitprice, numbers, receivableaccount, tctype, tcnumber, ccnumber, bdaccount, "
		+ "htwyjyqdays, htwyjbl, receivedaccount, htwyjaccount, cxdmaccount, htwyjjmaccount, yscdaccount, refundaccount, hxaccount, "
		+ "yjhkdate, createdby, createdtime, cid, bdzq, tjcdtkaccount, htwytkaccount, zfaccount, skstatus, sfscpz, spstate, "
		+ "contractname, costitemname, costitemid, operatordepartid, instanceid, projectcode, ysaccount, cxdmingaccount, "
		+ "yjaccount, refundingaccount, tjcdtkingaccount, htwytkingaccount, yjtkkingaccount, prerefundaccount, pretjcdtkingaccount, "
		+ "prehtwytkingaccount, preyjtkkingaccount, yjtkaccount, remark ,operatordepartname,mobiles)" + 
		"VALUES " + 
		"  (:v_objectid, :v_code, :v_sonid, :v_companyid, :v_companyname, :v_contractcode, :v_projectid, :v_projectname, :v_customerid, :v_customername, "
		+ ":v_expenditureid, :v_expenditurename, :v_yjexpenditureid, :v_yjexpenditurename, :v_costtype, :v_receivablecycle, :v_receivabledate, "
		+ ":v_costdate, :v_startdate, :v_enddate, :v_unitprice, :v_numbers, :v_receivableaccount, :v_tctype, :v_tcnumber, :v_ccnumber, :v_bdaccount,"
		+ ":v_htwyjyqdays, :v_htwyjbl, :v_receivedaccount, :v_htwyjaccount, :v_cxdmaccount, :v_htwyjjmaccount, :v_yscdaccount, :v_refundaccount, :v_hxaccount, "
		+ ":v_yjhkdate, :v_createdby, :v_createdtime, :v_cid, :v_bdzq, :v_tjcdtkaccount, :v_htwytkaccount, :v_zfaccount, :v_skstatus, :v_sfscpz, :v_spstate, "
		+ ":v_contractname, :v_costitemname, :v_costitemid, :v_operatordepartid, :v_instanceid, :v_projectcode, :v_ysaccount, :v_cxdmingaccount, "
		+ ":v_yjaccount, :v_refundingaccount, :v_tjcdtkingaccount, :v_htwytkingaccount, :v_yjtkkingaccount, :v_prerefundaccount, :v_pretjcdtkingaccount, "
		+ ":v_prehtwytkingaccount, :v_preyjtkkingaccount, :v_yjtkaccount, :v_remark,:v_operatordepartname,:v_mobiles) ";
		int count=0;
		try {
			count = sqlDbHelper.executeNonQuery(strSql, parameters);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return count;
	}
	/**
	 * 功能：遍历sheet，获取插入【能源表】需要的参数
	 * 功能：循环调用插入能源表方法
	 * @param sheet
	 * @return
	 */
	private String energyGetParams(XSSFSheet sheet,List<String> planidList) {
		String result = "";
		List<Parameter> paraInfo = new ArrayList<Parameter>();//插入能源表所有的参数
		XSSFRow row = null;
		XSSFCell cell = null;
		int count = 0;
		int countOne = 0;
		String planid = "";
		int i = 0;
		int j = 0;
		try {
			for (i = 0; i < sheet.getLastRowNum()+1; i++) {
				paraInfo.clear();
				planid = planidList.get(i);
				row = sheet.getRow(i);
				if (row == null || i == 0) {
					continue;
				}
				for (j = 0; j < row.getLastCellNum(); j++) {
					cell = row.getCell(j);
					if (cell == null) {
						continue;
					}
					if (j==7) 
						paraInfo.add(new Parameter("v_tablecode", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));
					if (j==8) 
						paraInfo.add(new Parameter("v_startnum", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));
					if (j==9) 
						paraInfo.add(new Parameter("v_endnum", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));
					if (j==10) 
						paraInfo.add(new Parameter("v_quantity", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));
					if (j==11) 
						paraInfo.add(new Parameter("v_rate", DbType.Double, cell.getCellType()==3?null:cell.getNumericCellValue()));
					if (j==14) 
						paraInfo.add(new Parameter("v_startdate", DbType.Date, cell.getCellType()==3?null:cell.getDateCellValue()));
					if (j==15) 
						paraInfo.add(new Parameter("v_enddate", DbType.Date, cell.getCellType()==3?null:cell.getDateCellValue()));
					if (j==16) 
						paraInfo.add(new Parameter("v_companyid", DbType.String, cell.getCellType()==3?null:cell.getStringCellValue()));
				}
				paraInfo.add(new Parameter("v_objectid", DbType.String, UUID.randomUUID().toString()));
				paraInfo.add(new Parameter("v_planid", DbType.String, planid));
				paraInfo.add(new Parameter("v_operatordepartid", DbType.String, this.getUserValidator()==null?null:this.getUserValidator().getDepartment()));
				paraInfo.add(new Parameter("v_userid", DbType.String, this.getUserValidator()==null?null:this.getUserValidator().getUserID()));
				paraInfo.add(new Parameter("v_createdtime", DbType.Date, new Date()));
				countOne = energyInsert(paraInfo.toArray(new Parameter[]{}));
				count += countOne;
			}
		} catch (Exception e) {
			e.printStackTrace();
			result = "第" + (i+1) +"行第"+(j+1)+"列格式不对！";
			return result;
		}
		if (count == sheet.getLastRowNum()) {
			return "done";
		}else {
			result = "第" + (i+1) +"行第"+(j+1)+"列格式不对！";
			return result;
		}
	}
	/**
	 * 功能：插入一条数据到能源表
	 * @param parameters
	 * @return
	 */
	private int energyInsert(Parameter[] parameters) {
		String strSql = "INSERT INTO sys_con_energybuildinfo " +
				"  (objectid, planid, tablecode, startnum, endnum, quantity, startdate, enddate, rate, "
				+ "companyid, operatordepartid, userid, createdtime) " + 
				"VALUES " + 
				"  (:v_objectid, :v_planid, :v_tablecode, :v_startnum, :v_endnum, :v_quantity, :v_startdate, :v_enddate, :v_rate, "
				+ ":v_companyid, :v_operatordepartid, :v_userid, :v_createdtime) ";

		int count=0;
		try {
			count = sqlDbHelper.executeNonQuery(strSql, parameters);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return count;
	}
	
	
	
	
	/**
	 * 能源搭接模板导出
	 * @param request
	 * @param response
	 * @return
	 */
	@ResponseBody
	@RequestMapping(value = "/ExportEnergyBuild",method = {RequestMethod.POST,RequestMethod.GET},produces = "application/json; charset=UTF-8")
	public String ExportEnergyBuild(HttpServletRequest request, HttpServletResponse response) {
		// 定义返回结果
		List<Map<String, String>> list = new ArrayList<>();
		// 获取前台传过来的參數
		//TODO:这里要对参数公司处理
//		String Company = request.getParameter("Company");
		String projectID = request.getParameter("projectID");
		String StartDate = request.getParameter("StartDate");
		String EndDate = request.getParameter("EndDate");
		SqlDbHelper helper = new SqlDbHelper();
		// 定义sql语句
		String sqlCommand = "SELECT C.COMPANYNAME,F.PRONAME,F.CONTRACTCODE,F.CUSNAME,F.COSTITEMNAME,F.UNITPRICE,F.ACCOUNT, " +
						"C.COMPANYID,F.PROID,F.CID,F.CUSID,F.COSTITEMID " + 
						"FROM SYS_CON_CONTRACTFIXEDCOSTITEM F " + 
						"LEFT JOIN SYS_CON_CONTRACTINFO C ON F.CID=C.OBJECTID " + 
						"WHERE F.COSTTYPE='费项类型2' AND C.STATUS='2' ";//-----states应为1，costtype应为“4”
		// TODO: 修改sql条件
		if (!DotNetToJavaStringHelper.isNullOrEmpty(projectID)) {
			sqlCommand += "AND F.PROID = '" + projectID +"' ";
		}
		try {
			list = helper.getInfoMapList(sqlCommand, null);
		} catch (Exception e) {
			e.printStackTrace();
		}
		//--------------------------改-------
		String docsPath = request.getSession().getServletContext()
				.getRealPath("/Portal/WFRes/_ExclTemplate");
		String fileName = "export_energyPlan.xlsx";
		String filePath = docsPath + FILE_SEPARATOR + fileName;
		//导出文件
		String fileNametemp = "导出能源搭接类收款计划创建_" + System.currentTimeMillis() + ".xlsx";
		String fileNameOut = "";
		try {
			fileNameOut = URLEncoder.encode(fileNametemp, "UTF-8");
		} catch (UnsupportedEncodingException e1) {
			e1.printStackTrace();
		} 
		FileInputStream fis;
		XSSFWorkbook workBook = null;
		try {
			fis = new FileInputStream(filePath);
			workBook=new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}  //  输入流
		if (workBook==null) {
			return "";
		}
		XSSFSheet sheet = workBook.getSheetAt(0);
		if(list.size() != 0){
			int k = 0;
			XSSFRow row1 = sheet.getRow(1);
			for(int j=0;j<list.size();j++){
				k = 0;
				POIUtils.copyRow(workBook,row1,sheet.createRow(j+2),false);  
				XSSFRow row = sheet.getRow(j+2);
				row.getCell(k++).setCellValue(j+1);
				row.getCell(k++).setCellValue(list.get(j).get("COMPANYNAME"));
				row.getCell(k++).setCellValue(list.get(j).get("PRONAME"));
				row.getCell(k++).setCellValue(list.get(j).get("CONTRACTCODE"));
				row.getCell(k++).setCellValue(list.get(j).get("CUSNAME"));
				row.getCell(k++).setCellValue(list.get(j).get("COSTITEMNAME"));
				double unitprice = 0;
				if(list.get(j).get("UNITPRICE")!="") {
					unitprice = Double.parseDouble(list.get(j).get("UNITPRICE"));
				}
				row.getCell(k++).setCellValue(unitprice);
				row.getCell(k++).setCellType(XSSFCell.CELL_TYPE_STRING);
				k = k + 4;           
				double account = 0;
				if(list.get(j).get("ACCOUNT")!="") {
					account = Double.parseDouble(list.get(j).get("ACCOUNT"));
				}
				row.getCell(k++).setCellValue(account);
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				try {
					row.getCell(k++).setCellValue(sdf.parse(StartDate));
					row.getCell(k++).setCellValue(sdf.parse(StartDate));//开始日期
					row.getCell(k++).setCellValue(sdf.parse(EndDate));
				} catch (ParseException e) {
					e.printStackTrace();
				}//应收日期
				row.getCell(k++).setCellValue(list.get(j).get("COMPANYID"));
				row.getCell(k++).setCellValue(list.get(j).get("PROID"));
				row.getCell(k++).setCellValue(list.get(j).get("CID"));
				row.getCell(k++).setCellValue(list.get(j).get("CUSID"));
				row.getCell(k++).setCellValue(list.get(j).get("COSTITEMID"));
			}
			sheet.shiftRows(2, sheet.getLastRowNum() , -1);
		}
        response.reset();  
		response.setContentType("application/msexcel");//设置生成的文件类型  
		response.setCharacterEncoding("UTF-8");//设置文件头编码方式和文件名  
		response.setHeader("Content-Disposition", "attachment; filename=" + fileNameOut);  
		OutputStream os = null;
		try {
			os = response.getOutputStream();
			workBook.write(os);
			os.flush();
			os.close(); 
		} catch (IOException e) {
			e.printStackTrace();
		}  
		//返回结果
		Map<String, Object> map = new HashMap<String, Object>();
		// 把集合转换成json对象集合字符串输出到前台
		map.put("list", list);
		String result = JSONArray.fromObject(map).toString();
		return result;
	}
	
	 /**
     * 表格样式设置
     * @param workbook
     * @param fontsize
     * @return
     */
    @SuppressWarnings("unused")
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
		return null;
	}
	
}
