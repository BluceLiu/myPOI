package OThinker.H3.Controller.BizSys.EnergyBuildManager;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.util.Region;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIUtils {  
//  /**   
//   * 把一个excel中的cellstyletable复制到另一个excel，这里会报错，不能用这种方法，不明白呀？？？？？  
//   * @param fromBook   
//   * @param toBook   
//   */   
//  public static void copyBookCellStyle(XSSFWorkbook fromBook,XSSFWorkbook toBook){  
//      for(short i=0;i<fromBook.getNumCellStyles();i++){  
//          XSSFCellStyle fromStyle=fromBook.getCellStyleAt(i);  
//          XSSFCellStyle toStyle=toBook.getCellStyleAt(i);  
//          if(toStyle==null){  
//              toStyle=toBook.createCellStyle();  
//          }   
//          copyCellStyle(fromStyle,toStyle);  
//      }   
//  }   
	/**
	 * 删除空白行
	 */
	public XSSFSheet deleteBlankRows(XSSFSheet sheet) {
		boolean flag = false;  
		XSSFRow r = null;
		XSSFCell c = null;
        System.out.println("总行数："+(sheet.getLastRowNum()+1));  
        System.out.println("非空行："+(sheet.getLastRowNum()+1));  
        for (int i = 0; i <= sheet.getLastRowNum();) {  
            r = sheet.getRow(i);  
            if(r == null){  
                // 如果是空行（即没有任何数据、格式），直接把它以下的数据往上移动  
                sheet.shiftRows(i+1, sheet.getLastRowNum(),-1);  
                continue;  
            }  
            flag = false;  
            for (int j = 0; j < r.getLastCellNum(); j++) {
            	c = r.getCell(j);
                if(c.getCellType() != XSSFCell.CELL_TYPE_BLANK){  
                    flag = true;  
                    break;  
                }  
            }  
            if(flag){  
                i++;  
                continue;  
            }  
            else{//如果是空白行（即可能没有数据，但是有一定格式）  
                if(i == sheet.getLastRowNum())//如果到了最后一行，直接将那一行remove掉  
                    sheet.removeRow(r);  
                else//如果还没到最后一行，则数据往上移一行  
                    sheet.shiftRows(i+1, sheet.getLastRowNum(),-1);  
            }  
        } 
        return sheet;
	}
    /** 
     * 复制一个单元格样式到目的单元格样式 
     * @param fromStyle 
     * @param toStyle 
     */  
    public static void copyCellStyle(XSSFCellStyle fromStyle,  
            XSSFCellStyle toStyle) {  
        toStyle.setAlignment(fromStyle.getAlignment());  
        //边框和边框颜色   
        toStyle.setBorderBottom(fromStyle.getBorderBottom());  
        toStyle.setBorderLeft(fromStyle.getBorderLeft());  
        toStyle.setBorderRight(fromStyle.getBorderRight());  
        toStyle.setBorderTop(fromStyle.getBorderTop());  
        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());  
        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());  
        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());  
        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());  
          
        //背景和前景   
        toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundXSSFColor());  
        toStyle.setFillForegroundColor(fromStyle.getFillForegroundXSSFColor());  
//        toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());  
//        toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());  
          
        toStyle.setDataFormat(fromStyle.getDataFormat());  
        toStyle.setFillPattern(fromStyle.getFillPattern());  
//      toStyle.setFont(fromStyle.getFont(null));  
        toStyle.setHidden(fromStyle.getHidden());  
        toStyle.setIndention(fromStyle.getIndention());//首行缩进  
        toStyle.setLocked(fromStyle.getLocked());  
        toStyle.setRotation(fromStyle.getRotation());//旋转  
        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());  
        toStyle.setWrapText(fromStyle.getWrapText());  
    }  
    /** 
     * Sheet复制 
     * @param fromSheet 
     * @param toSheet 
     * @param copyValueFlag 
     */  
//    public static void copySheet(XSSFWorkbook wb,XSSFSheet fromSheet, XSSFSheet toSheet,  
//            boolean copyValueFlag) {  
//        //合并区域处理   
//        mergerRegion(fromSheet, toSheet);  
//        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {  
//            XSSFRow tmpRow = (XSSFRow) rowIt.next();  
//            XSSFRow newRow = toSheet.createRow(tmpRow.getRowNum());  
//            //行复制   
//            copyRow(wb,tmpRow,newRow,copyValueFlag);  
//        }  
//    }  
    /** 
     * 行复制功能 
     * @param fromRow 
     * @param toRow 
     */  
    public static void copyRow(XSSFWorkbook wb,XSSFRow fromRow,XSSFRow toRow,boolean copyValueFlag){  
        for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext();) {  
            XSSFCell tmpCell = (XSSFCell) cellIt.next();  
            XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());  //.getCellNum()
            copyCell(wb,tmpCell, newCell, copyValueFlag);  
        }  
    }  
    /** 
    * 复制原有sheet的合并单元格到新创建的sheet 
    *  
    * @param sheetCreat 新创建sheet 
    * @param sheet      原有的sheet 
    */  
//    public static void mergerRegion(XSSFSheet fromSheet, XSSFSheet toSheet) {  
//       int sheetMergerCount = fromSheet.getNumMergedRegions();  
//       for (int i = 0; i < sheetMergerCount; i++) {  
//        Region mergedRegionAt = fromSheet.getMergedRegion(i);  
//        toSheet.addMergedRegion(mergedRegionAt);  
//       }  
//    }  
    /** 
     * 复制单元格 
     *  
     * @param srcCell 
     * @param distCell 
     * @param copyValueFlag 
     *            true则连同cell的内容一起复制 
     */  
    public static void copyCell(XSSFWorkbook wb,XSSFCell srcCell, XSSFCell distCell,  
            boolean copyValueFlag) {  
        XSSFCellStyle newstyle=wb.createCellStyle();  
        copyCellStyle(srcCell.getCellStyle(), newstyle);  
//        distCell.setEncoding((srcCell.getEncoding());  
        //样式   
        distCell.setCellStyle(newstyle);  
        //评论   
        if (srcCell.getCellComment() != null) {  
            distCell.setCellComment(srcCell.getCellComment());  
        }  
        // 不同数据类型处理   
        int srcCellType = srcCell.getCellType();  
        distCell.setCellType(srcCellType);  
        if (copyValueFlag) {  
            if (srcCellType == XSSFCell.CELL_TYPE_NUMERIC) {  
                if (HSSFDateUtil.isCellDateFormatted(srcCell)) {  
                    distCell.setCellValue(srcCell.getDateCellValue());  
                } else {  
                    distCell.setCellValue(srcCell.getNumericCellValue());  
                }  
            } else if (srcCellType == XSSFCell.CELL_TYPE_STRING) {  
                distCell.setCellValue(srcCell.getRichStringCellValue());  
            } else if (srcCellType == XSSFCell.CELL_TYPE_BLANK) {  
                // nothing21   
            } else if (srcCellType == XSSFCell.CELL_TYPE_BOOLEAN) {  
                distCell.setCellValue(srcCell.getBooleanCellValue());  
            } else if (srcCellType == XSSFCell.CELL_TYPE_ERROR) {  
                distCell.setCellErrorValue(srcCell.getErrorCellValue());  
            } else if (srcCellType == XSSFCell.CELL_TYPE_FORMULA) {  
                distCell.setCellFormula(srcCell.getCellFormula());  
            } else { // nothing29  
            }  
        }  
    }  
}  