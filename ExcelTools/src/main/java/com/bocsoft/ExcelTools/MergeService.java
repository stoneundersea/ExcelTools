package com.bocsoft.ExcelTools;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import javafx.scene.control.Alert.AlertType;


@Service
//@Component
public class MergeService  {
	static final Logger logger = LoggerFactory.getLogger(MergeService.class);
	//@Autowired
	//private ToolUtils toolUtils;
	int replaceRecords=0,appendRecords=0;
	public void excelMerge(String targetFileName, 
			String sourceFileName,
			String targetSheetName,String sourceSheetName,
			TreeSet<Integer> marchColumns, 
			int keyColumn,String keyTxt) throws Exception {
		 //String fileType = targetFileName.substring(sourceFileName.lastIndexOf(".") + 1, targetFileName.length());
		 FileInputStream inputStreamTarget = new FileInputStream(targetFileName);
		 FileInputStream inputStreamSource = new FileInputStream(sourceFileName);
		 Workbook targetWb = new XSSFWorkbook(inputStreamTarget);
		 Workbook sourceWb = new XSSFWorkbook(inputStreamSource);
		 Sheet targetSheet = targetWb.getSheet(targetSheetName);
		 if(null == targetSheet) {
			 targetWb.close();
			 sourceWb.close();
			 throw new ToolsException("目标Sheet不存在");
		 }			 
		 Sheet sourceSheet = sourceWb.getSheet(sourceSheetName);
		 if(null == sourceSheet) {
			 sourceWb.close();
			 targetWb.close();
			 throw new ToolsException("源Sheet不存在");
		 }
			 
		 replaceRecords=0;
		 appendRecords=0;
		 mergeSheet(targetSheet,sourceSheet,marchColumns,keyColumn,keyTxt);
		 inputStreamTarget.close();
		 inputStreamSource.close();
		 FileOutputStream excelFileOutPutStream = new FileOutputStream(targetFileName);
		 targetWb.write(excelFileOutPutStream);
		 excelFileOutPutStream.flush();
		 excelFileOutPutStream.close();
		 targetWb.close();
		 sourceWb.close();
		 ToolUtils.showMsg(AlertType.INFORMATION, 
				 "运行结果说明","数据统计",
				 "目标文件："+targetFileName  + " Sheet ："	+ targetSheet.getSheetName() +"\n"+
				 "源文件："+ sourceFileName + " Sheet ："	+ sourceSheet.getSheetName() +"\n"+
				 "合并内容： "+keyTxt+"\n"+
				 "合并记录数："+String.valueOf(replaceRecords)+"\n"+
				 " 新增记录数： "+String.valueOf(appendRecords));
		 logger.info("运行结果说明:");
		 logger.info("目标文件："+targetFileName  + " Sheet ："	+ targetSheet.getSheetName());
		 logger.info("源文件："+ sourceFileName + " Sheet ："	+ sourceSheet.getSheetName());	
		 logger.info("合并内容： "+keyTxt);
		 logger.info("合并记录数："+String.valueOf(replaceRecords)+" 新增记录数： "+String.valueOf(appendRecords));
	}

	/**
	 * 
	 * @Author : YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description:  合并两个Excel Sheet 
	  *     先从源Sheet中根据keyColumn,keyTxt选出需要合并到目标Sheet中的Row；
	 *	再用源Sheet中选出的Row用marchColumn和目标Sheet的Row进行匹配
	  *      如果匹配上，用源Shee的这Row替换目标Sheet对应的Row，如果匹配不上
	  *      则将源Sheet的Row追加到目标Sheet的最后Row后面。
	 * @version ：1.0
	 * @param: 
	 *    @param targetSheet  目标Sheet
	 *    @param SourceSheet  源Sheet
	 *    @param marchColumns 目标Sheet和源Sheet用来匹配的Column 
	 *    @param keyColumn    源Sheet用来筛选的Column 
	 *    @param keyTxt       源Sheet用来筛选的Column对应的文本内容
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */

	public void mergeSheet(Sheet targetSheet,Sheet sourceSheet,TreeSet<Integer> marchColumns,int keyColumn,String keyTxt) {
		Row sourceRow;
		// 遍历源Sheet的每一Row，满足条件(非空,及筛选列其文本内容满足条件值)，进入目标Sheet进行匹配处理。
		for (int i=0; i<=sourceSheet.getLastRowNum(); i++) {
			sourceRow = sourceSheet.getRow(i);
			if(null!=sourceRow) {
				if((null!=sourceRow.getCell(keyColumn))&&(keyTxt.equals(sourceRow.getCell(keyColumn).getStringCellValue()))) {
					mergeRow(targetSheet,sourceSheet.getRow(i),marchColumns);
				}			
			}
		}
		
	}
	
	
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 将目标Row合并或追加到目标Sheet中  
	 * @version ：1.0
	 * @param: 
	 *    @param targetSheet  目标Sheet
	 *    @param sourceRow    需要合并的源Sheet中的一Row
	 *    @param marchColumns 用来匹配目标Sheet中合并Row的关键列
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public void mergeRow(Sheet targetSheet,Row sourceRow, TreeSet<Integer> marchColumns) {
		int replaceFlag = 0;
		Iterator<Integer> itSet = marchColumns.iterator();
		boolean marchFlag;
		Integer marchColumn;
		// 遍历目标Sheet中的Row，通过关键列和源Row匹配	
		for(int i=0;i<=targetSheet.getLastRowNum();i++) {
			marchFlag = true;
			//必须重新设置迭代器，以便下次匹配时重新开始迭代
			itSet = marchColumns.iterator();
			while(itSet.hasNext()) {
				marchColumn=(Integer)itSet.next();
				//对marchColumn集合中的每一列进行匹配
				if(!marchCell(targetSheet,sourceRow,i,marchColumn.intValue())) {
					marchFlag = false;
					break;
				}
			}
			//如果存在关键列都能匹配上目标Sheet中的Row，则用源Row进行覆盖
			if(marchFlag) {
				replaceFlag = 1;
				replaceRow(targetSheet.getRow(i), sourceRow);
				replaceRecords++;
				break;
			}
		}
		//如果Sheet中的所有Row都和源Row的关键列匹配不上，则用源Row在目标Sheet追加Row
		if(0==replaceFlag) {
				appendRow(targetSheet,sourceRow);
				appendRecords++;							
		}	
	}
	
	
	
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description:  对两个Excel单元格内容进行比较
	  *                为简化操作，都将单元格内容转化为字符串再进行比较
	 * @version ：1.0
	 * @param: 
	 *    @param targetSheet 待比较的单元格1所属Sheet  
	 *    @param sourceRow   待比较的单元格2所属Row
	 *    @param marchLine   待比较单元格1所在的Row编号
	 *    @param marchColumn 待比较两个单元格的列编号
	  *      只所以没有直接使用Cell作为参数，是应为存在为Null的情况，防止例外的处理。
	 *    @return
	 * @return boolean
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public boolean marchCell(Sheet targetSheet,Row sourceRow,int marchLine,int marchColumn)
	{
		boolean marchFlag = false;
		String targetValue,  sourceValue;
		CellType sourceCellType, targetCellType;
		if((null!=targetSheet.getRow(marchLine))&&
				(null!=targetSheet.getRow(marchLine).getCell(marchColumn))) {
			targetValue = getCellValue(targetSheet.getRow(marchLine).getCell(marchColumn));
		}else targetValue = null;
		
		if(null!=sourceRow.getCell(marchColumn)) sourceValue = getCellValue(sourceRow.getCell(marchColumn));
		else sourceValue = null;
			
		if(null!=sourceValue && null!=targetValue) {
			if(sourceValue.equals(targetValue)) marchFlag = true;
			else marchFlag = false;				
		}else if(null==sourceValue && null==targetValue) marchFlag = true;
		else marchFlag = false;
		
		return marchFlag;
	}
	
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 获取单元格的内容，统一转换为字符串
	 * @version ：1.0
	 * @param: 
	 *    @param cell
	 *    @return 单元格内容字符串 
	 * @return String
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public String getCellValue(Cell cell) {
		String returnString = null;
		switch(cell.getCellType())
		{ 
			case  BLANK: 
				returnString = null;
				break;
			case  BOOLEAN:  
				returnString = String.valueOf(cell.getBooleanCellValue());
				break;
			case  ERROR:  
				returnString = Byte.toString(cell.getErrorCellValue());
				break;
			case  FORMULA: 
				returnString = null;
				break;
			case  NUMERIC: 
				returnString = String.valueOf(cell.getNumericCellValue());
				break;
			case  STRING: 
				returnString = cell.getStringCellValue();
				break;
		}
		return returnString;
		
	}
	
	
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 用源Row单元格的格式及内容覆盖目标Row单元格的格式及内容
	 * @version ：1.0
	 * @param: 
	 *    @param targetRow     目标Row
	 *    @param sourceRow     源Row
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public void replaceRow(Row targetRow, Row sourceRow) {
		for(int i =0 ;i<sourceRow.getLastCellNum();i++) {
			if(null == targetRow.getCell(i)) targetRow.createCell(i);
			replaceCell(targetRow.getCell(i),sourceRow.getCell(i));
		}
		logger.info("replace one line: "+formTxtRow(sourceRow));
	}
	
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 用源单元格覆盖目标单元格的内容
	 * @version ： 1.0
	 * @param: 
	 *    @param targetCell  目标单元格
	 *    @param sourceCell  源单元格
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public void replaceCell(Cell targetCell, Cell sourceCell){
		CellType sourceCellType;
		CellStyle targetCellStyle;
		
		if(null!=sourceCell) {
			targetCellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
			targetCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
			targetCell.setCellStyle(targetCellStyle);
			sourceCellType = sourceCell.getCellType();
			if(null!=sourceCell.getCellComment()) replaceCellComment(targetCell,sourceCell.getCellComment());
			else if(null!=targetCell.getCellComment()) targetCell.setCellComment(null);
			switch(sourceCellType)
			{ 
				case  BLANK: 
					targetCell.setBlank();
					break;
				case  BOOLEAN:  
					targetCell.setCellValue(sourceCell.getBooleanCellValue());
					break;
				case  ERROR:  
					targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
					break;
				case  FORMULA: 
					targetCell.setCellFormula(sourceCell.getCellFormula());
					break;
				case  NUMERIC: 
					targetCell.setCellValue(sourceCell.getNumericCellValue());
					break;
				case  STRING: 
					targetCell.setCellValue(sourceCell.getRichStringCellValue());
			}
		}else targetCell = null;
	}
	
	
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 将源单元格的批注复制到目标单元格中 
	 * @version ：1.0
	 * @param: 
	 *    @param targetCell    目标单元格
	 *    @param sourceComment 源单元的批注
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public void replaceCellComment(Cell targetCell,Comment sourceComment) {
		if(null!=targetCell.getCellComment()) targetCell.setCellComment(null);
		ClientAnchor sourceClientAnchor = sourceComment.getClientAnchor();
		ClientAnchor targetClientAnchor 
			= targetCell.getSheet().createDrawingPatriarch().createAnchor(
					sourceClientAnchor.getDx1(), 
					sourceClientAnchor.getDy1(),
					sourceClientAnchor.getDx2(), 
					sourceClientAnchor.getDy2(),
					targetCell.getColumnIndex()+sourceComment.getClientAnchor().getCol1()-sourceComment.getColumn(),
					targetCell.getRowIndex()+sourceComment.getClientAnchor().getRow1()-sourceComment.getRow(),
					targetCell.getColumnIndex()+sourceComment.getClientAnchor().getCol2()-sourceComment.getColumn(),
					targetCell.getRowIndex()+sourceComment.getClientAnchor().getRow2()-sourceComment.getRow()
				);
		Comment targetComment = targetCell.getSheet().createDrawingPatriarch().createCellComment(targetClientAnchor);
		targetComment.setColumn(targetCell.getColumnIndex());
		targetComment.setRow(targetCell.getRowIndex());
		targetComment.setAuthor(sourceComment.getAuthor());
		targetComment.setString(sourceComment.getString());
		targetComment.setVisible(sourceComment.isVisible()); 
	}
	
	/**
	 * 
	 * @Author : YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 将源Row追加到目标Sheet的最后
	 * @version ：1.0
	 * @param: 
	 *    @param targetSheet  目标Sheet
	 *    @param sourceRow    源Row
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public void appendRow(Sheet targetSheet, Row sourceRow)
	{
		int appendRowNum=targetSheet.getLastRowNum();
		appendRowNum++;
		//在目标Sheet的最后Row后追加一Row,再用源Row覆盖
		Row targetRow = targetSheet.createRow(appendRowNum);
		for(int i =0 ;i<sourceRow.getLastCellNum();i++)
		{
			targetRow.createCell(i);
			replaceCell(targetRow.createCell(i), sourceRow.getCell(i));
		}
		logger.info("append one line: "+formTxtRow(sourceRow));
	}

	/**
	 * 
	 * @Author : YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 将一Row的内容转换为字符串，用于调试。
	 * @version ：1.0
	 * @param: 
	 *    @param iRow  Sheet中一Row
	 *    @return
	 * @return String
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public String formTxtRow(Row iRow)
	{
		StringBuffer  iStringBuffer = new StringBuffer();
		Cell sourceCell;
		if(null!=iRow) {
			for(int i =0 ;i<iRow.getLastCellNum();i++)
			{
				iStringBuffer.append("("+String.valueOf(i)+")");
				sourceCell = iRow.getCell(i);
				if(null!=sourceCell) {
					switch(iRow.getCell(i).getCellType())
					{
			    		case BLANK: 
			    			iStringBuffer.append(" ");
			    			break;
			    		case  BOOLEAN:   
			    			iStringBuffer.append(String.valueOf(iRow.getCell(i).getBooleanCellValue()));
			    			break;
			    		case  ERROR:   
			    			iStringBuffer.append("#Error");
			    			break;
			    		case  FORMULA: 
			    			iStringBuffer.append("#FORMULA");
			    			break;
			    		case  NUMERIC: 
			    			iStringBuffer.append(String.valueOf(iRow.getCell(i).getNumericCellValue()));
			    			break;
			    		case  STRING: iStringBuffer.append(iRow.getCell(i).getStringCellValue());			

					}
				}else iStringBuffer.append("#NULL");
			}
		}else 	iStringBuffer.append("This Row is NULL");
		return iStringBuffer.toString();
	}	
}
