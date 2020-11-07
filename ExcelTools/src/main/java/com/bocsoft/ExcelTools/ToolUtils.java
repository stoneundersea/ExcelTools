package com.bocsoft.ExcelTools;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import javafx.scene.control.Alert;
/**
* @Author : YangNJ
* @Create Date: 2020-11-6 22:54:47
* @Description:工具类，提供一些静态函数。
* @version ：1.0
* @Update Date :
* @Update By : 
* @Update Description:
*/
public class ToolUtils {
	static final Logger logger = LoggerFactory.getLogger(ToolUtils.class);
	/**
	 * 
	 * @Author :  YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: 利用Alert显示提示框
	 * @version ：1.0
	 * @param: 
	 *    @param alertType
	 *    @param msgTitle
	 *    @param msgHeadText
	 *    @param msgContenetText
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public static void showMsg(Alert.AlertType alertType,
			String msgTitle,String msgHeadText,String msgContenetText) {
		Alert alert = new Alert(alertType);
		alert.setTitle(msgTitle);
		alert.setHeaderText(msgHeadText);
		alert.setContentText(msgContenetText);
		alert.showAndWait();
	}
	
	public static boolean isNumeric(String str){   
		   Pattern pattern = Pattern.compile("[0-9]*");   
		   Matcher isNum = pattern.matcher(str);  
		   if( !isNum.matches() ){  
		       return false;   
		   }   
		   return true;   
	}  
	
	/**
	 * 
	 * @Author :
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description:
	  *   将字母组合的EXCEL列TITLE转换为int数字
	 * @version ：
	 * @param: 
	 *    @param columnTitle
	 *    @return
	 * @return int
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	public static int columnTitleToInt(String columnTitle) {
		int columnNum = 0;
		int tempNum;
		for(int i=0 ; i<columnTitle.length(); i++) {
			//大写字母转数字 A-1 B-2
			tempNum = columnTitle.charAt(i) - 'A' + 1;
			columnNum = columnNum+powerSum(tempNum,columnTitle.length()-i);		
		}
		//logger.info(columnTitle+"<->"+ String.valueOf(columnNum));
		return columnNum-1;
	}

	public static int powerSum(int sumNum,int powerNum) {		
		if(powerNum==1) {
			return sumNum; 
		}else{
			sumNum = sumNum*26;
			return(powerSum(sumNum,powerNum-1));			
		}
	}

}


