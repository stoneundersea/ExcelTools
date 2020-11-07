package com.bocsoft.ExcelTools;
/**
* @Author : YangNJ
* @Create Date: 2020-11-7 10:15:01
* @Description: 自定一异常类
* @version ：1.0
* @Update Date :
* @Update By : 
* @Update Description:
*/
public class ToolsException extends Exception{
    public ToolsException() {
    	super();
    }
 
    public ToolsException(String message) {
    	super(message);
    }
 
    public ToolsException(String message, Throwable cause) {
        super(message, cause);
    }
 
    public ToolsException(Throwable cause) {
        super(cause);
    }
	

}
