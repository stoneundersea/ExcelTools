package com.bocsoft.ExcelTools;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.TreeSet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.image.Image;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.AnchorPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
/**
* @Author :  YangNJ
* @Create Date: 2020-11-14 16:49:36
* @Description: MergerTabPane Controller,合并Sheet内容
* @version ： V1.0
* @Update Date :
* @Update By : 
* @Update Description:
*/



@Component
//@FXMLController
public class MergeTabController {
	static final Logger logger = LoggerFactory.getLogger(MainController.class);
	@FXML
	private AnchorPane rootLayout;
	@FXML
	private TextField priFileName;
    @FXML
	private TextField secFileName;
    @FXML
    private Button buttonPriFileSelect;
    @FXML
    private Button buttonSecFileSelect;
    @FXML
    private TextField targetSheetName;
    @FXML
    private TextField sourceSheetName;
    @FXML
    private TextField mergeColumn1;
    @FXML
    private TextField mergeColumn2;
    @FXML
    private TextField mergeColumn3;
    @FXML
    private TextField mergeColumn4;
    @FXML
    private TextField mergeColumn5;
    @FXML
    private TextField mergeColumn6;
    @FXML
    private TextField keyColumn;
    @FXML
    private TextField keyTxt;    
    @FXML
    private Button selectSourceSheet;
    @FXML
    private Button selectTargetSheet;   
    @FXML
    private Button buttonMerge; 
    @Autowired
    private MergeService mergeService;
    
    final FileChooser fileChooser = new FileChooser();
	public void initialize() {
		//文本框属性绑定，控制文本变化的对应行为
		this.mergeColumn1.textProperty().addListener((observable,oldvalue,newvalue)->{
			checkLetterTxt(newvalue);
		});
		this.mergeColumn2.textProperty().addListener((observable,oldvalue,newvalue)->{
			checkLetterTxt(newvalue);
		});
		this.mergeColumn3.textProperty().addListener((observable,oldvalue,newvalue)->{
			checkLetterTxt(newvalue);
		});
		this.mergeColumn4.textProperty().addListener((observable,oldvalue,newvalue)->{
			checkLetterTxt(newvalue);
		});
		this.mergeColumn5.textProperty().addListener((observable,oldvalue,newvalue)->{
			checkLetterTxt(newvalue);
		});
		this.mergeColumn6.textProperty().addListener((observable,oldvalue,newvalue)->{
			checkLetterTxt(newvalue);
		});
		this.mergeColumn1.setText("A");
		this.sourceSheetName.setText("Sheet1");
		this.targetSheetName.setText("Sheet1");
	}
	
    @FXML
	public void selectButtonClicked(ActionEvent actionEvent) {
	   	//在Controller中获取到Stage
		Stage primaryStage = (Stage) rootLayout.getScene().getWindow();
		configureFileChooser(fileChooser);
		File file = fileChooser.showOpenDialog(primaryStage);
		if(buttonPriFileSelect ==(Button)actionEvent.getSource()) {
			priFileName.setText(file.getAbsolutePath());
		}
		if(buttonSecFileSelect ==(Button)actionEvent.getSource()) {
			secFileName.setText(file.getAbsolutePath());
		}
		
	}
    
    @FXML    
    public void mergeButtonClicked(ActionEvent actionEvent) {
		TreeSet<Integer> marchColumns = new TreeSet<Integer>();
		try {
			checkInput();
		} catch (Exception e1) {
			logger.info("警示信息",e1);
			ToolUtils.showMsg(AlertType.WARNING,"警示信息","输入错误",e1.getMessage());
			return;
		}
		marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn1.getText().trim()));
		if(checkTextInput(mergeColumn2.getText()))
			marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn2.getText().trim().toUpperCase()));
		if(checkTextInput(mergeColumn3.getText()))
			marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn3.getText().trim().toUpperCase()));
		if(checkTextInput(mergeColumn4.getText()))
			marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn4.getText().trim().toUpperCase()));
		if(checkTextInput(mergeColumn5.getText()))
			marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn5.getText().trim().toUpperCase()));
		if(checkTextInput(mergeColumn6.getText()))
			marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn6.getText().trim().toUpperCase()));
		try {
			mergeService.excelMerge(priFileName.getText(),
				secFileName.getText(),
				targetSheetName.getText(),
				sourceSheetName.getText(),
				marchColumns,
				ToolUtils.columnTitleToInt(keyColumn.getText().trim()),
				keyTxt.getText());
			} catch (Exception e) {
				logger.error("异常",e);
				ToolUtils.showMsg(AlertType.ERROR,"系统错误","例外信息",e.getMessage());
				
		}    	
    }
       	

	private void configureFileChooser(FileChooser fileChooser) {
		fileChooser.setTitle("Select xlsx File");
        fileChooser.setInitialDirectory(
            new File(System.getProperty("user.home"))
        ); 
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("xlsx", "*.xlsx")
         );
		
	}
	
	//Check all input Control box context.还需要补充
	//Update with exception replace if else.
	public void checkInput() throws Exception {
		if(!checkTextInput(targetSheetName.getText())) 
			throw new ToolsException("选择目标SheetName必须输入!");			
		if(!checkTextInput(sourceSheetName.getText())) 
			throw new ToolsException("选择源SheetName必须输入!");			
		if(!checkTextInput(priFileName.getText()))
			throw new ToolsException("选择合并主文件不能为空!");
		if(!checkTextInput(secFileName.getText()))
			throw new ToolsException("选择合并从文件不能为空!");
		if(!checkTextInput(mergeColumn1.getText())||!checkLetterTxt(mergeColumn1.getText()))
			throw new ToolsException("合并匹配列1必须输入字母!");
		if(!checkLetterTxt(mergeColumn2.getText())) 
			throw new ToolsException("合并匹配列2必须输入字母或空!");
		if(!checkLetterTxt(mergeColumn3.getText()))
			throw new ToolsException("合并匹配列3必须输入字母或空!");	
		if(!checkLetterTxt(mergeColumn4.getText()))
			throw new ToolsException("合并匹配列4必须输入字母或空!");
		if(!checkLetterTxt(mergeColumn5.getText()))
			throw new ToolsException("合并匹配列5必须输入字母或空!");
		if(!checkLetterTxt(mergeColumn6.getText()))
			throw new ToolsException("合并匹配列6必须输入字母或空!");
		if(!checkTextInput(keyColumn.getText()))
			throw new ToolsException("合并筛选列必须输入数字!");			
		if(!checkTextInput(keyTxt.getText()))
			throw new ToolsException("合并筛选内容不能为空!");
	}
	
	//Check digit input,must not null,length great 0,all context is numeric.
	public boolean checkDigitInput(String digitInput) {
		boolean checkReturn=true;
		if((null == digitInput)||(digitInput.length()<= 0)||
     			(!ToolUtils.isNumeric(digitInput.trim())))
			checkReturn = false;
		return checkReturn;		
	}	
	
	
	/**
	 * @Author :YangNJ
	 * @Create Date: 2020-11-1 11:06:44
	 * @Description: Excel Sheet select function,使用弹出子窗口的方式。
	 * @version ：V1.0
	 * @param: 
	 *    @param actionEvent
	 * @return void
	 * @throws 
	 *-------------------------------
	 * @Update Date :
	 * @Update By : 
	 * @Update Description:
	 */
	
	@FXML
	public void selectSheetButtonClicked(ActionEvent actionEvent)  {
		String execlFileName = null;
		String choicedSheet = null;
		ArrayList<String> sheetNames;
		FXMLLoader loader = new FXMLLoader(getClass().getResource("/selectWindows.fxml"));
		Stage choiceStage=new Stage();//创建舞台；
		choiceStage.setAlwaysOnTop(true);
		try {
			choiceStage.setScene(new Scene(loader.load()));//将场景载入舞台
		} catch (IOException e) {
			logger.error("异常",e);
			ToolUtils.showMsg(AlertType.ERROR,"系统错误","例外信息",e.getMessage());
			return;
		}  
		if(selectSourceSheet==(Button)actionEvent.getSource()) {
			choiceStage.setTitle("选择源Sheet");
			execlFileName = secFileName.getText();
		}
		if(selectTargetSheet==(Button)actionEvent.getSource()) {
			choiceStage.setTitle("选择目标Sheet");
			execlFileName = priFileName.getText();
		}
		ChoiceController choiceController =  loader.getController();
		try {
			sheetNames = mergeService.excelQuerySheets(execlFileName);
			//Caller传递参数给Controller
			choiceController.initData(sheetNames);
		} catch (Exception e) {
			logger.error("异常",e);
			ToolUtils.showMsg(AlertType.ERROR,"系统错误","例外信息",e.getMessage());
			return;
		}
		choiceStage.getIcons().add(new Image("/ExcelTools.jpg"));
		choiceStage.showAndWait(); //显示窗口；
		//从子窗口获取选择的sheet，如果非空则更新相应的文本框内容
		choicedSheet = choiceController.getChoicedSheet();
		if(null!=choicedSheet) {
			if(selectSourceSheet==(Button)actionEvent.getSource())
				sourceSheetName.setText(choicedSheet);
			if(selectTargetSheet==(Button)actionEvent.getSource())
				targetSheetName.setText(choicedSheet);
		}
	}
	
	
	//Check Text Inout,must not null,length great 0.
	public boolean checkTextInput(String textInput) {
		boolean checkReturn=true;
		if((null == textInput)||(textInput.length()<= 0))
			checkReturn = false;
		return checkReturn;		
	}	
	
	public boolean checkLetterTxt(String txtString) {
		return(0==txtString.length()||txtString.matches("[a-zA-Z]+"));
	}

}
