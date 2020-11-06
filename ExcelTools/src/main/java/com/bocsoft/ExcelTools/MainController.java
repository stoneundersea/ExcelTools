package com.bocsoft.ExcelTools;

import java.io.File;
import java.util.TreeSet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.Pane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

@Component
//@FXMLController
public class MainController {
	static final Logger logger = LoggerFactory.getLogger(MainController.class);
	@FXML
	private Pane rootLayout;
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
    private Button buttonMerge;    
    @Autowired
    private MergeService mergeService;

    final FileChooser fileChooser = new FileChooser();
	public void initialize() {
		//this.priFileName.setText("0000000");
		//this.priFileName.setText("D:/YangJX_Doc/物理竞赛学生跟踪/2020物理国集 - target.xlsx");
        //this.secFileName.setText("D:/YangJX_Doc/物理竞赛学生跟踪/2020物理国集.xlsx");
		this.mergeColumn1.setText("A");
		this.sourceSheetName.setText("Sheet1");
		this.targetSheetName.setText("Sheet1");
		//this.mergeColumn.setText("0");
		//System.out.println(mergeService);
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
		if(checkInput()) {
			marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn1.getText().trim()));
			if(checkTextInput(mergeColumn2.getText()))
				marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn2.getText().trim()));
			if(checkTextInput(mergeColumn3.getText()))
				marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn3.getText().trim()));
			if(checkTextInput(mergeColumn4.getText()))
				marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn4.getText().trim()));
			if(checkTextInput(mergeColumn5.getText()))
				marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn5.getText().trim()));
			if(checkTextInput(mergeColumn6.getText()))
				marchColumns.add(ToolUtils.columnTitleToInt(mergeColumn6.getText().trim()));
			try {
				mergeService.excelMerge(priFileName.getText(),
					secFileName.getText(),
					targetSheetName.getText(),
					sourceSheetName.getText(),
					marchColumns,
					ToolUtils.columnTitleToInt(keyColumn.getText().trim()),
					keyTxt.getText());
				} catch (Exception e) {
				// TODO Auto-generated catch block
					logger.error("异常",e);
					ToolUtils.showMsg(AlertType.ERROR,"系统错误","例外信息",e.getMessage());
				
			}    	
    	}
       	
    }

	private void configureFileChooser(FileChooser fileChooser) {
		// TODO Auto-generated method stub
		fileChooser.setTitle("Select xlsx File");
        fileChooser.setInitialDirectory(
            new File(System.getProperty("user.home"))
        ); 
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("xlsx", "*.xlsx")
         );
		
	}
	
	//Check all input Control box context.还需要补充
	public boolean checkInput() {
		boolean checkReturn=true;
		if(!checkTextInput(targetSheetName.getText())) {
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","选择目标SheetNum必须输入数字!");
		}else if(!checkTextInput(sourceSheetName.getText())) {
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","选择源SheetNum必须输入数字!");
		}else if(!checkTextInput(priFileName.getText())){
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","选择合并主文件不能为空!");
		}else if(!checkTextInput(secFileName.getText())){
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","选择合并从文件不能为空!");
		}else if(!checkTextInput(mergeColumn1.getText())){
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","合并匹配列必须输入数字!");
		}else if(!checkTextInput(keyColumn.getText())){
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","合并筛选列必须输入数字!");
		}else if(!checkTextInput(keyTxt.getText())){
			checkReturn=false;
			ToolUtils.showMsg(AlertType.WARNING, "警示信息","输入错误","合并筛选内容不能为空！");
		}
     	return checkReturn;	
	}
	
	//Check digit input,must not null,length great 0,all context is numeric.
	public boolean checkDigitInput(String digitInput) {
		boolean checkReturn=true;
		if((null == digitInput)||(digitInput.length()<= 0)||
     			(!ToolUtils.isNumeric(digitInput.trim())))
			checkReturn = false;
		return checkReturn;		
	}	
	
	//Check Text Inout,must not null,length great 0.
	public boolean checkTextInput(String textInput) {
		boolean checkReturn=true;
		if((null == textInput)||(textInput.length()<= 0))
			checkReturn = false;
		return checkReturn;		
	}	
	
	
}