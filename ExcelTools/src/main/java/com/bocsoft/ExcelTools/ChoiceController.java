package com.bocsoft.ExcelTools;

import java.util.ArrayList;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;

import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.layout.Pane;
import javafx.stage.Stage;


/**
* @Author : YangNJ
* @Create Date: 2020-11-14 7:59:42
* @Description: 弹出选择框处理，利用父窗口的数据设置选择框的选项，
*      并将选择结果保存，待父窗口Controller获取使用
* @version ：V1.0
* @Update Date :
* @Update By : 
* @Update Description:
*/
@Component
public class ChoiceController {
	static final Logger logger = LoggerFactory.getLogger(ChoiceController.class);
	@FXML
	private Pane rootLayout;
    @FXML
    private Button buttonOK;
    @FXML
    private Button buttonCancel;
    @FXML
    private ChoiceBox<String> choiceBoxSelect;
    
    private String choicedSheet;
    
   	public void initialize() {
    	
    }
    
   	//使用父窗口传递过来的数据设置选择款的选项
    public void initData(ArrayList<String> choiceList) {
    	ObservableList<String> observableList = FXCollections.observableList(choiceList);
    	choiceBoxSelect.setItems(observableList);
    	choiceBoxSelect.getSelectionModel().select(0);
    }
    
    public String getChoicedSheet() {
    	return choicedSheet;
    }
    
    @FXML
   	public void choiceBoxClicked(ActionEvent actionEvent) {
    	
    }
    
    @FXML
   	public void buttonClicked(ActionEvent actionEvent) {
    	Stage stage;
    	if(buttonOK==(Button)actionEvent.getSource()) {
    		choicedSheet = choiceBoxSelect.getSelectionModel().getSelectedItem();
    		stage = (Stage)buttonOK.getScene().getWindow(); 
    		stage.close();
    	}
    	if(buttonCancel==(Button)actionEvent.getSource()) {
    		choicedSheet = null;
    		stage = (Stage)buttonOK.getScene().getWindow(); 
    		stage.close();
    	}
    }

}
