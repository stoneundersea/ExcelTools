package com.bocsoft.ExcelTools;
import java.io.IOException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.image.Image;
import javafx.scene.control.MenuItem;
import javafx.scene.layout.Pane;
import javafx.stage.Stage;

@Component
//@FXMLController
public class MainController {
	static final Logger logger = LoggerFactory.getLogger(MainController.class);
	@FXML
	private Pane rootLayout;
	@FXML
	private MenuItem helpAbout;
     
  	public void initialize() {
	
	}
	
    @FXML
	public void helpAboutClicked(ActionEvent actionEvent) {
    	FXMLLoader loader = new FXMLLoader(getClass().getResource("/helpAbout.fxml"));
    	Stage aboutStage=new Stage();
    	try {
			aboutStage.setScene(new Scene(loader.load()));
		} catch (IOException e) {
			logger.error("异常",e);
			ToolUtils.showMsg(AlertType.ERROR,"系统错误","例外信息",e.getMessage());
			return;			
		}
    	aboutStage.setAlwaysOnTop(true);
    	aboutStage.setTitle("About Excel Tools");
    	aboutStage.getIcons().add(new Image("/ExcelTools.jpg"));
    	aboutStage.showAndWait();;  
    }
}

