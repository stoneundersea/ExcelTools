package com.bocsoft.ExcelTools;

import java.io.IOException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

@SpringBootApplication
// 设置Spring boot配置参数位置
//@PropertySource(value={"file:./config/application.properties"})
public class ExcelToolsApplication extends Application {
	static final Logger logger = LoggerFactory.getLogger(ExcelToolsApplication.class);

	//private static ConfigurableApplicationContext applicationContext;
	private static ApplicationContext applicationContext;
	public static void main(String[] args) {
		//SpringApplication.run(ExcelToolsApplication.class, args);
		applicationContext = SpringApplication.run(ExcelToolsApplication.class, args);
		Application.launch(ExcelToolsApplication.class,args);
	}

	@Override
	public void start(Stage primaryStage) throws Exception {
		primaryStage.setTitle("Excel Tools");
		primaryStage.setScene(new Scene(createRoot()));
		primaryStage.getIcons().add(new Image("/ExcelTools.jpg"));
		primaryStage.show();
		
	}

	private Parent createRoot() throws IOException {
		FXMLLoader fxmlLoader = new FXMLLoader();
		fxmlLoader.setLocation(getClass().getResource("/main.fxml"));
		fxmlLoader.setControllerFactory(applicationContext::getBean);
		return fxmlLoader.load();
	}

}
