package com;

import com.niemiec.logic.SheetsManager;

public class Main {

	public static void main(String[] args) {
		
		String path = "C:\\Users\\a.niemiec\\Desktop\\mieszkancy ulic\\";
		String firstFile = "1";
		String lastFile = "185";
		
		
		SheetsManager manager = new SheetsManager(path, firstFile, lastFile);
		manager.loadData();
	}

}
