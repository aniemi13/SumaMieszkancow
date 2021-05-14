package com.niemiec;

import com.niemiec.logic.SheetsManager;

public class Main {

	public static void main(String[] args) {
		
		String path = "C:\\Users\\a.niemiec\\Desktop\\mieszkancy ulic\\";
		String firstFile = "1";
		String lastFile = "10";
		
		
		SheetsManager manager = new SheetsManager(path, firstFile, lastFile);
		manager.loadData();
	}

}
