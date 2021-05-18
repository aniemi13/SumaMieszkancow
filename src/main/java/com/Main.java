package com;

import com.niemiec.logic.SheetsManager;
import com.niemiec.logic.SheetsManager2;

public class Main {

	public static void main(String[] args) {
		
		String path = "C:\\Users\\a.niemiec\\Desktop\\mieszkancy ulic\\";
		String firstFile = "1";
		String lastFile = "185";
		
		
		SheetsManager2 manager = new SheetsManager2(path, firstFile, lastFile);
		manager.loadData();
	}

}
