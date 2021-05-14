package com.niemiec.logic;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.niemiec.model.Town;

public class SheetsManager {
	private List<Town> towns;
	private String path;
	private String firstFile;
	private String lastFile;
	
	public SheetsManager(String path, String firstFile, String lastFile) {
		towns = new ArrayList<Town>();
		this.path = path;
		this.firstFile = firstFile;
		this.lastFile = lastFile;
	}
	
	public void loadData() {
		
		for (int i = Integer.parseInt(firstFile); i <= Integer.parseInt(lastFile); i++) {
			Town town = new Town();
			FileInputStream file = null;
			HSSFWorkbook workbook = null;
			try {
				file = new FileInputStream(new File(path + i + ".xls"));
				workbook = new HSSFWorkbook(file);
			} catch (Exception e) {
				System.err.println(e);
			}
			
			HSSFSheet sheet = workbook.getSheetAt(0);
			Row row = sheet.getRow(1);
			if (row != null) {
				if (row.getCell(0) != null) {
					town.setName(row.getCell(0).toString());
					
					if (row.getCell(2) != null) {
						town.setStreet(row.getCell(2).toString());
					} else {
						town.setStreet("");
					}
				}
	
				if (town.getName() != null) {
					towns.add(town);
				}
			}
			
			Row mieszkancy = sheet.getRow(3);
			Row numer = sheet.getRow(5);
			Row licznik = sheet.getRow(7);
			
			if (numer != null) {
				int wiersze = sheet.getLastRowNum();
				int ludzie = 0;
				String dom = null;
				//System.out.println(town.getName() + " " + town.getStreet() + " " + wiersze);
				for (int j = 1; j <= wiersze; j++, ludzie++) {
					
					if (sheet.getRow(j).getCell(5) != null) {
						
						if (j > 1 || j == wiersze) {
							town.setHomeAndResidents(dom, Integer.toString(ludzie));
							
						}
						dom = sheet.getRow(j).getCell(5).toString();
						System.out.println("Dom: " + dom + " ");
						ludzie = 0;
					}
					//System.out.print("Dom: " + sheet.getRow(j).getCell(5) + " ");
					//System.out.println("Ludzie: " + ludzie);
				}
				
				//System.out.println(town.getName() + " " + town.getStreet() + " " + wiersze);
			}
			try {
				file.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		for (Town t: towns) {
			if (t.getName() != null)
				System.out.println(t.getName() + " " + t.getStreet() + " " + t.getHomes().getHomesNumbers() + " " + t.getHomes().getNumberOfResidents());
		}
	}
	
	
	
}
