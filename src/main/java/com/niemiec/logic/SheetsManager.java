package com.niemiec.logic;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
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
				double ludzie = 0.0;
				String dom = null;
				//System.out.println(town.getName() + " " + town.getStreet() + " " + wiersze);
				for (int j = 1; j <= wiersze; j++) {
					
					if (sheet.getRow(j).getCell(5) != null) {
						
						if (j > 1 || j == wiersze) {
							town.setHomeAndResidents(dom, "", Double.toString(ludzie));
							
						}
						dom = sheet.getRow(j).getCell(5).toString();
						//System.out.println("Dom: " + dom + " ");
						ludzie = 0;
					}
					//System.out.print("Dom: " + sheet.getRow(j).getCell(5) + " ");
					//System.out.println("Ludzie: " + ludzie);
					if (sheet.getRow(j).getCell(3) != null) {
						ludzie += Double.parseDouble(sheet.getRow(j).getCell(3).toString());
					}
					if (j == wiersze) {
						town.setHomeAndResidents(dom, "", Double.toString(ludzie));
					}
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
		
		saveInFile1(towns);
	}
	
	private void saveInFile1(List<Town> towns) {
		HSSFWorkbook docelowy = new HSSFWorkbook();
		HSSFSheet arkusz = docelowy.createSheet("Arkusz 1");
		
		int rowCount = 0;

		for (Town t: towns) {
			String city = t.getName();
			String street = t.getStreet();
			
			for (int i = 0; i < t.getHomes().getHomesNumbers().size(); i++) {
				int countCell = 0;
				Row row = arkusz.createRow(rowCount++);
				Cell cell = row.createCell(countCell++);
				cell.setCellValue(city);
				cell = row.createCell(countCell++);
				cell.setCellValue(street);
				cell = row.createCell(countCell++);
				cell.setCellValue(t.getHomes().getHomesNumbers().get(i));
				cell = row.createCell(countCell);
				cell.setCellValue(t.getHomes().getNumberOfResidents().get(i));
			}
			
			if (t.getHomes().getHomesNumbers().isEmpty()) {
				int countCell = 0;
				Row row = arkusz.createRow(rowCount++);
				Cell cell = row.createCell(countCell++);
				cell.setCellValue(city);
				cell = row.createCell(countCell++);
				cell.setCellValue(street);
				cell = row.createCell(countCell);
			}
			
		}
		
		try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\a.niemiec\\Desktop\\mieszkancy ulic\\Calosc.xls")) {
			docelowy.write(outputStream);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private void saveInFile2(List<Town> towns) {
		HSSFWorkbook docelowy = new HSSFWorkbook();
		HSSFSheet arkusz = docelowy.createSheet("Arkusz 1");
		
		int rowCount = 0;

		for (Town t: towns) {
			String city = t.getName();
			String street = t.getStreet();
			int countCell = 0;
			Row row = arkusz.createRow(rowCount++);
			int rezidents = 0;
			for (int i = 0; i < t.getHomes().getNumberOfResidents().size(); i++) {
				rezidents += Integer.parseInt(t.getHomes().getNumberOfResidents().get(i));
				
			}
			
			Cell cell = row.createCell(countCell++);
			cell.setCellValue(city);
			cell = row.createCell(countCell++);
			cell.setCellValue(street);
			cell = row.createCell(countCell);
			cell.setCellValue(rezidents);
			
			
		}
		
		try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\a.niemiec\\Desktop\\mieszkancy ulic\\Calosc2.xls")) {
			docelowy.write(outputStream);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
