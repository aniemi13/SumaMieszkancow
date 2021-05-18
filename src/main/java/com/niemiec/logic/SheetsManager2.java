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

public class SheetsManager2 {
	private List<Town> towns;
	private String path;
	private String firstFile;
	private String lastFile;
	
	public SheetsManager2(String path, String firstFile, String lastFile) {
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
			
			
			//if (sheet.getRow(1).getCell(5) != null) {
				int wiersze = sheet.getLastRowNum();
				String dom = "";
				
				//TODO tutaj je¿eli tylko jeden dom
				//int ileNiepustych = counta(wiersze, sheet);
				//System.out.println(sheet.getRow(1).getCell(0) + " " + ileNiepustych);
				
				for (int j = 1; j <= wiersze; j++) {
					double ludzie = 0.0;
					if (sheet.getRow(j).getCell(5) != null) {
						dom = sheet.getRow(j).getCell(5).toString();
						
						
						
						if (sheet.getRow(j).getCell(6) != null) {
							int licznik = 0;
							try {
								while (sheet.getRow(j + licznik + 1).getCell(5) == null && sheet.getRow(j + licznik + 1).getCell(3) != null) {
									licznik++;
								}
							} catch (Exception e) {
								// TODO: handle exception
							}
							for (int k = j; k < j + licznik + 1; k++) {
								String ulica = "";
								if (sheet.getRow(k).getCell(3) != null) {
									if (sheet.getRow(k).getCell(6) != null) {
										ulica = sheet.getRow(k).getCell(6).toString();
									}
									ludzie = Double.parseDouble(sheet.getRow(k).getCell(3).toString());
									town.setHomeAndResidents(dom, ulica, Double.toString(ludzie));
								}
							}
							
							
						} else {
							int licznik = 0;
							try {
								while (sheet.getRow(j + licznik + 1).getCell(5) == null && sheet.getRow(j + licznik + 1).getCell(3) != null) {
									licznik++;
								}
							} catch (Exception e) {
								// TODO: handle exception
							}
							for (int k = j; k < j + licznik + 1; k++) {
								if (sheet.getRow(k).getCell(3) != null)
									ludzie += Double.parseDouble(sheet.getRow(k).getCell(3).toString());
							}
							town.setHomeAndResidents(dom, "", Double.toString(ludzie));
							j = j + licznik;
							
						}
					}
				}
			//}
		}
		
		for (Town t: towns) {
			if (t.getName() != null)
				System.out.println(t.getName() + " " + t.getStreet() + " " + t.getHomes().getHomesNumbers() + " " + t.getHomes().getNumberOfResidents());
		}
		
		saveInFile1(towns);
		saveInFile2(towns);
	}
	
	private int counta(int wiersze, HSSFSheet sheet) {
		// TODO Auto-generated method stub
		int licznik = 0;
		for (int i = 0; i < wiersze; i++) {
			if (sheet.getRow(i).getCell(5) != null) {
				if (sheet.getRow(i).getCell(5).toString().length() > 0)
					licznik++;
			}
		}
		return licznik;
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
				cell = row.createCell(countCell++);
				cell.setCellValue(t.getHomes().getFlatNumbers().get(i));
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
		
		try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\a.niemiec\\Desktop\\mieszkancy ulic\\Calosc3.xls")) {
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
				rezidents += Double.parseDouble(t.getHomes().getNumberOfResidents().get(i));
				
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
