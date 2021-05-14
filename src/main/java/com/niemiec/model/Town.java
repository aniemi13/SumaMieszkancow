package com.niemiec.model;

public class Town {
	private String name;
	private String street;
	private Homes homes;
	
	public Town() {
		homes = new Homes();
	}
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getStreet() {
		return street;
	}
	public void setStreet(String street) {
		this.street = street;
	}
	
	public void setHomeAndResidents(String home, String residents) {
		homes.setHome(home, residents);
	}
	
	public Homes getHomes() {
		return homes;
	}
}
