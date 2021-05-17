package com.niemiec.model;

import java.util.ArrayList;
import java.util.List;

public class Homes {
	private List<String> homesNumbers;
	private List<String> flatNumbers;
	private List<String> numberOfResidents;
	
	public Homes() {
		homesNumbers = new ArrayList<String>();
		numberOfResidents = new ArrayList<String>();
	}
	
	public void setHome(String homeNumber, String flatNumber, String numerOfResidents) {
		homesNumbers.add(homeNumber);
		flatNumbers.add(flatNumber);
		numberOfResidents.add(numerOfResidents);
	}

	public List<String> getHomesNumbers() {
		return homesNumbers;
	}
	
	public List<String> getFlatNumbers() {
		return flatNumbers;
	}

	public List<String> getNumberOfResidents() {
		return numberOfResidents;
	}
}
