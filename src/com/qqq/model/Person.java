package com.qqq.model;

public class Person {
	String name;
	String department;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getDepartment() {
		return department;
	}

	public void setDepartment(String department) {
		this.department = department;
	}

	@Override
	public String toString() {
		return "Person [name=" + name + ", department=" + department + "]";
	}

	
}
