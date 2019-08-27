package com.insert_excel;

import java.util.Date;

class Employee {
    private String name;
    private String email;
    private int userid;
    private double salary;

    public Employee(String name, String email, int userid, double salary) {
        this.name = name;
        this.email = email;
        this.userid = userid;
        this.salary = salary;
    }

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public int getuserId() {
		return userid;
	}

	public void setuserId(int userid) {
		this.userid = userid;
	}

	public double getSalary() {
		return salary;
	}

	public void setSalary(double salary) {
		this.salary = salary;
	}

	
}