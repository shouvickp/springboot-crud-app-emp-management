package com.example.springboot.service;

import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.web.multipart.MultipartFile;

import com.example.springboot.model.Employee;

public interface EmployeeService {
	
	List<Employee> getAllEmployees();
	
	void saveEmployee(Employee employee);
	
	Employee getEmployeeById(long id);
	
    void deleteEmployeeById(long id);

	boolean createPdfFile(List<Employee> employees, ServletContext servletContext, HttpServletRequest request,
			HttpServletResponse response);

	boolean createExcelFile(List<Employee> employees, ServletContext servletContext, HttpServletRequest request,
			HttpServletResponse response);

	boolean createIDCard(Employee employee, ServletContext servletContext, HttpServletRequest request,
			HttpServletResponse response);

	void store(MultipartFile file);
	
}
