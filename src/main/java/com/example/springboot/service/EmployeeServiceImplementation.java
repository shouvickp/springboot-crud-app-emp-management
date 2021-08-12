package com.example.springboot.service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Optional;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.springboot.model.Employee;
import com.example.springboot.repository.EmployeeRepository;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

@Service
public class EmployeeServiceImplementation implements EmployeeService{

	@Autowired
	private EmployeeRepository employeeRepository;
	
	@Override
	public List<Employee> getAllEmployees() {
		return employeeRepository.findAll();
	}

	@Override
	public void saveEmployee(Employee employee) {
		this.employeeRepository.save(employee);
		
	}
	
	@Override
    public Employee getEmployeeById(long id) {
        Optional < Employee > optional = employeeRepository.findById(id);
        Employee employee = null;
        if (optional.isPresent()) {
            employee = optional.get();
        } else {
            throw new RuntimeException(" Employee not found for id :: " + id);
        }
        return employee;
    }

    @Override
    public void deleteEmployeeById(long id) {
        this.employeeRepository.deleteById(id);
    }
    
    @Override
    public boolean createPdfFile(List<Employee> employees, ServletContext context, HttpServletRequest request, HttpServletResponse response) {
    	 
    	Document document = new Document(PageSize.A4,15,15,45,30);
		try
		{
	       String filePath = context.getRealPath("/resources/reports");
	       File file = new File(filePath);
	       boolean exists = new File(filePath).exists();
	       if(!exists)
	       {
	    	   new File(filePath).mkdirs();
	    	   
	       }
	       
	       PdfWriter writer = PdfWriter.getInstance(document,
	    		   new FileOutputStream(file+"/"+"empList"+".pdf"));
	       document.open();
	       Font mainFont = FontFactory.getFont("Employee List",10,BaseColor.BLACK);
	       Paragraph paragraph = new Paragraph("Employee Details ", mainFont);
	       paragraph.setAlignment(Element.ALIGN_CENTER);
	       paragraph.setIndentationLeft(50);
	       paragraph.setIndentationRight(50);
	       paragraph.setSpacingAfter(10);
	       document.add(paragraph);
	       
	       PdfPTable table = new PdfPTable(3);
	       table.setWidthPercentage(100);
	       table.setSpacingBefore(10f);
	       table.setSpacingAfter(10);
	       
	       Font tableHeader = FontFactory.getFont("Arial", 10, BaseColor.BLACK);
	       Font tableBody = FontFactory.getFont("Arial", 9,BaseColor.BLACK);
	       
	       float[] columnWidths = {4f, 4f, 4f};
	       table.setWidths(columnWidths);
	       
	       PdfPCell  firstname= new PdfPCell(new Paragraph("First Name",tableHeader));
	       firstname.setBorderColor(BaseColor.BLACK);
	       firstname.setPaddingLeft(10);
	       firstname.setHorizontalAlignment(Element.ALIGN_CENTER);
	       firstname.setVerticalAlignment(Element.ALIGN_CENTER);
	       firstname.setBackgroundColor(BaseColor.GRAY);
	       firstname.setExtraParagraphSpace(5f);
	       table.addCell(firstname);
	       
	       PdfPCell  lastname= new PdfPCell(new Paragraph("Last Name",tableHeader));
	       lastname.setBorderColor(BaseColor.BLACK);
	       lastname.setPaddingLeft(10);
	       lastname.setHorizontalAlignment(Element.ALIGN_CENTER);
	       lastname.setVerticalAlignment(Element.ALIGN_CENTER);
	       lastname.setBackgroundColor(BaseColor.GRAY);
	       lastname.setExtraParagraphSpace(5f);
	       table.addCell(lastname);
	       
	       PdfPCell  email= new PdfPCell(new Paragraph("Email",tableHeader));
	       email.setBorderColor(BaseColor.BLACK);
	       email.setPaddingLeft(10);
	       email.setHorizontalAlignment(Element.ALIGN_CENTER);
	       email.setVerticalAlignment(Element.ALIGN_CENTER);
	       email.setBackgroundColor(BaseColor.GRAY);
	       email.setExtraParagraphSpace(5f);
	       table.addCell(email);
	       
	       for(Employee employee: employees) {
	    	   PdfPCell firstNamevalue = new PdfPCell(new Paragraph(employee.getFirstname(), tableHeader));
	    	   firstNamevalue.setBorderColor(BaseColor.BLACK);
	    	   firstNamevalue.setPaddingLeft(10);
	    	   firstNamevalue.setHorizontalAlignment(Element.ALIGN_LEFT);
	    	   firstNamevalue.setVerticalAlignment(Element.ALIGN_CENTER);
	    	   firstNamevalue.setBackgroundColor(BaseColor.WHITE);
	    	   firstNamevalue.setExtraParagraphSpace(5f);
			   table.addCell(firstNamevalue);
			   
			   PdfPCell lastNameValue = new PdfPCell(new Paragraph(employee.getLastname(), tableHeader));
			   lastNameValue.setBorderColor(BaseColor.BLACK);
			   lastNameValue.setPaddingLeft(10);
			   lastNameValue.setHorizontalAlignment(Element.ALIGN_LEFT);
			   lastNameValue.setVerticalAlignment(Element.ALIGN_CENTER);
			   lastNameValue.setBackgroundColor(BaseColor.WHITE);
			   lastNameValue.setExtraParagraphSpace(5f);
			   table.addCell(lastNameValue);
			   
			   PdfPCell emailValue = new PdfPCell(new Paragraph(employee.getEmail(), tableHeader));
			   emailValue.setBorderColor(BaseColor.BLACK);
			   emailValue.setPaddingLeft(10);
			   emailValue.setHorizontalAlignment(Element.ALIGN_LEFT);
			   emailValue.setVerticalAlignment(Element.ALIGN_CENTER);
			   emailValue.setBackgroundColor(BaseColor.WHITE);
			   emailValue.setExtraParagraphSpace(5f);
			   table.addCell(emailValue);
	       }
	      
	       document.add(table);
	       document.close();
	       return true;
			
	
		}
		catch(Exception ex)
		{
			return false;
		}
    }
    @Override
    public boolean createExcelFile(List<Employee> employees, ServletContext servletContext, HttpServletRequest request,
			HttpServletResponse response) {
    	String filePath = servletContext.getRealPath("/resources/reports");
	       File file = new File(filePath);
	       boolean exists = new File(filePath).exists();
	       if(!exists)
	       {
	    	   new File(filePath).mkdirs();
	    	   
	       }
	       try {
	    	   FileOutputStream outputStream = new FileOutputStream(file+"/"+"empList"+".xls");
	    	   HSSFWorkbook workbook = new HSSFWorkbook();
	    	   HSSFSheet worksheet = workbook.createSheet("Employees");
	    	   worksheet.setDefaultColumnWidth(30);
	    	   
	    	   HSSFCellStyle headerCellStyle = workbook.createCellStyle();
	    	   headerCellStyle.setFillForegroundColor(HSSFColor.BLUE.index);
//			   headerCellStyle.setFillPattern((short)HSSFCellStyle.SOLID_FOREGROUND);

			   HSSFRow headerRow = worksheet.createRow(0);
			   
			   HSSFCell firstName = headerRow.createCell(0);
			   firstName.setCellValue("First Name");
			   firstName.setCellStyle(headerCellStyle);
			   
			   HSSFCell lastName = headerRow.createCell(1);
			   lastName.setCellValue("Last Name");
			   lastName.setCellStyle(headerCellStyle);
			   
			   HSSFCell email = headerRow.createCell(2);
			   email.setCellValue("Email");
			   email.setCellStyle(headerCellStyle);
			   
			   int i=1;
			   for(Employee employee : employees) {
				   HSSFRow bodyRow = worksheet.createRow(i);
				   HSSFCellStyle bodyCellStyle = workbook.createCellStyle();
				   bodyCellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
				   
				   HSSFCell firstNameValue = bodyRow.createCell(0);
				   firstNameValue.setCellValue(employee.getFirstname());
				   firstNameValue.setCellStyle(bodyCellStyle);
				   
				   HSSFCell lastNameValue = bodyRow.createCell(1);
				   lastNameValue.setCellValue(employee.getLastname());
				   lastNameValue.setCellStyle(bodyCellStyle);
				   
				   HSSFCell emailValue = bodyRow.createCell(2);
				   emailValue.setCellValue(employee.getEmail());
				   emailValue.setCellStyle(bodyCellStyle);
				   
				   i++;
			   }
			   workbook.write(outputStream);
			   outputStream.flush();
			   outputStream.close();
			   return true;
	       }
	       catch(Exception ex)
			{
				return false;
			}
    }

}
