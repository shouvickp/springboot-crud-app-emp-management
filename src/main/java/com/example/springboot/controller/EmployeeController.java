package com.example.springboot.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.io.IOException;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.WebContext;

import com.example.springboot.model.Employee;
import com.example.springboot.service.EmployeeService;
import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.io.source.ByteArrayOutputStream;
import org.springframework.http.MediaType;


@Controller
public class EmployeeController {
	
	@Autowired
	private EmployeeService employeeService;
	@Autowired
    private ServletContext servletContext;

    private final TemplateEngine templateEngine;

    public EmployeeController(TemplateEngine templateEngine) {
        this.templateEngine = templateEngine;
    }
	//	display list of employees
	@GetMapping("/")
	public String viewHomePage(Model model) {
		model.addAttribute("listEmployees", employeeService.getAllEmployees());
		return "index";
	}
	
	@GetMapping("/showNewEmployeeForm")
	public String showNewEmployeeform(Model model) {
		//create model attribute to bind form data
		Employee employee =new Employee();
		model.addAttribute("employee",employee);
		return"new_employee";
	}
	
	@PostMapping("/saveEmployee")
	public String saveEmployee(@ModelAttribute("employee") Employee employee) {
		//save employee to database
		employeeService.saveEmployee(employee);
		return "redirect:/";
	}
	
	@GetMapping("/showFormForUpdate/{id}")
    public String showFormForUpdate(@PathVariable(value = "id") long id, Model model) {

        // get employee from the service
        Employee employee = employeeService.getEmployeeById(id);

        // set employee as a model attribute to pre-populate the form
        model.addAttribute("employee", employee);
        return "update_employee";
    }
	
	@GetMapping("/deleteEmployee/{id}")
    public String deleteEmployee(@PathVariable(value = "id") long id) {

        // call delete employee method 
        this.employeeService.deleteEmployeeById(id);
        return "redirect:/";
    }
	
	@RequestMapping(value="/download", method= RequestMethod.GET)
	public void createPdf(HttpServletRequest request, HttpServletResponse response) {
		List<Employee> employees = employeeService.getAllEmployees();
		boolean isFlag = employeeService.createPdfFile(employees, servletContext , request, response);
		 if(isFlag)
		  {
			  System.out.println("file create hochche");
			  String fullPath = request.getServletContext().getRealPath("resources/reports/"+"empList"+".pdf");
			  System.out.println(fullPath);
			  filedownload(fullPath,response,"empList.pdf");
		  }
	}
	
	private void filedownload(String fullPath, HttpServletResponse response, String fileName) {
		File file = new File(fullPath);
		final int BUFFER_SIZE = 4096;
		if(file.exists())
		{
			try
			{
				FileInputStream  fis= new FileInputStream(file);
				String mimeType= servletContext.getMimeType(fullPath);
				response.setContentType(mimeType);
				response.setHeader("content-disposition", "attachment;fileName="+fileName);

				OutputStream os= response.getOutputStream();
				byte[] buffer= new byte[BUFFER_SIZE];
				int bytesRead = -1;
				while((bytesRead=fis.read(buffer))!=-1)
				{
					os.write(buffer, 0, bytesRead);
				}
				
				fis.close();
				os.close();
				file.delete();
			}
			catch(Exception ex)
			{
				ex.printStackTrace();
			}
		}
		
	}
	
	@RequestMapping(value="/downloadExcel", method= RequestMethod.GET)
	public void createExcel(HttpServletRequest request, HttpServletResponse response) {
		List<Employee> employees = employeeService.getAllEmployees();
		boolean isFlag = employeeService.createExcelFile(employees, servletContext , request, response);
		 if(isFlag)
		  {
			  System.out.println("file create hochche");
			  String fullPath = request.getServletContext().getRealPath("resources/reports/"+"empList"+".xls");
			  System.out.println(fullPath);
			  filedownload(fullPath,response,"empList.xls");
		  }
	}
	
	@RequestMapping(path = "/exportEmployeeList")
    public ResponseEntity<?> getPDF(HttpServletRequest request, HttpServletResponse response) throws IOException {

        /* Do Business Logic*/
		List<Employee> employees = employeeService.getAllEmployees();
        /* Create HTML using Thymeleaf template Engine */

        WebContext context = new WebContext(request, response, servletContext);
        context.setVariable("listEmployees", employees);
        String empHtml = templateEngine.process("index", context);

        /* Setup Source and target I/O streams */

        ByteArrayOutputStream target = new ByteArrayOutputStream();

        /*Setup converter properties. */
        ConverterProperties converterProperties = new ConverterProperties();
        converterProperties.setBaseUri("http://localhost:8080");

        /* Call convert method */
        HtmlConverter.convertToPdf(empHtml, target, converterProperties);  

        /* extract output as bytes */
        byte[] bytes = target.toByteArray();


        /* Send the response as downloadable PDF */

        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_PDF)
                .body(bytes);

    }
}
