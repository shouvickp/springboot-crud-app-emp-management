package com.example.springboot.controller;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;
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
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
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
	
	@RequestMapping(value="/icardDownload/{id}", method= RequestMethod.GET)
	public void icardDownload(@PathVariable(value = "id") long id, HttpServletRequest request, HttpServletResponse response) {
		Employee employee = employeeService.getEmployeeById(id);
		boolean isFlag = employeeService.createIDCard(employee, servletContext , request, response);
		 if(isFlag)
		  {
			  System.out.println("icard create hochche");
			  String fullPath = request.getServletContext().getRealPath("resources/reports/"+"emp_"+id+".pdf");
			  System.out.println(fullPath);
			  filedownload(fullPath,response,"emp_"+id+"_IDCARD.pdf");
		  }
	}
	
	@RequestMapping(value="/downloadAllICard", method= RequestMethod.GET)
	public void downloadAll(HttpServletRequest req, HttpServletResponse res) throws IOException {
		List<String> files = new ArrayList<>();
		List<Employee> employees = employeeService.getAllEmployees();
		for(Employee employee : employees) {
			boolean isFlag = employeeService.createIDCard(employee, servletContext , req, res);
			 if(isFlag)
			  {
				  System.out.println("icard create hochche");
				  String fullPath = req.getServletContext().getRealPath("resources/reports/"+"emp_"+employee.getId()+".pdf");
				  files.add(fullPath);
//				  filedownload(fullPath,res,"emp_"+employee.getId()+"_IDCARD.pdf");
			  }
		}
	
		if(files!=null && files.size()>0) {
			downloadZipFile(files,"IDCards.zip",res);
		}	        
	    
	}
	
	private void downloadZipFile(List<String> files, String zipFile, HttpServletResponse res) {
		try {
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			ZipOutputStream zos = new ZipOutputStream(baos);
			byte[] bytes = new byte[4096];
			for(String file :files) {
				FileInputStream fis = new FileInputStream(file);
				BufferedInputStream bis = new BufferedInputStream(fis);
				zos.putNextEntry(new ZipEntry(file.substring(file.lastIndexOf("\\")+1)));
				int bytesRead;
				while((bytesRead = bis.read(bytes)) != -1) {
					zos.write(bytes, 0, bytesRead);
				}
				zos.closeEntry();
				bis.close();
				fis.close();
			}
			zos.flush();
			baos.flush();
			zos.close();
			baos.close();
			
			byte[] zip = baos.toByteArray();
			ServletOutputStream sos = res.getOutputStream();
			res.setContentType("application/zip");
			res.setHeader("content-disposition", "attachment;fileName="+zipFile);
			sos.write(zip);
			sos.flush();
			sos.close();
			
		}
		catch(Exception e) {
			e.printStackTrace();
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
			  String fullPath = request.getServletContext().getRealPath("resources/reports/"+"empList"+".xlsx");
			  System.out.println(fullPath);
			  filedownload(fullPath,response,"empList.xlsx");
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
	
	
	@PostMapping("/uploadEmployeeList")
	public String uploadMultipartFile(@RequestParam("xlFile") MultipartFile file, Model model) {
		try {
			employeeService.store(file);
//			model.addAttribute("message", "File uploaded successfully!");
		} catch (Exception e) {
			model.addAttribute("message", "Fail! -> uploaded filename: " + file.getOriginalFilename());
		}
        return "redirect:/";
    }
}
