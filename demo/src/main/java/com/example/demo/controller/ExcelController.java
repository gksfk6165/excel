package com.example.demo.controller;

import java.util.HashMap;

import org.apache.commons.io.FilenameUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.service.Excelservice2;

@Controller
public class ExcelController {
	
	@Autowired
	Excelservice2 service;
	
	@GetMapping(value="/")
	public String test() throws Exception{
		return "test";
	}
	
	@GetMapping(value="/test")
	public String excelselect(Model model) throws Exception{
		model.addAttribute("list",service.excelselect());
		return "excel";
	}
	
	
	
	@PostMapping(value="/excel/read")
	public String readExcel(@RequestParam("file") MultipartFile file,
							@RequestParam("group_code") String group_code,
							@RequestParam("sample_code") String sample_code,
							@RequestParam("depth1") String depth1,
							@RequestParam("depth2") String depth2,
							@RequestParam("depth3") String depth3) throws Exception{
		String filename="";
		try {
			filename = file.getOriginalFilename();
		} catch (Exception e) {
			System.out.println("파일 이름 에러");
		}
		
		System.out.println("파일 이름 : " + filename);
		System.out.println("그룹 코드 이름 : " + group_code);
		System.out.println("샘플 코드 이름 : " + sample_code);
		System.out.println("depth1 이름 : " + depth1);
		System.out.println("depth2 이름 : " + depth2);
		System.out.println("depth3 이름 : " + depth3);
		
		//cell_sample_type 에들어갈 데이터들 
		HashMap<String, Object> hmap = new HashMap<>();
		hmap.put("group_code", group_code);
		hmap.put("sample_code", sample_code);
		hmap.put("depth1", depth1);
		hmap.put("depth2", depth2);
		hmap.put("depth3", depth3);
		hmap.put("sample_file_name",filename);
		service.insertName(hmap);
		
		
		//excel 파일 안 데이터들 적재
		
		
		//우선 여기서 파일 확장자 xlsx / xls 구분해서 해당하는 service로 전달
		String extension = FilenameUtils.getExtension(file.getOriginalFilename());
		
		if(extension.equals("xlsx")) {
			System.out.println("xlsx확장자//////////////");
			service.insertExcelDatexlsx2(file,filename);
			
		}else if(extension.equals("xls")) {
			System.out.println("xls확장자//////////////");
			service.insertExcelDatexlsx2(file,filename);
		}
			
		
		
		
		
	return "redirect:/";
	}
	
	

}
