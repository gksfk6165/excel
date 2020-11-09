package com.example.demo.service;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import com.example.demo.dao.Exceldao;
import com.example.demo.vo.Excelvo;
import com.example.demo.vo.celldatavo;

@Service
public class Excelservice2 {

	@Autowired
	Exceldao dao;

	public List<Excelvo> excelselect() throws Exception {
		return dao.excelselect();
	}

	// cell_sample_type 적재
	public void insertName(HashMap<String, String> hmap) throws Exception {
		dao.insertName(hmap);
	}
	
	
	

	// insertExcelDatexlsx : xlsx확장자의 cell_data 적재
	 public void insertExcelDatexlsx2(MultipartFile file, String filename) throws Exception {

	      
	      int seq =dao.selectseq(filename);
	      XSSFWorkbook workbook = null;
	      
	      OPCPackage opc = OPCPackage.open(file.getInputStream());
	      XSSFReader xssfReader = new XSSFReader(opc);
	      
	      XSSFReader.SheetIterator itr =(XSSFReader.SheetIterator) xssfReader.getSheetsData();
	      
	      StylesTable styles = xssfReader.getStylesTable();
	      
	      ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opc);
	      
	      List<celldatavo> dataList = new ArrayList<>();
	      
	      int count =0;
	      while(itr.hasNext()) {
	    	  System.out.println("sheet : "+ ((count++)+1));
	    	  InputStream sheetStream = itr.next();
	    	  InputSource sheetSource = new InputSource(sheetStream);
	    	  
	    	  Sheet2ListHandler sheet2ListHandler = new Sheet2ListHandler(dataList,17,seq);
	    	  
	    	  ContentHandler handler = new XSSFSheetXMLHandler(styles, strings, sheet2ListHandler, false);
	    	  
	    	  
	    	  SAXParserFactory saxFactory = SAXParserFactory.newInstance();
	    	  SAXParser saxParser = saxFactory.newSAXParser();
	    	  
	    	  XMLReader sheetParser = saxParser.getXMLReader();
	    	  
	    	  sheetParser.setContentHandler(handler);
	    	  
	    	  
	    	  sheetParser.parse(sheetSource);
	    	  
	    	  sheetStream.close();
	    	  System.out.println("sheet끝");
	    	  
	      }
	      opc.close();
	      
	      System.out.println("dataList 의 사이즈 : "+dataList.size());
	      
	      List<List<celldatavo>> ret = split(dataList, 5000);
	      for(int i=0;i<ret.size();i++) {
	    	  dao.insertExcel(ret.get(i));
	      }
	      
	   }
	 
	 public static <T> List<List<T>> split(List<T> resList, int count) {
	        if (resList == null || count <1)
	            return null;
	        List<List<T>> ret = new ArrayList<List<T>>();
	        int size = resList.size();
	        if (size <= count) {
	            // 데이터 부족 count 지정 크기
	            ret.add(resList);
	        } else {
	            int pre = size / count;
	            int last = size % count;
	            // 앞 pre 개 집합, 모든 크기 다 count 가지 요소
	            for (int i = 0; i <pre; i++) {
	                List<T> itemList = new ArrayList<T>();
	                for (int j = 0; j <count; j++) {
	                    itemList.add(resList.get(i * count + j));
	                }
	                ret.add(itemList);
	            }
	            // last 진행이 처리
	            if (last > 0) {
	                List<T> itemList = new ArrayList<T>();
	                for (int i = 0; i <last; i++) {
	                    itemList.add(resList.get(pre * count + i));
	                }
	                ret.add(itemList);
	            }
	        }
	        return ret;
	    }


}
