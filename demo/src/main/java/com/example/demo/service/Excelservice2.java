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

	//cell_data  데이터들의 seq 번호
	int seq_num ;
	
	//cell_sample_type에 seq에 들어갈 번호
	int count_seq_num;
	
	//cell_sample_type 에 들어갈 데이터들
	HashMap<String, Object> hmap =new HashMap<>();
	
	
	public List<Excelvo> excelselect() throws Exception {
		return dao.excelselect();
	}

	// Excellcontroller에 저장된 hmap 을 excelservice2에서 사용하기 위해 가져옴
	public void insertName(HashMap<String, Object> hmap) throws Exception {
		this.hmap =hmap;
	}
	
	
	

	// insertExcelDatexlsx : xlsx확장자의 cell_data 적재
	 public void insertExcelDatexlsx2(MultipartFile file, String filename) throws Exception {
		 
		 //cell_sample_type의 마지막 seq 값 가지고와서 + 1 해줌으로서 현재 excel의 seq 값 정해주기
		 int lastseqnum = dao.lastseqnum()+1;
		 
		 //datalist 에 들어갈 seq 
		 seq_num=lastseqnum;
		 
		 //cell_sample_type 에 들어갈 seq
		 count_seq_num=lastseqnum;
		 
		 //sheet가 이어질때마다 이어지는 첫번째 row 삭제
		 boolean cut_row = false;
		 
		 
		 
		 
	      
//	      int seq =dao.selectseq(filename);
	      XSSFWorkbook workbook = null;
	      
	      OPCPackage opc = OPCPackage.open(file.getInputStream());
	      XSSFReader xssfReader = new XSSFReader(opc);
	      
	      XSSFReader.SheetIterator itr =(XSSFReader.SheetIterator) xssfReader.getSheetsData();

	      StylesTable styles = xssfReader.getStylesTable();
	      
	      ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opc);
	      
	      List<celldatavo> dataList = new ArrayList<>();
	      
	      //arraylist 2번
	      List<List> dataList2 = new ArrayList<>();
	      
	      
	      //sheet 반복 횟수
	      int count =1;
	      
	      
	      //각 sheet의 연관성  true(하나의 sheet를 여러개로 나누어 excel을 만든경우)/ false(각각의 sheet들은 연관성이 없다)
	      boolean sheetsame=false;
	      //이전 SHEET 이름
	      String SheetPrevName="";
	      //SHEET 비교할때
	      String PN="";
	      //현재 SHEET 이름
    	  String SheetNowName="";
    	  //SHEET 비교할때
    	  String NN="";
	      while(itr.hasNext()) {
	    	 // System.out.println("현재 sheet : "+count);
	    	  
	    	  
	    	  
	    	  InputStream sheetStream = itr.next();
	    	  InputSource sheetSource = new InputSource(sheetStream);
	    	  
	    	  Sheet2ListHandler sheet2ListHandler = null;
	    	  
	    	  
	    	  //sheet 이름으로 시트들이 각각 따로따로 구성되었는지..
	    	  //같은 구성인데 양이 많아 따로따로 sheet 구분한건지 확인하는 작업 진행.
	    	  if(count ==1) {
	    		  SheetPrevName=itr.getSheetName();
	    		  if(!SheetPrevName.equals("Info") && !SheetPrevName.equals("Channel_Chart") && !(SheetPrevName.substring(0, SheetPrevName.indexOf("_"))).equals("Statistics")) {
//	    			  System.out.println("Info가 아니다");
//	    			  System.out.println("count : " + count);
	    			  sheet2ListHandler=new Sheet2ListHandler(dataList,seq_num,false,0,cut_row);
	    			  sheetsame=true;
	    			  //info가 아니니까 1번째 sheet 값도 넣어야한다...
	    		  }else {
//	    			  System.out.println("Info다");
	    			  SheetPrevName="";
	    			  count--;
	    		  }
	    				  
	    	  }else {
	    		  SheetNowName=itr.getSheetName();
	    		  
//	    		  System.out.println("count : " + count);
//	    		  System.out.println("이전sheet : "+SheetPrevName + "현재 sheet :  "+SheetNowName);
	    		  //Sheet 이름이  Sheet  <-> Channel_1-005_1 구분
	    		  
	    		  if(SheetNowName.charAt(0)=='C') {
	    			  //Sheet 명이 C로 시작할 때
//	    			  System.out.println("확인 작업");
	    			  
	    			  //나눠져있는 sheet인지 아닌디 -구분
	    			  if(SheetPrevName.lastIndexOf("_")==7) {
	    				  PN=SheetPrevName.substring(SheetPrevName.indexOf("-")+1);
//	    				  System.out.println("PN 단독 sheet : "+ PN);
	    			  }else if(SheetPrevName.equals("")) {
	    				  PN="";
	    			  }
	    			  else {
	    				  PN=SheetPrevName.substring(SheetPrevName.indexOf("-")+1,SheetPrevName.lastIndexOf("_"));
//	    				  System.out.println("PN 여러개 sheet : "+ PN);
	    			  }
	    			  
	    			  if(SheetNowName.lastIndexOf("_")==7) {
	    				  NN=SheetNowName.substring(SheetNowName.indexOf("-")+1);
//	    				  System.out.println("NN 단독 sheet : "+ NN);
	    			  }else {
	    				  NN=SheetNowName.substring(SheetNowName.indexOf("-")+1,SheetNowName.lastIndexOf("_"));
//	    				  System.out.println("NN 여러개 sheet : "+ NN);
	    			  }
	    			  
	    			  
	    			  //같은지 비교
	    			  if(PN.equals(NN)) {
	    				  //같다
	    				  //하나의 arraylist 에 데이터를 쌓는다.
//	    				  System.out.println("PN 과 NN 이 같다");
	    				  cut_row=true;
	    				  sheet2ListHandler=new Sheet2ListHandler(dataList,seq_num,false,0,cut_row);
	    				  sheetsame=true;
	    			  }else {
	    				  //다르다
	    				  //또 다른 arraylist를 만들어 데이터를 쌓는다
//	    				  System.out.println("PN 과 NN 이 다르다");
	    				  cut_row=false;
	    				  seq_num++;
	    				  if(count ==2) {
//	    					  System.out.println("1번째 sheet");
	    					  dataList2.add(dataList);
	    					  sheetsame=false;
	    				  }
	    				  dataList= new ArrayList<celldatavo>();
	    				  sheet2ListHandler=new Sheet2ListHandler(dataList,seq_num,false,0,cut_row);
	    				  dataList2.add(dataList);
	    				  sheetsame=false;
	    			  }
	    			  
	    			  
	    		  }else {
	    			  //Sheet 명이 'C' 가아닌것들 (ex. 'S'로 시작)
	    			  //원래 하던 대로 한번의 데이터를 적재하면 될듯싶습니다.
//	    			  System.out.println("sheet 명이 s로 시작합니다.");
//	    			  System.out.println("count : " + count);
	    			  if(count ==1) {
//    					  System.out.println("1번째 sheet");
//    					  System.out.println("count : " + count);
//    					  System.out.println("dataList 크기 : " +dataList.size());
    					  if(dataList.size()!=0) {
    						  dataList2.add(dataList);
    					  }
    				  }
	    			  cut_row=false;
	    			  sheet2ListHandler=new Sheet2ListHandler(dataList,seq_num,false,0,cut_row);
	    			  sheetsame=false;
	    			  if(dataList.size()!=0) {
//	    				  System.out.println("dataList 크기 : " +dataList.size());
						  dataList2.add(dataList);
					  }
	    		  }
	    		  
	    		  if(count!=1) {
	    			  SheetPrevName =SheetNowName;
	    			  SheetNowName="";
	    		  }
	    	  }//if~else  count ==1
	    	  
//	    	  System.out.println("count : " + count);
	    	  
	    	  
	    	  
	    	  
	    	  if(sheet2ListHandler!=null) {
//	    		  System.out.println("데이터 들어 갑니다~~");
	    		  ContentHandler handler = new XSSFSheetXMLHandler(styles, strings, sheet2ListHandler, false);
	    		  SAXParserFactory saxFactory = SAXParserFactory.newInstance();
	    		  SAXParser saxParser = saxFactory.newSAXParser();
	    		  
	    		  XMLReader sheetParser = saxParser.getXMLReader();
	    		  
	    		  sheetParser.setContentHandler(handler);
	    		  
	    		  
	    		  sheetParser.parse(sheetSource);
	    		  
	    		  sheetStream.close();
//	    		  System.out.println("sheet끝");
	    	  }
	    	  
	    	  
	    	  count++;
	      }
	      
	      //처음부터 끝까지 같은 sheet 일 경우!!!
	      if(sheetsame) {
	    	  dataList2.add(dataList);
    	  }
	      
	      opc.close();
	      
//	      System.out.println("dataList2 의 사이즈 : "+dataList2.size());
	      if(dataList2.size() !=0) {
	    	  for(int i=0;i<dataList2.size();i++) {
//	    		  System.out.println("count_seq_num : "+count_seq_num+",  i  : "+i);
					hmap.put("seq", count_seq_num++);
					dao.insertName(hmap);
					List<List<celldatavo>> ret = split(dataList2.get(i), 5000);
					for (int j = 0; j < ret.size(); j++) {
//						System.out.println("ret의 size : " + ret.size());
						dao.insertExcel(ret.get(j));
					}
	      }
	    	  
	      }else {
	    	  System.out.println("필요하지 않은 excel 입니다!");
	      }
	     
	      
	   }
	 //데이터를 count * n 인 n개 list 생성
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
