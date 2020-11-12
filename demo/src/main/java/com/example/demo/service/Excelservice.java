package com.example.demo.service;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.dao.Exceldao;
import com.example.demo.vo.Excelvo;
import com.example.demo.vo.celldatavo;

//@Service
public class Excelservice {

	@Autowired
	Exceldao dao;

	public List<Excelvo> excelselect() throws Exception {
		return dao.excelselect();
	}

	// cell_sample_type 적재
	public void insertName(HashMap<String, Object> hmap) throws Exception {
		dao.insertName(hmap);
	}
	
	
	String [] columns= {"Data_Point","Test_Time(s)","Date_Time","Step_Time(s)","Step_Index",
			"Cycle_Index","Current(A)","Voltage(V)","Charge_Capacity(Ah)","Discharge_Capacity(Ah)",
			"Charge_Energy(Wh)","Discharge_Energy(Wh)","dV/dt(V/s)",
			"Internal_Resistance(Ohm)","Is_FC_Data","AC_Impedance(Ohm)",
			"ACI_Phase_Angle(Deg)","temperature_1","temperature_2"};

	// insertExcelDatexlsx : xlsx확장자의 cell_data 적재
	 public void insertExcelDatexlsx(MultipartFile file, String filename) throws Exception {

	      
	      int seq =dao.selectseq(filename);
	      XSSFWorkbook workbook = null;
	      
	      
	      try {
	         workbook = new XSSFWorkbook(file.getInputStream());
	         
	        
	         
	         XSSFSheet curSheet;
	         XSSFRow curRow;
	         XSSFCell curCell;
	         celldatavo vo=null;

	         // Sheet 탐색 for 문
	         for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
	            // 현재 Sheet 반환
	            curSheet = workbook.getSheetAt(sheetIndex);

	            // row의 첫번째 cell 값이 원하는 값이 아닐경우 sheet pass
	            List<celldatavo> datalist = new ArrayList<celldatavo>();
	            
	            boolean columnname =false;
	            // row 탐색 for 문
	            
	            int rowIndex;
	            for (rowIndex = 0; rowIndex < curSheet.getPhysicalNumberOfRows(); rowIndex++) {
	            	// 반활할 객체를 생성
	      	     
	               if (rowIndex != 0) {
	                  // 현재 row 반환
	                  curRow = curSheet.getRow(rowIndex);

	                  vo = new celldatavo();
	                  String value;
	                 
	                  
	                  //원하는 row 값이 없으면 다음 sheet 
	                  if(columnname==false) {
	                     break;
	                  }
	                  for (int cellIndex = 0; cellIndex < curRow.getPhysicalNumberOfCells(); cellIndex++) {
	                     
	                     curCell = curRow.getCell(cellIndex);
	                     if (true) {
	                        value = "";
	                        // cell 스타일이 다르더라도 String으로 반환 받음
	                        switch (curCell.getCellType()) {
	                        case Cell.CELL_TYPE_FORMULA:
	                           value = curCell.getCellFormula();
	                           break;
	                        case HSSFCell.CELL_TYPE_NUMERIC:
	                           if (DateUtil.isCellDateFormatted(curCell)) {
	                              Date date = curCell.getDateCellValue();
	                              value = new SimpleDateFormat("MM-dd-yyyy HH:mm:ss").format(date);
	                           } else
	                              value = String.valueOf(curCell.getNumericCellValue());
	                           break;
	                        case HSSFCell.CELL_TYPE_STRING:
	                           value = curCell.getStringCellValue() + "";
	                           break;
	                        case HSSFCell.CELL_TYPE_BLANK:
	                           value = curCell.getBooleanCellValue() + "";
	                           break;
	                        case HSSFCell.CELL_TYPE_ERROR:
	                           value = curCell.getErrorCellValue() + "";
	                           break;
	                        default:
	                           value = new String();
	                           break;
	                        }
	                        switch (cellIndex) {
	                        case 0: // datapoint
	                           vo.setData_point((int) Double.parseDouble(value));
	                           break;
	                        case 1: // Test_Time
//	                              vo.setTest_time(Math.round(Double.parseDouble(value)*1000000000)/1000000000.0);
	                           vo.setTest_time(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 2: // Date_Time
	//   DATE 타입 변환                           vo.setDate_time(tf);
	                           vo.setDate_time(value);
	                           break;
	                        case 3: // Step_Time
	                           vo.setStep_time(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 4: // Step_Index
	                           vo.setStep_index((int) Double.parseDouble(value));
	                           break;
	                        case 5: // Cycle_Index
	                           vo.setCycle_index((int) Double.parseDouble(value));
	                           break;
	                        case 6:// Current(A)
	                           vo.setCurrent(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 7: // Voltage(V)
	                           vo.setVoltage(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 8: // Charge_Capacity(Ah)
	                           vo.setCharge_capacity(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 9: // Discharge_Capacity(Ah)
	                           vo.setDischarge_capacity(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 10: // Charge_Energy(Wh)
	                           vo.setCharge_energy(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 11: // Discharge_Energy(Wh)
	                           vo.setDischarge_energy(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 12: // dv/dt
	                           vo.setDv_dt(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 13: // Internal_Resistance(Ohm)
	                           vo.setInternal_resistance(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 14: // Is_FC_Data
	                           vo.setIs_fc_data(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 15: // AC_Impedance(Ohm)
	                           vo.setAc_impedance(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        case 16: // ACI_Phase_Angel(Deg)
	                           vo.setAci_phase_angle(
	                                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	                           break;
	                        default:
	                           break;
	                        }
	                     }
	                  }//cell-for구문
	                        vo.setType_seq(seq);
	                        datalist.add(vo);
	                   
	               }else {
	            	 //colunm 검사 
		                  for (int cellIndex = 0; cellIndex < curSheet.getRow(rowIndex).getPhysicalNumberOfCells(); cellIndex++) {
		                     String column;
		                     try {
		                        column = curSheet.getRow(0).getCell(cellIndex).getStringCellValue();
		                     }catch(NullPointerException e) {
		                        break;
		                     }
		                     for (int i = 0; i < columns.length; i++) {
		                        if (column.equals(columns[i])) {
		                           columnname=true;
		                        }else {
		                        }
		                     }

		                  }//colunm 검사
	               }
	               
	             //n개씩 데이터 적재 (n = 내가 설정한 개수 만큼 )
		            if((rowIndex!=0 && rowIndex % 30000 == 0 ) || (rowIndex ==curSheet.getPhysicalNumberOfRows()-1)) {
		            	System.out.println("datalist 크기 : "+datalist.size());
		            	dao.insertExcel(datalist);
		            	datalist.clear();
					} // row - for 구문
				}

			} //sheet-for구문
	      } catch (FileNotFoundException e) {
	         // TODO Auto-generated catch block
	         e.printStackTrace();
	      } catch (IOException e) {
	         // TODO Auto-generated catch block
	         e.printStackTrace();

	      } finally {
	         try {
	            // 사용한 자원은 finally에서 해제
	            if (workbook != null)
	               workbook.close();
	         } catch (IOException e) {
	            // TODO Auto-generated catch block
	            e.printStackTrace();
	         }
	      }

	   }

	// insertExcelDate : xls 확장자의 cell_data 적재
	/*
	 public void insertExcelDatexls(MultipartFile file,String filename) throws Exception {

			// 반활할 객체를 생성
			List<celldatavo> datalistxls = new ArrayList<celldatavo>();

			HSSFWorkbook workbook = null;
			try {
				workbook =  new HSSFWorkbook(file.getInputStream());
				
				HSSFSheet curSheet;
		        HSSFRow   curRow;
		        HSSFCell  curCell;
				celldatavo vo;

				// Sheet 탐색 for 문
				for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
					// 현재 Sheet 반환
					curSheet = workbook.getSheetAt(sheetIndex);

					System.out.println("현재 시트 : " + curSheet);

					// row 탐색 for 문
					for (int rowIndex = 0; rowIndex < curSheet.getPhysicalNumberOfRows(); rowIndex++) {
						if (rowIndex != 0) {
							// 현재 row 반환
							curRow = curSheet.getRow(rowIndex);

							vo = new celldatavo();
							String value;
							// row의 첫번째 cell 값이 비어있지 않은 경우 만 cell 탐색
							
							if (curRow.getCell(0).getStringCellValue()!=null) {

								// cell 탐색 for 문
								for (int cellIndex = 0; cellIndex < curRow.getPhysicalNumberOfCells(); cellIndex++) {
									curCell = curRow.getCell(cellIndex);

									if (true) {
										value = "";
										// cell 스타일이 다르더라도 String으로 반환 받음
										switch (curCell.getCellType()) {
										case HSSFCell.CELL_TYPE_FORMULA:
											value = curCell.getCellFormula();
											break;
										case HSSFCell.CELL_TYPE_NUMERIC:
											value = curCell.getNumericCellValue() + "";
											break;
										case HSSFCell.CELL_TYPE_STRING:
											value = curCell.getStringCellValue() + "";
											break;
										case HSSFCell.CELL_TYPE_BLANK:
											value = curCell.getBooleanCellValue() + "";
											break;
										case HSSFCell.CELL_TYPE_ERROR:
											value = curCell.getErrorCellValue() + "";
											break;
										default:
											value = new String();
											break;
										}

										switch (cellIndex) {
										case 0: // datapoint
											vo.setData_point((int)Double.parseDouble(value));
											break;
										case 1: // Test_Time
											vo.setTest_time(Double.parseDouble(value));
											break;
										case 2: // Date_Time
//		DATE 타입 변환									vo.setDate_time(tf);
											break;
										case 3: // Step_Time
											vo.setStep_time(Double.parseDouble(value));
											break;
										case 4: // Step_Index
											vo.setStep_index((int)Double.parseDouble(value));
											break;
										case 5: // Cycle_Index
											vo.setCycle_index((int)Double.parseDouble(value));
											break;
										case 6:// Current(A)
											vo.setCurrent(Double.parseDouble(value));
											break;
										case 7: // Voltage(V)
											vo.setVoltage(Double.parseDouble(value));
											break;
										case 8: // Charge_Capacity(Ah)
											vo.setCharge_capacity(Double.parseDouble(value));
											break;
										case 9: // Discharge_Capacity(Ah)
											vo.setDischarge_capacity(Double.parseDouble(value));
											break;
										case 10: // Charge_Energy(Wh)
											vo.setCharge_energy(Double.parseDouble(value));
											break;
										case 11: // Discharge_Energy(Wh)
											vo.setDischarge_energy(Double.parseDouble(value));
											break;
										case 12: // dv/dt
											vo.setDv_dt(Double.parseDouble(value));
											break;
										case 13: // Internal_Resistance(Ohm)
											vo.setInternal_resistance(Double.parseDouble(value));
											break;
										case 14: // Is_FC_Data
											vo.setIs_fc_data(Double.parseDouble(value));
											break;
										case 15: // AC_Impedance(Ohm)
											 vo.setAc_impedance(Double.parseDouble(value));
											break;
										case 16: // ACI_Phase_Angel(Deg)
											vo.setAci_phase_angle(Double.parseDouble(value));
											break;
										case 17: // temperature_1
											vo.setTemperature_1(Double.parseDouble(value));
											break;
										case 18: // temperature_2
											vo.setTemperature_2(Double.parseDouble(value));
											break;
										default:
											break;
										}
									}
								}
								vo.setType_seq(dao.selectseq(filename));
								datalistxls.add(vo);
							}else {
								break;
							}
						}
					}
					}
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();

			} finally {
				try {
					// 사용한 자원은 finally에서 해제
					if (workbook != null)
						workbook.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			 dao.insertExcel(datalistxls);
		}
*/
}
