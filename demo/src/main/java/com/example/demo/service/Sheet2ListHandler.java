package com.example.demo.service;

import java.util.List;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import com.example.demo.vo.celldatavo;

public class Sheet2ListHandler implements SheetContentsHandler {
	
	private List<celldatavo> datalist;
	
	private celldatavo vo;
	
	private int columnCnt;
	
	private int seq;
	
	//현재 읽고 있는 cell의 col
	private int currColNum;
	
	private boolean rowback=false;
	

	public Sheet2ListHandler(List<celldatavo> datalist,int columnsCnt,int seq) {
		// TODO Auto-generated constructor stub
		this.datalist= datalist;
		this.columnCnt = columnsCnt;
		this.seq=seq;
	}
	
	
	//ROW 시작 부분에서 발생하는 이벤트를 처리하는 method
	@Override
	public void startRow(int rowNum) {
		// TODO Auto-generated method stub
		//rowNum  첫행 비교
		System.out.println("rowNum : "+ rowNum);
		if(rowback) {
			return;
		}else {
			if(rowNum ==0) {
				rowback=false;
			}else {
				this.vo = new celldatavo();
			}
		}
		currColNum=0;
		System.out.println("행 시작");

	}

	//Row의 끝에서 발생하는 이벤트를 처리하는 method
	@Override
	public void endRow(int rowNum) {
		// TODO Auto-generated method stub
		//비교해서 추가 하는 부분\
		if(rowNum !=0) {
			vo.setType_seq(seq);
			datalist.add(vo);
		}
		System.out.println("행 끝");
		
	}
	

	//cell 이벤트 발생시 해당 cell의 주소와 값을 받아옴.
	//입맛에 맞게 처리하면됨.
	@Override
	public void cell(String columnNum, String value, XSSFComment comment) {
		System.out.println("cell 들어옴");
		// TODO Auto-generated method stub
		//우선 columnNum 의 앞 문자 확인! A인지 확인 A이면 1인지 확인
		currColNum++;
		char columncell = columnNum.charAt(0); //
		String columnlast = columnNum.substring(1);
		String A1="";
		System.out.println("currColNum : "+ currColNum);
		if(currColNum ==1) {
			A1=value;
		}
		
//		System.out.println("columncell : "+ columncell +", columnlast : " + columnlast);
		if(Integer.parseInt(columnlast)==1) {
			System.out.println("columnlast : "+ columnlast);
			//excel에 첫번째 행은 column 명이므로 제외
			//여기 다시 생각해보기
			if((value==null) || (!A1.equals("Data_Point"))) {
				System.out.println("1차검증 :  value : "+ value +"  , A1 : "+ A1);
				System.out.println("1 : "+(value==null));
				System.out.println("2 : "+(!A1.equals("Data_Point")));
				return;
			}
		}
		else {
			System.out.println("switch문 들어감");
			switch (columncell) {
	        case 'A': // datapoint
	           vo.setData_point((int) Double.parseDouble(value));
	           break;
	        case 'B': // Test_Time
//	              vo.setTest_time(Math.round(Double.parseDouble(value)*1000000000)/1000000000.0);
	           vo.setTest_time(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'C': // Date_Time
	//   DATE 타입 변환                           vo.setDate_time(tf);
	           vo.setDate_time(value);
	           break;
	        case 'D': // Step_Time
	           vo.setStep_time(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'E': // Step_Index
	           vo.setStep_index((int) Double.parseDouble(value));
	           break;
	        case 'F': // Cycle_Index
	           vo.setCycle_index((int) Double.parseDouble(value));
	           break;
	        case 'G':// Current(A)
	           vo.setCurrent(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'H': // Voltage(V)
	           vo.setVoltage(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'I': // Charge_Capacity(Ah)
	           vo.setCharge_capacity(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'J': // Discharge_Capacity(Ah)
	           vo.setDischarge_capacity(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'K': // Charge_Energy(Wh)
	           vo.setCharge_energy(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'L': // Discharge_Energy(Wh)
	           vo.setDischarge_energy(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'M': // dv/dt
	           vo.setDv_dt(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'N': // Internal_Resistance(Ohm)
	           vo.setInternal_resistance(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'O': // Is_FC_Data
	           vo.setIs_fc_data(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'P': // AC_Impedance(Ohm)
	           vo.setAc_impedance(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	           break;
	        case 'Q': // ACI_Phase_Angel(Deg)
	           vo.setAci_phase_angle(
	                 Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
	         
	           break;
	        default:
	           break;
	        }
			System.out.println("switch문 나감");

		
		}
		
		System.out.println("cell 나감");
	}

	@Override
	public void headerFooter(String paramString1, boolean paramBoolean, String paramString2) {
		// TODO Auto-generated method stub
		//우선 columnNum 의 앞 문자 확인! A인지 확인 A이면 1인지 확인
	}



}
