package com.example.demo.service;

import java.util.List;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import com.example.demo.vo.celldatavo;

public class Sheet2ListHandler implements SheetContentsHandler {
	
	private List<celldatavo> datalist;
	
	private celldatavo vo;
	
	
	private int seq;
	
	//현재 읽고 있는 cell의 col
	private int currColNum;
	
	private boolean rowback;
	
	private int currAllNum;
	
	private String firstcolumnname;
	
	private boolean cut_row;

	public Sheet2ListHandler(List<celldatavo> datalist, int seq,boolean rowback,int currAllNum,boolean cut_row) {
		// TODO Auto-generated constructor stub
		this.datalist= datalist;
		this.seq=seq;
		//새로운 sheet 열때 마다 rowback 값 false로 초기화
		this.rowback=rowback;
		this.currAllNum=currAllNum;
		this.firstcolumnname = "";
		this.cut_row= cut_row;
	}
	
	
	//ROW 시작 부분에서 발생하는 이벤트를 처리하는 method
	@Override
	public void startRow(int rowNum) {
		// TODO Auto-generated method stub
		//rowNum  첫행 비교
//		System.out.println("rowNum : "+ rowNum);
		if(rowback) {
			//System.out.println("원하지 않는 sheet 입니다////startRow");
			return;
		}else {
			if(rowNum ==0) {
				rowback=false;
			}else {
				this.vo = new celldatavo();
			}
		}
		currColNum=0;
		
		//System.out.println("행 시작");

	}

	//Row의 끝에서 발생하는 이벤트를 처리하는 method
	@Override
	public void endRow(int rowNum) {
		// TODO Auto-generated method stub
		//비교해서 추가 하는 부분\
		if(rowback) {
			//System.out.println("원하지 않는 sheet 입니다////endRow");
			return;
		}else {
			if(rowNum !=0) {
				vo.setType_seq(seq);
				// 데이터가 잘 들어갔는지 확인
				//System.out.println(vo.toString());
				if(rowNum==1) {
					if(cut_row) {
						
					}else {
						datalist.add(vo);
					}
				}else {
					datalist.add(vo);
				}
				
			}
			//System.out.println("행 끝");
		}
	}
	

	//cell 이벤트 발생시 해당 cell의 주소와 값을 받아옴.
	//입맛에 맞게 처리하면됨.
	@Override
	public void cell(String columnNum, String value, XSSFComment comment) {
		//System.out.println("cell 들어옴");
		// TODO Auto-generated method stub
		// 우선 columnNum 의 앞 문자 확인! A인지 확인 A이면 1인지 확인
		currColNum++;
		currAllNum++;
		char columncell = columnNum.charAt(0); //
		String columnlast = columnNum.substring(1);
		
		if (currAllNum == 1) {
			firstcolumnname = value;
		}
		//System.out.println("sheet의 첫 cell  값이 Data_Point가 아닐때 내가 찾는 EXCEL 인지 확인 작업");
		if (firstcolumnname.equals("Data_Point")) {
			//System.out.println("////////원하는 sheet////////");
			// excel에 첫번째 행은 column 명이므로 제외
			// 여기 다시 생각해보기*****************************************
			// 1. A1의 값이 NULL(빈 값)
			// 2. A1의 값이 Data_Point가 아니거나..
			if (Integer.parseInt(columnlast) != 1) {
				//System.out.println("switch문 들어감");
				switch (columncell) {
				case 'A': // datapoint
					vo.setData_point((int) Double.parseDouble(value));
					break;
				case 'B': // Test_Time
//			              vo.setTest_time(Math.round(Double.parseDouble(value)*1000000000)/1000000000.0);
					vo.setTest_time(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'C': // Date_Time
					// DATE 타입 변환 vo.setDate_time(tf);
					vo.setDate_time(value);
					break;
				case 'D': // Step_Time
					vo.setStep_time(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'E': // Step_Index
					vo.setStep_index((int) Double.parseDouble(value));
					break;
				case 'F': // Cycle_Index
					vo.setCycle_index((int) Double.parseDouble(value));
					break;
				case 'G':// Current(A)
					vo.setCurrent(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'H': // Voltage(V)
					vo.setVoltage(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'I': // Charge_Capacity(Ah)
					vo.setCharge_capacity(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'J': // Discharge_Capacity(Ah)
					vo.setDischarge_capacity(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'K': // Charge_Energy(Wh)
					vo.setCharge_energy(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'L': // Discharge_Energy(Wh)
					vo.setDischarge_energy(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'M': // dv/dt
					vo.setDv_dt(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'N': // Internal_Resistance(Ohm)
					vo.setInternal_resistance(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'O': // Is_FC_Data
					vo.setIs_fc_data(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'P': // AC_Impedance(Ohm)
					vo.setAc_impedance(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'Q': // ACI_Phase_Angel(Deg)
					vo.setAci_phase_angle(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'R':
					vo.setTemperature_1(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				case 'S':
					vo.setTemperature_2(Double.parseDouble(String.format("%20.9f", Double.parseDouble(value))));
					break;
				default:
					break;
				}
				//System.out.println("switch문 나감");
			}
		} else {
			//System.out.println("XXXXXXXX원하지 않는 sheetXXXXXXXX");
			rowback = true;
			return;
		}
		//System.out.println("cell 나감");
	}


	@Override
	public void headerFooter(String text, boolean isHeader, String tagName) {
		// TODO Auto-generated method stub
		//사용하지 않습니다.....(사실 사용할줄 모릅니다...ㅠㅠ); 아마도 excel에 header footer 해주는 작업 되어있는 excel 만 가능한 것 같습니다...
		
	}
}




//빈 sheet 일경우 cell 값이 아예없어서 실행하지도 않음
