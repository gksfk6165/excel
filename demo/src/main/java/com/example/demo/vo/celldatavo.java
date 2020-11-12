package com.example.demo.vo;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class celldatavo {

	//cell_data에 들어갈 데이터들
			
			@JsonProperty(value="type_seq")
			private int type_seq;
			@JsonProperty(value="Data_Point")
			private int data_point;
			@JsonProperty(value="Test_Time(s)")
			private double test_time;
			@JsonProperty(value="Date_Time")
			private String date_time;
			@JsonProperty(value="Step_Time(s)")
			private double step_time;
			@JsonProperty(value="Step_Index")
			private int step_index;
			@JsonProperty(value="Cycle_Index")
			private int cycle_index;
			@JsonProperty(value="Current(A)")
			private double current;
			@JsonProperty(value="Voltage(V)")
			private double voltage;
			@JsonProperty(value="Charge_Capacity(Ah)")
			private double charge_capacity; 
			@JsonProperty(value="Discharge_Capacity(Ah)")
			private double discharge_capacity; 
			@JsonProperty(value="Charge_Energy(Wh)")
			private double charge_energy;
			@JsonProperty(value="Discharge_Energy(Wh)")
			private double discharge_energy;
			@JsonProperty(value="dV/dt(V/s)")
			private double dv_dt;
			@JsonProperty(value="Internal_Resistance(Ohm)")
			private double internal_resistance;
			@JsonProperty(value="Is_FC_Data")
			private double is_fc_data;
			@JsonProperty(value="AC_Impedance(Ohm)")
			private double ac_impedance;
			@JsonProperty(value="ACI_Phase_Angle(Deg)")
			private double aci_phase_angle;
			@JsonProperty(value = "Temperature (C)_1")
			private double temperature_1;
			@JsonProperty(value = "Temperature (C)_2")
			private double temperature_2;
			
			
			
			@Override
			public String toString() {
				StringBuffer sb = new StringBuffer();
				sb.append("type_seq : " +type_seq);
				sb.append(" .Date_point :" +data_point);
				sb.append(" ,Test_Time :" +test_time);
				sb.append(" ,Date_Time :" +date_time);
				sb.append(" ,Step_Time :" +step_time);
				sb.append(" ,Step_Index :" +step_index);
				sb.append(" ,Cycle_Index :" +cycle_index);
				sb.append(" ,Current(A) :" +current);
				sb.append(" ,Voltage(V) :" +voltage);
				sb.append(" ,Charge_Capacity(Ah) :" +charge_capacity);
				sb.append(" ,Discharge_Capacity(Ah) :" +discharge_capacity);
				sb.append(" ,Charge_Energy(Wh) :" +charge_energy);
				sb.append(" ,Discharge_Energy(Wh) :" +discharge_energy);
				sb.append(" ,dv/dt : " +dv_dt);
				sb.append(" ,Internal_Resistance(Ohm) :" +internal_resistance);
				sb.append(" ,Is_FC_Data :" +is_fc_data);
				sb.append(" ,AC_Impedance(Ohm) :" +ac_impedance);
				sb.append(" ,ACI_Phase_Angel(Deg) :" +aci_phase_angle);
				sb.append(" ,temperature_1 :" +temperature_1);
				sb.append(" ,temperature_2 :" +temperature_2);
				return sb.toString();
			}
}
