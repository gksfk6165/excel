package com.example.demo.dao;

import java.util.HashMap;
import java.util.List;

import org.apache.ibatis.session.SqlSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import com.example.demo.vo.Excelvo;
import com.example.demo.vo.celldatavo;

@Repository
public class Exceldao {

	@Autowired
	SqlSession sql;
	
	
	public List<Excelvo> excelselect() {
		return sql.selectList("mappers.excelMapper.excelselect");
	}
	
	
	
	public void insertName(HashMap<String, String> hmap) {
	sql.insert("mappers.excelMapper.excelinsert",hmap); 
	}
	
	
	public void insertExcel(List<celldatavo> list){
		sql.insert("mappers.excelMapper.excelinsertcelldate",list) ;
	}
	
	public int selectseq(String filename) {
		return sql.selectOne("mappers.excelMapper.selectseq",filename);
	}
	
	 
}

