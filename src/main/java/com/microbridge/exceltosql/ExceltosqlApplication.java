package com.microbridge.exceltosql;

import com.microbridge.exceltosql.util.ExcelToSQL;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExceltosqlApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExceltosqlApplication.class, args);

		ExcelToSQL excelToSQL = new ExcelToSQL();
		excelToSQL.writeExcelToSQL("D:\\ExcelFiles\\bece_2023_schools_data.xlsx");
		System.out.println((excelToSQL.composeQuery("centre", excelToSQL.tableColumns())).toString());

	}

}
