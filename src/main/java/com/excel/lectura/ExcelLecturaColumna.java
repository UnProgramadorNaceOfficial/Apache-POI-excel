package com.excel.lectura;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.simple.SimpleLoggerContextFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLecturaColumna {
	

	public static void main(String[] args) {
		
		File archivo = new File("Datos.xlsx");
		
		try {
			
			InputStream input = new FileInputStream(archivo);
			
			XSSFWorkbook libro = new XSSFWorkbook(input);

			XSSFSheet hoja = libro.getSheetAt(0);
			
			Iterator<Row> filas = hoja.rowIterator();
			
			Cell columna = null;
			
			while(filas.hasNext()) {
				columna = filas.next().getCell(0);
				
				System.out.println(columna.getStringCellValue());
			}
			
			input.close();
			libro.close();
					
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
