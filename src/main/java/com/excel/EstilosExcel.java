package com.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EstilosExcel {

	public static void main(String[] args) {
		
		// 1) Crear el libro
		XSSFWorkbook libro = new XSSFWorkbook();
		
		// 2) Crear la hoja
		XSSFSheet hoja = libro.createSheet();
		
		// 3) Crear filas
		XSSFRow fila = hoja.createRow(1);
		
		// 4) Crear celdas
		XSSFCell celda = fila.createCell(1);
		XSSFCellStyle estiloCelda = libro.createCellStyle();
				
		/* Configuracion de estilos */
		estiloCelda.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estiloCelda.setBorderBottom(BorderStyle.THIN);
		estiloCelda.setBorderTop(BorderStyle.THIN);
		estiloCelda.setBorderLeft(BorderStyle.THIN);
		estiloCelda.setBorderRight(BorderStyle.THIN);
		
		
		/* Configuracion de celda */
		celda.setCellValue("Estilos con apache POI");
		celda.setCellStyle(estiloCelda);
		
		
		/* Configuracion Hoja */
		hoja.autoSizeColumn(1);
		
		
		try {
			OutputStream output = new FileOutputStream("EstilosExcel.xlsx");
			libro.write(output);
			
			libro.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
