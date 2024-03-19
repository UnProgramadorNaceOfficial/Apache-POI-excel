package com.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FuentesExcel {

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
		
		XSSFFont fuente = libro.createFont();
		fuente.setFontName("Franklin Gothic Book");
		fuente.setBold(true);
		fuente.setItalic(true);
		fuente.setFontHeightInPoints((short) 14);
		fuente.setColor(IndexedColors.RED.getIndex());
		fuente.setUnderline(FontUnderline.SINGLE);

		/* Configuracion de estilos */
		estiloCelda.setFont(fuente);
		estiloCelda.setAlignment(HorizontalAlignment.CENTER);
		estiloCelda.setVerticalAlignment(VerticalAlignment.CENTER);

		/* Configuracion de celda */
		celda.setCellValue("Apache POI - Fuentes");
		celda.setCellStyle(estiloCelda);

		/* Configuracion Hoja */
		hoja.autoSizeColumn(1);

		try {
			OutputStream output = new FileOutputStream("FuentesExcel.xlsx");
			libro.write(output);

			libro.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
