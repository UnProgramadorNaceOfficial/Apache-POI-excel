package com.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatosExcel {

	public static void main(String[] args) {

		// 1) Crear el libro
		XSSFWorkbook libro = new XSSFWorkbook();

		// 2) Crear la hoja
		XSSFSheet hoja = libro.createSheet();

		// 3) Crear filas
		XSSFRow fila = hoja.createRow(3);

		// 4) Crear celdas
		XSSFCell celda = fila.createCell(3);
		XSSFCellStyle estiloCelda = libro.createCellStyle();

		/* Configuracion de estilos */
		estiloCelda.setDataFormat(libro.createDataFormat().getFormat("h:mm:ss"));

		/* Configuracion de celda */
		celda.setCellValue(LocalDateTime.now());
		celda.setCellStyle(estiloCelda);

		/* Configuracion Hoja */

		try {
			OutputStream output = new FileOutputStream("FormatosExcel.xlsx");
			libro.write(output);

			libro.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
