package com.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RangosExcel {

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
		CellRangeAddress rango = new CellRangeAddress(2, 7, 2, 5);

		/* Configuracion de estilos */
		estiloCelda.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		/* Configuracion de celda */
		celda.setCellValue("Prueba de rangos");

		/* Configuracion Hoja */
		hoja.addMergedRegion(rango);

		try {
			OutputStream output = new FileOutputStream("RangosExcel.xlsx");
			libro.write(output);

			libro.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
