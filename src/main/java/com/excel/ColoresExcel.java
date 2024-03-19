package com.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ColoresExcel {

	public static void main(String[] args) {
		
		/* Colores */
		XSSFColor verdeClaro = crearColor("62F744");

		// 1) Crear el libro
		XSSFWorkbook libro = new XSSFWorkbook();

		// 2) Crear la hoja
		XSSFSheet hoja = libro.createSheet();

		// 3) Crear filas
		XSSFRow fila = hoja.createRow(1);

		// 4) Crear celdas
		XSSFCell celda = fila.createCell(1);
		XSSFCellStyle estiloCelda = libro.createCellStyle();
		
		XSSFCell celda2 = fila.createCell(2);
		XSSFCellStyle estilosCelda2 = libro.createCellStyle();

		/* Configuracion de estilos */
		estiloCelda.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		estilosCelda2.setFillForegroundColor(verdeClaro);
		estilosCelda2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		/* Configuracion de celda */
		celda.setCellValue("Color Predeterminado");
		celda.setCellStyle(estiloCelda);
		
		celda2.setCellValue("Color Personalizado");
		celda2.setCellStyle(estilosCelda2);

		/* Configuracion Hoja */
		hoja.autoSizeColumn(1);
		hoja.autoSizeColumn(2);

		try {
			OutputStream output = new FileOutputStream("ColoresExcel.xlsx");
			libro.write(output);

			libro.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static XSSFColor crearColor(String colorHexadecimal) {
		try {
			byte[] rgb = Hex.decodeHex(colorHexadecimal);
			
			return new XSSFColor(rgb);
		} catch (DecoderException e) {
			e.printStackTrace();
			throw new RuntimeException("Error al crear el color.");
		}
		
	}
}
