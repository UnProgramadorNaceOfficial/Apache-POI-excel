package com.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PrincipalExcel {

	public static void main(String[] args) {

		// 1) Crear un libro
		Workbook libro = new XSSFWorkbook();

		// 2) Crear las hojas
		Sheet hoja = libro.createSheet("Personas");

		// 3) Crear las filas
		Row cabecera = hoja.createRow(2);
		Row registro1 = hoja.createRow(3);
		Row registro2 = hoja.createRow(4);

		// 4) Crear las columnas
		Cell nombre = cabecera.createCell(1);
		Cell edad = cabecera.createCell(2);
		Cell ciudad = cabecera.createCell(3);

		nombre.setCellValue("Nombre");
		edad.setCellValue("Edad");
		ciudad.setCellValue("Ciudad");
		
		Cell nombreRegistro1 = registro1.createCell(1);
		Cell edadRegistro1 = registro1.createCell(2);
		Cell ciudadRegistro1 = registro1.createCell(3);

		nombreRegistro1.setCellValue("Santiago");
		edadRegistro1.setCellValue("23");
		ciudadRegistro1.setCellValue("Medellin");

		Cell nombreRegistro2 = registro2.createCell(1);
		Cell edadRegistro2 = registro2.createCell(2);
		Cell ciudadRegistro2 = registro2.createCell(3);

		nombreRegistro2.setCellValue("Anyi");
		edadRegistro2.setCellValue("22");
		ciudadRegistro2.setCellValue("Bogota");

		try {
			OutputStream output = new FileOutputStream("ArchivoExcel.xlsx");
			libro.write(output);

			libro.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}