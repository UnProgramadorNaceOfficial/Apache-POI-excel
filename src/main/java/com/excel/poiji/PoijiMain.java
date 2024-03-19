package com.excel.poiji;

import java.io.File;
import java.util.List;

import com.poiji.bind.Poiji;
import com.poiji.option.PoijiOptions;

public class PoijiMain {

	public static void main(String[] args) {
		
		File archivo = new File("DatosClientes.xlsx");
		
		PoijiOptions options = PoijiOptions.PoijiOptionsBuilder
										   .settings()
//										   .sheetIndex(0) 		  // Indice de hoja
										   .sheetName("Personas") // Nombre de hoja
//										   .skip(3)				  // Salto de filas
//										   .limit(6)			  // Limite de registros
										   .trimCellValue(true)   // Eliminar espacios en blanco	
										   .password("1234")      // Contrasenia del archivo
										   .build();
		
		List<Persona> personas = Poiji.fromExcel(archivo, Persona.class, options);
		
		personas.forEach(System.out::println);
		
	}

}
