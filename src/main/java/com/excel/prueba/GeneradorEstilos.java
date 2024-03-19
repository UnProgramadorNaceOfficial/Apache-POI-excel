package com.excel.prueba;

import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GeneradorEstilos {

	public static class Builder {
		
		private short colorDefecto;
		private XSSFColor colorPersonalizado;
		private FillPatternType tipoPatron;
		private XSSFFont fuente;
		private String formato;
		private HorizontalAlignment alineacionHorizontal;
		private VerticalAlignment alineacionVertical;
		private BorderStyle bordeArriba;
		private BorderStyle bordeAbajo;
		private BorderStyle bordeDerecho;
		private BorderStyle bordeIzquierdo;
		
		public Builder setColorDefecto(short colorDefecto) {
			this.colorDefecto = colorDefecto;
			return this;
		}
		
		public Builder setColorPersonalizado(String colorPersonalizado) {
			try {
				byte[] rgb = Hex.decodeHex(colorPersonalizado);
				this.colorPersonalizado = new XSSFColor(rgb);
				
				return this;
				
			} catch (Exception e) {
				throw new RuntimeException("Error al decodificar el color.");
			}
		}
		
		public Builder setTipoPatron(FillPatternType tipoPatron) {
			this.tipoPatron = tipoPatron;
			return this;
		}
		
		public Builder setFuente(XSSFFont fuente) {
			this.fuente = fuente;
			return this;
		}
		
		public Builder setFormato(String formato) {
			this.formato = formato;
			return this;
		}
		
		public Builder setAlineacionHorizontal(HorizontalAlignment alineacionHorizontal) {
			this.alineacionHorizontal = alineacionHorizontal;
			return this;
		}
		
		public Builder setAlineacionVertical(VerticalAlignment alineacionVertical) {
			this.alineacionVertical = alineacionVertical;
			return this;
		}
		
		public Builder setBordeArriba(BorderStyle bordeArriba) {
			this.bordeArriba = bordeArriba;
			return this;
		}
		
		public Builder setBordeAbajo(BorderStyle bordeAbajo) {
			this.bordeAbajo = bordeAbajo;
			return this;
		}
		
		public Builder setBordeDerecho(BorderStyle bordeDerecho) {
			this.bordeDerecho = bordeDerecho;
			return this;
		}
		
		public Builder setBordeIzquierdo(BorderStyle bordeIzquierdo) {
			this.bordeIzquierdo = bordeIzquierdo;
			return this;
		}
		
		public XSSFCellStyle build(XSSFWorkbook libro) {
			
			XSSFCellStyle estiloCelda = libro.createCellStyle();
			
			if(this.colorDefecto != 0) {
				estiloCelda.setFillForegroundColor(colorDefecto);
			}
			
			if(this.colorPersonalizado != null) {
				estiloCelda.setFillForegroundColor(colorPersonalizado);
			}
			
			if(this.tipoPatron != null) {
				estiloCelda.setFillPattern(tipoPatron);
			}
			
			if(this.fuente != null) {
				estiloCelda.setFont(fuente);
			}
			
			if(this.formato != null) {
				estiloCelda.setDataFormat(libro.createDataFormat().getFormat(formato));
			}
			
			if(this.alineacionHorizontal != null) {
				estiloCelda.setAlignment(alineacionHorizontal);
			}
			
			if(this.alineacionVertical != null) {
				estiloCelda.setVerticalAlignment(alineacionVertical);
			}
			
			if(this.bordeArriba != null) {
				estiloCelda.setBorderTop(bordeArriba);
			}
			
			if(this.bordeAbajo != null) {
				estiloCelda.setBorderBottom(bordeAbajo);
			}
			
			if(this.bordeDerecho != null) {
				estiloCelda.setBorderRight(bordeDerecho);
			}
			
			if(this.bordeIzquierdo != null) {
				estiloCelda.setBorderLeft(bordeIzquierdo);
			}
			
			return estiloCelda;
		}
	}
}
