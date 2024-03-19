package com.excel;

import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GeneradorFuente {
	
	public static class Builder {
		
		private String nombreFuente;
		private boolean conNegrilla = false;
		private boolean conItalica = false;
		private short tamanioFuente;
		private short colorDefecto;
		private XSSFColor colorPersonalizado;
		private FontUnderline tipoUnderline;
		
		public Builder setNombreFuente(String nombreFuente) {
			this.nombreFuente = nombreFuente;
			return this;
		}
		
		public Builder setNegrilla(boolean conNegrilla) {
			this.conNegrilla = conNegrilla;
			return this;
		}
		
		public Builder setItalica(boolean conItalica) {
			this.conItalica = conItalica;
			return this;
		}
		
		public Builder setTamanioFuente(short tamanioFuente) {
			this.tamanioFuente = tamanioFuente;
			return this;
		}
		
		public Builder setColorDefecto(short colorDefecto) {
			this.colorDefecto = colorDefecto;
			return this;
		}
		
		public Builder setColorPersonalizado(String colorHex) {
			
			try {
				byte[] rgb = Hex.decodeHex(colorHex);
				this.colorPersonalizado = new XSSFColor(rgb);
				
				return this;
			} catch (Exception e) {
				throw new RuntimeException("Error al decodificar el color personalizado.");
			}
		}
		
		public Builder setUnderline(FontUnderline tipoUnderline) {
			this.tipoUnderline = tipoUnderline;
			return this;
		}
		
		public XSSFFont build(XSSFWorkbook libro) {
			
			XSSFFont fuente = libro.createFont();
			
			if(nombreFuente != null) {
				fuente.setFontName(nombreFuente);
			}
			
			if(conNegrilla) {
				fuente.setBold(conNegrilla);
			}
			
			if(conItalica) {
				fuente.setItalic(conItalica);
			}
			
			if(tamanioFuente != 0) {
				fuente.setFontHeightInPoints(tamanioFuente);
			}
			
			if(colorDefecto != 0) {
				fuente.setColor(colorDefecto);
			}
			
			if(colorPersonalizado != null) {
				fuente.setColor(colorPersonalizado);
			}
			
			if(tipoUnderline != null) {
				fuente.setUnderline(tipoUnderline);
			}
			
			return fuente;
		}
	}
}
