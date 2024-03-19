package com.excel.poiji;

import com.poiji.annotation.ExcelCellName;

public class Persona {

	@ExcelCellName("id")
	private Long id;

	@ExcelCellName(value = "documento", mandatoryCell = true)
	private String documento;

	@ExcelCellName("nombre")
	private String nombre;

	@ExcelCellName("apellido")
	private String apellido;

	@ExcelCellName("ciudad")
	private String ciudad;

	@ExcelCellName("fecha nacimiento")
	private String fechaNacimiento;

	public Persona(Long id, String documento, String nombre, String apellido, String ciudad, String fechaNacimiento) {
		this.id = id;
		this.documento = documento;
		this.nombre = nombre;
		this.apellido = apellido;
		this.ciudad = ciudad;
		this.fechaNacimiento = fechaNacimiento;
	}

	public Persona() {
	}

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public String getDocumento() {
		return documento;
	}

	public void setDocumento(String documento) {
		this.documento = documento;
	}

	public String getNombre() {
		return nombre;
	}

	public void setNombre(String nombre) {
		this.nombre = nombre;
	}

	public String getApellido() {
		return apellido;
	}

	public void setApellido(String apellido) {
		this.apellido = apellido;
	}

	public String getCiudad() {
		return ciudad;
	}

	public void setCiudad(String ciudad) {
		this.ciudad = ciudad;
	}

	public String getFechaNacimiento() {
		return fechaNacimiento;
	}

	public void setFechaNacimiento(String fechaNacimiento) {
		this.fechaNacimiento = fechaNacimiento;
	}

	@Override
	public String toString() {
		return "Persona [id=" + id + ", documento=" + documento + ", nombre=" + nombre + ", apellido=" + apellido
				+ ", ciudad=" + ciudad + ", fechaNacimiento=" + fechaNacimiento + "]";
	}

}
