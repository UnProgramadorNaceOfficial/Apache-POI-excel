package com.excel.prueba;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class Cliente {

	private Long id;
	private String nombre;
	private String apellido;
	private String documento;
	private String email;
	private LocalDate fechaNacimiento;

	public Cliente() {
	}

	public Cliente(Long id, String nombre, String apellido, String documento, String email, LocalDate fechaNacimiento) {
		this.id = id;
		this.nombre = nombre;
		this.apellido = apellido;
		this.documento = documento;
		this.email = email;
		this.fechaNacimiento = fechaNacimiento;
	}
	
	public List<Object> obtenerAtributos(){
		
		List<Object> atributos = new ArrayList<>();
		atributos.add(this.id);
		atributos.add(this.nombre);
		atributos.add(this.apellido);
		atributos.add(this.documento);
		atributos.add(this.email);
		atributos.add(this.fechaNacimiento);
		
		return atributos;
	}
	
	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
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

	public String getDocumento() {
		return documento;
	}

	public void setDocumento(String documento) {
		this.documento = documento;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public LocalDate getFechaNacimiento() {
		return fechaNacimiento;
	}

	public void setFechaNacimiento(LocalDate fechaNacimiento) {
		this.fechaNacimiento = fechaNacimiento;
	}

}
