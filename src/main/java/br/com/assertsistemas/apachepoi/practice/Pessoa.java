package br.com.assertsistemas.apachepoi.practice;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table (name = "tb_pessoa")
public class Pessoa {

	@Id
	@GeneratedValue (strategy = GenerationType.IDENTITY)
	@Column (name = "ID")
	private Long id;
	
	@Column (name = "NAME")
	private String nome;
	
	@Column (name = "AGE")
	private Integer idade;
	
	@Column (name = "VOCATION")
	private String cargo;
	
	@Column (name = "EARNINGS")
	private double salario;
	
	public Pessoa() {}
	
	public Pessoa (String nome, Integer idade, String cargo, double salario) {
		this.nome = nome;
		this.idade = idade;
		this.cargo = cargo;
		this.salario = salario;
	}

	public String getNome() {
		return nome;
	}

	public void setNome(String nome) {
		this.nome = nome;
	}

	public double getIdade() {
		return idade;
	}

	public void setIdade(Integer idade) {
		this.idade = idade;
	}

	public String getCargo() {
		return cargo;
	}

	public void setCargo(String cargo) {
		this.cargo = cargo;
	}

	public double getSalario() {
		return salario;
	}

	public void setSalario(double salario) {
		this.salario = salario;
	}
	
	
}
