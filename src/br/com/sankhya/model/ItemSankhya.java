package br.com.sankhya.model;

import java.math.BigDecimal;
import java.sql.Timestamp;

public class ItemSankhya {

    private BigDecimal nroUnico;
    private Timestamp data;
    private String nroCartao;
    private String motorista;
    private String placa;
    private String descricaoDoVeiculo;
    private BigDecimal horimetro;
    private BigDecimal distancia;
    private String produto;
    private BigDecimal quantidade;

    // NOVO campo para o importador (coluna “Descrição” do Excel)
    private String descricao;

    public BigDecimal getNroUnico() {
        return nroUnico;
    }

    public void setNroUnico(BigDecimal nroUnico) {
        this.nroUnico = nroUnico;
    }

    public Timestamp getData() {
        return data;
    }

    public void setData(Timestamp data) {
        this.data = data;
    }

    public String getNroCartao() {
        return nroCartao;
    }

    public void setNroCartao(String nroCartao) {
        this.nroCartao = nroCartao;
    }

    public String getMotorista() {
        return motorista;
    }

    public void setMotorista(String motorista) {
        this.motorista = motorista;
    }

    public String getPlaca() {
        return placa;
    }

    public void setPlaca(String placa) {
        this.placa = placa;
    }

    public BigDecimal getHorimetro() {
        return horimetro;
    }

    public void setHorimetro(BigDecimal horimetro) {
        this.horimetro = horimetro;
    }

    public BigDecimal getDistancia() {
        return distancia;
    }

    public void setDistancia(BigDecimal distancia) {
        this.distancia = distancia;
    }

    public String getProduto() {
        return produto;
    }

    public void setProduto(String produto) {
        this.produto = produto;
    }

    public BigDecimal getQuantidade() {
        return quantidade;
    }

    public void setQuantidade(BigDecimal quantidade) {
        this.quantidade = quantidade;
    }

    public String getDescricaoDoVeiculo() {
        return descricaoDoVeiculo;
    }

    public void setDescricaoDoVeiculo(String descricaoDoVeiculo) {
        this.descricaoDoVeiculo = descricaoDoVeiculo;
    }

    // NOVOS getters e setters
    public String getDescricao() {
        return descricao;
    }

    public void setDescricao(String descricao) {
        this.descricao = descricao;
    }

    @Override
    public String toString() {
        return "ItemSankhya{" +
                "nroUnico=" + nroUnico +
                ", data=" + data +
                ", nroCartao='" + nroCartao + '\'' +
                ", motorista='" + motorista + '\'' +
                ", placa='" + placa + '\'' +
                ", descricaoDoVeiculo='" + descricaoDoVeiculo + '\'' +
                ", horimetro=" + horimetro +
                ", distancia=" + distancia +
                ", produto='" + produto + '\'' +
                ", quantidade=" + quantidade +
                ", descricao='" + descricao + '\'' +
                '}';
    }
}
