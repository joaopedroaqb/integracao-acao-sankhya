package br.com.sankhya.model;

import java.util.HashMap;
import java.util.Map;

import br.com.sankhya.model.ModelExcelDTO.ColunaExcel;
import br.com.sankhya.model.ModelExcelDTO.EnumTipoColuna;

public class ModelExcelDTO {

    public static class ColunaExcel {
        private int index;
        private String nomeColuna;
        private EnumTipoColuna tipoColuna;
        private Object conteudo;

        public ColunaExcel() {}

        public ColunaExcel(String nomeColuna, EnumTipoColuna tipoColuna, int index) {
            this.index = index;
            this.nomeColuna = nomeColuna;
            this.tipoColuna = tipoColuna;
            System.out.println("Index: "+ index + " - nomecoluna: "+ nomeColuna + " - tipoColuna: " + tipoColuna );
        }

        public int getIndex() { return index; }
        public void setIndex(int index) { this.index = index; }

        public String getNomeColuna() { return nomeColuna; }
        public void setNomeColuna(String nomeColuna) { this.nomeColuna = nomeColuna; }

        public EnumTipoColuna getTipoColuna() { return tipoColuna; }
        public void setTipoColuna(EnumTipoColuna tipoColuna) { this.tipoColuna = tipoColuna; }

        public Object getConteudo() { return conteudo; }
        public void setConteudo(Object conteudo) { this.conteudo = conteudo; }

        @Override
        public String toString() {
            return "ColunaExcel{" +
                    "index=" + index +
                    ", nomeColuna='" + nomeColuna + '\'' +
                    ", tipoColuna=" + tipoColuna +
                    ", conteudo=" + conteudo +
                    '}';
        }
    }

    public enum EnumTipoColuna {
        DATA("DT"),
        NUMERO("B"),
        DECIMAL("D"),
        STRING("S");

        private String valor;
        private static final Map<String, EnumTipoColuna> funcaoPorValor = new HashMap<>();

        EnumTipoColuna(String valor) { this.valor = valor; }

        public String getValor() { return valor; }

        public static EnumTipoColuna getEnum(String descricao) {
            return funcaoPorValor.get(descricao);
        }

        static {
            for (EnumTipoColuna e : EnumTipoColuna.values()) {
                funcaoPorValor.put(e.getValor(), e);
            }
        }
    }

    public static class ValidaColuna {
        private boolean isAcionado;
        private String colunaRecebida;
        private ColunaExcel coluna;

        public ValidaColuna(ColunaExcel coluna) { this.coluna = coluna; }
        public ValidaColuna(boolean isAcionado) { this.isAcionado = isAcionado; }
        public ValidaColuna(boolean isAcionado, ColunaExcel coluna, String colunaRecebida, int indice) {
            this.isAcionado = isAcionado;
            this.coluna = coluna;
            this.colunaRecebida = colunaRecebida;
        }

        public boolean isAcionado() { return isAcionado; }
        public void setAcionado(boolean acionado) { isAcionado = acionado; }

        public String getColunaRecebida() { return colunaRecebida; }
        public void setColunaRecebida(String colunaRecebida) { this.colunaRecebida = colunaRecebida; }

        public ColunaExcel getColuna() { return coluna; }
        public void setColuna(ColunaExcel coluna) { this.coluna = coluna; }

        @Override
        public String toString() {
            return "ValidaColuna{" +
                    "isAcionado=" + isAcionado +
                    ", colunaRecebida='" + colunaRecebida + '\'' +
                    ", coluna=" + coluna +
                    '}';
        }
    }
}
