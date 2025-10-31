package br.com.sankhya.contants;

import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelMap {

    /** Cabeçalho Excel (aliases) → campo VO (AD_IMPTESTEITE) */
    public static Map<String, String> aliases() {
        Map<String, String> m = new LinkedHashMap<>();

        m.put("NUTAB", "NUTAB");

        // CODPROD com várias variações comuns
        m.put("CÓD PROD", "CODPROD");
        m.put("COD PROD", "CODPROD");
        m.put("CÓD PROD.", "CODPROD");
        m.put("COD. PROD", "CODPROD");
        m.put("COD. PROD.", "CODPROD");
        m.put("CODPROD", "CODPROD");
        m.put("CÓDIGO DO PRODUTO", "CODPROD");
        m.put("CODIGO DO PRODUTO", "CODPROD");
        m.put("CÓDIGO PRODUTO", "CODPROD");
        m.put("CODIGO PRODUTO", "CODPROD");
        // fallbacks (cautela)
        m.put("CÓDIGO", "CODPROD");
        m.put("CODIGO", "CODPROD");

        // valores
        m.put("VAREJO (CX)", "VLRVENDA");
        m.put("VAREJO CX", "VLRVENDA");
        m.put("VAREJO", "VLRVENDA");

        m.put("601 A CARGA FECHADA (CX)", "VLRVENDAV2");
        m.put("CARGA FECHADA (CX)", "VLRVENDAV2");

        m.put("CARRETA (CX)", "VLRVENDAV3");

        m.put("% CONTRATO", "PERCCON");
        m.put("PERCENTUAL CONTRATO", "PERCCON");

        // Datas (vamos remover no service, mas deixamos o alias caso apareça)
        m.put("DATA DE AGENDAMENTO", "DATAAGEND");
        m.put("DATA AGENDAMENTO", "DATAAGEND");
        m.put("VIGÊNCIA", "DTVIGOR");
        m.put("VIGENCIA", "DTVIGOR");

        // Campo sequencial (se não vier, geramos)
        m.put("SEQUENCIAL", "AD_IDEXTERNO");

        return m;
    }

    /** Tipos esperados para conversão (N=int, D=decimal, DT=Timestamp, S=string) */
    public static Map<String, String> tiposVo() {
        Map<String, String> t = new LinkedHashMap<>();
        t.put("NUTAB", "N");
        t.put("CODPROD", "N");
        t.put("VLRVENDA", "D");
        t.put("VLRVENDAV2", "D");
        t.put("VLRVENDAV3", "D");
        t.put("PERCCON", "D");
        t.put("AD_IDEXTERNO", "S");

        // datas (convertidas, mas removidas no service para ficar em branco)
        t.put("DATAAGEND", "DT");
        t.put("DTVIGOR", "DT");
        return t;
    }
}
