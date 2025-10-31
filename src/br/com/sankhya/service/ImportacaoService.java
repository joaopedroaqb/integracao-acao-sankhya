package br.com.sankhya.service;

import br.com.sankhya.helper.ExcelHelper;
import br.com.sankhya.jape.EntityFacade;
import br.com.sankhya.jape.vo.DynamicVO;
import br.com.sankhya.modelcore.util.EntityFacadeFactory;
import br.com.sankhya.repository.ItemRespository;
import br.com.sankhya.utils.ArquivoUtils;

import java.io.ByteArrayInputStream;
import java.math.BigDecimal;
import java.util.*;

public class ImportacaoService {

    private static final String ENTIDADE_CAB = "AD_IMPTESTE";
    private static final String CAMPO_ARQUIVO = "ARQUIVO";

    private final ItemRespository itemRepo;

    public ImportacaoService() {
        this.itemRepo = new ItemRespository();
    }

    public void gerenciarImportacao(BigDecimal nunico) throws Exception {
        EntityFacade dwf = EntityFacadeFactory.getDWFFacade();

        DynamicVO cab = (DynamicVO) dwf.findEntityByPrimaryKeyAsVO(ENTIDADE_CAB, new Object[]{ nunico });
        if (cab == null) throw new Exception("Cabeçalho não encontrado (NUNICO=" + nunico + ").");

        byte[] blob = cab.asBlob(CAMPO_ARQUIVO);
        if (blob == null || blob.length == 0)
            throw new Exception("Nenhum arquivo anexado no campo '" + CAMPO_ARQUIVO + "'.");

        // Lê e decodifica a planilha
        List<Map<String, Object>> linhas = ExcelHelper.processarArquivo(
                ArquivoUtils.getLerArquivo(new ByteArrayInputStream(blob))
        );
        if (linhas.isEmpty()) throw new Exception("A planilha não possui linhas válidas.");

        // Validação de obrigatórios
        List<Integer> faltantesCodProd = new ArrayList<>();
        List<Integer> faltantesNutab   = new ArrayList<>();

        for (Map<String, Object> r : linhas) {
            Object rowNum = r.get("_ROW");
            int row = (rowNum instanceof Number) ? ((Number) rowNum).intValue() : -1;

            if (!r.containsKey("CODPROD") || r.get("CODPROD") == null || r.get("CODPROD").toString().trim().isEmpty())
                faltantesCodProd.add(row);
            if (!r.containsKey("NUTAB") || r.get("NUTAB") == null || r.get("NUTAB").toString().trim().isEmpty())
                faltantesNutab.add(row);
        }

        if (!faltantesCodProd.isEmpty() || !faltantesNutab.isEmpty()) {
            StringBuilder sb = new StringBuilder("Planilha inválida. Campos obrigatórios ausentes:\n");
            if (!faltantesCodProd.isEmpty()) sb.append(" - CODPROD vazio nas linhas: ").append(faltantesCodProd).append("\n");
            if (!faltantesNutab.isEmpty()) sb.append(" - NUTAB vazio nas linhas: ").append(faltantesNutab).append("\n");
            sb.append("Ajuste o arquivo e importe novamente.");
            throw new Exception(sb.toString());
        }

        // <<< Regra: NUNCA gravar datas (deixa DATAAGEND e DTVIGOR em branco)
        for (Map<String, Object> r : linhas) {
            r.remove("DATAAGEND");
            r.remove("DTVIGOR");
        }

        // Substitui itens
        itemRepo.deletarItens(nunico);

        int sequencial = 1;
        for (Map<String, Object> registro : linhas) {
            registro.remove("_ROW"); // metadado
            registro.putIfAbsent("AD_IDEXTERNO", String.valueOf(sequencial++));
            itemRepo.incluirLinha(nunico, registro);
        }
    }
}
