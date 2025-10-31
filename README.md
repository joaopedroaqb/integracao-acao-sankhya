#  Projeto de Importação de Planilhas Excel para Sankhya

Este projeto implementa uma **Rotina Java** personalizada para o ERP **Sankhya-Om**, responsável por **importar dados de planilhas Excel (.xls e .xlsx)** diretamente para tabelas customizadas, garantindo controle de validações, integridade dos dados e integração direta com o *framework JAPE* da Sankhya.

---

## Estrutura Geral do Projeto

###  Pacotes e Classes Principais

| Pacote | Classe | Função |
|--------|---------|--------|
| `br.com.sankhya.view` | **ImportarExcelBT** | Classe de botão de ação da rotina (executa via interface do Sankhya) |
| `br.com.sankhya.service` | **ImportacaoService** | Lógica de controle principal, leitura do arquivo, validações e orquestração de inserções |
| `br.com.sankhya.repository` | **ItemRespository** | CRUD dos itens da planilha na tabela `AD_IMPTESTEITE` |
| `br.com.sankhya.helper` | **ExcelHelper** | Leitura e conversão dos dados do arquivo Excel para um formato processável |
| `br.com.sankhya.contants` | **ExcelMap** | Mapeamento de colunas da planilha → campos da entidade Sankhya |
| `br.com.sankhya.utils` | **ArquivoUtils** | Funções utilitárias para leitura de arquivos, formatação de células e manipulação de InputStreams |

---

## Fluxo de Execução

1. O usuário **anexa o arquivo Excel** no cabeçalho (`AD_IMPTESTE.ARQUIVO`).
2. A rotina é executada através do botão **Importar Arquivo** (`ImportarExcelBT`).
3. A classe `ImportacaoService`:
   - Valida se o arquivo foi anexado e salvo;
   - Lê o conteúdo binário (BLOB);
   - Decodifica a planilha via `ExcelHelper`;
   - Remove campos de data (`DATAAGEND`, `DTVIGOR`);
   - Executa validações obrigatórias (`CODPROD`, `NUTAB`);
   - Reinsere as linhas válidas na tabela `AD_IMPTESTEITE`.

---

## Estrutura do Código

Os códigos completos utilizados estão descritos abaixo.

---

### ImportarExcelBT.java

```java
package br.com.sankhya.view;

import br.com.sankhya.extensions.actionbutton.AcaoRotinaJava;
import br.com.sankhya.extensions.actionbutton.ContextoAcao;
import br.com.sankhya.extensions.actionbutton.Registro;
import br.com.sankhya.service.ImportacaoService;

import java.math.BigDecimal;

public class ImportarExcelBT implements AcaoRotinaJava {

    @Override
    public void doAction(ContextoAcao ctx) throws Exception {
        Registro[] linhas = ctx != null ? ctx.getLinhas() : null;

        if (linhas == null || linhas.length == 0) {
            ctx.mostraErro("Selecione uma linha do cabeçalho para importar.");
            return;
        }

        for (Registro r : linhas) {
            if (r == null) continue;

            BigDecimal nunico = (BigDecimal) r.getCampo("NUNICO");
            if (nunico == null) {
                ctx.mostraErro("Campo NUNICO não encontrado na linha selecionada.");
                continue;
            }

            new ImportacaoService().gerenciarImportacao(nunico);
        }

        ctx.setMensagemRetorno("Importação executada com sucesso!");
    }
}

```

### ImportacaoService.java

```java
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

```

### ItemRespository.java

```java
package br.com.sankhya.repository;

import br.com.sankhya.jape.EntityFacade;
import br.com.sankhya.jape.core.JapeSession;
import br.com.sankhya.jape.vo.DynamicVO;
import br.com.sankhya.jape.vo.EntityVO;
import br.com.sankhya.jape.wrapper.JapeFactory;
import br.com.sankhya.jape.wrapper.JapeWrapper;
import br.com.sankhya.modelcore.MGEModelException;
import br.com.sankhya.modelcore.util.EntityFacadeFactory;
import org.apache.commons.lang3.exception.ExceptionUtils;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Map;

public class ItemRespository {

    private static final String ENTIDADE_ITEM = "AD_IMPTESTEITE";

    public void deletarItens(BigDecimal nunico) throws MGEModelException {
        JapeSession.SessionHandle hnd = null;
        try {
            hnd = JapeSession.open();
            JapeWrapper dao = JapeFactory.dao(ENTIDADE_ITEM);
            dao.deleteByCriteria("NUNICO = ?", nunico);
        } catch (Exception e) {
            MGEModelException.throwMe(e);
        } finally {
            JapeSession.close(hnd);
        }
    }

    public void incluirLinha(BigDecimal nunico, Map<String, Object> campos) throws Exception {
        try {
            EntityFacade dwf = EntityFacadeFactory.getDWFFacade();
            DynamicVO vo = (DynamicVO) dwf.getDefaultValueObjectInstance(ENTIDADE_ITEM);

            // FK do cabeçalho
            vo.setProperty("NUNICO", nunico);

            // Atribui campos (converte Date -> Timestamp; ignora meta "_ROW")
            for (Map.Entry<String, Object> e : campos.entrySet()) {
                String campo = e.getKey();
                if ("_ROW".equals(campo)) continue;

                Object valor = e.getValue();
                if (valor instanceof java.util.Date && !(valor instanceof Timestamp)) {
                    valor = new Timestamp(((java.util.Date) valor).getTime());
                }
                try {
                    vo.setProperty(campo, valor);
                } catch (Exception ignore) {
                    // ignora campos inexistentes no VO (segurança)
                }
            }

            dwf.createEntity(ENTIDADE_ITEM, (EntityVO) vo);
        } catch (Exception e) {
            throw new Exception("Erro ao incluir item da planilha: " + ExceptionUtils.getStackTrace(e));
        }
    }
}

```

### ExcelHelper.java

```java
package br.com.sankhya.helper;

import br.com.sankhya.contants.ExcelMap;
import br.com.sankhya.utils.ArquivoUtils;
import com.sankhya.util.TimeUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;

/**
 * Lê a primeira planilha e devolve registros como Map<campoVO, valor>.
 * - Casa por NOME do cabeçalho (aliases), em qualquer ordem.
 * - Converte tipos (N, D, DT, S). DT é Timestamp, mas as datas serão removidas no service.
 * - Anota a linha de origem com "_ROW" (1-based).
 */
public class ExcelHelper {

    public static List<Map<String, Object>> processarArquivo(InputStream file) throws Exception {
        byte[] all = toBytes(file);
        try {
            return decode(new HSSFWorkbook(new ByteArrayInputStream(all)));
        } catch (Exception e1) {
            try {
                return decode(new XSSFWorkbook(new ByteArrayInputStream(all)));
            } catch (Exception e2) {
                throw new Exception("Erro ao abrir Excel: " + e2.getMessage(), e2);
            }
        }
    }

    private static byte[] toBytes(InputStream in) throws Exception {
        byte[] buf = new byte[8192];
        int len;
        java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
        while ((len = in.read(buf)) != -1) baos.write(buf, 0, len);
        return baos.toByteArray();
    }

    private static List<Map<String, Object>> decode(Workbook wb) throws Exception {
        if (wb == null || wb.getNumberOfSheets() == 0) throw new Exception("Arquivo sem planilhas.");
        Sheet sheet = wb.getSheetAt(0);
        Row header = sheet.getRow(0);
        if (header == null) throw new Exception("Cabeçalho (linha 1) ausente.");

        // nome normalizado -> índice da coluna
        Map<String, Integer> idxExcel = new LinkedHashMap<>();
        for (int c = header.getFirstCellNum(); c <= header.getLastCellNum(); c++) {
            Cell cell = header.getCell(c);
            String name = normalize(ArquivoUtils.formatarCelulaComoString(cell));
            if (!name.isEmpty()) idxExcel.put(name, c);
        }

        // campoVO -> índice da coluna no Excel
        Map<String, String> aliases = ExcelMap.aliases();
        Map<String, Integer> voToIdx = new LinkedHashMap<>();
        for (Map.Entry<String, String> e : aliases.entrySet()) {
            String excelName = normalize(e.getKey());
            if (idxExcel.containsKey(excelName)) {
                voToIdx.put(e.getValue(), idxExcel.get(excelName));
            }
        }

        Map<String, String> tipos = ExcelMap.tiposVo();
        List<Map<String, Object>> out = new ArrayList<>();

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Map<String, Object> registro = new LinkedHashMap<>();
            boolean vazio = true;

            for (Map.Entry<String, Integer> ent : voToIdx.entrySet()) {
                String campoVo = ent.getKey();
                int col = ent.getValue();
                String tipo = tipos.get(campoVo);
                Object valor = converter(row.getCell(col), tipo);

                if (valor != null && !(valor instanceof String && ((String) valor).trim().isEmpty())) {
                    registro.put(campoVo, valor);
                    vazio = false;
                }
            }

            if (!vazio) {
                registro.put("_ROW", r + 1);
                out.add(registro);
            }
        }

        return out;
    }

    private static Object converter(Cell cell, String tipo) throws Exception {
        if (cell == null) return null;

        String raw = ArquivoUtils.formatarCelulaComoString(cell, 6);
        if (raw == null) return null;
        raw = raw.trim();
        if (raw.isEmpty()) return null;

        // "nulos" comuns
        String rawNorm = raw.replaceAll("[\\s\\-–—_]+", "");
        if (rawNorm.equalsIgnoreCase("null") || rawNorm.equalsIgnoreCase("na") || rawNorm.equalsIgnoreCase("n/a")) {
            return null;
        }

        if ("N".equals(tipo)) { // inteiro
            return new BigDecimal(raw.replace(",", ".")).setScale(0, BigDecimal.ROUND_HALF_UP);
        }
        if ("D".equals(tipo)) { // decimal
            return new BigDecimal(raw.replace(",", ".")).setScale(6, BigDecimal.ROUND_HALF_UP);
        }
        if ("DT".equals(tipo)) { // Timestamp (se formatado; o service remove de qualquer forma)
            if (!raw.matches("\\d{1,2}/\\d{1,2}/\\d{4}")) return null;
            if ("00/00/0000".equals(raw)) return null;
            return TimeUtils.toTimestamp(raw, "dd/MM/yyyy");
        }
        return raw; // string
    }

    private static String normalize(String s) {
        if (s == null) return "";
        s = s.trim().toUpperCase();
        s = java.text.Normalizer.normalize(s, java.text.Normalizer.Form.NFD);
        s = s.replaceAll("\\p{InCombiningDiacriticalMarks}+", "");
        s = s.replace(".", "").replace("  ", " ");
        return s;
    }
}

```

---

## Bibliotecas Utilizadas

| Biblioteca | Finalidade |
|-------------|-------------|
| **Apache POI (poi, poi-ooxml)** | Leitura e interpretação de planilhas Excel (`.xls` e `.xlsx`) |
| **Sankhya JAPE API** | Operações de persistência e manipulação de entidades (`DynamicVO`, `EntityFacade`) |
| **Sankhya Action Button API** | Execução de rotinas Java via botões personalizados no ERP |
| **Apache Commons Lang3** | Utilitário para exceptions e manipulação de strings |
| **Sankhya ModelCore Util** | Criação e manipulação de `EntityFacade` padrão |

---

## Observações Técnicas

- Todas as operações de escrita ocorrem dentro de **sessões JAPE controladas**.
- As colunas da planilha podem vir em **qualquer ordem**, pois são mapeadas pelos **aliases** definidos em `ExcelMap.java`.
- Campos de data (`DATAAGEND`, `DTVIGOR`) são **removidos automaticamente** e **não gravados**.
- O sistema atribui valores automáticos a `AD_IDEXTERNO` de forma **sequencial**.
- A importação **substitui** os itens existentes vinculados ao mesmo `NUNICO`.
