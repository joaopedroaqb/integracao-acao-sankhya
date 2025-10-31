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
