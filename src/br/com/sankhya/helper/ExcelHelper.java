package br.com.sankhya.helper;

import br.com.sankhya.contants.ExcelMap;
import br.com.sankhya.utils.ArquivoUtils;
import com.sankhya.util.TimeUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

/**
 * Lê a primeira planilha e devolve registros como Map<campoVO, valor>.
 * - Casa por nome do cabeçalho (aliases) em qualquer ordem.
 * - Converte tipos (N=int, D=decimal, DT=Timestamp, S=string).
 * - Marca a origem com "_ROW" (1-based).
 * - IGNORA linhas ocultas (zeroHeight/hidden) se habilitado.
 * - NÃO interpreta critérios do AutoFilter (sem dependência ooxml-schemas).
 */
public class ExcelHelper {

    /** Ignorar linhas ocultas (ocultadas manualmente/AutoFilter visual). */
    private static final boolean IMPORTAR_APENAS_LINHAS_VISIVEIS = true;

    public static List<Map<String, Object>> processarArquivo(InputStream file) throws Exception {
        byte[] all = toBytes(file);
        try { // .xls
            return decode(new HSSFWorkbook(new ByteArrayInputStream(all)));
        } catch (Exception e1) { // .xlsx
            return decode(new XSSFWorkbook(new ByteArrayInputStream(all)));
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

        // campoVO -> índice da coluna no Excel (por alias)
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

            // Respeita visibilidade (sem CT*)
            if (IMPORTAR_APENAS_LINHAS_VISIVEIS && isRowHidden(sheet, row)) continue;

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

    // ---------- utilitários ----------
    private static boolean isRowHidden(Sheet sheet, Row row) {
        boolean hidden = false;
        try { hidden = row.getZeroHeight(); } catch (Exception ignore) {}
        try { if (row instanceof XSSFRow) hidden = hidden || ((XSSFRow) row).getCTRow().getHidden(); } catch (Exception ignore) {}
        try { if (row instanceof HSSFRow) hidden = hidden || ((HSSFRow) row).getZeroHeight(); } catch (Exception ignore) {}
        return hidden;
    }

    private static String normalize(String s) {
        if (s == null) return "";
        s = s.trim().toUpperCase();
        s = java.text.Normalizer.normalize(s, java.text.Normalizer.Form.NFD);
        s = s.replaceAll("\\p{InCombiningDiacriticalMarks}+", "");
        s = s.replace(".", "").replace("  ", " ");
        return s;
    }

    private static Object converter(Cell cell, String tipo) throws Exception {
        if (cell == null) return null;

        String raw = ArquivoUtils.formatarCelulaComoString(cell, 6);
        if (raw == null) return null;
        raw = raw.trim();
        if (raw.isEmpty()) return null;

        // “nulos” comuns
        String rawNorm = raw.replaceAll("[\\s\\-–—_]+", "");
        if (rawNorm.equalsIgnoreCase("null") || rawNorm.equalsIgnoreCase("na") || rawNorm.equalsIgnoreCase("n/a")) {
            return null;
        }

        if ("N".equals(tipo)) { // inteiro
            return new BigDecimal(raw.replace(",", ".")).setScale(0, RoundingMode.HALF_UP);
        }
        if ("D".equals(tipo)) { // decimal
            return new BigDecimal(raw.replace(",", ".")).setScale(6, RoundingMode.HALF_UP);
        }
        if ("DT".equals(tipo)) { // Timestamp (service remove depois, se necessário)
            if (!raw.matches("\\d{1,2}/\\d{1,2}/\\d{4}")) return null;
            if ("00/00/0000".equals(raw)) return null;
            return com.sankhya.util.TimeUtils.toTimestamp(raw, "dd/MM/yyyy");
        }
        return raw; // texto
    }
}
