package br.com.sankhya.utils;

import br.com.sankhya.jape.dao.JdbcWrapper;
import br.com.sankhya.jape.sql.NativeSql;
import br.com.sankhya.modelcore.util.EntityFacadeFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ArquivoUtils {

    public static String formatarCelulaComoString(Cell cell, int... casasDecimais) {
        if (cell == null) {
            return "";
        }

        int cellType = cell.getCellType();
        int decimal = (casasDecimais.length > 0) ? casasDecimais[0] : 0;

        try {
            switch (cellType) {
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
                        Date date = cell.getDateCellValue();
                        return df.format(date);
                    } else {
                        BigDecimal valorDecimal = BigDecimal.valueOf(cell.getNumericCellValue())
                                .setScale(decimal, RoundingMode.HALF_UP);
                        return valorDecimal.toPlainString();
                    }

                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();

                case Cell.CELL_TYPE_BOOLEAN:
                    return Boolean.toString(cell.getBooleanCellValue());

                case Cell.CELL_TYPE_FORMULA:
                    return cell.getCellFormula();

                case Cell.CELL_TYPE_ERROR:
                    return Byte.toString(cell.getErrorCellValue());

                case Cell.CELL_TYPE_BLANK:
                    return "";

                default:
                    return "Tipo de célula não reconhecido.";
            }
        } catch (Exception e) {
            return String.format("Erro ao formatar célula: %s", e.getMessage());
        }
    }

    public static String decodificarColunaPeloTipo(Cell cell) {
        String mensagem = "";
        switch (cell.getCellType()) {
            case 0: mensagem = "Númerico ou Data"; break;
            case 1: mensagem = "Texto"; break;
            case 2: mensagem = "Formula"; break;
            case 3: mensagem = "Vazio"; break;
            case 4: mensagem = "Booleano"; break;
            case 5: mensagem = "Error"; break;
        }
        return mensagem;
    }

    public static InputStream getLerArquivo(InputStream inputStream) throws Exception {
        JdbcWrapper jdbc = null;
        NativeSql query = null;

        try {
            jdbc = EntityFacadeFactory.getDWFFacade().getJdbcWrapper();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            StringBuffer buf = new StringBuffer();
            InputStream in = inputStream;

            byte[] b = new byte[2048];
            int length;
            boolean hasFileInfo = false;
            boolean writeDirectly = false;
            int offset = 0;
            while ((length = in.read(b)) > 0) {
                if (writeDirectly) {
                    baos.write(b, 0, length);
                } else {
                    offset = buf.length();
                    buf.append(new String(b));
                    if (!hasFileInfo && "__start_fileinformation__".equals(buf.substring(0, 25))) {
                        hasFileInfo = true;
                    }

                    if (hasFileInfo) {
                        int i = buf.indexOf("__end_fileinformation__");
                        if (i > -1) {
                            i += 23;// tamanho do "__end_fileinformation__"
                            i -= offset; // O quanto ja havia sido lido antes
                            baos.write(b, i, length - i);
                            writeDirectly = true;
                        }
                    } else {
                        baos.write(b, 0, length);
                        writeDirectly = true;
                    }
                }
            }
            baos.flush();
            in.close();
            inputStream = new ByteArrayInputStream(baos.toByteArray());
        } finally {
            NativeSql.releaseResources(query);
            JdbcWrapper.closeSession(jdbc);
        }
        return inputStream;
    }
}
