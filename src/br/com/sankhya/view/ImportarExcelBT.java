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
