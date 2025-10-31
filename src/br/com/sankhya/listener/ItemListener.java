package br.com.sankhya.listener;

import br.com.sankhya.extensions.eventoprogramavel.EventoProgramavelJava;
import br.com.sankhya.jape.event.PersistenceEvent;
import br.com.sankhya.jape.event.TransactionContext;
import br.com.sankhya.jape.vo.DynamicVO;
import com.sankhya.util.TimeUtils;

import java.math.BigDecimal;

public class ItemListener implements EventoProgramavelJava {

    public ItemListener() {}

    @Override
    public void beforeInsert(PersistenceEvent event) throws Exception {
        DynamicVO itemVO = (DynamicVO) event.getVo();
        itemVO.setProperty("DTALTER", TimeUtils.getNow());
    }

    @Override
    public void beforeUpdate(PersistenceEvent event) throws Exception {
        DynamicVO itemVO = (DynamicVO) event.getVo();
        bloquearAlteracaoOuDeletacao(itemVO);
    }

    @Override
    public void beforeDelete(PersistenceEvent event) throws Exception {
        DynamicVO itemVO = (DynamicVO) event.getVo();
        bloquearAlteracaoOuDeletacao(itemVO);
    }

    private void bloquearAlteracaoOuDeletacao(DynamicVO itemVO) throws Exception {
        BigDecimal nuNota = itemVO.asBigDecimal("NUNOTA");
        if (nuNota != null) {
            throw new Exception("Nota já foi gerada; não é possível editar ou excluir o item.");
        }
    }

    @Override public void afterInsert(PersistenceEvent event) throws Exception {}
    @Override public void afterUpdate(PersistenceEvent event) throws Exception {}
    @Override public void afterDelete(PersistenceEvent event) throws Exception {}
    @Override public void beforeCommit(TransactionContext tranCtx) throws Exception {}
}
