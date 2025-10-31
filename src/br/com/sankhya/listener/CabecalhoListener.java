package br.com.sankhya.listener;

import br.com.sankhya.extensions.eventoprogramavel.EventoProgramavelJava;
import br.com.sankhya.jape.event.ModifingFields;
import br.com.sankhya.jape.event.PersistenceEvent;
import br.com.sankhya.jape.event.TransactionContext;
import br.com.sankhya.jape.vo.DynamicVO;
import br.com.sankhya.modelcore.auth.AuthenticationInfo;
import com.sankhya.util.TimeUtils;

public class CabecalhoListener implements EventoProgramavelJava {
    @Override
    public void beforeInsert(PersistenceEvent event) throws Exception {
        DynamicVO cabecalhoVO  = (DynamicVO) event.getVo();
        cabecalhoVO.setProperty("STATUS", "P");
        cabecalhoVO.setProperty("DTIMP", TimeUtils.getNow());
        cabecalhoVO.setProperty("CODUSU", AuthenticationInfo.getCurrent().getUserID());
    }

    @Override
    public void beforeUpdate(PersistenceEvent event) throws Exception {
        DynamicVO cabecalhoVO  = (DynamicVO) event.getVo();
        if(!(cabecalhoVO.asString("STATUS").equals("P") || cabecalhoVO.asString("STATUS").equals("I"))){
            ModifingFields modifingFields = event.getModifingFields();
            if(modifingFields.isModifing("ARQUIVO")){
                throw new Exception("Ops! O status é diferente de pendente, caso deseje alterar exclua a importação pelo botão de ação.");
            }
        }
    }

    @Override public void beforeDelete(PersistenceEvent event) throws Exception {}
    @Override public void afterInsert(PersistenceEvent event) throws Exception {}
    @Override public void afterUpdate(PersistenceEvent event) throws Exception {}
    @Override public void afterDelete(PersistenceEvent event) throws Exception {}
    @Override public void beforeCommit(TransactionContext tranCtx) throws Exception {}
}
