package br.com.sankhya.repository;

import br.com.sankhya.jape.EntityFacade;
import br.com.sankhya.jape.bmp.PersistentLocalEntity;
import br.com.sankhya.jape.vo.DynamicVO;
import br.com.sankhya.jape.vo.EntityVO;
import br.com.sankhya.modelcore.util.EntityFacadeFactory;

import java.math.BigDecimal;

public class CabecalhoRepository {

    private final String entidadeCabecalho;

    public CabecalhoRepository() {
        this("AD_IMPTESTE");
    }

    public CabecalhoRepository(String entidadeCabecalho) {
        this.entidadeCabecalho = entidadeCabecalho;
    }

    public void atualizarStatusDoCabecalho(BigDecimal nunico, String status) throws Exception {
        EntityFacade dwfFacade = EntityFacadeFactory.getDWFFacade();
        PersistentLocalEntity finEntity = dwfFacade.findEntityByPrimaryKey(entidadeCabecalho, new Object[]{nunico});
        DynamicVO finVO = (DynamicVO) finEntity.getValueObject();
        finVO.setProperty("STATUS", status);
        finEntity.setValueObject((EntityVO) finVO);
    }
}
