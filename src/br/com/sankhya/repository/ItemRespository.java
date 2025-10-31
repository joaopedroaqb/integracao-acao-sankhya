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
