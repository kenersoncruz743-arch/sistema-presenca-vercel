// ADICIONE ESTAS FUNÇÕES NO FINAL DO lib/sheets.js, ANTES DO module.exports

  // ==================== FUNÇÕES COM BUSCA POR ABA ====================
  
  async function removerBufferPorAba(chave, matricula) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const grupo = String(row.get('Grupo') || '').trim();
        const mat = String(row.get('matricula') || '').trim();
        
        // Busca por Supervisor OU por Grupo (aba)
        if ((supervisor === chave || grupo === chave) && mat === String(matricula)) {
          await row.delete();
          console.log(`[SHEETS] Removido colaborador ${matricula} da chave ${chave}`);
          return { ok: true };
        }
      }
      
      console.log(`[SHEETS] Colaborador ${matricula} não encontrado para chave ${chave}`);
      return { ok: false };
    } catch (error) {
      console.error('[SHEETS] Erro ao remover buffer por aba:', error);
      return { ok: false };
    }
  }

  async function atualizarStatusBufferPorAba(chave, matricula, status) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const grupo = String(row.get('Grupo') || '').trim();
        const mat = String(row.get('matricula') || '').trim();
        
        // Busca por Supervisor OU por Grupo (aba)
        if ((supervisor === chave || grupo === chave) && mat === String(matricula)) {
          row.set('status', status);
          await row.save();
          console.log(`[SHEETS] Status atualizado para ${matricula}: ${status}`);
          return { ok: true };
        }
      }
      
      console.log(`[SHEETS] Colaborador ${matricula} não encontrado para chave ${chave}`);
      return { ok: false };
    } catch (error) {
      console.error('[SHEETS] Erro ao atualizar status por aba:', error);
      return { ok: false };
    }
  }

  async function atualizarDesvioBufferPorAba(chave, matricula, desvio) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const grupo = String(row.get('Grupo') || '').trim();
        const mat = String(row.get('matricula') || '').trim();
        
        // Busca por Supervisor OU por Grupo (aba)
        if ((supervisor === chave || grupo === chave) && mat === String(matricula)) {
          row.set('desvio', desvio);
          await row.save();
          console.log(`[SHEETS] Desvio atualizado para ${matricula}: ${desvio}`);
          return { ok: true };
        }
      }
      
      console.log(`[SHEETS] Colaborador ${matricula} não encontrado para chave ${chave}`);
      return { ok: false };
    } catch (error) {
      console.error('[SHEETS] Erro ao atualizar desvio por aba:', error);
      return { ok: false };
    }
  }

// IMPORTANTE: Adicione estas linhas DENTRO da classe, antes do final:

  // Expõe as novas funções
  this.removerBufferPorAba = removerBufferPorAba;
  this.atualizarStatusBufferPorAba = atualizarStatusBufferPorAba;
  this.atualizarDesvioBufferPorAba = atualizarDesvioBufferPorAba;
