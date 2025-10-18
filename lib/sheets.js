// lib/sheets.js - VERSÃO COM CAMPO DESVIO
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsService {
  constructor() {
    this.doc = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return this.doc;
    try {
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      this.initialized = true;
      return this.doc;
    } catch (error) {
      throw new Error('Falha na conexão: ' + error.message);
    }
  }

  async validarLogin(usuario, senha) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Usuarios'];
      if (!sheet) return { ok: false, msg: 'Aba Usuarios não encontrada' };
      
      const rows = await sheet.getRows();
      const abas = [];

      for (const row of rows) {
        const u = String(row.get('Usuario') || '').trim();
        const s = String(row.get('Senha') || '').trim();
        const aba = String(row.get('Aba') || '').trim();
        
        if (u === usuario && s === senha && aba) {
          abas.push(aba);
        }
      }

      const unicas = [...new Set(abas)];
      return unicas.length > 0 ? { ok: true, usuario, abas: unicas } : { ok: false, msg: 'Login inválido' };
    } catch (error) {
      return { ok: false, msg: error.message };
    }
  }

  async buscarColaboradores(filtro = '') {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Quadro'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const lista = [];

      for (const row of rows) {
        const matricula = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcao = String(row.get('Função que atua') || row.get('FUNÇÃO NO RM') || '').trim();
        
        if (nome && (!filtro || nome.toLowerCase().includes(filtro.toLowerCase()) || matricula.toLowerCase().includes(filtro.toLowerCase()))) {
          lista.push({ matricula, nome, funcao });
        }
      }

      return lista;
    } catch (error) {
      console.error('Erro busca:', error);
      return [];
    }
  }

  async getBuffer(supervisor, aba) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const buffer = [];

      for (const row of rows) {
        const sup = String(row.get('Supervisor') || '').toLowerCase();
        const rowAba = String(row.get('Grupo') || '').toLowerCase();
        
        if (sup === supervisor.toLowerCase() && (!aba || rowAba === aba.toLowerCase())) {
          buffer.push({
            matricula: row.get('matricula'),
            nome: row.get('Nome'),
            funcao: row.get('Função'),
            status: row.get('status') || '',
            desvio: row.get('desvio') || ''
          });
        }
      }

      return buffer;
    } catch (error) {
      console.error('Erro buffer:', error);
      return [];
    }
  }

  async adicionarBuffer(supervisor, aba, colaborador) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false, msg: 'Aba Lista não encontrada' };
      
      await sheet.addRow({
        'Supervisor': supervisor,
        'Grupo': aba,
        'matricula': colaborador.matricula,
        'Nome': colaborador.nome,
        'Função': colaborador.funcao,
        'status': '',
        'desvio': ''
      });

      return { ok: true };
    } catch (error) {
      console.error('Erro adicionar:', error);
      return { ok: false, msg: error.message };
    }
  }

  async removerBuffer(supervisor, matricula) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && String(row.get('matricula')) === String(matricula)) {
          await row.delete();
          return { ok: true };
        }
      }
      
      return { ok: false };
    } catch (error) {
      console.error('Erro remover:', error);
      return { ok: false };
    }
  }

  async atualizarStatusBuffer(supervisor, matricula, status) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && String(row.get('matricula')) === String(matricula)) {
          row.set('status', status);
          await row.save();
          return { ok: true };
        }
      }
      
      return { ok: false };
    } catch (error) {
      console.error('Erro atualizar status:', error);
      return { ok: false };
    }
  }

  async atualizarDesvioBuffer(supervisor, matricula, desvio) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) return { ok: false };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && String(row.get('matricula')) === String(matricula)) {
          row.set('desvio', desvio);
          await row.save();
          return { ok: true };
        }
      }
      
      return { ok: false };
    } catch (error) {
      console.error('Erro atualizar desvio:', error);
      return { ok: false };
    }
  }

  // ==================== FUNÇÕES MAPA DE CARGA ====================

  async getMapaCarga(filtros = {}) {
    try {
      console.log('[SHEETS] Iniciando getMapaCarga...');
      console.log('[SHEETS] Filtros recebidos:', filtros);
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        console.error('[SHEETS] Aba "Mapa de Carga" não encontrada');
        console.error('[SHEETS] Abas disponíveis:', Object.keys(this.doc.sheetsByTitle));
        return [];
      }
      
      console.log('[SHEETS] Aba encontrada:', sheet.title);
      
      const rows = await sheet.getRows();
      console.log('[SHEETS] Total de linhas:', rows.length);
      
      const dados = [];
  
      for (const row of rows) {
        try {
          const m3Raw = String(row.get('M³') || '0').trim();
          const m3Value = parseFloat(m3Raw.replace(',', '.')) || 0;
          
          const visitasPendenteRaw = String(row.get('Visitas Pendente') || '0').trim();
          const visitasPendenteNum = parseInt(visitasPendenteRaw) || 0;
          
          const item = {
            segmento: String(row.get('Segmento') || '').trim(),
            tipoLoja: String(row.get('Tipo Loja') || '').trim(),
            empresa: String(row.get('Empresa') || '').trim(),
            carga: String(row.get('Carga') || '').trim(),
            descricao: String(row.get('Descrição') || '').trim(),
            volume: String(row.get('Volume') || '').trim(),
            valor: String(row.get('Valor') || '').trim(),
            dataRot: String(row.get('Data Rot') || '').trim(),
            statusSep: String(row.get('Separação st') || '').trim(),
            statusConf: String(row.get('conf. St') || '').trim(),
            loja: String(row.get('Loja') || '').trim(),
            ton: String(row.get('Ton') || '').trim(),
            m3: m3Value,
            visitasPendente: visitasPendenteNum,
            prioridade: String(row.get('Prioridade') || '').trim(),
            box: String(row.get('BOX') || '').trim()
          };
          
          if (!item.carga) continue;
          
          let incluir = true;
          
          if (filtros.loja) {
            const lojaFiltro = filtros.loja.toLowerCase();
            incluir = incluir && (
              item.loja.toLowerCase().includes(lojaFiltro) ||
              item.descricao.toLowerCase().includes(lojaFiltro)
            );
          }
          
          if (filtros.tipoLoja) {
            incluir = incluir && item.tipoLoja === filtros.tipoLoja;
          }
          
          if (filtros.segmento) {
            incluir = incluir && item.segmento === filtros.segmento;
          }
          
          if (filtros.dataRot) {
            incluir = incluir && item.dataRot === filtros.dataRot;
          }
          
          if (filtros.statusSep) {
            incluir = incluir && item.statusSep === filtros.statusSep;
          }
          
          if (incluir) {
            dados.push(item);
          }
          
        } catch (rowError) {
          console.error('[SHEETS] Erro ao processar linha:', rowError);
          continue;
        }
      }
  
      console.log('[SHEETS] Registros processados:', dados.length);
      
      if (dados.length > 0) {
        console.log('[SHEETS] Exemplo de registro:', dados[0]);
      }
      
      return dados;
      
    } catch (error) {
      console.error('[SHEETS] Erro em getMapaCarga:', error);
      console.error('[SHEETS] Stack:', error.stack);
      throw error;
    }
  }
  
  async atualizarMapaCarga(carga, campos) {
    try {
      console.log('[SHEETS] Iniciando atualizarMapaCarga...');
      console.log('[SHEETS] Carga:', carga);
      console.log('[SHEETS] Campos:', campos);
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        console.error('[SHEETS] Aba "Mapa de Carga" não encontrada');
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      const rows = await sheet.getRows();
      console.log('[SHEETS] Buscando carga em', rows.length, 'linhas');
      
      for (const row of rows) {
        const cargaRow = String(row.get('Carga') || '').trim();
        
        if (cargaRow === String(carga)) {
          console.log('[SHEETS] Carga encontrada, atualizando...');
          
          for (const [campo, valor] of Object.entries(campos)) {
            row.set(campo, valor);
            console.log(`[SHEETS] ${campo} = ${valor}`);
          }
          
          await row.save();
          console.log('[SHEETS] Carga atualizada com sucesso');
          
          return { ok: true, msg: 'Carga atualizada com sucesso' };
        }
      }
      
      console.log('[SHEETS] Carga não encontrada');
      return { ok: false, msg: 'Carga não encontrada' };
      
    } catch (error) {
      console.error('[SHEETS] Erro em atualizarMapaCarga:', error);
      return { ok: false, msg: error.message };
    }
  }
  
  // ==================== FUNÇÕES ALOCAÇÃO BOX ====================

  async getCargasSemBox(filtros = {}) {
    try {
      console.log('[SHEETS] Iniciando getCargasSemBox...');
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const cargas = [];

      for (const row of rows) {
        try {
          const segmento = String(row.get('Segmento') || '').trim();
          const tipoLoja = String(row.get('Tipo Loja') || '').trim();
          const box = String(row.get('BOX') || '').trim();
          const carga = String(row.get('Carga') || '').trim();
          
          if (!carga) continue;
          
          if (!box && (segmento === 'LJ COMPER' || segmento === 'LJ FORT')) {
            if (tipoLoja !== 'LJ COMPER' && tipoLoja !== 'LJ FORT') continue;
            
            const m3Raw = String(row.get('M³') || '0').trim();
            const m3Value = parseFloat(m3Raw.replace(',', '.')) || 0;
            
            const visitasPendenteRaw = String(row.get('Visitas Pendente') || '0').trim();
            const visitasPendenteNum = parseInt(visitasPendenteRaw) || 0;
            
            const item = {
              carga: carga,
              descricao: String(row.get('Descrição') || '').trim(),
              tipoLoja: tipoLoja,
              segmento: segmento,
              loja: String(row.get('Loja') || '').trim(),
              m3: m3Value,
              dataRot: String(row.get('Data Rot') || '').trim(),
              valor: String(row.get('Valor') || '').trim(),
              statusSep: String(row.get('Separação st') || '').trim(),
              statusConf: String(row.get('conf. St') || '').trim(),
              visitasPendente: visitasPendenteNum,
              prioridade: String(row.get('Prioridade') || '').trim()
            };
            
            let incluir = true;
            
            if (filtros.loja) {
              const lojaFiltro = filtros.loja.toLowerCase();
              incluir = incluir && item.loja.toLowerCase().includes(lojaFiltro);
            }
            
            if (filtros.tipoLoja) {
              incluir = incluir && item.tipoLoja === filtros.tipoLoja;
            }
            
            if (filtros.segmento) {
              incluir = incluir && item.segmento === filtros.segmento;
            }
            
            if (filtros.dataRot) {
              incluir = incluir && item.dataRot === filtros.dataRot;
            }
            
            if (incluir) {
              cargas.push(item);
            }
          }
          
        } catch (rowError) {
          console.error('[SHEETS] Erro ao processar linha:', rowError);
          continue;
        }
      }

      console.log('[SHEETS] Cargas sem BOX encontradas:', cargas.length);
      return cargas;
      
    } catch (error) {
      console.error('[SHEETS] Erro em getCargasSemBox:', error);
      throw error;
    }
  }

  async getEstadoBoxes() {
    try {
      console.log('[SHEETS] Iniciando getEstadoBoxes...');
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const boxesOcupados = [];
  
      for (const row of rows) {
        try {
          const box = String(row.get('BOX') || '').trim();
          const segmento = String(row.get('Segmento') || '').trim();
          const tipoLoja = String(row.get('Tipo Loja') || '').trim();
          
          if (box && (segmento === 'LJ COMPER' || segmento === 'LJ FORT')) {
            const m3Raw = String(row.get('M³') || '0').trim();
            const m3Value = parseFloat(m3Raw.replace(',', '.')) || 0;
            
            const visitasPendenteRaw = String(row.get('Visitas Pendente') || '0').trim();
            const visitasPendenteNum = parseInt(visitasPendenteRaw) || 0;
            
            boxesOcupados.push({
              box: box,
              carga: String(row.get('Carga') || '').trim(),
              descricao: String(row.get('Descrição') || '').trim(),
              loja: String(row.get('Loja') || '').trim(),
              m3: m3Value,
              dataRot: String(row.get('Data Rot') || '').trim(),
              tipoLoja: tipoLoja,
              segmento: segmento,
              visitasPendente: visitasPendenteNum,
              prioridade: String(row.get('Prioridade') || '').trim()
            });
          }
          
        } catch (rowError) {
          console.error('[SHEETS] Erro ao processar linha:', rowError);
          continue;
        }
      }

      console.log('[SHEETS] Boxes ocupados encontrados:', boxesOcupados.length);
      return boxesOcupados;
      
    } catch (error) {
      console.error('[SHEETS] Erro em getEstadoBoxes:', error);
      throw error;
    }
  }

  async alocarCargaBox(boxNum, cargaId) {
    try {
      console.log('[SHEETS] Alocando carga:', cargaId, 'no BOX:', boxNum);
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const carga = String(row.get('Carga') || '').trim();
        
        if (carga === String(cargaId)) {
          row.set('BOX', String(boxNum));
          await row.save();
          
          console.log('[SHEETS] BOX alocado com sucesso');
          return { ok: true, msg: 'BOX alocado com sucesso' };
        }
      }
      
      return { ok: false, msg: 'Carga não encontrada' };
      
    } catch (error) {
      console.error('[SHEETS] Erro ao alocar BOX:', error);
      return { ok: false, msg: error.message };
    }
  }

  async liberarBox(boxNum) {
    try {
      console.log('[SHEETS] Liberando BOX:', boxNum);
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      const rows = await sheet.getRows();
      let encontrado = false;
      
      for (const row of rows) {
        const box = String(row.get('BOX') || '').trim();
        
        if (box === String(boxNum)) {
          row.set('BOX', '');
          await row.save();
          encontrado = true;
          console.log('[SHEETS] BOX liberado:', boxNum);
        }
      }
      
      if (encontrado) {
        return { ok: true, msg: 'BOX liberado com sucesso' };
      } else {
        return { ok: false, msg: 'BOX não encontrado ou já estava vazio' };
      }
      
    } catch (error) {
      console.error('[SHEETS] Erro ao liberar BOX:', error);
      return { ok: false, msg: error.message };
    }
  }

  async salvarAlocacoes(dados) {
    try {
      console.log('[SHEETS] Salvando múltiplas alocações:', dados.length);
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      const rows = await sheet.getRows();
      let totalSalvos = 0;
      
      for (const alocacao of dados) {
        const { carga, box } = alocacao;
        
        for (const row of rows) {
          const cargaRow = String(row.get('Carga') || '').trim();
          
          if (cargaRow === String(carga)) {
            row.set('BOX', String(box));
            await row.save();
            totalSalvos++;
            break;
          }
        }
      }
      
      console.log('[SHEETS] Alocações salvas:', totalSalvos);
      
      return { 
        ok: true, 
        msg: `${totalSalvos} alocações salvas com sucesso`,
        total: totalSalvos
      };
      
    } catch (error) {
      console.error('[SHEETS] Erro ao salvar alocações:', error);
      return { ok: false, msg: error.message };
    }
  }

  // ==================== FUNÇÃO SALVAR NA BASE COM DESVIO ====================
  
  async salvarNaBase(dados) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) return { ok: false, msg: 'Aba Base não encontrada' };
      
      const hoje = new Date().toLocaleDateString('pt-BR');
      
      console.log('[SHEETS] Iniciando salvamento na Base...');
      console.log('[SHEETS] Total de linhas para processar:', dados.length);
      
      const rowsExistentes = await sheet.getRows();
      console.log('[SHEETS] Linhas existentes na Base:', rowsExistentes.length);
      
      let totalAtualizados = 0;
      let totalNovos = 0;
      
      for (const linha of dados) {
        const [sup, aba, matricula, nome, funcao, status, desvio] = linha;
        
        if (!matricula && !nome) continue;
        
        let registroExistente = null;
        
        for (const row of rowsExistentes) {
          const rowSup = String(row.get('Supervisor') || '').trim();
          const rowData = String(row.get('Data') || '').trim();
          const rowMat = String(row.get('Matricula') || '').trim();
          
          if (rowSup === sup && rowData === hoje && rowMat === matricula) {
            registroExistente = row;
            break;
          }
        }
        
        if (registroExistente) {
          console.log(`[SHEETS] Atualizando registro: ${nome} (${matricula})`);
          registroExistente.set('Status', status);
          registroExistente.set('Desvio', desvio || '');
          await registroExistente.save();
          totalAtualizados++;
        } else {
          console.log(`[SHEETS] Inserindo novo registro: ${nome} (${matricula})`);
          await sheet.addRow({
            'Supervisor': sup,
            'Aba': aba,
            'Matricula': matricula,
            'Nome': nome,
            'Função': funcao,
            'Status': status,
            'Data': hoje,
            'Desvio': desvio || ''
          });
          totalNovos++;
        }
      }
      
      console.log('[SHEETS] Salvamento concluído!');
      console.log(`[SHEETS] ${totalNovos} novos registros inseridos`);
      console.log(`[SHEETS] ${totalAtualizados} registros atualizados`);
      
      return { 
        ok: true, 
        msg: `Dados salvos com sucesso! ${totalNovos} novos, ${totalAtualizados} atualizados.`,
        totais: {
          novos: totalNovos,
          atualizados: totalAtualizados,
          total: totalNovos + totalAtualizados
        }
      };
    } catch (error) {
      console.error('[SHEETS] Erro ao salvar:', error);
      return { ok: false, msg: error.message };
    }
  }
}

module.exports = new SheetsService();
