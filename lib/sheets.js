// lib/sheets.js - VERSÃO COMPLETA COM TODAS AS FUNÇÕES
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
            status: row.get('status') || ''
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
        'status': ''
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

  // ==================== FUNÇÕES MAPA DE CARGA ====================
  
  async getMapaCarga(filtros = {}) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const dados = [];

      for (const row of rows) {
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
          m3: String(row.get('M³') || '').trim(),
          rup: String(row.get('Rup') || '').trim(),
          prioridade: String(row.get('Prioridade') || '').trim(),
          box: String(row.get('BOX') || '').trim() // Campo para alocação de BOX
        };
        
        // Aplica filtros
        if (filtros.loja && !item.loja.toLowerCase().includes(filtros.loja.toLowerCase())) {
          continue;
        }
        if (filtros.tipoLoja && item.tipoLoja !== filtros.tipoLoja) {
          continue;
        }
        if (filtros.dataRot && item.dataRot !== filtros.dataRot) {
          continue;
        }
        
        dados.push(item);
      }

      return dados;
    } catch (error) {
      console.error('Erro ao buscar Mapa de Carga:', error);
      return [];
    }
  }

  async atualizarMapaCarga(carga, campos) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (String(row.get('Carga') || '').trim() === String(carga)) {
          // Atualiza os campos fornecidos
          for (const [campo, valor] of Object.entries(campos)) {
            row.set(campo, valor);
          }
          await row.save();
          return { ok: true, msg: 'Carga atualizada com sucesso' };
        }
      }
      
      return { ok: false, msg: 'Carga não encontrada' };
    } catch (error) {
      console.error('Erro ao atualizar Mapa de Carga:', error);
      return { ok: false, msg: error.message };
    }
  }

  // ==================== FUNÇÕES ALOCAÇÃO BOX ====================
  
  async getCargasSemBox(filtros = {}) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const cargas = [];

      for (const row of rows) {
        const box = String(row.get('BOX') || '').trim();
        const tipoLoja = String(row.get('Tipo Loja') || '').trim();
        
        // Filtra apenas cargas sem BOX e dos tipos corretos
        if (!box && (tipoLoja === 'LJ COMPER' || tipoLoja === 'LJ FORT')) {
          const item = {
            carga: String(row.get('Carga') || '').trim(),
            descricao: String(row.get('Descrição') || '').trim(),
            tipoLoja: tipoLoja,
            loja: String(row.get('Loja') || '').trim(),
            m3: parseFloat(String(row.get('M³') || '0').replace(',', '.')),
            dataRot: String(row.get('Data Rot') || '').trim(),
            valor: String(row.get('Valor') || '').trim(),
            segmento: String(row.get('Segmento') || '').trim()
          };
          
          // Aplica filtros adicionais
          if (filtros.loja && !item.loja.toLowerCase().includes(filtros.loja.toLowerCase())) {
            continue;
          }
          if (filtros.tipoLoja && item.tipoLoja !== filtros.tipoLoja) {
            continue;
          }
          
          cargas.push(item);
        }
      }

      return cargas;
    } catch (error) {
      console.error('Erro ao buscar cargas sem BOX:', error);
      return [];
    }
  }

  async getEstadoBoxes() {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return [];
      
      const rows = await sheet.getRows();
      const boxesOcupados = [];

      for (const row of rows) {
        const box = String(row.get('BOX') || '').trim();
        
        if (box) {
          boxesOcupados.push({
            box: box,
            carga: String(row.get('Carga') || '').trim(),
            descricao: String(row.get('Descrição') || '').trim(),
            loja: String(row.get('Loja') || '').trim(),
            m3: parseFloat(String(row.get('M³') || '0').replace(',', '.')),
            dataRot: String(row.get('Data Rot') || '').trim()
          });
        }
      }

      return boxesOcupados;
    } catch (error) {
      console.error('Erro ao buscar estado dos boxes:', error);
      return [];
    }
  }

  async alocarCargaBox(boxNum, cargaId) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (String(row.get('Carga') || '').trim() === String(cargaId)) {
          row.set('BOX', String(boxNum));
          await row.save();
          return { ok: true, msg: `Carga ${cargaId} alocada no BOX ${boxNum}` };
        }
      }
      
      return { ok: false, msg: 'Carga não encontrada' };
    } catch (error) {
      console.error('Erro ao alocar carga:', error);
      return { ok: false, msg: error.message };
    }
  }

  async liberarBox(boxNum) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      
      const rows = await sheet.getRows();
      let liberadas = 0;
      
      for (const row of rows) {
        if (String(row.get('BOX') || '').trim() === String(boxNum)) {
          row.set('BOX', '');
          await row.save();
          liberadas++;
        }
      }
      
      return { 
        ok: true, 
        msg: `BOX ${boxNum} liberado (${liberadas} carga(s) desalocada(s))` 
      };
    } catch (error) {
      console.error('Erro ao liberar BOX:', error);
      return { ok: false, msg: error.message };
    }
  }

  async salvarAlocacoes(dados) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      
      const rows = await sheet.getRows();
      let alocadas = 0;
      
      for (const alocacao of dados) {
        const { cargaId, boxNum } = alocacao;
        
        for (const row of rows) {
          if (String(row.get('Carga') || '').trim() === String(cargaId)) {
            row.set('BOX', String(boxNum));
            await row.save();
            alocadas++;
            break;
          }
        }
      }
      
      return { 
        ok: true, 
        msg: `${alocadas} alocação(ões) salva(s) com sucesso` 
      };
    } catch (error) {
      console.error('Erro ao salvar alocações:', error);
      return { ok: false, msg: error.message };
    }
  }

  // ==================== FUNÇÃO SALVAR NA BASE ====================
  
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
        const [sup, aba, matricula, nome, funcao, status] = linha;
        
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
          console.log(`[SHEETS] Atualizando status do registro: ${nome} (${matricula})`);
          registroExistente.set('Status', status);
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
            'Data': hoje
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
