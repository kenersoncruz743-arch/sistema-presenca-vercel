// lib/sheets.js - VERSÃO COMPLETA COM BATCH UPDATE OTIMIZADO
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

  async adicionarBuffer(supervisor, aba, colaborador) {
    try {
      console.log('[SHEETS] Adicionando ao buffer:', { supervisor, aba, colaborador });
      
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        console.error('[SHEETS] Aba Lista não encontrada');
        return { ok: false, msg: 'Aba Lista não encontrada' };
      }
      
      const rows = await sheet.getRows();
      const jaExiste = rows.some(row => {
        const rowSup = String(row.get('Supervisor') || '').trim();
        const rowGrupo = String(row.get('Grupo') || '').trim();
        const rowMat = String(row.get('matricula') || '').trim();
        
        return rowSup === supervisor && rowGrupo === aba && rowMat === String(colaborador.matricula);
      });
      
      if (jaExiste) {
        console.log('[SHEETS] Colaborador já existe no buffer');
        return { ok: true, msg: 'Colaborador já está na lista' };
      }
      
      await sheet.addRow({
        'Supervisor': supervisor,
        'Grupo': aba,
        'matricula': String(colaborador.matricula),
        'Nome': colaborador.nome,
        'Função': colaborador.funcao,
        'status': '',
        'desvio': ''
      });

      console.log('[SHEETS] ✓ Colaborador adicionado ao buffer com sucesso');
      return { ok: true };
    } catch (error) {
      console.error('[SHEETS] Erro ao adicionar:', error);
      return { ok: false, msg: error.message };
    }
  }

  async getBuffer(supervisor, aba) {
    try {
      console.log('[SHEETS] Buscando buffer:', { supervisor, aba });
      
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        console.error('[SHEETS] Aba Lista não encontrada');
        return [];
      }
      
      const rows = await sheet.getRows();
      const buffer = [];

      for (const row of rows) {
        const sup = String(row.get('Supervisor') || '').trim();
        const rowGrupo = String(row.get('Grupo') || '').trim();
        
        if (sup === supervisor && rowGrupo === aba) {
          buffer.push({
            matricula: String(row.get('matricula') || '').trim(),
            nome: String(row.get('Nome') || '').trim(),
            funcao: String(row.get('Função') || '').trim(),
            status: String(row.get('status') || '').trim(),
            desvio: String(row.get('desvio') || '').trim()
          });
        }
      }

      console.log(`[SHEETS] ✓ Buffer carregado: ${buffer.length} colaboradores`);
      return buffer;
    } catch (error) {
      console.error('[SHEETS] Erro ao buscar buffer:', error);
      return [];
    }
  }

  async removerBuffer(supervisor, matricula) {
    try {
      console.log('[SHEETS] Removendo do buffer:', { supervisor, matricula });
      
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        console.error('[SHEETS] Aba Lista não encontrada');
        return { ok: false };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const rowSup = String(row.get('Supervisor') || '').trim();
        const rowMat = String(row.get('matricula') || '').trim();
        
        if (rowSup === supervisor && rowMat === String(matricula)) {
          await row.delete();
          console.log(`[SHEETS] ✓ Colaborador ${matricula} removido`);
          return { ok: true };
        }
      }
      
      console.log(`[SHEETS] ✗ Colaborador ${matricula} não encontrado`);
      return { ok: false };
    } catch (error) {
      console.error('[SHEETS] Erro ao remover:', error);
      return { ok: false };
    }
  }

  async atualizarStatusBuffer(supervisor, matricula, status) {
    try {
      console.log('[SHEETS] Atualizando status:', { supervisor, matricula, status });
      
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        console.error('[SHEETS] Aba Lista não encontrada');
        return { ok: false };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const rowSup = String(row.get('Supervisor') || '').trim();
        const rowMat = String(row.get('matricula') || '').trim();
        
        if (rowSup === supervisor && rowMat === String(matricula)) {
          row.set('status', status);
          await row.save();
          console.log(`[SHEETS] ✓ Status atualizado: ${status}`);
          return { ok: true };
        }
      }
      
      console.log(`[SHEETS] ✗ Colaborador ${matricula} não encontrado`);
      return { ok: false };
    } catch (error) {
      console.error('[SHEETS] Erro ao atualizar status:', error);
      return { ok: false };
    }
  }

  // ==================== MAPA DE CARGA - OTIMIZADO COM BATCH UPDATE ====================
    
  async limparColunasMapaCarga() {
    try {
      console.log('[SHEETS] Limpando colunas via BATCH UPDATE...');
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      console.log('[SHEETS] Headers encontrados:', headers);
      
      // Encontra índices das colunas
      const colunasParaLimpar = ['Empresa', 'SM', 'Deposito', 'BOX', 'Carga', 'Coluna 1', 'Descrição', 'sp', 'Ton', 'M³', 'Valor', 'Rup', 'Visita Picking', 'Volume', 'Coluna 2', 'inclusão', 'Roteirização', 'Geração Mesa', '"', 'Reposição', 'Palete_Box', 'Baixa', 'Separação', 'Final separação', 'Conferencia', 'seotr'];
      
      const indicesColunas = colunasParaLimpar
        .map(col => headers.indexOf(col))
        .filter(idx => idx !== -1);
      
      if (indicesColunas.length === 0) {
        return { ok: false, msg: 'Nenhuma coluna encontrada para limpar' };
      }
      
      console.log(`[SHEETS] Limpando ${indicesColunas.length} colunas`);
      
      // Carrega células
      await sheet.loadCells();
      
      const totalRows = sheet.rowCount;
      let linhasLimpas = 0;
      
      // Limpa células em batch (ignora header row 0)
      for (let row = 1; row < totalRows; row++) {
        for (const colIndex of indicesColunas) {
          const cell = sheet.getCell(row, colIndex);
          if (cell) {
            cell.value = '';
          }
        }
        linhasLimpas++;
        
        // Salva em lotes de 100 linhas
        if (linhasLimpas % 100 === 0) {
          await sheet.saveUpdatedCells();
          console.log(`[SHEETS] ${linhasLimpas}/${totalRows} linhas limpas`);
        }
      }
      
      // Salva células restantes
      await sheet.saveUpdatedCells();
      
      console.log('[SHEETS] ✓ Colunas limpas com sucesso');
      
      return { 
        ok: true, 
        msg: `${linhasLimpas} linhas limpas com sucesso!`,
        total: linhasLimpas
      };
      
    } catch (error) {
      console.error('[SHEETS] Erro ao limpar:', error);
      return { 
        ok: false, 
        msg: 'Erro ao limpar colunas: ' + error.message 
      };
    }
  }
    
  async processarMapaCargaColado(dadosColados) {
    try {
      console.log('[SHEETS] Processando dados colados...');
      console.log('[SHEETS] Total de linhas:', dadosColados.length);
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      console.log('[SHEETS] Headers:', headers);
      
      // Carrega linhas existentes
      const rows = await sheet.getRows();
      console.log(`[SHEETS] Linhas existentes: ${rows.length}`);
      
      // Se não tem linhas, adiciona novas
      if (rows.length === 0) {
        console.log('[SHEETS] Criando novas linhas...');
        
        const novasLinhas = [];
        
        for (const linha of dadosColados) {
          const campos = Array.isArray(linha) ? linha : String(linha).split('\t');
          
          if (campos.length < 10) continue;
          
          const empresa = String(campos[0] || '').trim();
          const carga = String(campos[4] || '').trim();
          const descricao = String(campos[6] || '').trim();
          
          // Detecta segmento e tipo de loja
          let segmento = '';
          let tipoLoja = '';
          
          if (descricao.toUpperCase().includes('COMPER')) {
            segmento = 'LJ COMPER';
            tipoLoja = 'LJ COMPER';
          } else if (descricao.toUpperCase().includes('FORT')) {
            segmento = 'LJ FORT';
            tipoLoja = 'LJ FORT';
          }
          
          // Extrai loja
          const matchLoja = descricao.match(/LJ\s+(\d+)/i);
          const loja = matchLoja ? `LJ ${matchLoja[1]}` : '';
          
          novasLinhas.push({
            'Segmento': segmento,
            'Tipo Loja': tipoLoja,
            'Empresa': empresa,
            'Carga': carga,
            'Descrição': descricao,
            'Ton': String(campos[7] || '').trim(),
            'Volume': String(campos[8] || '').trim(),
            'M³': String(campos[8] || '').trim(),
            'Valor': String(campos[9] || '').trim(),
            'Visitas Pendente': String(campos[12] || '').trim(),
            'Data Rot': String(campos[14] || '').trim(),
            'Separação st': String(campos[18] || '').trim(),
            'conf. St': String(campos[19] || '').trim(),
            'Loja': loja,
            'Prioridade': '',
            'BOX': ''
          });
        }
        
        console.log(`[SHEETS] Adicionando ${novasLinhas.length} novas linhas...`);
        
        // Adiciona em lotes
        const lote = 50;
        for (let i = 0; i < novasLinhas.length; i += lote) {
          const grupo = novasLinhas.slice(i, i + lote);
          await sheet.addRows(grupo);
          console.log(`[SHEETS] ${i + grupo.length}/${novasLinhas.length} adicionadas`);
        }
        
        return { 
          ok: true, 
          msg: `${novasLinhas.length} cargas adicionadas!`,
          total: novasLinhas.length
        };
      }
      
      // Se tem linhas, atualiza as existentes em BATCH
      console.log('[SHEETS] Atualizando linhas existentes via BATCH...');
      
      await sheet.loadCells();
      
      const headerMap = {};
      headers.forEach((h, idx) => { headerMap[h] = idx; });
      
      let linhasAtualizadas = 0;
      
      for (let i = 0; i < Math.min(rows.length, dadosColados.length); i++) {
        const linha = dadosColados[i];
        const campos = Array.isArray(linha) ? linha : String(linha).split('\t');
        
        if (campos.length < 10) continue;
        
        const rowIndex = i + 1; // +1 porque row 0 é header
        
        const empresa = String(campos[0] || '').trim();
        const carga = String(campos[4] || '').trim();
        const descricao = String(campos[6] || '').trim();
        
        // Detecta segmento e tipo de loja
        let segmento = '';
        let tipoLoja = '';
        
        if (descricao.toUpperCase().includes('COMPER')) {
          segmento = 'LJ COMPER';
          tipoLoja = 'LJ COMPER';
        } else if (descricao.toUpperCase().includes('FORT')) {
          segmento = 'LJ FORT';
          tipoLoja = 'LJ FORT';
        }
        
        // Extrai loja
        const matchLoja = descricao.match(/LJ\s+(\d+)/i);
        const loja = matchLoja ? `LJ ${matchLoja[1]}` : '';
        
        // Atualiza células
        if (headerMap['Empresa'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Empresa']).value = empresa;
        }
        if (headerMap['Carga'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Carga']).value = carga;
        }
        if (headerMap['Descrição'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Descrição']).value = descricao;
        }
        if (headerMap['Ton'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Ton']).value = String(campos[7] || '').trim();
        }
        if (headerMap['Volume'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Volume']).value = String(campos[8] || '').trim();
        }
        if (headerMap['M³'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['M³']).value = String(campos[8] || '').trim();
        }
        if (headerMap['Valor'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Valor']).value = String(campos[9] || '').trim();
        }
        if (headerMap['Visitas Pendente'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Visitas Pendente']).value = String(campos[12] || '').trim();
        }
        if (headerMap['Data Rot'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Data Rot']).value = String(campos[14] || '').trim();
        }
        if (headerMap['Separação st'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Separação st']).value = String(campos[18] || '').trim();
        }
        if (headerMap['conf. St'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['conf. St']).value = String(campos[19] || '').trim();
        }
        if (headerMap['Loja'] !== undefined) {
          sheet.getCell(rowIndex, headerMap['Loja']).value = loja;
        }
        
        linhasAtualizadas++;
        
        // Salva em lotes de 50
        if (linhasAtualizadas % 50 === 0) {
          await sheet.saveUpdatedCells();
          console.log(`[SHEETS] ${linhasAtualizadas}/${rows.length} atualizadas`);
        }
      }
      
      // Salva células restantes
      await sheet.saveUpdatedCells();
      
      console.log('[SHEETS] ✓ Processamento concluído');
      
      return { 
        ok: true, 
        msg: `${linhasAtualizadas} cargas atualizadas com sucesso!`,
        total: linhasAtualizadas
      };
      
    } catch (error) {
      console.error('[SHEETS] Erro no processamento:', error);
      console.error('[SHEETS] Stack:', error.stack);
      return { 
        ok: false, 
        msg: 'Erro ao processar: ' + error.message 
      };
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
        return [];
      }
      
      const rows = await sheet.getRows();
      console.log('[SHEETS] Total de linhas:', rows.length);
      
      const dados = [];
  
      for (const row of rows) {
        try {
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
            visitasPendente: String(row.get('Visitas Pendente') || '').trim(),
            prioridade: String(row.get('Prioridade') || '').trim(),
            box: String(row.get('BOX') || '').trim()
          };
          
          if (!item.carga) continue;
          
          // Aplica filtros
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
      return dados;
      
    } catch (error) {
      console.error('[SHEETS] Erro em getMapaCarga:', error);
      throw error;
    }
  }
  
  async atualizarMapaCarga(carga, campos) {
    try {
      console.log('[SHEETS] Atualizando Mapa de Carga:', { carga, campos });
      
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const cargaRow = String(row.get('Carga') || '').trim();
        
        if (cargaRow === String(carga)) {
          for (const [campo, valor] of Object.entries(campos)) {
            row.set(campo, valor);
          }
          
          await row.save();
          console.log('[SHEETS] Carga atualizada com sucesso');
          
          return { ok: true, msg: 'Carga atualizada' };
        }
      }
      
      return { ok: false, msg: 'Carga não encontrada' };
      
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
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
              visitasPendente: String(row.get('Visitas Pendente') || '').trim(),
              prioridade: String(row.get('Prioridade') || '').trim()
            };
            
            // Aplica filtros
            let incluir = true;
            
            if (filtros.loja) {
              incluir = incluir && item.loja.toLowerCase().includes(filtros.loja.toLowerCase());
            }
            
            if (filtros.tipoLoja) {
              incluir = incluir && item.tipoLoja === filtros.tipoLoja;
            }
            
            if (filtros.segmento) {
              incluir = incluir && item.segmento === filtros.segmento;
            }
            
            if (incluir) {
              cargas.push(item);
            }
          }
          
        } catch (rowError) {
          continue;
        }
      }
  
      return cargas;
      
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
      throw error;
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
        try {
          const box = String(row.get('BOX') || '').trim();
          const segmento = String(row.get('Segmento') || '').trim();
          const tipoLoja = String(row.get('Tipo Loja') || '').trim();
          
          if (box && (segmento === 'LJ COMPER' || segmento === 'LJ FORT')) {
            const m3Raw = String(row.get('M³') || '0').trim();
            const m3Value = parseFloat(m3Raw.replace(',', '.')) || 0;
            
            boxesOcupados.push({
              box: box,
              carga: String(row.get('Carga') || '').trim(),
              descricao: String(row.get('Descrição') || '').trim(),
              loja: String(row.get('Loja') || '').trim(),
              m3: m3Value,
              dataRot: String(row.get('Data Rot') || '').trim(),
              tipoLoja: tipoLoja,
              segmento: segmento,
              visitasPendente: String(row.get('Visitas Pendente') || '').trim(),
              prioridade: String(row.get('Prioridade') || '').trim()
            });
          }
          
        } catch (rowError) {
          continue;
        }
      }
  
      return boxesOcupados;
      
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
      throw error;
    }
  }

  async alocarCargaBox(boxNum, cargaId) {
    try {
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const cargaRow = String(row.get('Carga') || '').trim();
        
        if (cargaRow === String(cargaId)) {
          row.set('BOX', String(boxNum));
          await row.save();
          
          console.log(`[SHEETS] BOX ${boxNum} alocado para carga ${cargaId}`);
          return { ok: true, msg: `BOX ${boxNum} alocado` };
        }
      }
      
      return { ok: false, msg: 'Carga não encontrada' };
      
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
      return { ok: false, msg: error.message };
    }
  }
  
  async liberarBox(boxNum) {
    try {
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const boxRow = String(row.get('BOX') || '').trim();
        
        if (boxRow === String(boxNum)) {
          row.set('BOX', '');
          await row.save();
          
          console.log(`[SHEETS] BOX ${boxNum} liberado`);
          return { ok: true, msg: `BOX ${boxNum} liberado` };
        }
      }
      
      return { ok: false, msg: 'BOX não encontrado' };
      
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
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
      
      const rowsExistentes = await sheet.getRows();
      
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
          registroExistente.set('Status', status);
          if (desvio) registroExistente.set('Desvio', desvio);
          await registroExistente.save();
          totalAtualizados++;
        } else {
          await sheet.addRow({
            'Supervisor': sup,
            'Aba': aba,
            'Matricula': matricula,
            'Nome': nome,
            'Função': funcao,
            'Status': status,
            'Desvio': desvio || '',
            'Data': hoje
          });
          totalNovos++;
        }
      }
      
      return { 
        ok: true, 
        msg: `${totalNovos} novos, ${totalAtualizados} atualizados`,
        totais: { novos: totalNovos, atualizados: totalAtualizados }
      };
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
      return { ok: false, msg: error.message };
    }
  }

  // ==================== FUNÇÕES COM BUSCA POR ABA ====================

  async removerBufferPorAba(chave, matricula) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Lista não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const grupo = String(row.get('Grupo') || '').trim();
        const mat = String(row.get('matricula') || '').trim();
        
        if ((supervisor === chave || grupo === chave) && mat === String(matricula)) {
          await row.delete();
          return { ok: true };
        }
      }
      
      return { ok: false, msg: 'Colaborador não encontrado' };
    } catch (error) {
      return { ok: false, msg: error.message };
    }
  }

  async atualizarStatusBufferPorAba(chave, matricula, status) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Lista não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const grupo = String(row.get('Grupo') || '').trim();
        const mat = String(row.get('matricula') || '').trim();
        
        if ((supervisor === chave || grupo === chave) && mat === String(matricula)) {
          row.set('status', status);
          await row.save();
          return { ok: true };
        }
      }
      
      return { ok: false, msg: 'Colaborador não encontrado' };
    } catch (error) {
      return { ok: false, msg: error.message };
    }
  }

  async atualizarDesvioBufferPorAba(chave, matricula, desvio) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Lista'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Lista não encontrada' };
      }
      
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        const supervisor = String(row.get('Supervisor') || '').trim();
        const grupo = String(row.get('Grupo') || '').trim();
        const mat = String(row.get('matricula') || '').trim();
        
        if ((supervisor === chave || grupo === chave) && mat === String(matricula)) {
          row.set('desvio', desvio);
          await row.save();
          return { ok: true };
        }
      }
      
      return { ok: false, msg: 'Colaborador não encontrado' };
    } catch (error) {
      return { ok: false, msg: error.message };
    }
  }

}

module.exports = new SheetsService();
