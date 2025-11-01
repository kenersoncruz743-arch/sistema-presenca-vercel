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

  // ========== FUNÇÕES MAPA DE CARGA ==========

  async getMapaCarga(filtros = {}) {
    try {
      console.log('[SHEETS] Carregando Mapa de Carga...');
      await this.init();
      
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        console.error('[SHEETS] Aba "Mapa de Carga" não encontrada');
        return [];
      }

      const rows = await sheet.getRows();
      console.log(`[SHEETS] ${rows.length} linhas na aba Mapa de Carga`);

      const dados = [];

      rows.forEach(row => {
        const carga = String(row.get('Carga') || '').trim();
        if (!carga) return;

        const item = {
          empresa: String(row.get('Empresa') || '').trim(),
          sm: String(row.get('SM') || '').trim(),
          deposito: String(row.get('Deposito') || '').trim(),
          box: String(row.get('BOX') || '').trim(),
          carga: carga,
          descricao: String(row.get('Descrição') || '').trim(),
          ton: String(row.get('Ton') || '0').trim(),
          m3: parseFloat(String(row.get('M³') || row.get('Volume') || '0').replace(',', '.')) || 0,
          valor: String(row.get('Valor') || '0').trim(),
          rup: String(row.get('Rup') || '').trim(),
          visitasPendente: parseInt(String(row.get('Visita Picking') || row.get('Visitas Pendente') || '0')) || 0,
          inclusao: String(row.get('inclusão') || '').trim(),
          roteirizacao: String(row.get('Roteirização') || '').trim(),
          dataRot: String(row.get('Data Rot') || '').trim(),
          geracaoMesa: String(row.get('Geração Mesa') || '').trim(),
          reposicao: String(row.get('Reposição') || '').trim(),
          paleteBox: String(row.get('Palete_Box') || '').trim(),
          baixa: String(row.get('Baixa') || '').trim(),
          statusSep: String(row.get('Separação') || row.get('Separação st') || '').trim(),
          finalSeparacao: String(row.get('Final separação') || '').trim(),
          conferencia: String(row.get('Conferencia') || '').trim(),
          statusConf: String(row.get('conf. St') || '').trim(),
          loja: String(row.get('Loja') || '').trim(),
          diaOferta: String(row.get('Dia oferta') || '').trim(),
          prioridade: String(row.get('Prioridade') || '').trim(),
          totalVertical: String(row.get('Total_Vertical') || '').trim(),
          segmento: String(row.get('Segmento') || '').trim(),
          tipoLoja: String(row.get('Tipo Loja') || '').trim(),
          conjugada: String(row.get('Conjugada') || '').trim()
        };

        dados.push(item);
      });

      console.log(`[SHEETS] ✓ ${dados.length} cargas processadas`);
      return dados;

    } catch (error) {
      console.error('[SHEETS] Erro ao carregar Mapa de Carga:', error);
      throw error;
    }
  }

  async getCargasSemBox(filtros = {}) {
    try {
      console.log('[SHEETS] Buscando cargas sem BOX...');
      const todasCargas = await this.getMapaCarga(filtros);
      
      const cargasSemBox = todasCargas.filter(c => !c.box || c.box === '');
      
      console.log(`[SHEETS] ✓ ${cargasSemBox.length} cargas sem BOX`);
      return cargasSemBox;

    } catch (error) {
      console.error('[SHEETS] Erro ao buscar cargas sem BOX:', error);
      throw error;
    }
  }

  async getEstadoBoxes() {
    try {
      console.log('[SHEETS] Buscando estado dos boxes...');
      const todasCargas = await this.getMapaCarga();
      
      const boxesOcupados = todasCargas
        .filter(c => c.box && c.box !== '')
        .map(c => ({
          box: c.box,
          carga: c.carga,
          descricao: c.descricao,
          loja: c.loja,
          tipoLoja: c.tipoLoja,
          m3: c.m3,
          dataRot: c.dataRot,
          valor: c.valor
        }));

      console.log(`[SHEETS] ✓ ${boxesOcupados.length} boxes ocupados`);
      return boxesOcupados;

    } catch (error) {
      console.error('[SHEETS] Erro ao buscar estado dos boxes:', error);
      throw error;
    }
  }

  async alocarCargaBox(boxNum, cargaId) {
    try {
      console.log('[SHEETS] Alocando carga no BOX:', { boxNum, cargaId });
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
          console.log(`[SHEETS] ✓ Carga ${cargaId} alocada no BOX ${boxNum}`);
          return { ok: true, msg: `Carga alocada no BOX ${boxNum}` };
        }
      }

      console.log(`[SHEETS] ✗ Carga ${cargaId} não encontrada`);
      return { ok: false, msg: 'Carga não encontrada' };

    } catch (error) {
      console.error('[SHEETS] Erro ao alocar carga:', error);
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
      let liberados = 0;
      
      for (const row of rows) {
        const box = String(row.get('BOX') || '').trim();
        if (box === String(boxNum)) {
          row.set('BOX', '');
          await row.save();
          liberados++;
        }
      }

      console.log(`[SHEETS] ✓ BOX ${boxNum} liberado (${liberados} cargas)`);
      return { ok: true, msg: `BOX ${boxNum} liberado`, cargas: liberados };

    } catch (error) {
      console.error('[SHEETS] Erro ao liberar BOX:', error);
      return { ok: false, msg: error.message };
    }
  }

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
      const colunasParaLimpar = ['Empresa','SM','Deposito','BOX','Carga','Coluna 1','Descrição','sp','Ton','M³','Valor','Rup','Visita Picking','Volume','Coluna 2','inclusão','Roteirização','Geração Mesa','"','Reposição','Palete_Box','Baixa','Separação','Final separação','Conferencia','seotr'];
      const colunasProtegidas = ['Visitas Pendente','Separação st','conf. St','Loja','Dia oferta','Prioridade','Total_Vertical','Segmento','Tipo Loja','Data Rot','Conjugada'];
      console.log('[SHEETS] Colunas protegidas (NÃO serão limpas):', colunasProtegidas);
      console.log('[SHEETS] Buscando colunas para limpar:', colunasParaLimpar);
      const indicesColunas = [];
      colunasParaLimpar.forEach(col => {
        const idx = headers.indexOf(col);
        if (idx !== -1) {
          if (!colunasProtegidas.includes(col)) {
            indicesColunas.push(idx);
            console.log(`[SHEETS] ✓ Coluna "${col}" será limpa (índice ${idx})`);
          } else {
            console.warn(`[SHEETS] ⚠ Coluna "${col}" está protegida, pulando`);
          }
        } else {
          console.warn(`[SHEETS] ⚠ Coluna "${col}" NÃO encontrada`);
        }
      });
      if (indicesColunas.length === 0) {
        return { ok: false, msg: 'Nenhuma coluna encontrada para limpar' };
      }
      console.log(`[SHEETS] Limpando ${indicesColunas.length} de ${colunasParaLimpar.length} colunas especificadas`);
      console.log(`[SHEETS] ${colunasProtegidas.length} colunas protegidas não serão alteradas`);
      await sheet.loadCells();
      const totalRows = sheet.rowCount;
      let linhasLimpas = 0;
      for (let row = 1; row < totalRows; row++) {
        for (const colIndex of indicesColunas) {
          const cell = sheet.getCell(row, colIndex);
          if (cell) {
            cell.value = '';
          }
        }
        linhasLimpas++;
        if (linhasLimpas % 100 === 0) {
          await sheet.saveUpdatedCells();
          console.log(`[SHEETS] ${linhasLimpas}/${totalRows} linhas limpas`);
        }
      }
      await sheet.saveUpdatedCells();
      console.log('[SHEETS] ✓ Colunas limpas com sucesso');
      console.log(`[SHEETS] Colunas protegidas preservadas: ${colunasProtegidas.join(', ')}`);
      return { ok: true, msg: `${linhasLimpas} linhas limpas em ${indicesColunas.length} colunas! (${colunasProtegidas.length} colunas protegidas)`, total: linhasLimpas, colunasLimpas: indicesColunas.length, colunasProtegidas: colunasProtegidas.length };
    } catch (error) {
      console.error('[SHEETS] Erro ao limpar:', error);
      return { ok: false, msg: 'Erro ao limpar colunas: ' + error.message };
    }
  }

  async processarMapaCargaColado(dadosColados) {
    try {
      console.log('[SHEETS] ===== PROCESSAMENTO PURO - SEM CÁLCULOS =====');
      console.log('[SHEETS] Total de linhas recebidas:', dadosColados.length);
      if (!dadosColados || dadosColados.length === 0) {
        return { ok: false, msg: 'Nenhum dado fornecido' };
      }
      await this.init();
      const sheet = this.doc.sheetsByTitle['Mapa de Carga'];
      if (!sheet) {
        return { ok: false, msg: 'Aba Mapa de Carga não encontrada' };
      }
      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      console.log('[SHEETS] Headers da planilha:', headers);
      const headerMap = {};
      headers.forEach((h, idx) => { 
        if (h) headerMap[h.trim()] = idx; 
      });
      const linhasProcessadas = [];
      dadosColados.forEach((linha, idx) => {
        try {
          const campos = Array.isArray(linha) ? linha : String(linha).split('\t');
          if (campos.length < 10) {
            console.warn(`[SHEETS] Linha ${idx+1} ignorada: apenas ${campos.length} campos`);
            return;
          }
          const empresa = String(campos[0] || '').trim();
          const sm = String(campos[1] || '').trim();
          const deposito = String(campos[2] || '').trim();
          const box = String(campos[3] || '').trim();
          const carga = String(campos[4] || '').trim();
          const col1 = String(campos[5] || '').trim();
          const descricao = String(campos[6] || '').trim();
          const sp = String(campos[7] || '').trim();
          const ton = String(campos[8] || '').trim();
          const volume = String(campos[9] || '').trim();
          const valor = String(campos[10] || '').trim();
          const rup = String(campos[11] || '').trim();
          const visitasPendente = String(campos[12] || '').trim();
          const col2 = String(campos[13] || '').trim();
          const inclusao = String(campos[14] || '').trim();
          const roteirizacao = String(campos[15] || '').trim();
          const dataRot = String(campos[16] || '').trim();
          const geracaoMesa = String(campos[17] || '').trim();
          const aspas = String(campos[18] || '').trim();
          const reposicao = String(campos[19] || '').trim();
          const paleteBox = String(campos[20] || '').trim();
          const baixa = String(campos[21] || '').trim();
          const separacao = String(campos[22] || '').trim();
          const finalSeparacao = String(campos[23] || '').trim();
          const conferencia = String(campos[24] || '').trim();
          const seotr = String(campos[25] || '').trim();
          if (!carga || !descricao) {
            console.warn(`[SHEETS] Linha ${idx+1} ignorada: sem carga ou descrição`);
            return;
          }
          const dataRotFormatada = dataRot.includes(' ') ? dataRot.split(' ')[0] : dataRot;
          linhasProcessadas.push({
            'Empresa': empresa,
            'SM': sm,
            'Deposito': deposito,
            'BOX': box,
            'Carga': carga,
            'Coluna 1': col1,
            'Descrição': descricao,
            'sp': sp,
            'Ton': ton,
            'M³': volume,
            'Volume': volume,
            'Valor': valor,
            'Rup': rup,
            'Visita Picking': visitasPendente,
            'Coluna 2': col2,
            'inclusão': inclusao,
            'Roteirização': roteirizacao,
            'Data Rot': dataRotFormatada,
            'Geração Mesa': geracaoMesa,
            '"': aspas,
            'Reposição': reposicao,
            'Palete_Box': paleteBox,
            'Baixa': baixa,
            'Separação': separacao,
            'Final separação': finalSeparacao,
            'Conferencia': conferencia,
            'seotr': seotr
          });
          if ((idx + 1) % 100 === 0) {
            console.log(`[SHEETS] Processadas ${idx + 1}/${dadosColados.length} linhas`);
          }
        } catch (erroLinha) {
          console.error(`[SHEETS] Erro ao processar linha ${idx+1}:`, erroLinha);
        }
      });
      console.log(`[SHEETS] Total de linhas processadas: ${linhasProcessadas.length}`);
      if (linhasProcessadas.length === 0) {
        return { ok: false, msg: 'Nenhuma linha válida para processar' };
      }
      console.log('[SHEETS] Exemplo da primeira linha processada:', linhasProcessadas[0]);
      const rows = await sheet.getRows();
      console.log(`[SHEETS] Linhas existentes na planilha: ${rows.length}`);
      if (rows.length === 0) {
        console.log('[SHEETS] Planilha vazia, adicionando novas linhas...');
        const lote = 50;
        let adicionadas = 0;
        for (let i = 0; i < linhasProcessadas.length; i += lote) {
          const grupo = linhasProcessadas.slice(i, i + lote);
          await sheet.addRows(grupo);
          adicionadas += grupo.length;
          console.log(`[SHEETS] ${adicionadas}/${linhasProcessadas.length} linhas adicionadas`);
        }
        return { ok: true, msg: `${linhasProcessadas.length} cargas adicionadas com sucesso!`, total: linhasProcessadas.length };
      } else {
        console.log('[SHEETS] Atualizando linhas existentes via BATCH...');
        await sheet.loadCells();
        let linhasAtualizadas = 0;
        const maxLinhas = Math.min(rows.length, linhasProcessadas.length);
        for (let i = 0; i < maxLinhas; i++) {
          const dados = linhasProcessadas[i];
          const rowIndex = i + 1;
          Object.keys(dados).forEach(coluna => {
            const colIndex = headerMap[coluna];
            if (colIndex !== undefined) {
              const cell = sheet.getCell(rowIndex, colIndex);
              if (cell) {
                cell.value = dados[coluna];
              }
            }
          });
          linhasAtualizadas++;
          if (linhasAtualizadas % 50 === 0) {
            await sheet.saveUpdatedCells();
            console.log(`[SHEETS] ${linhasAtualizadas}/${maxLinhas} linhas atualizadas`);
          }
        }
        await sheet.saveUpdatedCells();
        if (linhasProcessadas.length > rows.length) {
          const novasLinhas = linhasProcessadas.slice(rows.length);
          console.log(`[SHEETS] Adicionando ${novasLinhas.length} novas linhas...`);
          const lote = 50;
          let adicionadas = 0;
          for (let i = 0; i < novasLinhas.length; i += lote) {
            const grupo = novasLinhas.slice(i, i + lote);
            await sheet.addRows(grupo);
            adicionadas += grupo.length;
            console.log(`[SHEETS] ${adicionadas}/${novasLinhas.length} novas linhas adicionadas`);
          }
          return { ok: true, msg: `${linhasAtualizadas} cargas atualizadas e ${novasLinhas.length} novas adicionadas!`, total: linhasAtualizadas + novasLinhas.length };
        }
        return { ok: true, msg: `${linhasAtualizadas} cargas atualizadas com sucesso!`, total: linhasAtualizadas };
      }
    } catch (error) {
      console.error('[SHEETS] ===== ERRO NO PROCESSAMENTO =====');
      console.error('[SHEETS] Erro:', error);
      return { ok: false, msg: error.message };
    }
  }

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
      return { ok: true, msg: `${totalNovos} novos, ${totalAtualizados} atualizados`, totais: { novos: totalNovos, atualizados: totalAtualizados } };
    } catch (error) {
      console.error('[SHEETS] Erro:', error);
      return { ok: false, msg: error.message };
    }
  }

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
