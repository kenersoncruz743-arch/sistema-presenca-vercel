const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsAvariaService {
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
      this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_ID_AVARIA, serviceAccountAuth);
      await this.doc.loadInfo();
      this.initialized = true;
      console.log(`[AVARIA] Conectado: ${this.doc.title}`);
      return this.doc;
    } catch (error) {
      console.error('[AVARIA] Erro na conexão:', error);
      throw new Error('Falha na conexão com planilha de avaria: ' + error.message);
    }
  }

  async obterDadosColaboradores() {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['quadro'];
      if (!sheet) {
        console.error('[AVARIA] Aba "quadro" não encontrada');
        return [];
      }
      const rows = await sheet.getRows();
      const colaboradores = [];
      rows.forEach(row => {
        const chapa = String(row.get('Coluna 1') || row.get('chapa') || '').trim();
        const nome = String(row.get('NOME') || row.get('nome') || '').trim();
        const funcao = String(row.get('Função que atua') || row.get('funcao') || '').trim();
        if (chapa && nome) {
          colaboradores.push({ chapa, nome, funcao });
        }
      });
      console.log(`[AVARIA] ${colaboradores.length} colaboradores carregados`);
      return colaboradores;
    } catch (error) {
      console.error('[AVARIA] Erro ao buscar colaboradores:', error);
      return [];
    }
  }

  async obterDadosProdutos() {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Endereços'];
      if (!sheet) {
        console.error('[AVARIA] Aba "Endereços" não encontrada');
        return [];
      }
      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      console.log('[AVARIA] Headers encontrados:', headers);
      const rows = await sheet.getRows();
      const produtos = [];
      rows.forEach(row => {
        let codigo = '';
        try {
          codigo = String(row.get(headers[13]) || '').trim();
          if (!codigo) {
            codigo = String(row.get(headers[12]) || '').trim();
          }
        } catch (e) {
          console.warn('[AVARIA] Erro ao ler código:', e.message);
        }
        const descricao = String(row.get(headers[14]) || '').trim();
        const embalagem = String(row.get(headers[15]) || '').trim();
        if (codigo) {
          produtos.push({ codigo, descricao, embalagem });
        }
      });
      console.log(`[AVARIA] ${produtos.length} produtos carregados`);
      return produtos;
    } catch (error) {
      console.error('[AVARIA] Erro ao buscar produtos:', error);
      return [];
    }
  }

  async salvarRegistro(usuarioLogado, codigoProduto, descricaoProduto, embalagemProduto, quantidade) {
    try {
      await this.init();
      let sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) {
        console.log('[AVARIA] Criando aba Base...');
        sheet = await this.doc.addSheet({
          title: 'Base',
          headerValues: ['Data/Hora', 'Usuário', 'Código Produto', 'Descrição', 'Embalagem', 'Quantidade']
        });
      }
      const agora = new Date();
      const dataHora = agora.toLocaleString('pt-BR', { timeZone: 'America/Manaus' });
      await sheet.addRow({
        'Data/Hora': dataHora,
        'Usuário': usuarioLogado,
        'Código Produto': String(codigoProduto),
        'Descrição': descricaoProduto,
        'Embalagem': embalagemProduto,
        'Quantidade': parseInt(quantidade, 10)
      });
      console.log(`[AVARIA] Registro salvo: ${usuarioLogado} - ${codigoProduto} - ${quantidade}`);
      return { ok: true, msg: 'Registro salvo com sucesso!' };
    } catch (error) {
      console.error('[AVARIA] Erro ao salvar:', error);
      return { ok: false, msg: 'Erro ao salvar: ' + error.message };
    }
  }

  async obterHistorico(filtros = {}) {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Base'];
      if (!sheet) {
        return { ok: true, dados: [] };
      }
      const rows = await sheet.getRows();
      let dados = rows.map(row => ({
        dataHora: String(row.get('Data/Hora') || ''),
        usuario: String(row.get('Usuário') || ''),
        codigoProduto: String(row.get('Código Produto') || ''),
        descricao: String(row.get('Descrição') || ''),
        embalagem: String(row.get('Embalagem') || ''),
        quantidade: parseInt(row.get('Quantidade') || 0)
      }));
      if (filtros.usuario) {
        dados = dados.filter(d => d.usuario.toLowerCase().includes(filtros.usuario.toLowerCase()));
      }
      if (filtros.codigoProduto) {
        dados = dados.filter(d => d.codigoProduto.includes(filtros.codigoProduto));
      }
      console.log(`[AVARIA] Histórico: ${dados.length} registros`);
      return { ok: true, dados };
    } catch (error) {
      console.error('[AVARIA] Erro ao buscar histórico:', error);
      return { ok: false, msg: error.message };
    }
  }
}

module.exports = new SheetsAvariaService();
