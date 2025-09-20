// lib/sheets.js
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
      
      console.log(`Conectado à planilha: ${this.doc.title}`);
      this.initialized = true;
      
      return this.doc;
    } catch (error) {
      console.error('Erro ao conectar com Google Sheets:', error);
      throw new Error('Falha na conexão com a planilha');
    }
  }

  async getSheet(sheetName) {
    const doc = await this.init();
    let sheet = doc.sheetsByTitle[sheetName];
    
    if (!sheet) {
      sheet = await doc.addSheet({ 
        title: sheetName,
        headerValues: this.getDefaultHeaders(sheetName)
      });
      console.log(`Aba '${sheetName}' criada`);
    }
    
    return sheet;
  }

  getDefaultHeaders(sheetName) {
    const headers = {
      'Usuarios': ['Usuario', 'Senha', 'Aba'],
      'Quadro': ['Matricula', 'Nome', 'Funcao'],
      'Lista': ['Supervisor', 'Aba', 'Matricula', 'Nome', 'Funcao', 'Status'],
      'Base': ['Supervisor', 'Aba', 'Matricula', 'Nome', 'Funcao', 'Status', 'Data']
    };
    return headers[sheetName] || [];
  }

  async validarLogin(usuario, senha) {
    try {
      const sheet = await this.getSheet('Usuarios');
      const rows = await sheet.getRows();
      
      const abas = [];
      usuario = String(usuario || '').trim();
      senha = String(senha || '').trim();

      for (const row of rows) {
        const u = String(row.get('Usuario') || '').trim();
        const s = String(row.get('Senha') || '').trim();
        const aba = String(row.get('Aba') || '').trim();
        
        if (u === usuario && s === senha && aba) {
          abas.push(aba);
        }
      }

      const unicas = [...new Set(abas)].sort((a,b) => a.localeCompare(b,'pt-BR'));
      
      if (unicas.length === 0) {
        return { ok: false, msg: 'Credenciais inválidas' };
      }
      
      return { ok: true, usuario, abas: unicas };
    } catch (error) {
      console.error('Erro no login:', error);
      return { ok: false, msg: 'Erro interno no servidor' };
    }
  }

  async buscarColaboradores(filtro = '') {
    try {
      const sheet = await this.getSheet('Quadro');
      const rows = await sheet.getRows();
      
      const lista = [];
      const f = String(filtro || '').toLowerCase();

      for (const row of rows) {
        const matricula = String(row.get('Matricula') || '');
        const nome = String(row.get('Nome') || '');
        const funcao = String(row.get('Funcao') || '');
        
        if (!f || matricula.toLowerCase().includes(f) || nome.toLowerCase().includes(f)) {
          lista.push({ matricula, nome, funcao });
        }
      }

      lista.sort((a, b) => {
        const fa = (a.funcao || '').toLowerCase();
        const fb = (b.funcao || '').toLowerCase();
        if (fa < fb) return -1;
        if (fa > fb) return 1;
        return (a.nome || '').toLowerCase().localeCompare((b.nome || '').toLowerCase(), 'pt-BR');
      });

      return lista;
    } catch (error) {
      console.error('Erro ao buscar colaboradores:', error);
      throw error;
    }
  }

  async getBuffer(supervisor, aba) {
    try {
      const sheet = await this.getSheet('Lista');
      const rows = await sheet.getRows();
      
      const buffer = [];
      const supLower = String(supervisor || '').toLowerCase();
      const abaLower = String(aba || '').toLowerCase();

      for (const row of rows) {
        const sup = String(row.get('Supervisor') || '').toLowerCase();
        const rowAba = String(row.get('Aba') || '').toLowerCase();
        
        const supOk = sup === supLower;
        const abaOk = !aba || rowAba === abaLower;
        
        if (supOk && abaOk) {
          buffer.push({
            matricula: row.get('Matricula'),
            nome: row.get('Nome'),
            funcao: row.get('Funcao'),
            status: row.get('Status') || ''
          });
        }
      }

      return buffer;
    } catch (error) {
      console.error('Erro ao buscar buffer:', error);
      throw error;
    }
  }

  async adicionarBuffer(supervisor, aba, colaborador) {
    try {
      const sheet = await this.getSheet('Lista');
      const rows = await sheet.getRows();
      
      const existe = rows.some(row => 
        row.get('Supervisor') === supervisor && 
        String(row.get('Matricula')) === String(colaborador.matricula)
      );
      
      if (existe) {
        return { ok: false, msg: 'Colaborador já está no buffer' };
      }

      await sheet.addRow({
        Supervisor: supervisor,
        Aba: aba,
        Matricula: colaborador.matricula,
        Nome: colaborador.nome,
        Funcao: colaborador.funcao,
        Status: ''
      });

      return { ok: true };
    } catch (error) {
      console.error('Erro ao adicionar no buffer:', error);
      throw error;
    }
  }

  async removerBuffer(supervisor, matricula) {
    try {
      const sheet = await this.getSheet('Lista');
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && 
            String(row.get('Matricula')) === String(matricula)) {
          await row.delete();
          return { ok: true };
        }
      }
      
      return { ok: false, msg: 'Item não encontrado' };
    } catch (error) {
      console.error('Erro ao remover do buffer:', error);
      throw error;
    }
  }

  async atualizarStatusBuffer(supervisor, matricula, status) {
    try {
      const sheet = await this.getSheet('Lista');
      const rows = await sheet.getRows();
      
      for (const row of rows) {
        if (row.get('Supervisor') === supervisor && 
            String(row.get('Matricula')) === String(matricula)) {
          row.set('Status', status);
          await row.save();
          return { ok: true };
        }
      }
      
      return { ok: false, msg: 'Item não encontrado' };
    } catch (error) {
      console.error('Erro ao atualizar status:', error);
      throw error;
    }
  }

  async salvarNaBase(dados) {
    try {
      const sheet = await this.getSheet('Base');
      
      if (!dados || dados.length === 0) {
        return { ok: false, msg: 'Nenhum dado para salvar' };
      }

      const supervisorLogado = dados[0][0];
      
      const rows = await sheet.getRows();
      const rowsToDelete = rows.filter(row => row.get('Supervisor') === supervisorLogado);
      
      for (const row of rowsToDelete) {
        await row.delete();
      }

      const hoje = new Date().toLocaleDateString('pt-BR');
      let salvos = 0;

      for (const linha of dados) {
        const [sup, aba, matricula, nome, funcao, status] = linha;
        
        if (matricula || nome || funcao) {
          await sheet.addRow({
            Supervisor: sup,
            Aba: aba,
            Matricula: matricula,
            Nome: nome,
            Funcao: funcao,
            Status: status,
            Data: hoje
          });
          salvos++;
        }
      }

      return { ok: true, msg: `${salvos} registros salvos na base` };
    } catch (error) {
      console.error('Erro ao salvar na base:', error);
      throw error;
    }
  }
}

const sheetsService = new SheetsService();
module.exports = sheetsService;
