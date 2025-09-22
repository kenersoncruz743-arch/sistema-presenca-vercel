// lib/sheets.js - ADAPTADO PARA SUA ESTRUTURA EXATA
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
      'Quadro': ['Coluna 1', 'NOME', 'FUNÇÃO NO RM', 'Função que atua', 'Coluna 2'],
      'Lista': ['Supervisor', 'Grupo', 'matricula', 'Nome', 'Função', 'status'],
      'Base': ['Supervisor', 'Aba', 'Matricula', 'Nome', 'Função', 'Status', 'Data']
    };
    return headers[sheetName] || [];
  }

  async validarLogin(usuario, senha) {
    try {
      const sheet = await this.getSheet('Usuarios');
      const rows = await sheet.getRows();
      
      console.log(`Verificando login para: ${usuario}`);
      console.log(`Total de linhas na aba Usuarios: ${rows.length}`);
      
      const abas = [];
      usuario = String(usuario || '').trim();
      senha = String(senha || '').trim();

      for (const row of rows) {
        const u = String(row.get('Usuario') || '').trim();
        const s = String(row.get('Senha') || '').trim();
        const aba = String(row.get('Aba') || '').trim();
        
        console.log(`Comparando: "${u}" === "${usuario}" && "${s}" === "${senha}"`);
        
        if (u === usuario && s === senha && aba) {
          abas.push(aba);
          console.log(`Match encontrado para aba: ${aba}`);
        }
      }

      const unicas = [...new Set(abas)].sort((a,b) => a.localeCompare(b,'pt-BR'));
      console.log(`Abas encontradas para o usuário: ${JSON.stringify(unicas)}`);
      
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
      
      console.log(`=== BUSCA DE COLABORADORES ===`);
      console.log(`Filtro aplicado: "${filtro}"`);
      console.log(`Total de linhas na aba Quadro: ${rows.length}`);
      
      const lista = [];
      const f = String(filtro || '').toLowerCase();

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        
        // Usando suas colunas exatas
        const matricula = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcaoRM = String(row.get('FUNÇÃO NO RM') || '').trim();
        const funcaoAtua = String(row.get('Função que atua') || '').trim();
        
        // Usa a função que atua como prioridade, senão a função no RM
        const funcao = funcaoAtua || funcaoRM;
        
        console.log(`Linha ${i + 1}: "${matricula}" | "${nome}" | "${funcao}"`);
        
        // Só inclui se tem nome preenchido
        if (nome) {
          const nomeMatch = nome.toLowerCase().includes(f);
          const matriculaMatch = matricula.toLowerCase().includes(f);
          const funcaoMatch = funcao.toLowerCase().includes(f);
          
          if (!f || nomeMatch || matriculaMatch || funcaoMatch) {
            lista.push({ matricula, nome, funcao });
            console.log(`✅ Incluído: ${matricula} - ${nome}`);
          } else {
            console.log(`❌ Filtro não bateu para: ${matricula} - ${nome}`);
          }
        } else {
          console.log(`⚠️ Nome vazio na linha ${i + 1}, pulando...`);
        }
      }

      console.log(`Total de colaboradores encontrados: ${lista.length}`);

      // Ordena por função e depois por nome
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
      
      console.log(`Buscando buffer para: ${supervisor} - ${aba}`);
      
      const buffer = [];
      const supLower = String(supervisor || '').toLowerCase();
      const abaLower = String(aba || '').toLowerCase();

      for (const row of rows) {
        const sup = String(row.get('Supervisor') || '').toLowerCase();
        // Nota: sua aba Lista usa 'Grupo' ao invés de 'Aba'
        const rowAba = String(row.get('Grupo') || '').toLowerCase();
        
        const supOk = sup === supLower;
        const abaOk = !aba || rowAba === abaLower;
        
        if (supOk && abaOk) {
          buffer.push({
            matricula: row.get('matricula'), // minúscula na sua planilha
            nome: row.get('Nome'),
            funcao: row.get('Função'),
            status: row.get('status') || '' // minúscula na sua planilha
          });
        }
      }

      console.log(`Buffer encontrado: ${buffer.length} itens`);
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
      
      // Verifica se já existe
      const existe = rows.some(row => 
        row.get('Supervisor') === supervisor && 
        String(row.get('matricula')) === String(colaborador.matricula)
      );
      
      if (existe) {
        return { ok: false, msg: 'Colaborador já está no buffer' };
      }

      // Adiciona usando os nomes corretos das colunas
      await sheet.addRow({
        'Supervisor': supervisor,
        'Grupo': aba, // sua aba Lista usa 'Grupo'
        'matricula': colaborador.matricula, // minúscula
        'Nome': colaborador.nome,
        'Função': colaborador.funcao,
        'status': '' // minúscula
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
            String(row.get('matricula')) === String(matricula)) {
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
            String(row.get('matricula')) === String(matricula)) {
          row.set('status', status); // minúscula
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
      
      // Remove dados existentes do supervisor
      const rows = await sheet.getRows();
      const rowsToDelete = rows.filter(row => row.get('Supervisor') === supervisorLogado);
      
      for (const row of rowsToDelete) {
        await row.delete();
      }

      // Adiciona novos dados usando os nomes corretos das colunas
      const hoje = new Date().toLocaleDateString('pt-BR');
      let salvos = 0;

      for (const linha of dados) {
        const [sup, aba, matricula, nome, funcao, status] = linha;
        
        if (matricula || nome || funcao) {
          await sheet.addRow({
            'Supervisor': sup,
            'Aba': aba,
            'Matricula': matricula, // maiúscula na Base
            'Nome': nome,
            'Função': funcao,
            'Status': status, // maiúscula na Base
            'Data': hoje
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
