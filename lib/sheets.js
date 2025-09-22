// lib/sheets.js - ADAPTADO PARA DUAS PLANILHAS SEPARADAS
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsService {
  constructor() {
    this.doc = null;
    this.docColetores = null; // Segunda planilha para coletores
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

      // Planilha principal
      this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      
      console.log(`Conectado à planilha principal: ${this.doc.title}`);
      this.initialized = true;
      
      return this.doc;
    } catch (error) {
      console.error('Erro ao conectar com Google Sheets:', error);
      throw new Error('Falha na conexão com a planilha');
    }
  }

  async getSheet(sheetName, useColetores = false) {
    const doc = useColetores ? await this.getColetoresDoc() : await this.init();
    let sheet = doc.sheetsByTitle[sheetName];
    
    if (!sheet) {
      sheet = await doc.addSheet({ 
        title: sheetName,
        headerValues: this.getDefaultHeaders(sheetName)
      });
      console.log(`Aba '${sheetName}' criada em ${useColetores ? 'planilha de coletores' : 'planilha principal'}`);
    }
    
    return sheet;
  }

  async getColetoresDoc() {
    if (!this.docColetores) {
      if (!process.env.GOOGLE_SHEETS_COLETORES_ID) {
        throw new Error('GOOGLE_SHEETS_COLETORES_ID não configurado. Configure a variável de ambiente para usar planilha separada.');
      }
      
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      this.docColetores = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_COLETORES_ID, serviceAccountAuth);
      await this.docColetores.loadInfo();
      console.log(`Conectado à planilha de coletores: ${this.docColetores.title}`);
    }
    return this.docColetores;
  }

  getDefaultHeaders(sheetName) {
    const headers = {
      // Headers da planilha principal (presença)
      'Usuarios': ['Usuario', 'Senha', 'Aba'],
      'Quadro': ['Coluna 1', 'NOME', 'FUNÇÃO NO RM', 'Função que atua', 'Coluna 2'],
      'Lista': ['Supervisor', 'Grupo', 'matricula', 'Nome', 'Função', 'status'],
      'Base': ['Supervisor', 'Aba', 'Matricula', 'Nome', 'Função', 'Status', 'Data'],
      
      // Headers da planilha de coletores (separada)
      'Base': ['Data', 'Chapa', 'Nome', 'Funcao', 'NumeroColetor', 'TipoOperacao', 'Situacao', 'Supervisor'],
      'Quadro': ['Chapa', 'Nome', 'Funcao'],
      'Presenca': ['Data', 'Supervisor', 'Matricula', 'Nome', 'Funcao', 'Status']
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

  // MÉTODOS EXISTENTES PARA PRESENÇA
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
        
        const matricula = String(row.get('Coluna 1') || '').trim();
        const nome = String(row.get('NOME') || '').trim();
        const funcaoRM = String(row.get('FUNÇÃO NO RM') || '').trim();
        const funcaoAtua = String(row.get('Função que atua') || '').trim();
        
        const funcao = funcaoAtua || funcaoRM;
        
        console.log(`Linha ${i + 1}: "${matricula}" | "${nome}" | "${funcao}"`);
        
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
        const rowAba = String(row.get('Grupo') || '').toLowerCase();
        
        const supOk = sup === supLower;
        const abaOk = !aba || rowAba === abaLower;
        
        if (supOk && abaOk) {
          buffer.push({
            matricula: row.get('matricula'),
            nome: row.get('Nome'),
            funcao: row.get('Função'),
            status: row.get('status') || ''
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
      
      const existe = rows.some(row => 
        row.get('Supervisor') === supervisor && 
        String(row.get('matricula')) === String(colaborador.matricula)
      );
      
      if (existe) {
        return { ok: false, msg: 'Colaborador já está no buffer' };
      }

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
          row.set('status', status);
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
            'Supervisor': sup,
            'Aba': aba,
            'Matricula': matricula,
            'Nome': nome,
            'Função': funcao,
            'Status': status,
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

  // NOVOS MÉTODOS PARA COLETORES (usando planilha separada)

  async obterDadosColetores() {
    try {
      const sheet = await this.getSheet('Quadro', true); // true = usar planilha de coletores
      const rows = await sheet.getRows();
      
      console.log(`Obtendo dados de coletores da planilha separada: ${rows.length} registros`);
      
      const dados = [];
      for (const row of rows) {
        const chapa = String(row.get('Chapa') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const funcao = String(row.get('Funcao') || '').trim();
        
        if (chapa && nome) {
          dados.push({ chapa, nome, funcao });
        }
      }
      
      console.log(`Total de colaboradores na planilha de coletores: ${dados.length}`);
      return dados;
    } catch (error) {
      console.error('Erro ao obter dados de coletores:', error);
      throw error;
    }
  }

  async salvarRegistroColetor(dados) {
    try {
      const { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor } = dados;
      
      if (!chapa || !numeroColetor || !situacoes || situacoes.length === 0) {
        return { ok: false, msg: 'Preencha todos os campos obrigatórios.' };
      }

      const sheet = await this.getSheet('Base', true); // true = usar planilha de coletores
      const rows = await sheet.getRows();
      
      const agora = new Date();
      const situacoesTexto = Array.isArray(situacoes) ? situacoes.join(', ') : situacoes;

      // Verifica duplicidade (últimos 2 minutos)
      if (rows.length > 0) {
        const ultimaEntrada = rows[rows.length - 1];
        const ultimaData = new Date(ultimaEntrada.get('Data'));
        const ultimaChapa = ultimaEntrada.get('Chapa');
        const ultimoColetor = ultimaEntrada.get('NumeroColetor');
        const ultimoTipo = ultimaEntrada.get('TipoOperacao');

        const mesmaChapa = ultimaChapa === chapa;
        const mesmoColetor = ultimoColetor === numeroColetor;
        const mesmoTipo = ultimoTipo === tipoOperacao;
        const tempoDecorrido = (agora - ultimaData) / 1000;

        if (mesmaChapa && mesmoColetor && mesmoTipo && tempoDecorrido < 120) {
          return { ok: false, msg: 'Registro já enviado recentemente. Aguarde alguns minutos antes de reenviar.' };
        }
      }

      // Salva o registro na planilha de coletores
      await sheet.addRow({
        'Data': agora.toISOString(),
        'Chapa': chapa,
        'Nome': nome,
        'Funcao': funcao,
        'NumeroColetor': numeroColetor,
        'TipoOperacao': tipoOperacao,
        'Situacao': situacoesTexto,
        'Supervisor': supervisor || ''
      });

      console.log(`Registro de coletor salvo na planilha separada: Coletor ${numeroColetor} - ${nome}`);
      return { ok: true, msg: 'Dados salvos com sucesso!' };
    } catch (error) {
      console.error('Erro ao salvar registro de coletor:', error);
      return { ok: false, msg: 'Erro ao salvar registro' };
    }
  }

  async obterColetorStatus() {
    try {
      const sheet = await this.getSheet('Base', true); // true = usar planilha de coletores
      const rows = await sheet.getRows();
      
      console.log(`Obtendo status de coletores da planilha separada: ${rows.length} registros`);
      
      const mapa = {};

      // Varre de trás para frente para pegar último status
      for (let i = rows.length - 1; i >= 0; i--) {
        const row = rows[i];
        const data = new Date(row.get('Data'));
        const chapa = row.get('Chapa');
        const nome = row.get('Nome');
        const funcao = row.get('Funcao');
        const coletor = row.get('NumeroColetor');
        const tipo = row.get('TipoOperacao');
        const situacao = row.get('Situacao');
        const supervisor = row.get('Supervisor') || '';

        if (!mapa[coletor]) {
          mapa[coletor] = {
            data: data.toLocaleDateString('pt-BR') + ' ' + data.toLocaleTimeString('pt-BR'),
            tipo,
            situacao,
            chapa,
            nome,
            supervisor
          };
        }
      }

      console.log(`Status de ${Object.keys(mapa).length} coletores processado`);
      return mapa;
    } catch (error) {
      console.error('Erro ao obter status dos coletores:', error);
      throw error;
    }
  }

  async gerarResumoColetores() {
    try {
      const statusPorColetor = await this.obterColetorStatus();

      let disponiveis = 0;
      let indisponiveis = 0;
      let quebrados = 0;

      for (const coletor in statusPorColetor) {
        const status = statusPorColetor[coletor];

        if (status.tipo === 'Entrega' && status.situacao === 'OK') {
          disponiveis++;
        } else if (status.tipo === 'Retirada') {
          indisponiveis++;
        }

        if (status.situacao !== 'OK') {
          quebrados++;
        }
      }

      return {
        disponiveis,
        indisponiveis,
        quebrados,
        total: Object.keys(statusPorColetor).length
      };
    } catch (error) {
      console.error('Erro ao gerar resumo de coletores:', error);
      throw error;
    }
  }
}

const sheetsService = new SheetsService();
module.exports = sheetsService;
