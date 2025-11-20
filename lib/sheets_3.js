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

  /**
   * Normaliza código removendo zeros à esquerda e espaços
   * SEMPRE retorna string
   */
  normalizarCodigo(codigo) {
    if (!codigo) return '';
    
    // Converte para string e remove espaços
    let normalizado = String(codigo).trim();
    
    // Se começa com zeros, remove (exceto se for só "0")
    if (normalizado !== '0') {
      normalizado = normalizado.replace(/^0+/, '');
    }
    
    // Se ficou vazio após remover zeros, retorna '0'
    if (normalizado === '') normalizado = '0';
    
    return normalizado;
  }

  async obterDadosProdutos() {
    try {
      await this.init();
      console.log('[AVARIA] Buscando produtos na aba Endereços...');
      
      const sheet = this.doc.sheetsByTitle['Endereços'];
      if (!sheet) {
        console.error('[AVARIA] Aba "Endereços" não encontrada');
        return [];
      }
      
      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      console.log('[AVARIA] Headers:', headers);
      
      const rows = await sheet.getRows();
      console.log(`[AVARIA] Total de linhas: ${rows.length}`);
      
      const produtos = [];
      const mapaCodigos = new Map(); // Para controle de duplicatas
      
      rows.forEach((row, index) => {
        try {
          const codigo = String(row.get('COD') || '').trim();
          const descricao = String(row.get('DESC') || '').trim();
          const ean = String(row.get('Ean') || '').trim();
          const dun = String(row.get('Dun') || '').trim();
          const codDesc = String(row.get('Cod. desc.') || '').trim();
          const endereco = String(row.get('Endereço') || '').trim();
          
          // Debug primeiras 5 linhas
          if (index < 5) {
            console.log(`[AVARIA] Linha ${index + 1}:`, {
              codigo,
              descricao,
              ean,
              dun,
              codDesc
            });
          }
          
          // Valida se tem pelo menos código ou descrição
          if (!codigo && !descricao) {
            return;
          }
          
          // Normaliza códigos
          const codigoNorm = this.normalizarCodigo(codigo);
          const eanNorm = this.normalizarCodigo(ean);
          const dunNorm = this.normalizarCodigo(dun);
          
          // Cria objeto do produto
          const produto = {
            codigo: codigoNorm,
            codigoOriginal: codigo,
            descricao: descricao || 'Sem descrição',
            embalagem: codDesc || 'N/A',
            ean: eanNorm,
            eanOriginal: ean,
            dun: dunNorm,
            dunOriginal: dun,
            endereco: endereco || 'N/A'
          };
          
          // Adiciona ao mapa usando TODOS os códigos possíveis como chave
          const chavesParaBusca = [];
          
          // Adiciona código original e normalizado
          if (codigo) {
            chavesParaBusca.push(codigo);
            if (codigoNorm !== codigo) {
              chavesParaBusca.push(codigoNorm);
            }
          }
          
          // Adiciona EAN original e normalizado
          if (ean) {
            chavesParaBusca.push(ean);
            if (eanNorm !== ean) {
              chavesParaBusca.push(eanNorm);
            }
          }
          
          // Adiciona DUN original e normalizado
          if (dun) {
            chavesParaBusca.push(dun);
            if (dunNorm !== dun) {
              chavesParaBusca.push(dunNorm);
            }
          }
          
          // Remove duplicatas e códigos vazios/zero
          const chavesUnicas = [...new Set(chavesParaBusca)]
            .filter(c => c && c !== '0' && c !== '');
          
          // Adiciona produto ao mapa para cada chave única
          chavesUnicas.forEach(chave => {
            if (!mapaCodigos.has(chave)) {
              mapaCodigos.set(chave, produto);
            }
          });
          
          // Adiciona à lista de produtos (sem duplicação por descrição)
          const jaExiste = produtos.some(p => 
            p.descricao === produto.descricao && 
            p.codigo === produto.codigo
          );
          
          if (!jaExiste) {
            produtos.push(produto);
          }
          
        } catch (erro) {
          console.error(`[AVARIA] Erro ao processar linha ${index + 1}:`, erro);
        }
      });
      
      console.log(`[AVARIA] ✓ ${produtos.length} produtos processados`);
      console.log(`[AVARIA] ✓ ${mapaCodigos.size} códigos indexados`);
      
      // Debug: mostra alguns exemplos
      if (mapaCodigos.size > 0) {
        const exemplos = Array.from(mapaCodigos.entries()).slice(0, 5);
        console.log('[AVARIA] Exemplos de códigos indexados:');
        exemplos.forEach(([chave, prod]) => {
          console.log(`  "${chave}" → ${prod.descricao}`);
        });
      }
      
      // Retorna objeto com lista de produtos E mapa de busca
      return {
        produtos: produtos,
        mapaBusca: mapaCodigos
      };
      
    } catch (error) {
      console.error('[AVARIA] Erro ao buscar produtos:', error);
      return { produtos: [], mapaBusca: new Map() };
    }
  }

  async salvarRegistro(usuarioLogado, codigoProduto, descricaoProduto, embalagemProduto, motivo, quantidade) {
    try {
      await this.init();
      let sheet = this.doc.sheetsByTitle['Base'];
      
      if (!sheet) {
        console.log('[AVARIA] Criando aba Base...');
        sheet = await this.doc.addSheet({
          title: 'Base',
          headerValues: ['Data/Hora', 'Usuário', 'Código Produto', 'Descrição', 'Embalagem', 'Motivo', 'Quantidade']
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
        'Motivo': motivo,
        'Quantidade': parseInt(quantidade, 10)
      });
      
      console.log(`[AVARIA] Registro salvo: ${usuarioLogado} - ${codigoProduto} - ${motivo} - ${quantidade}`);
      return { ok: true, msg: '✓ Registro salvo com sucesso!' };
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
        motivo: String(row.get('Motivo') || ''),
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
