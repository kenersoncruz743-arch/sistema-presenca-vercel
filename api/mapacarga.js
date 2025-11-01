<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Mapa de Carga - Grid Visual</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container { max-width: 100%; margin: 0 auto; }
        
        .header {
            background: white;
            padding: 20px 30px;
            border-radius: 15px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
        }
        
        .header-left h1 { color: #333; font-size: 24px; display: flex; align-items: center; gap: 10px; }
        .header-right { display: flex; gap: 10px; flex-wrap: wrap; }
        
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        
        .btn-primary { background: #667eea; color: white; }
        .btn-success { background: #28a745; color: white; }
        .btn-danger { background: #dc3545; color: white; }
        .btn-warning { background: #ffc107; color: #212529; }
        .btn-secondary { background: #6c757d; color: white; }
        .btn:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
        .btn:disabled { opacity: 0.6; cursor: not-allowed; transform: none; }
        
        .alert {
            padding: 15px 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
            animation: slideIn 0.3s;
        }
        
        .alert-info { background: #d1ecf1; color: #0c5460; border-left: 4px solid #17a2b8; }
        .alert-success { background: #d4edda; color: #155724; border-left: 4px solid #28a745; }
        .alert-danger { background: #f8d7da; color: #721c24; border-left: 4px solid #dc3545; }
        .alert-warning { background: #fff3cd; color: #856404; border-left: 4px solid #ffc107; }
        
        @keyframes slideIn {
            from { transform: translateY(-20px); opacity: 0; }
            to { transform: translateY(0); opacity: 1; }
        }
        
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 1000;
            align-items: center;
            justify-content: center;
        }
        
        .modal.active { display: flex; }
        
        .modal-content {
            background: white;
            padding: 30px;
            border-radius: 15px;
            max-width: 98vw;
            width: 1800px;
            max-height: 95vh;
            overflow-y: auto;
        }
        
        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #f0f0f0;
        }
        
        .close-modal {
            background: none;
            border: none;
            font-size: 28px;
            cursor: pointer;
            color: #999;
            line-height: 1;
            padding: 0;
            width: 30px;
            height: 30px;
        }
        
        .instructions {
            background: #e7f3ff;
            border-left: 4px solid #2196F3;
            padding: 15px 20px;
            margin-bottom: 20px;
            border-radius: 8px;
        }
        
        .instructions h4 {
            color: #1976d2;
            margin-bottom: 10px;
            font-size: 14px;
        }
        
        .instructions ol {
            margin-left: 20px;
            color: #333;
            font-size: 13px;
            line-height: 1.8;
        }
        
        .instructions ol li {
            margin-bottom: 6px;
        }
        
        .grid-info {
            background: #f8f9fa;
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            gap: 30px;
            flex-wrap: wrap;
            align-items: center;
        }
        
        .grid-stat {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .grid-stat i {
            font-size: 20px;
            color: #667eea;
        }
        
        .grid-stat .label {
            font-size: 12px;
            color: #6c757d;
        }
        
        .grid-stat .value {
            font-size: 18px;
            font-weight: bold;
            color: #333;
        }
        
        .spreadsheet-container {
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            overflow: hidden;
            background: white;
            margin-bottom: 20px;
        }
        
        .spreadsheet-wrapper {
            overflow: auto;
            max-height: 600px;
        }
        
        .spreadsheet {
            border-collapse: collapse;
            width: 100%;
            font-size: 11px;
            font-family: 'Courier New', monospace;
        }
        
        .spreadsheet th {
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            padding: 8px 4px;
            text-align: center;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
            border-right: 1px solid rgba(255,255,255,0.2);
            font-size: 10px;
            min-width: 80px;
        }
        
        .spreadsheet th.row-number {
            background: #495057;
            min-width: 50px;
            font-size: 11px;
        }
        
        .spreadsheet td {
            border: 1px solid #e0e0e0;
            padding: 6px 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 150px;
            font-size: 11px;
        }
        
        .spreadsheet td.row-number {
            background: #f8f9fa;
            text-align: center;
            font-weight: bold;
            color: #495057;
            position: sticky;
            left: 0;
            z-index: 5;
            border-right: 2px solid #dee2e6;
        }
        
        .spreadsheet tbody tr:hover td:not(.row-number) {
            background: #e3f2fd;
        }
        
        .spreadsheet tbody tr:nth-child(even) td:not(.row-number) {
            background: #fafafa;
        }
        
        .spreadsheet tbody tr:nth-child(even):hover td:not(.row-number) {
            background: #e3f2fd;
        }
        
        .col-empresa { min-width: 60px !important; }
        .col-sm { min-width: 40px !important; }
        .col-deposito { min-width: 70px !important; }
        .col-box { min-width: 60px !important; }
        .col-carga { min-width: 90px !important; max-width: 90px !important; font-weight: 600; }
        .col-descricao { min-width: 250px !important; max-width: 250px !important; }
        .col-ton { min-width: 60px !important; text-align: right; }
        .col-volume { min-width: 70px !important; max-width: 70px !important; text-align: right; }
        .col-valor { min-width: 80px !important; text-align: right; }
        .col-visitas { min-width: 60px !important; text-align: center; }
        .col-data { min-width: 90px !important; }
        .col-status { min-width: 80px !important; }
        .col-loja { min-width: 70px !important; font-weight: 600; }
        
        .highlight-col {
            background: #fff3cd !important;
        }
        
        .paste-area {
            width: 100%;
            min-height: 200px;
            max-height: 300px;
            padding: 15px;
            border: 2px dashed #e0e0e0;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 11px;
            resize: vertical;
            background: #f8f9fa;
            overflow-y: auto;
            margin-bottom: 15px;
        }
        
        .paste-area:focus {
            outline: none;
            border-color: #667eea;
            background: white;
        }
        
        .loading {
            display: inline-block;
            animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        @media (max-width: 768px) {
            .header { flex-direction: column; text-align: center; }
            .modal-content { padding: 20px; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-left">
                <h1><i class="fas fa-truck-loading"></i> Gestão de Mapa de Carga</h1>
            </div>
            <div class="header-right">
                <button class="btn btn-primary" onclick="abrirModalGrid()"><i class="fas fa-table"></i> Importar com Grid Visual</button>
                <button class="btn btn-danger" onclick="limparColunas()"><i class="fas fa-eraser"></i> Limpar Tabela</button>
                <button class="btn btn-success" onclick="carregarDados()"><i class="fas fa-sync-alt"></i> Atualizar</button>
                <button class="btn btn-secondary" onclick="voltarMenu()"><i class="fas fa-arrow-left"></i> Voltar</button>
            </div>
        </div>

        <div id="statusDiv"></div>
        
        <div style="background: white; padding: 20px; border-radius: 15px; text-align: center;">
            <i class="fas fa-table" style="font-size: 64px; color: #667eea; margin-bottom: 15px;"></i>
            <h2 style="color: #333; margin-bottom: 10px;">Use o Grid Visual para Importar Dados</h2>
            <p style="color: #666; margin-bottom: 20px;">Visualize os dados em formato de planilha antes de importar</p>
            <button class="btn btn-primary" onclick="abrirModalGrid()" style="font-size: 16px; padding: 15px 30px;">
                <i class="fas fa-table"></i> Abrir Grid de Importação
            </button>
        </div>
    </div>

    <!-- Modal Grid -->
    <div class="modal" id="modalGrid">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-table"></i> Importação com Grid Visual - Estilo Planilha Excel</h3>
                <button class="close-modal" onclick="fecharModal()">×</button>
            </div>

            <div class="instructions">
                <h4><i class="fas fa-info-circle"></i> Como Usar o Grid Visual</h4>
                <ol>
                    <li>Abra o <strong>Excel</strong> e copie TODAS as linhas (COM ou SEM cabeçalho)</li>
                    <li>Cole no campo de texto abaixo com <strong>Ctrl+V</strong></li>
                    <li>Os dados aparecerão automaticamente no <strong>grid visual</strong> estilo Excel</li>
                    <li>Verifique se as colunas estão corretas (Empresa, Carga, Descrição, etc.)</li>
                    <li>Colunas importantes: <span style="background: #fff3cd; padding: 2px 6px; border-radius: 3px;">Empresa, Carga, Descrição, Ton, Volume, Valor, Visitas, Data Rot, Status Sep</span></li>
                    <li>Clique em <strong>"Processar e Salvar"</strong> para importar</li>
                    <li><strong>Importante:</strong> Grid suporta até 1000 linhas</li>
                </ol>
            </div>
            
            <textarea class="paste-area" id="pasteArea" placeholder="Cole aqui os dados do Excel (Ctrl+V)...&#10;&#10;O grid visual aparecerá automaticamente abaixo..."></textarea>
            
            <div id="gridInfo" class="grid-info" style="display: none;">
                <div class="grid-stat">
                    <i class="fas fa-table"></i>
                    <div>
                        <div class="label">Total de Linhas</div>
                        <div class="value" id="totalLinhas">0</div>
                    </div>
                </div>
                <div class="grid-stat">
                    <i class="fas fa-columns"></i>
                    <div>
                        <div class="label">Total de Colunas</div>
                        <div class="value" id="totalColunas">0</div>
                    </div>
                </div>
                <div class="grid-stat">
                    <i class="fas fa-check-circle"></i>
                    <div>
                        <div class="label">Status</div>
                        <div class="value" id="statusGrid" style="color: #28a745;">Pronto</div>
                    </div>
                </div>
            </div>
            
            <div id="gridContainer" style="display: none;">
                <div class="spreadsheet-container">
                    <div class="spreadsheet-wrapper">
                        <table class="spreadsheet" id="gridTable">
                            <thead>
                                <tr>
                                    <th class="row-number">#</th>
                                    <th class="col-empresa">A<br>Empresa</th>
                                    <th class="col-sm">B<br>SM</th>
                                    <th class="col-deposito">C<br>Depósito</th>
                                    <th class="col-box">D<br>BOX</th>
                                    <th class="col-carga highlight-col">E<br>Carga</th>
                                    <th>F<br>Col 1</th>
                                    <th class="col-descricao highlight-col">G<br>Descrição</th>
                                    <th>H<br>SP</th>
                                    <th class="col-ton highlight-col">I<br>Ton</th>
                                    <th class="col-volume highlight-col">J<br>Volume (m³)</th>
                                    <th class="col-valor highlight-col">K<br>Valor</th>
                                    <th>L<br>RUP</th>
                                    <th class="col-visitas highlight-col">M<br>Visitas</th>
                                    <th>N<br>Col 2</th>
                                    <th>O<br>Inclusão</th>
                                    <th>P<br>Rot</th>
                                    <th class="col-data highlight-col">Q<br>Data Rot</th>
                                    <th>R<br>Geração</th>
                                    <th>S<br>Aspas</th>
                                    <th>T<br>Repos</th>
                                    <th>U<br>Palete</th>
                                    <th>V<br>Baixa</th>
                                    <th class="col-status highlight-col">W<br>Status Sep</th>
                                    <th>X<br>Fim Sep</th>
                                    <th>Y<br>Conf</th>
                                    <th>Z<br>SEOTR</th>
                                </tr>
                            </thead>
                            <tbody id="gridBody"></tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div style="margin-top: 20px; display: flex; gap: 10px;">
                <button class="btn btn-success" onclick="processarDados()" style="flex: 1;" id="btnProcessar" disabled>
                    <i class="fas fa-check"></i> Processar e Salvar
                </button>
                <button class="btn btn-secondary" onclick="limparArea()">
                    <i class="fas fa-eraser"></i> Limpar
                </button>
            </div>
            <div id="uploadStatus" style="margin-top: 20px;"></div>
        </div>
    </div>

    <script>
        const API_URL = '/api/mapacarga';
        let dadosGrid = [];

        function showStatus(msg, type = 'info') {
            const div = document.getElementById('statusDiv');
            const icons = { info: 'info-circle', success: 'check-circle', danger: 'exclamation-triangle', warning: 'exclamation-circle' };
            div.innerHTML = `<div class="alert alert-${type}"><i class="fas fa-${icons[type]}"></i><span>${msg}</span></div>`;
            setTimeout(() => div.innerHTML = '', 5000);
        }

        function abrirModalGrid() {
            document.getElementById('modalGrid').classList.add('active');
            document.getElementById('pasteArea').focus();
        }

        function fecharModal() {
            document.getElementById('modalGrid').classList.remove('active');
        }

        function limparArea() {
            document.getElementById('pasteArea').value = '';
            document.getElementById('uploadStatus').innerHTML = '';
            document.getElementById('gridContainer').style.display = 'none';
            document.getElementById('gridInfo').style.display = 'none';
            document.getElementById('btnProcessar').disabled = true;
            dadosGrid = [];
        }

        document.addEventListener('DOMContentLoaded', function() {
            const pasteArea = document.getElementById('pasteArea');
            
            if (pasteArea) {
                pasteArea.addEventListener('input', function() {
                    gerarGrid();
                });
            }
        });

        function gerarGrid() {
            const texto = document.getElementById('pasteArea').value.trim();
            
            if (!texto) {
                document.getElementById('gridContainer').style.display = 'none';
                document.getElementById('gridInfo').style.display = 'none';
                document.getElementById('btnProcessar').disabled = true;
                dadosGrid = [];
                return;
            }

            try {
                const linhas = texto.split('\n').filter(l => l.trim());
                
                if (linhas.length === 0) {
                    return;
                }

                // Limita a 1000 linhas
                const linhasLimitadas = linhas.slice(0, 1000);
                
                dadosGrid = linhasLimitadas.map(linha => {
                    const campos = linha.split('\t');
                    // Garante 26 colunas
                    while (campos.length < 26) campos.push('');
                    return campos.slice(0, 26);
                });

                const totalColunas = Math.max(...dadosGrid.map(l => l.length));

                document.getElementById('totalLinhas').textContent = dadosGrid.length;
                document.getElementById('totalColunas').textContent = totalColunas;
                document.getElementById('gridInfo').style.display = 'flex';
                
                if (linhas.length > 1000) {
                    document.getElementById('statusGrid').textContent = 'Limitado';
                    document.getElementById('statusGrid').style.color = '#ffc107';
                    showStatus(`⚠️ Apenas as primeiras 1000 linhas serão exibidas (total: ${linhas.length})`, 'warning');
                } else {
                    document.getElementById('statusGrid').textContent = 'Pronto';
                    document.getElementById('statusGrid').style.color = '#28a745';
                }

                renderGrid();
                
                document.getElementById('gridContainer').style.display = 'block';
                document.getElementById('btnProcessar').disabled = false;

            } catch (error) {
                console.error('[GRID] Erro:', error);
                document.getElementById('uploadStatus').innerHTML = `
                    <div class="alert alert-danger">
                        <i class="fas fa-exclamation-triangle"></i>
                        Erro ao processar dados. Verifique o formato.
                    </div>
                `;
            }
        }

        function renderGrid() {
            const tbody = document.getElementById('gridBody');
            tbody.innerHTML = '';

            dadosGrid.forEach((linha, idx) => {
                const tr = document.createElement('tr');
                
                // Célula de número da linha
                const tdNum = document.createElement('td');
                tdNum.className = 'row-number';
                tdNum.textContent = idx + 1;
                tr.appendChild(tdNum);

                // 26 colunas de dados
                for (let i = 0; i < 26; i++) {
                    const td = document.createElement('td');
                    const valor = linha[i] || '';
                    td.textContent = valor;
                    td.title = valor; // Tooltip com valor completo
                    
                    // Classes especiais
                    if (i === 0) td.className = 'col-empresa';
                    else if (i === 4) td.className = 'col-carga highlight-col';
                    else if (i === 6) td.className = 'col-descricao highlight-col';
                    else if (i === 8) td.className = 'col-ton highlight-col';
                    else if (i === 9) td.className = 'col-volume highlight-col';
                    else if (i === 10) td.className = 'col-valor highlight-col';
                    else if (i === 12) td.className = 'col-visitas highlight-col';
                    else if (i === 16) td.className = 'col-data highlight-col';
                    else if (i === 22) td.className = 'col-status highlight-col';
                    
                    tr.appendChild(td);
                }

                tbody.appendChild(tr);
            });
        }

        async function processarDados() {
            if (dadosGrid.length === 0) {
                showStatus('Nenhum dado para processar', 'danger');
                return;
            }

            const statusDiv = document.getElementById('uploadStatus');
            const btnProcessar = document.getElementById('btnProcessar');
            
            statusDiv.innerHTML = '<div class="alert alert-info"><i class="fas fa-spinner loading"></i> Processando e salvando...</div>';
            btnProcessar.disabled = true;
            
            try {
                console.log('[MAPACARGA] Processando', dadosGrid.length, 'linhas');
                
                // Converte de volta para formato de linhas com tabs
                const linhasTexto = dadosGrid.map(linha => linha.join('\t'));
                
                const response = await fetch(API_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        action: 'importar',
                        dadosImportacao: linhasTexto
                    })
                });
                
                const result = await response.json();
                
                if (result.ok) {
                    statusDiv.innerHTML = `<div class="alert alert-success"><i class="fas fa-check-circle"></i> ${result.msg}</div>`;
                    setTimeout(() => {
                        fecharModal();
                        showStatus(result.msg, 'success');
                    }, 2000);
                } else {
                    throw new Error(result.msg || 'Erro ao processar');
                }
            } catch (error) {
                console.error('[MAPACARGA] Erro:', error);
                statusDiv.innerHTML = `<div class="alert alert-danger"><i class="fas fa-exclamation-triangle"></i> Erro: ${error.message}</div>`;
                btnProcessar.disabled = false;
            }
        }

        async function limparColunas() {
            if (!confirm('⚠️ ATENÇÃO: Esta ação irá limpar TODAS as colunas editáveis da planilha.\n\nColunas protegidas serão preservadas.\n\nDeseja continuar?')) {
                return;
            }

            const btn = event.target;
            const originalText = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = '<i class="fas fa-spinner loading"></i> Limpando...';

            try {
                showStatus('Iniciando limpeza...', 'info');

                const response = await fetch(API_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action: 'limpar' })
                });

                const result = await response.json();

                if (result.ok) {
                    showStatus(result.msg, 'success');
                } else {
                    throw new Error(result.msg || 'Erro ao limpar');
                }
            } catch (error) {
                console.error('[MAPACARGA] Erro:', error);
                showStatus('Erro ao limpar: ' + error.message, 'danger');
            } finally {
                btn.disabled = false;
                btn.innerHTML = originalText;
            }
        }

        async function carregarDados() {
            showStatus('Função carregarDados não implementada nesta versão', 'info');
        }

        function voltarMenu() {
            if (confirm('Deseja voltar ao menu?')) {
                window.location.href = '/menu';
            }
        }

        // Fechar modal ao clicar fora
        document.getElementById('modalGrid').addEventListener('click', function(e) {
            if (e.target === this) fecharModal();
        });

        // Atalho ESC
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') fecharModal();
        });

        window.onload = function() {
            const userName = document.getElementById('userName');
            if (userName) {
                const usuario = sessionStorage.getItem('usuario') || 'Usuário';
                userName.textContent = usuario;
            }
        };
    </script>
</body>
</html>
