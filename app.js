// Dados das planilhas e configurações
const planilhaData = {
    estoque: {
        nome: "Controle de Estoque",
        campos: ["codigo", "produto", "estoque_inicial", "entradas", "saidas", "saldo_atual", "estoque_minimo", "status"],
        camposNomes: {
            codigo: "Código",
            produto: "Produto",
            estoque_inicial: "Estoque Inicial",
            entradas: "Entradas",
            saidas: "Saídas", 
            saldo_atual: "Saldo Atual",
            estoque_minimo: "Estoque Mínimo",
            status: "Status"
        },
        formulas: {
            saldo_atual: "=C{row}+D{row}-E{row}",
            status: "=IF(F{row}<=G{row},\"CRÍTICO\",IF(F{row}<=G{row}*1.5,\"BAIXO\",\"OK\"))"
        }
    },
    financeiro: {
        nome: "Controle Financeiro",
        campos: ["data", "categoria", "descricao", "valor", "tipo", "saldo_acumulado"],
        camposNomes: {
            data: "Data",
            categoria: "Categoria", 
            descricao: "Descrição",
            valor: "Valor",
            tipo: "Tipo",
            saldo_acumulado: "Saldo Acumulado"
        },
        formulas: {
            saldo_acumulado: "=IF(E{row}=\"Receita\",F{prevRow}+D{row},F{prevRow}-D{row})"
        }
    },
    vendas: {
        nome: "Vendas e Comissões",
        campos: ["vendedor", "produto", "quantidade", "valor_unitario", "total", "comissao", "meta"],
        camposNomes: {
            vendedor: "Vendedor",
            produto: "Produto",
            quantidade: "Quantidade",
            valor_unitario: "Valor Unitário",
            total: "Total",
            comissao: "Comissão",
            meta: "Meta"
        },
        formulas: {
            total: "=C{row}*D{row}",
            comissao: "=E{row}*0.05"
        }
    },
    clientes: {
        nome: "Lista de Clientes",
        campos: ["nome", "email", "telefone", "endereco", "segmento", "status", "ultima_compra"],
        camposNomes: {
            nome: "Nome",
            email: "Email",
            telefone: "Telefone",
            endereco: "Endereço",
            segmento: "Segmento",
            status: "Status",
            ultima_compra: "Última Compra"
        }
    },
    projetos: {
        nome: "Controle de Projetos",
        campos: ["projeto", "tarefa", "responsavel", "inicio", "fim", "status", "progresso"],
        camposNomes: {
            projeto: "Projeto",
            tarefa: "Tarefa",
            responsavel: "Responsável",
            inicio: "Início",
            fim: "Fim",
            status: "Status",
            progresso: "Progresso (%)"
        }
    },
    agenda: {
        nome: "Agenda/Cronograma",
        campos: ["data", "hora", "evento", "local", "participantes", "status", "observacoes"],
        camposNomes: {
            data: "Data",
            hora: "Hora",
            evento: "Evento",
            local: "Local",
            participantes: "Participantes",
            status: "Status",
            observacoes: "Observações"
        }
    }
};

// Estado da aplicação
let currentTipo = '';
let currentFormData = {};
let dadosUsuarioReais = [];
let customSettings = {
    headerColor: '#2196F3',
    tableStyle: 'default'
};

// Elementos DOM
const menuToggle = document.getElementById('menuToggle');
const nav = document.getElementById('nav');
const sections = document.querySelectorAll('.section');
const navLinks = document.querySelectorAll('.nav-link');
const planilhaCards = document.querySelectorAll('.planilha-card');
const backBtn = document.getElementById('backBtn');
const backToFormBtn = document.getElementById('backToFormBtn');
const formularioTitulo = document.getElementById('formularioTitulo');
const planilhaForm = document.getElementById('planilhaForm');
const tipoAtual = document.getElementById('tipoAtual');
const camposEspecificos = document.getElementById('camposEspecificos');
const addCampoBtn = document.getElementById('addCampo');
const camposPersonalizados = document.getElementById('camposPersonalizados');
const previewBtn = document.getElementById('previewBtn');
const previewTable = document.getElementById('previewTable');
const customizeBtn = document.getElementById('customizeBtn');
const generateBtn = document.getElementById('generateBtn');
const customModal = document.getElementById('customModal');
const modalClose = document.getElementById('modalClose');
const cancelCustom = document.getElementById('cancelCustom');
const applyCustom = document.getElementById('applyCustom');
const loading = document.getElementById('loading');
const validationAlert = document.getElementById('validationAlert');

// Inicialização
document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
    showSection('home');
});

// Event Listeners
function setupEventListeners() {
    // Menu toggle
    menuToggle.addEventListener('click', toggleMobileMenu);
    
    // Navigation
    navLinks.forEach(link => {
        link.addEventListener('click', handleNavigation);
    });
    
    // Cards de planilhas
    planilhaCards.forEach(card => {
        card.addEventListener('click', function() {
            const tipo = this.dataset.tipo;
            initFormulario(tipo);
        });
    });
    
    // Botões de navegação
    backBtn.addEventListener('click', () => showSection('home'));
    backToFormBtn.addEventListener('click', () => showSection('formulario'));
    
    // Formulário
    addCampoBtn.addEventListener('click', addCampoPersonalizado);
    previewBtn.addEventListener('click', showPreviewComDadosReais);
    planilhaForm.addEventListener('submit', handleFormSubmit);
    
    // Preview e geração
    customizeBtn.addEventListener('click', () => customModal.classList.remove('hidden'));
    generateBtn.addEventListener('click', gerarPlanilhaComDadosReais);
    
    // Modal
    modalClose.addEventListener('click', closeModal);
    cancelCustom.addEventListener('click', closeModal);
    applyCustom.addEventListener('click', applyCustomization);
    
    // Fechar modal clicando fora
    customModal.addEventListener('click', function(e) {
        if (e.target === this) closeModal();
    });
    
    // Campos personalizados e dinâmicos
    document.addEventListener('click', function(e) {
        if (e.target.classList.contains('remove-campo')) {
            removeCampoPersonalizado(e.target);
        }
        if (e.target.classList.contains('remove-item')) {
            removeItemDinamico(e.target);
        }
    });
}

// ===== FUNÇÕES DE CAPTURA DE DADOS REAIS =====

function coletarDadosFormulario() {
    const tipo = document.getElementById('tipoAtual').value;
    const dados = [];
    
    console.log('Coletando dados do tipo:', tipo);
    
    if (tipo === 'estoque') {
        const produtos = document.querySelectorAll('#produtosList .campo-dinamico');
        produtos.forEach((produto, index) => {
            const codigo = produto.querySelector('input[name="produto_codigo[]"]');
            const nome = produto.querySelector('input[name="produto_nome[]"]');
            const inicial = produto.querySelector('input[name="produto_inicial[]"]');
            const minimo = produto.querySelector('input[name="produto_minimo[]"]');
            
            if (nome && nome.value.trim()) {
                dados.push({
                    codigo: codigo ? codigo.value.trim() : `PROD${(index + 1).toString().padStart(3, '0')}`,
                    nome: nome.value.trim(),
                    estoqueInicial: inicial ? parseInt(inicial.value) || 0 : 0,
                    entradas: 0,
                    saidas: 0,
                    estoqueMinimo: minimo ? parseInt(minimo.value) || 0 : 0
                });
            }
        });
    } else if (tipo === 'financeiro') {
        const categorias = document.querySelectorAll('#categoriasList .campo-dinamico');
        categorias.forEach((categoria, index) => {
            const nome = categoria.querySelector('input[name="categoria_nome[]"]');
            
            if (nome && nome.value.trim()) {
                dados.push({
                    data: new Date().toLocaleDateString('pt-BR'),
                    categoria: nome.value.trim(),
                    descricao: `Lançamento inicial - ${nome.value.trim()}`,
                    valor: 0,
                    tipo: index % 2 === 0 ? 'Receita' : 'Despesa'
                });
            }
        });
    } else if (tipo === 'vendas') {
        const vendedores = document.querySelectorAll('#vendedoresList .campo-dinamico');
        vendedores.forEach((vendedor, index) => {
            const nome = vendedor.querySelector('input[name="vendedor_nome[]"]');
            const comissao = vendedor.querySelector('input[name="vendedor_comissao[]"]');
            const meta = vendedor.querySelector('input[name="vendedor_meta[]"]');
            
            if (nome && nome.value.trim()) {
                dados.push({
                    vendedor: nome.value.trim(),
                    produto: 'A definir',
                    quantidade: 0,
                    valorUnitario: 0,
                    comissaoPercent: comissao ? parseFloat(comissao.value) || 5 : 5,
                    meta: meta ? parseFloat(meta.value) || 0 : 0
                });
            }
        });
    } else if (tipo === 'clientes') {
        const segmentos = document.querySelectorAll('#segmentosList .campo-dinamico');
        segmentos.forEach((segmento, index) => {
            const nome = segmento.querySelector('input[name="segmento_nome[]"]');
            
            if (nome && nome.value.trim()) {
                dados.push({
                    nome: `Cliente ${nome.value.trim()} ${index + 1}`,
                    email: `cliente${index + 1}@exemplo.com`,
                    telefone: '(00) 00000-0000',
                    endereco: 'Endereço a definir',
                    segmento: nome.value.trim(),
                    status: 'Ativo',
                    ultimaCompra: new Date().toLocaleDateString('pt-BR')
                });
            }
        });
    } else if (tipo === 'projetos') {
        const projetos = document.querySelectorAll('#projetosList .campo-dinamico');
        projetos.forEach((projeto, index) => {
            const nome = projeto.querySelector('input[name="projeto_nome[]"]');
            const responsavel = projeto.querySelector('input[name="projeto_responsavel[]"]');
            
            if (nome && nome.value.trim()) {
                dados.push({
                    projeto: nome.value.trim(),
                    tarefa: `Tarefa inicial - ${nome.value.trim()}`,
                    responsavel: responsavel ? responsavel.value.trim() : 'A definir',
                    inicio: new Date().toLocaleDateString('pt-BR'),
                    fim: new Date(Date.now() + 30*24*60*60*1000).toLocaleDateString('pt-BR'),
                    status: 'Em planejamento',
                    progresso: 0
                });
            }
        });
    } else if (tipo === 'agenda') {
        // Para agenda, criar eventos baseados nas configurações
        dados.push({
            data: new Date().toLocaleDateString('pt-BR'),
            hora: '09:00',
            evento: 'Evento exemplo',
            local: 'Local a definir',
            participantes: 'Participantes a definir',
            status: 'Agendado',
            observacoes: 'Observações a definir'
        });
    }
    
    console.log('Dados coletados:', dados);
    return dados;
}

function validarCamposObrigatorios() {
    const camposObrigatorios = document.querySelectorAll('input[required], select[required]');
    let todosValidos = true;
    
    // Remover classes de erro anteriores
    camposObrigatorios.forEach(campo => {
        campo.classList.remove('invalid');
    });
    
    // Validar cada campo
    camposObrigatorios.forEach(campo => {
        if (!campo.value.trim()) {
            campo.classList.add('invalid');
            todosValidos = false;
        }
    });
    
    // Validar se há pelo menos um item nos campos dinâmicos
    const tipo = document.getElementById('tipoAtual').value;
    if (tipo === 'estoque') {
        const produtos = document.querySelectorAll('#produtosList input[name="produto_nome[]"]');
        let temProdutoValido = false;
        produtos.forEach(produto => {
            if (produto.value.trim()) {
                temProdutoValido = true;
            }
        });
        if (!temProdutoValido) {
            todosValidos = false;
        }
    }
    
    // Mostrar/ocultar alerta
    if (!todosValidos) {
        validationAlert.classList.remove('hidden');
        validationAlert.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else {
        validationAlert.classList.add('hidden');
    }
    
    return todosValidos;
}

// ===== NAVEGAÇÃO =====

function toggleMobileMenu() {
    menuToggle.classList.toggle('active');
    nav.classList.toggle('active');
}

function handleNavigation(e) {
    e.preventDefault();
    const targetSection = e.target.dataset.section;
    showSection(targetSection);
    
    // Fechar menu mobile
    menuToggle.classList.remove('active');
    nav.classList.remove('active');
    
    // Atualizar nav ativo
    navLinks.forEach(link => link.classList.remove('active'));
    e.target.classList.add('active');
}

function showSection(sectionId) {
    sections.forEach(section => {
        section.classList.remove('active');
    });
    document.getElementById(sectionId).classList.add('active');
}

// ===== FORMULÁRIO =====

function initFormulario(tipo) {
    currentTipo = tipo;
    tipoAtual.value = tipo;
    const data = planilhaData[tipo];
    
    formularioTitulo.textContent = `Criar ${data.nome}`;
    generateCamposEspecificos(tipo);
    showSection('formulario');
    
    // Limpar dados anteriores
    dadosUsuarioReais = [];
    validationAlert.classList.add('hidden');
}

function generateCamposEspecificos(tipo) {
    const data = planilhaData[tipo];
    camposEspecificos.innerHTML = '';
    
    // Campos específicos baseados no tipo
    if (tipo === 'estoque') {
        generateEstoqueCampos();
    } else if (tipo === 'financeiro') {
        generateFinanceiroCampos();
    } else if (tipo === 'vendas') {
        generateVendasCampos();
    } else if (tipo === 'clientes') {
        generateClientesCampos();
    } else if (tipo === 'projetos') {
        generateProjetosCampos();
    } else if (tipo === 'agenda') {
        generateAgendaCampos();
    }
}

function generateEstoqueCampos() {
    const html = `
        <div class="form-group">
            <label class="form-label campo-obrigatorio">Produtos do seu estoque</label>
            <p class="data-source-indicator">Digite os produtos reais do seu estoque. Estes dados aparecerão exatamente como digitados na planilha.</p>
            <div id="produtosList">
                <div class="campo-dinamico">
                    <div class="campo-dinamico-header">
                        <h4>Produto 1</h4>
                        <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
                    </div>
                    <div class="campo-grid">
                        <input type="text" class="form-control" placeholder="Nome do produto *" name="produto_nome[]" required>
                        <input type="text" class="form-control" placeholder="Código (ex: PROD001)" name="produto_codigo[]">
                        <input type="number" class="form-control" placeholder="Estoque inicial" name="produto_inicial[]" min="0">
                        <input type="number" class="form-control" placeholder="Estoque mínimo" name="produto_minimo[]" min="0">
                    </div>
                </div>
            </div>
            <button type="button" class="btn btn--secondary btn--sm" onclick="addProduto()">+ Adicionar Produto</button>
        </div>
    `;
    camposEspecificos.innerHTML = html;
}

function generateFinanceiroCampos() {
    const html = `
        <div class="form-group">
            <label class="form-label">Tipo de Controle</label>
            <select class="form-control" name="tipo_financeiro" required>
                <option value="">Selecione...</option>
                <option value="completo">Receitas e Despesas</option>
                <option value="receitas">Apenas Receitas</option>
                <option value="despesas">Apenas Despesas</option>
            </select>
        </div>
        <div class="form-group">
            <label class="form-label campo-obrigatorio">Suas categorias financeiras</label>
            <p class="data-source-indicator">Digite as categorias financeiras do seu negócio. Aparecerão exatamente assim na planilha.</p>
            <div id="categoriasList">
                <div class="campo-dinamico">
                    <div class="campo-dinamico-header">
                        <h4>Categoria 1</h4>
                        <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
                    </div>
                    <input type="text" class="form-control" placeholder="Nome da categoria (ex: Vendas, Marketing)" name="categoria_nome[]" required>
                </div>
            </div>
            <button type="button" class="btn btn--secondary btn--sm" onclick="addCategoria()">+ Adicionar Categoria</button>
        </div>
    `;
    camposEspecificos.innerHTML = html;
}

function generateVendasCampos() {
    const html = `
        <div class="form-group">
            <label class="form-label campo-obrigatorio">Seus vendedores</label>
            <p class="data-source-indicator">Digite os nomes reais dos seus vendedores. Aparecerão exatamente assim na planilha.</p>
            <div id="vendedoresList">
                <div class="campo-dinamico">
                    <div class="campo-dinamico-header">
                        <h4>Vendedor 1</h4>
                        <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
                    </div>
                    <div class="campo-grid">
                        <input type="text" class="form-control" placeholder="Nome do vendedor *" name="vendedor_nome[]" required>
                        <input type="number" class="form-control" placeholder="% Comissão (ex: 5)" name="vendedor_comissao[]" min="0" max="100" step="0.1">
                        <input type="number" class="form-control" placeholder="Meta mensal (R$)" name="vendedor_meta[]" min="0">
                    </div>
                </div>
            </div>
            <button type="button" class="btn btn--secondary btn--sm" onclick="addVendedor()">+ Adicionar Vendedor</button>
        </div>
    `;
    camposEspecificos.innerHTML = html;
}

function generateClientesCampos() {
    const html = `
        <div class="form-group">
            <label class="form-label campo-obrigatorio">Segmentos de clientes</label>
            <p class="data-source-indicator">Digite os segmentos reais dos seus clientes (ex: Pessoa Física, Empresas, etc.).</p>
            <div id="segmentosList">
                <div class="campo-dinamico">
                    <div class="campo-dinamico-header">
                        <h4>Segmento 1</h4>
                        <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
                    </div>
                    <input type="text" class="form-control" placeholder="Nome do segmento (ex: Pessoa Física)" name="segmento_nome[]" required>
                </div>
            </div>
            <button type="button" class="btn btn--secondary btn--sm" onclick="addSegmento()">+ Adicionar Segmento</button>
        </div>
        <div class="form-group">
            <label class="form-label">Campos Opcionais</label>
            <div class="campo-grid">
                <label class="checkbox-label">
                    <input type="checkbox" name="incluir_endereco" checked> Incluir Endereço
                </label>
                <label class="checkbox-label">
                    <input type="checkbox" name="incluir_observacoes" checked> Incluir Observações
                </label>
            </div>
        </div>
    `;
    camposEspecificos.innerHTML = html;
}

function generateProjetosCampos() {
    const html = `
        <div class="form-group">
            <label class="form-label campo-obrigatorio">Seus projetos</label>
            <p class="data-source-indicator">Digite os nomes reais dos seus projetos e responsáveis.</p>
            <div id="projetosList">
                <div class="campo-dinamico">
                    <div class="campo-dinamico-header">
                        <h4>Projeto 1</h4>
                        <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
                    </div>
                    <div class="campo-grid">
                        <input type="text" class="form-control" placeholder="Nome do projeto *" name="projeto_nome[]" required>
                        <input type="text" class="form-control" placeholder="Responsável" name="projeto_responsavel[]">
                    </div>
                </div>
            </div>
            <button type="button" class="btn btn--secondary btn--sm" onclick="addProjeto()">+ Adicionar Projeto</button>
        </div>
    `;
    camposEspecificos.innerHTML = html;
}

function generateAgendaCampos() {
    const html = `
        <div class="form-group">
            <label class="form-label">Tipo de Agenda</label>
            <select class="form-control" name="tipo_agenda" required>
                <option value="">Selecione...</option>
                <option value="pessoal">Agenda Pessoal</option>
                <option value="trabalho">Agenda de Trabalho</option>
                <option value="eventos">Eventos/Reuniões</option>
            </select>
        </div>
        <div class="form-group">
            <label class="form-label">Período</label>
            <div class="campo-grid">
                <input type="date" class="form-control" name="data_inicio">
                <input type="date" class="form-control" name="data_fim">
            </div>
        </div>
    `;
    camposEspecificos.innerHTML = html;
}

// ===== FUNÇÕES PARA ADICIONAR ITENS DINÂMICOS =====

function addProduto() {
    const container = document.getElementById('produtosList');
    const count = container.children.length + 1;
    const html = `
        <div class="campo-dinamico">
            <div class="campo-dinamico-header">
                <h4>Produto ${count}</h4>
                <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
            </div>
            <div class="campo-grid">
                <input type="text" class="form-control" placeholder="Nome do produto *" name="produto_nome[]" required>
                <input type="text" class="form-control" placeholder="Código (ex: PROD${count.toString().padStart(3, '0')})" name="produto_codigo[]">
                <input type="number" class="form-control" placeholder="Estoque inicial" name="produto_inicial[]" min="0">
                <input type="number" class="form-control" placeholder="Estoque mínimo" name="produto_minimo[]" min="0">
            </div>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', html);
}

function addCategoria() {
    const container = document.getElementById('categoriasList');
    const count = container.children.length + 1;
    const html = `
        <div class="campo-dinamico">
            <div class="campo-dinamico-header">
                <h4>Categoria ${count}</h4>
                <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
            </div>
            <input type="text" class="form-control" placeholder="Nome da categoria" name="categoria_nome[]" required>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', html);
}

function addVendedor() {
    const container = document.getElementById('vendedoresList');
    const count = container.children.length + 1;
    const html = `
        <div class="campo-dinamico">
            <div class="campo-dinamico-header">
                <h4>Vendedor ${count}</h4>
                <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
            </div>
            <div class="campo-grid">
                <input type="text" class="form-control" placeholder="Nome do vendedor *" name="vendedor_nome[]" required>
                <input type="number" class="form-control" placeholder="% Comissão (ex: 5)" name="vendedor_comissao[]" min="0" max="100" step="0.1">
                <input type="number" class="form-control" placeholder="Meta mensal (R$)" name="vendedor_meta[]" min="0">
            </div>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', html);
}

function addSegmento() {
    const container = document.getElementById('segmentosList');
    const count = container.children.length + 1;
    const html = `
        <div class="campo-dinamico">
            <div class="campo-dinamico-header">
                <h4>Segmento ${count}</h4>
                <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
            </div>
            <input type="text" class="form-control" placeholder="Nome do segmento" name="segmento_nome[]" required>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', html);
}

function addProjeto() {
    const container = document.getElementById('projetosList');
    const count = container.children.length + 1;
    const html = `
        <div class="campo-dinamico">
            <div class="campo-dinamico-header">
                <h4>Projeto ${count}</h4>
                <button type="button" class="btn btn--outline btn--sm remove-item">×</button>
            </div>
            <div class="campo-grid">
                <input type="text" class="form-control" placeholder="Nome do projeto *" name="projeto_nome[]" required>
                <input type="text" class="form-control" placeholder="Responsável" name="projeto_responsavel[]">
            </div>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', html);
}

function removeItemDinamico(button) {
    const item = button.closest('.campo-dinamico');
    item.remove();
}

// ===== CAMPOS PERSONALIZADOS =====

function addCampoPersonalizado() {
    const html = `
        <div class="campo-personalizado">
            <input type="text" class="form-control" placeholder="Nome do campo personalizado">
            <button type="button" class="btn btn--outline btn--sm remove-campo">×</button>
        </div>
    `;
    camposPersonalizados.insertAdjacentHTML('beforeend', html);
}

function removeCampoPersonalizado(button) {
    button.parentElement.remove();
}

// ===== PREVIEW COM DADOS REAIS =====

function showPreviewComDadosReais(e) {
    e.preventDefault();
    
    // Validar campos obrigatórios
    if (!validarCamposObrigatorios()) {
        return;
    }
    
    // Coletar dados reais do formulário
    dadosUsuarioReais = coletarDadosFormulario();
    
    if (dadosUsuarioReais.length === 0) {
        alert('Por favor, adicione pelo menos um item antes de continuar.');
        return;
    }
    
    // Capturar outros dados do formulário
    const formData = new FormData(planilhaForm);
    currentFormData = Object.fromEntries(formData.entries());
    
    // Renderizar preview com dados reais
    renderPreviewComDadosReais();
    showSection('preview');
}

function renderPreviewComDadosReais() {
    const data = planilhaData[currentTipo];
    let html = '<thead><tr>';
    
    // Cabeçalhos
    data.campos.forEach(campo => {
        const nome = data.camposNomes[campo] || campo.replace(/_/g, ' ').toUpperCase();
        html += `<th>${nome}</th>`;
    });
    
    // Campos personalizados
    const camposPersonalizadosInputs = document.querySelectorAll('#camposPersonalizados input[type="text"]');
    camposPersonalizadosInputs.forEach(input => {
        if (input.value.trim()) {
            html += `<th>${input.value.trim()}</th>`;
        }
    });
    
    html += '</tr></thead><tbody>';
    
    // Dados reais do usuário
    dadosUsuarioReais.forEach((item, index) => {
        html += '<tr class="user-data">';
        
        data.campos.forEach(campo => {
            let valor = '';
            const rowNum = index + 2;
            
            switch (campo) {
                case 'codigo':
                    valor = item.codigo || `ITEM${(index + 1).toString().padStart(3, '0')}`;
                    break;
                case 'produto':
                case 'nome':
                    valor = item.nome || item.produto || '';
                    break;
                case 'estoque_inicial':
                    valor = item.estoqueInicial || 0;
                    break;
                case 'entradas':
                    valor = item.entradas || 0;
                    break;
                case 'saidas':
                    valor = item.saidas || 0;
                    break;
                case 'saldo_atual':
                    valor = `=C${rowNum}+D${rowNum}-E${rowNum}`;
                    break;
                case 'estoque_minimo':
                    valor = item.estoqueMinimo || 0;
                    break;
                case 'status':
                    valor = data.formulas && data.formulas.status ? data.formulas.status.replace(/\{row\}/g, rowNum) : 'OK';
                    break;
                case 'data':
                    valor = item.data || new Date().toLocaleDateString('pt-BR');
                    break;
                case 'categoria':
                    valor = item.categoria || '';
                    break;
                case 'descricao':
                    valor = item.descricao || '';
                    break;
                case 'valor':
                    valor = item.valor || 0;
                    break;
                case 'tipo':
                    valor = item.tipo || '';
                    break;
                case 'vendedor':
                    valor = item.vendedor || '';
                    break;
                case 'quantidade':
                    valor = item.quantidade || 0;
                    break;
                case 'valor_unitario':
                    valor = item.valorUnitario || 0;
                    break;
                case 'total':
                    valor = data.formulas && data.formulas.total ? data.formulas.total.replace(/\{row\}/g, rowNum) : 0;
                    break;
                case 'comissao':
                    valor = data.formulas && data.formulas.comissao ? data.formulas.comissao.replace(/\{row\}/g, rowNum) : 0;
                    break;
                default:
                    valor = item[campo] || '';
            }
            
            const cellClass = typeof valor === 'string' && valor.startsWith('=') ? 'formula-cell' : '';
            html += `<td class="${cellClass}">${valor}</td>`;
        });
        
        // Campos personalizados
        camposPersonalizadosInputs.forEach(() => {
            html += `<td class="user-input">Personalizado ${index + 1}</td>`;
        });
        
        html += '</tr>';
    });
    
    html += '</tbody>';
    previewTable.innerHTML = html;
    
    // Aplicar estilo customizado
    applyTableStyle();
}

// ===== PERSONALIZAÇÃO =====

function applyCustomization() {
    const headerColor = document.getElementById('headerColor').value;
    const tableStyle = document.getElementById('tableStyle').value;
    
    customSettings.headerColor = headerColor;
    customSettings.tableStyle = tableStyle;
    
    applyTableStyle();
    closeModal();
}

function applyTableStyle() {
    previewTable.className = `preview-table table-${customSettings.tableStyle}`;
    const headers = previewTable.querySelectorAll('th');
    headers.forEach(th => {
        th.style.backgroundColor = customSettings.headerColor;
    });
}

function closeModal() {
    customModal.classList.add('hidden');
}

// ===== GERAÇÃO DO EXCEL COM DADOS REAIS =====

function handleFormSubmit(e) {
    e.preventDefault();
    showPreviewComDadosReais(e);
}

function gerarPlanilhaComDadosReais() {
    if (dadosUsuarioReais.length === 0) {
        alert('Nenhum dado encontrado. Por favor, preencha o formulário primeiro.');
        return;
    }
    
    loading.classList.remove('hidden');
    
    setTimeout(() => {
        try {
            const wb = XLSX.utils.book_new();
            const data = planilhaData[currentTipo];
            
            // Preparar dados reais para o Excel
            const wsData = prepareExcelDataComDadosReais();
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            
            // Aplicar formatação
            applyExcelFormatting(ws, wsData.length);
            
            // Aplicar fórmulas
            applyExcelFormulasComDadosReais(ws, wsData.length);
            
            // Adicionar worksheet
            XLSX.utils.book_append_sheet(wb, ws, data.nome);
            
            // Gerar e baixar arquivo
            const nomeEmpresa = currentFormData.empresa || 'MinhaEmpresa';
            const fileName = `${data.nome.replace(/\s+/g, '_')}_${nomeEmpresa.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(wb, fileName);
            
            loading.classList.add('hidden');
            alert('Planilha gerada com sucesso com seus dados reais!');
            showSection('home');
            
        } catch (error) {
            console.error('Erro ao gerar planilha:', error);
            loading.classList.add('hidden');
            alert('Erro ao gerar planilha. Tente novamente.');
        }
    }, 1000);
}

function prepareExcelDataComDadosReais() {
    const data = planilhaData[currentTipo];
    const wsData = [];
    
    // Cabeçalho com nomes amigáveis
    const headers = data.campos.map(campo => 
        data.camposNomes[campo] || campo.replace(/_/g, ' ').toUpperCase()
    );
    
    // Adicionar campos personalizados
    const camposPersonalizadosInputs = document.querySelectorAll('#camposPersonalizados input[type="text"]');
    camposPersonalizadosInputs.forEach(input => {
        if (input.value.trim()) {
            headers.push(input.value.trim());
        }
    });
    
    wsData.push(headers);
    
    // Adicionar dados reais do usuário
    dadosUsuarioReais.forEach((item, index) => {
        const row = [];
        
        data.campos.forEach(campo => {
            switch (campo) {
                case 'codigo':
                    row.push(item.codigo || `ITEM${(index + 1).toString().padStart(3, '0')}`);
                    break;
                case 'produto':
                case 'nome':
                    row.push(item.nome || item.produto || '');
                    break;
                case 'estoque_inicial':
                    row.push(parseInt(item.estoqueInicial) || 0);
                    break;
                case 'entradas':
                    row.push(parseInt(item.entradas) || 0);
                    break;
                case 'saidas':
                    row.push(parseInt(item.saidas) || 0);
                    break;
                case 'saldo_atual':
                    row.push(''); // Será preenchido com fórmula
                    break;
                case 'estoque_minimo':
                    row.push(parseInt(item.estoqueMinimo) || 0);
                    break;
                case 'status':
                    row.push(''); // Será preenchido com fórmula
                    break;
                case 'data':
                    row.push(item.data || new Date().toLocaleDateString('pt-BR'));
                    break;
                case 'categoria':
                    row.push(item.categoria || '');
                    break;
                case 'descricao':
                    row.push(item.descricao || '');
                    break;
                case 'valor':
                    row.push(parseFloat(item.valor) || 0);
                    break;
                case 'tipo':
                    row.push(item.tipo || '');
                    break;
                case 'vendedor':
                    row.push(item.vendedor || '');
                    break;
                case 'quantidade':
                    row.push(parseInt(item.quantidade) || 0);
                    break;
                case 'valor_unitario':
                    row.push(parseFloat(item.valorUnitario) || 0);
                    break;
                case 'total':
                    row.push(''); // Será preenchido com fórmula
                    break;
                case 'comissao':
                    row.push(''); // Será preenchido com fórmula
                    break;
                default:
                    row.push(item[campo] || '');
            }
        });
        
        // Campos personalizados
        camposPersonalizadosInputs.forEach((input, i) => {
            if (input.value.trim()) {
                row.push(`Valor personalizado ${index + 1}`);
            }
        });
        
        wsData.push(row);
    });
    
    return wsData;
}

function applyExcelFormatting(ws, dataLength) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    // Formatação do cabeçalho
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddr = XLSX.utils.encode_cell({r: 0, c: col});
        if (!ws[cellAddr]) continue;
        
        ws[cellAddr].s = {
            fill: {
                fgColor: { rgb: customSettings.headerColor.replace('#', '') }
            },
            font: {
                bold: true,
                color: { rgb: "FFFFFF" }
            },
            alignment: {
                horizontal: "center",
                vertical: "center"
            }
        };
    }
    
    // Auto-ajustar colunas
    const cols = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        cols.push({ wch: 15 });
    }
    ws['!cols'] = cols;
}

function applyExcelFormulasComDadosReais(ws, dataLength) {
    const data = planilhaData[currentTipo];
    if (!data.formulas) return;
    
    const headers = Object.keys(data.camposNomes);
    
    for (let row = 1; row < dataLength; row++) {
        headers.forEach((campo, colIndex) => {
            const formula = data.formulas[campo];
            
            if (formula) {
                const cellAddr = XLSX.utils.encode_cell({r: row, c: colIndex});
                const processedFormula = formula
                    .replace(/\{row\}/g, row + 1)
                    .replace(/\{prevRow\}/g, row);
                
                ws[cellAddr] = {
                    f: processedFormula.substring(1) // Remove o =
                };
            }
        });
    }
}

// Expor funções globais para os botões dinâmicos
window.addProduto = addProduto;
window.addCategoria = addCategoria;
window.addVendedor = addVendedor;
window.addSegmento = addSegmento;
window.addProjeto = addProjeto;