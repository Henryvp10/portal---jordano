let slideIndex = 0;
let carrosselTimer = null; // Guarda o temporizador do carrossel

// IP fixo do servidor backend em Python (Flask)
const SERVER_IP = '192.168.0.60';
const PORT = '5000';

// Endereços do Backend em Python (Flask)
const API_URL_MOTORISTA = `http://${SERVER_IP}:${PORT}/api/enviar-motorista`;
const API_URL_VEICULO = `http://${SERVER_IP}:${PORT}/api/enviar-veiculo`;
const API_URL_LICENCAS = `http://${SERVER_IP}:${PORT}/api/enviar-licencas`;
const API_URL_VENCIMENTOS = `http://${SERVER_IP}:${PORT}/api/enviar-vencimentos`;

// Novas URLs para consulta de vencimentos (ACRESCENTADO)
const API_URL_CONSULTA_VENC_MOTORISTAS = `http://${SERVER_IP}:${PORT}/api/vencimentos/motoristas`;
const API_URL_CONSULTA_VENC_VEICULOS = `http://${SERVER_IP}:${PORT}/api/vencimentos/veiculos`;

// Executa ao carregar a página
document.addEventListener('DOMContentLoaded', () => {
    // Só inicia o carrossel se a tela possuir slides (ex: página index.html)
    if (document.querySelectorAll('.carousel-slide').length > 0) {
        iniciarCarrossel();
    }
});

// FUNÇÃO PARA ABRIR E FECHAR O SUBMENU (ACESSO DOCUMENTAÇÃO)
function toggleSubmenu(button) {
    const dropdown = button.closest('.nav-dropdown');
    if (dropdown) {
        dropdown.classList.toggle('open');
    }
}

// Autenticação e Redirecionamento da Tela de Login
function autenticarUsuario(event) {
    event.preventDefault();
    
    const usuario = document.getElementById('usuarioLogin').value.trim();
    const senha = document.getElementById('senhaLogin').value.trim();

    if (usuario !== "" && senha !== "") {
        // Redireciona diretamente para o painel principal
        window.location.href = "index.html";
    } else {
        alert("Por favor, preencha os campos de usuário e senha.");
    }
}

// Alternar exibição das abas (HOME, MOTORISTA, VEÍCULOS, LICENÇAS, VENCIMENTOS MOTORISTA, VENCIMENTOS VEÍCULO, etc.)
function selecionarAba(aba) {
    const painelHome = document.getElementById('painelHome');
    const painelBusca = document.getElementById('painelBusca');
    const painelVeiculo = document.getElementById('painelVeiculo');
    const painelLicencas = document.getElementById('painelLicencas');
    
    // Novos painéis de Vencimento
    const painelVencMotorista = document.getElementById('painelVencimentoMotorista');
    const painelVencVeiculo = document.getElementById('painelVencimentoVeiculo');
    
    // Painéis legados/interditados
    const painelVencimentos = document.getElementById('painelVencimentos');
    const painelPerdasSobras = document.getElementById('painelPerdasSobras');
    const painelEstadia = document.getElementById('painelEstadia');
    
    // Remove classe ativa de todos os botões do menu
    document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('active'));

    // Oculta todos os painéis
    if (painelHome) painelHome.style.display = 'none';
    if (painelBusca) painelBusca.style.display = 'none';
    if (painelVeiculo) painelVeiculo.style.display = 'none';
    if (painelLicencas) painelLicencas.style.display = 'none';
    if (painelVencMotorista) painelVencMotorista.style.display = 'none';
    if (painelVencVeiculo) painelVencVeiculo.style.display = 'none';
    if (painelVencimentos) painelVencimentos.style.display = 'none';
    if (painelPerdasSobras) painelPerdasSobras.style.display = 'none';
    if (painelEstadia) painelEstadia.style.display = 'none';

    // Ativa o painel correspondente ao clique no menu
    if (aba === 'home') {
        const btnHome = document.getElementById('btnHome');
        if (btnHome) btnHome.classList.add('active');
        if (painelHome) {
            painelHome.style.display = 'block';
            iniciarCarrossel();
        }
    } else if (aba === 'motorista') {
        const btnMotorista = document.getElementById('btnMotorista');
        if (btnMotorista) btnMotorista.classList.add('active');
        if (painelBusca) painelBusca.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'veiculo') {
        const btnVeiculo = document.getElementById('btnVeiculo');
        if (btnVeiculo) btnVeiculo.classList.add('active');
        if (painelVeiculo) painelVeiculo.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'licencas') {
        const btnLicencas = document.getElementById('btnLicencas');
        if (btnLicencas) btnLicencas.classList.add('active');
        if (painelLicencas) painelLicencas.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'vencMotorista') {
        const btnVencMotorista = document.getElementById('btnVencMotorista');
        if (btnVencMotorista) btnVencMotorista.classList.add('active');
        if (painelVencMotorista) painelVencMotorista.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'vencVeiculo') {
        const btnVencVeiculo = document.getElementById('btnVencVeiculo');
        if (btnVencVeiculo) btnVencVeiculo.classList.add('active');
        if (painelVencVeiculo) painelVencVeiculo.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'vencimentos') {
        const btnVencimentos = document.getElementById('btnVencimentos');
        if (btnVencimentos) btnVencimentos.classList.add('active');
        if (painelVencimentos) painelVencimentos.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'perdasSobras') {
        const btnPerdasSobras = document.getElementById('btnPerdasSobras');
        if (btnPerdasSobras) btnPerdasSobras.classList.add('active');
        if (painelPerdasSobras) painelPerdasSobras.style.display = 'block';
        pararCarrosselAuto();
    } else if (aba === 'estadia') {
        const btnEstadia = document.getElementById('btnEstadia');
        if (btnEstadia) btnEstadia.classList.add('active');
        if (painelEstadia) painelEstadia.style.display = 'block';
        pararCarrosselAuto();
    }
}

// Máscara de CPF para o formulário do Motorista
function mascaraCPF(input) {
    let value = input.value.replace(/\D/g, "");
    if (value.length > 11) value = value.slice(0, 11);
    
    value = value.replace(/(\d{3})(\d)/, "$1.$2");
    value = value.replace(/(\d{3})(\d)/, "$1.$2");
    value = value.replace(/(\d{3})(\d{1,2})$/, "$1-$2");
    
    input.value = value;
}

// Selecionar todos os documentos do Motorista
function marcarTodosDocs(source) {
    const checkboxes = document.querySelectorAll('.doc-check');
    checkboxes.forEach(chk => chk.checked = source.checked);
}

// Selecionar todos os documentos do Veículo
function marcarTodosDocsVeiculo(source) {
    const checkboxes = document.querySelectorAll('.doc-check-veiculo');
    checkboxes.forEach(chk => chk.checked = source.checked);
}

// Selecionar todos os documentos das Licenças
function marcarTodosDocsLicencas(source) {
    const checkboxes = document.querySelectorAll('.doc-check-licencas');
    checkboxes.forEach(chk => chk.checked = source.checked);
}

// Função para Limpar o Formulário de Motorista
function limparFormularioMotorista() {
    document.getElementById('nomeMotorista').value = '';
    document.getElementById('cpfMotorista').value = '';
    
    const checkboxes = document.querySelectorAll('.doc-check');
    checkboxes.forEach(chk => chk.checked = false);
    
    const chkTodos = document.getElementById('chkTodosDocs');
    if (chkTodos) chkTodos.checked = false;
}

// Função para Limpar o Formulário de Veículo
function limparFormularioVeiculo() {
    document.getElementById('placaVeiculo').value = '';
    
    const checkboxes = document.querySelectorAll('.doc-check-veiculo');
    checkboxes.forEach(chk => chk.checked = false);
    
    const chkTodos = document.getElementById('chkTodosDocsVeiculo');
    if (chkTodos) chkTodos.checked = false;
}

// Função para Limpar o Formulário de Licenças
function limparFormularioLicencas() {
    if (document.getElementById('estadoLicenca')) document.getElementById('estadoLicenca').selectedIndex = 0;
    
    const checkboxes = document.querySelectorAll('.doc-check-licencas');
    checkboxes.forEach(chk => chk.checked = false);
    
    const chkTodos = document.getElementById('chkTodosDocsLicencas');
    if (chkTodos) chkTodos.checked = false;
}

// FUNÇÕES DE LIMPEZA PARA A CONSULTA DE VENCIMENTOS (ACRESCENTADO)
function limparVencimentoMotorista() {
    if (document.getElementById('nomeVencimentoMotorista')) {
        document.getElementById('nomeVencimentoMotorista').value = '';
    }
    if (document.getElementById('cpfVencimentoMotorista')) {
        document.getElementById('cpfVencimentoMotorista').value = '';
    }
    const divResultado = document.getElementById('resultadoVencimentos');
    if (divResultado) {
        divResultado.innerHTML = '';
    }
}

function limparVencimentoVeiculo() {
    if (document.getElementById('placaVencimentoVeiculo')) {
        document.getElementById('placaVencimentoVeiculo').value = '';
    }
    const divResultado = document.getElementById('resultadoVencimentosVeiculo');
    if (divResultado) {
        divResultado.innerHTML = '';
    }
}

// Envio do formulário Motorista (Comunicação com o Servidor Python via API)
async function enviarSolicitacaoMotorista() {
    const nome = document.getElementById('nomeMotorista').value.trim();
    const cpf = document.getElementById('cpfMotorista').value.trim();
    const email = document.getElementById('emailSolicitante').value.trim();
    const docs = Array.from(document.querySelectorAll('.doc-check:checked')).map(c => c.value);

    if (!cpf && !nome) {
        alert('Por favor, informe o Nome ou o CPF do condutor.');
        return;
    }

    if (!email) {
        alert('Por favor, informe o e-mail do destinatário.');
        return;
    }

    if (docs.length === 0) {
        alert('Por favor, selecione ao menos um documento.');
        return;
    }

    const btnSubmit = document.querySelector('#formMotorista .btn-send-email');
    const textoOriginal = btnSubmit ? btnSubmit.innerHTML : '';

    try {
        if (btnSubmit) {
            btnSubmit.disabled = true;
            btnSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buscando e Enviando...';
        }

        const resposta = await fetch(API_URL_MOTORISTA, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                nome: nome,
                cpf: cpf,
                email: email,
                documentos: docs
            })
        });

        const resultado = await resposta.json();

        if (resposta.ok && resultado.sucesso) {
            alert(`Sucesso! ${resultado.mensagem}`);
        } else {
            alert(`Atenção: ${resultado.erro || 'Ocorreu um erro na requisição.'}`);
        }
    } catch (erro) {
        console.error('Erro na requisição:', erro);
        alert(`Não foi possível conectar ao servidor Python em (${API_URL_MOTORISTA}). Verifique se o app.py está em execução.`);
    } finally {
        if (btnSubmit) {
            btnSubmit.disabled = false;
            btnSubmit.innerHTML = textoOriginal;
        }
    }
}

// Envio do formulário Veículo (Comunicação com o Servidor Python via API)
async function enviarSolicitacaoVeiculo() {
    const placa = document.getElementById('placaVeiculo').value.trim().toUpperCase();
    const email = document.getElementById('emailSolicitanteVeiculo').value.trim();
    const docs = Array.from(document.querySelectorAll('.doc-check-veiculo:checked')).map(c => c.value);

    if (!placa) {
        alert('Por favor, informe a placa do veículo.');
        return;
    }

    if (!email) {
        alert('Por favor, informe o e-mail do destinatário.');
        return;
    }

    if (docs.length === 0) {
        alert('Por favor, selecione ao menos um documento do veículo.');
        return;
    }

    const btnSubmit = document.querySelector('#formVeiculo .btn-send-email');
    const textoOriginal = btnSubmit ? btnSubmit.innerHTML : '';

    try {
        if (btnSubmit) {
            btnSubmit.disabled = true;
            btnSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buscando e Enviando...';
        }

        const resposta = await fetch(API_URL_VEICULO, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                placa: placa,
                email: email,
                documentos: docs
            })
        });

        const resultado = await resposta.json();

        if (resposta.ok && resultado.sucesso) {
            alert(`Sucesso! ${resultado.mensagem}`);
        } else {
            alert(`Atenção: ${resultado.erro || 'Ocorreu um erro na requisição.'}`);
        }
    } catch (erro) {
        console.error('Erro na requisição:', erro);
        alert(`Não foi possível conectar ao servidor Python em (${API_URL_VEICULO}). Verifique se o app.py está em execução.`);
    } finally {
        if (btnSubmit) {
            btnSubmit.disabled = false;
            btnSubmit.innerHTML = textoOriginal;
        }
    }
}

// Envio do formulário Licenças (Comunicação com o Servidor Python via API)
async function enviarSolicitacaoLicencas() {
    const estado = document.getElementById('estadoLicenca') ? document.getElementById('estadoLicenca').value : '';
    const email = document.getElementById('emailSolicitanteLicencas') ? document.getElementById('emailSolicitanteLicencas').value.trim() : '';
    const docs = Array.from(document.querySelectorAll('.doc-check-licencas:checked')).map(c => c.value);

    if (!estado || estado === "") {
        alert('Por favor, selecione o estado (UF).');
        return;
    }

    if (!email) {
        alert('Por favor, informe o e-mail do destinatário.');
        return;
    }

    if (docs.length === 0) {
        alert('Por favor, selecione ao menos uma licença.');
        return;
    }

    const btnSubmit = document.querySelector('#formLicencas .btn-send-email');
    const textoOriginal = btnSubmit ? btnSubmit.innerHTML : '';

    try {
        if (btnSubmit) {
            btnSubmit.disabled = true;
            btnSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buscando e Enviando...';
        }

        const resposta = await fetch(API_URL_LICENCAS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                estado: estado,
                email: email,
                documentos: docs
            })
        });

        const resultado = await resposta.json();

        if (resposta.ok && resultado.sucesso) {
            alert(`Sucesso! ${resultado.mensagem}`);
        } else {
            alert(`Atenção: ${resultado.erro || 'Ocorreu um erro na requisição.'}`);
        }
    } catch (erro) {
        console.error('Erro na requisição:', erro);
        alert(`Não foi possível conectar ao servidor Python em (${API_URL_LICENCAS}). Verifique se o app.py está em execução.`);
    } finally {
        if (btnSubmit) {
            btnSubmit.disabled = false;
            btnSubmit.innerHTML = textoOriginal;
        }
    }
}

// Envio do formulário Vencimentos (Comunicação com o Servidor Python via API)
async function enviarSolicitacaoVencimentos() {
    const tipo = document.getElementById('tipoVencimento') ? document.getElementById('tipoVencimento').value : '';
    const email = document.getElementById('emailSolicitanteVencimentos') ? document.getElementById('emailSolicitanteVencimentos').value.trim() : '';

    if (!tipo) {
        alert('Por favor, selecione o tipo de relatório.');
        return;
    }

    if (!email) {
        alert('Por favor, informe o e-mail do destinatário.');
        return;
    }

    const btnSubmit = document.querySelector('#formVencimentos .btn-send-email');
    const textoOriginal = btnSubmit ? btnSubmit.innerHTML : '';

    try {
        if (btnSubmit) {
            btnSubmit.disabled = true;
            btnSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Gerando Relatório...';
        }

        const resposta = await fetch(API_URL_VENCIMENTOS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                tipo: tipo,
                email: email
            })
        });

        const resultado = await resposta.json();

        if (resposta.ok && resultado.sucesso) {
            alert(`Sucesso! ${resultado.mensagem}`);
        } else {
            alert(`Atenção: ${resultado.erro || 'Ocorreu um erro na requisição.'}`);
        }
    } catch (erro) {
        console.error('Erro na requisição:', erro);
        alert(`Não foi possível conectar ao servidor Python em (${API_URL_VENCIMENTOS}). Verifique se o app.py está em execução.`);
    } finally {
        if (btnSubmit) {
            btnSubmit.disabled = false;
            btnSubmit.innerHTML = textoOriginal;
        }
    }
}

// --- CONSULTAS DE VENCIMENTOS (ACRESCENTADAS) ---

// Consulta Vencimentos Motorista por CPF/Nome
async function consultarVencimentoMotorista() {
    const cpf = document.getElementById('cpfVencimentoMotorista') ? document.getElementById('cpfVencimentoMotorista').value.trim() : '';
    const nome = document.getElementById('nomeVencimentoMotorista') ? document.getElementById('nomeVencimentoMotorista').value.trim() : '';
    const divResultado = document.getElementById('resultadoVencimentos');

    if (!cpf) {
        alert('Por favor, informe o CPF do condutor.');
        return;
    }

    try {
        if (divResultado) divResultado.innerHTML = '<p class="carregando"><i class="fa-solid fa-spinner fa-spin"></i> Buscando vencimentos...</p>';

        const resposta = await fetch(`${API_URL_CONSULTA_VENC_MOTORISTAS}?cpf=${encodeURIComponent(cpf)}&nome=${encodeURIComponent(nome)}`);
        const dados = await resposta.json();

        if (resposta.ok) {
            renderizarTabelaVencimentos(dados.documentos, `Condutor: ${dados.condutor} (CPF: ${dados.cpf})`, divResultado);
        } else {
            if (divResultado) divResultado.innerHTML = `<p class="msg-erro">Atenção: ${dados.erro || 'Nenhum registro localizado.'}</p>`;
        }
    } catch (erro) {
        console.error('Erro na busca de vencimentos:', erro);
        if (divResultado) divResultado.innerHTML = '<p class="msg-erro">Erro ao conectar com o servidor Python.</p>';
    }
}

// Consulta Vencimentos Veículo por Placa
async function consultarVencimentoVeiculo() {
    const placa = document.getElementById('placaVencimentoVeiculo') ? document.getElementById('placaVencimentoVeiculo').value.trim().toUpperCase() : '';
    const divResultado = document.getElementById('resultadoVencimentosVeiculo');

    if (!placa) {
        alert('Por favor, informe a placa do veículo.');
        return;
    }

    try {
        if (divResultado) divResultado.innerHTML = '<p class="carregando"><i class="fa-solid fa-spinner fa-spin"></i> Buscando vencimentos...</p>';

        const resposta = await fetch(`${API_URL_CONSULTA_VENC_VEICULOS}?placa=${encodeURIComponent(placa)}`);
        const dados = await resposta.json();

        if (resposta.ok) {
            renderizarTabelaVencimentos(dados.documentos, `Veículo Placa: ${dados.placa}`, divResultado);
        } else {
            if (divResultado) divResultado.innerHTML = `<p class="msg-erro">Atenção: ${dados.erro || 'Nenhum registro localizado.'}</p>`;
        }
    } catch (erro) {
        console.error('Erro na busca de vencimentos:', erro);
        if (divResultado) divResultado.innerHTML = '<p class="msg-erro">Erro ao conectar com o servidor Python.</p>';
    }
}

// Renderiza a tabela de resultados com Badges de Cores
function renderizarTabelaVencimentos(documentos, titulo, container = document.getElementById('resultadoVencimentos')) {
    if (!container) return;

    if (!documentos || documentos.length === 0) {
        container.innerHTML = '<p class="msg-alerta">Nenhum documento encontrado na pasta.</p>';
        return;
    }

    let html = `
        <div class="card-resultado">
            <h3 style="margin-bottom: 15px; color: #005a2b;">${titulo}</h3>
            <table class="tabela-vencimentos">
                <thead>
                    <tr>
                        <th>Documento</th>
                        <th>Data de Validade</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
    `;

    documentos.forEach(doc => {
        let badgeClasse = 'badge-verde';
        if (doc.cor === 'vermelho') badgeClasse = 'badge-vermelho';
        if (doc.cor === 'amarelo') badgeClasse = 'badge-amarelo';
        if (doc.cor === 'cinza') badgeClasse = 'badge-cinza';

        html += `
            <tr>
                <td>${doc.documento}</td>
                <td>${doc.validade}</td>
                <td><span class="badge ${badgeClasse}">${doc.status}</span></td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
        </div>
    `;

    container.innerHTML = html;
}

// Funções do Carrossel (HOME)
function iniciarCarrossel() {
    exhibirSlide(slideIndex);
    iniciarCarrosselAuto();
}

function mudarSlide(n) {
    exhibirSlide(slideIndex += n);
    iniciarCarrosselAuto();
}

function irParaSlide(n) {
    exhibirSlide(slideIndex = n);
    iniciarCarrosselAuto();
}

function exhibirSlide(n) {
    const slides = document.querySelectorAll('.carousel-slide');
    const dots = document.querySelectorAll('.dot');
    
    if (slides.length === 0) return;

    if (n >= slides.length) slideIndex = 0;
    if (n < 0) slideIndex = slides.length - 1;

    slides.forEach(slide => slide.style.display = 'none');
    dots.forEach(dot => dot.classList.remove('active'));

    slides[slideIndex].style.display = 'block';
    if (dots[slideIndex]) dots[slideIndex].classList.add('active');
}

// Transição automática das imagens a cada 3 segundos
function iniciarCarrosselAuto() {
    pararCarrosselAuto();
    carrosselTimer = setInterval(() => {
        mudarSlide(1);
    }, 3000);
}

function pararCarrosselAuto() {
    if (carrosselTimer) {
        clearInterval(carrosselTimer);
    }
}

// Alternar exibição do Menu Lateral em telas menores
function toggleMenu() {
    const sidebar = document.getElementById('sidebar');
    if (sidebar.style.display === 'none' || sidebar.style.display === '') {
        sidebar.style.display = 'flex';
    } else {
        sidebar.style.display = 'none';
    }
}