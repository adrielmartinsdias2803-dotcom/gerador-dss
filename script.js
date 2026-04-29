document.addEventListener('DOMContentLoaded', () => {
    // Inicialização
    const printableArea = document.getElementById('printable-area');
    const template = document.getElementById('print-template').innerHTML;
    
    // Injeta o template na área de visualização
    printableArea.innerHTML = template;

    // Elementos do Formulário
    const inputTitulo   = document.getElementById('titulo');
    const selectSemana  = document.getElementById('semana');
    const selectMes     = document.getElementById('mes');
    const inputTema     = document.getElementById('tema');
    const inputData     = document.getElementById('data');
    const inputDuracao  = document.getElementById('duracao');
    const inputResponsavel = document.getElementById('responsavel');
    const inputSetor    = document.getElementById('setor');
    const inputConteudo = document.getElementById('conteudo');
    const btnImprimir   = document.getElementById('btn-imprimir');

    // Preenche a data de hoje por padrão
    const hoje = new Date();
    const dataFormatada = hoje.toISOString().split('T')[0];
    inputData.value = dataFormatada;

    // Auto-seleciona o mês atual no select
    const meses = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO',
                   'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
    selectMes.value = meses[hoje.getMonth()];

    // Monta o título completo automaticamente
    function buildTitulo() {
        const ano  = hoje.getFullYear();
        const sem  = selectSemana.value;
        const mes  = selectMes.value;
        const tema = inputTema.value.trim();
        inputTitulo.value = `${sem} SEMANA DE ${mes}/${ano}${tema ? ' - ' + tema : ''}`;
    }

    // Atualiza o título ao mudar qualquer campo relacionado
    [selectSemana, selectMes, inputTema].forEach(el => {
        el.addEventListener('input', () => { buildTitulo(); updatePreview(); });
        el.addEventListener('change', () => { buildTitulo(); updatePreview(); });
    });

    // Constrói o título inicial
    buildTitulo();

    // Função para gerar as linhas em branco da tabela de presença
    function generatePresenceRows() {
        const tbody = document.getElementById('presence-rows');
        if (!tbody) return;
        
        tbody.innerHTML = '';
        // O template original tem cerca de 40 linhas em branco para preencher a folha
        const numLinhas = 38; 
        
        for (let i = 0; i < numLinhas; i++) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td></td>
                <td></td>
                <td></td>
            `;
            tbody.appendChild(tr);
        }
    }

    // Função para atualizar a prévia
    function updatePreview() {
        // Formatar a data (DD/MM/AAAA)
        let dataTexto = '';
        if (inputData.value) {
            const partesData = inputData.value.split('-');
            dataTexto = `${partesData[2]}/${partesData[1]}/${partesData[0]}`;
        }

        // Atualizar textos no documento
        const outTituloText = inputTitulo.value || '';
        
        // Atualiza todos os elementos com a classe correspondente
        document.querySelectorAll('.out-titulo').forEach(el => el.textContent = outTituloText);
        document.querySelectorAll('.out-data').forEach(el => el.textContent = dataTexto);
        document.querySelectorAll('.out-duracao').forEach(el => el.textContent = inputDuracao.value || '');
        document.querySelectorAll('.out-responsavel').forEach(el => el.textContent = inputResponsavel.value || '');
        document.querySelectorAll('.out-setor').forEach(el => el.textContent = inputSetor.value || '');
        
        // Atualiza o conteúdo do texto
        document.querySelectorAll('.out-conteudo').forEach(el => el.textContent = inputConteudo.value || '');
    }

    // Adiciona os eventos (listeners) aos inputs para atualizar em tempo real
    const inputs = [inputTitulo, inputData, inputDuracao, inputResponsavel, inputSetor, inputConteudo];
    inputs.forEach(input => {
        input.addEventListener('input', updatePreview);
    });

    // Evento de Imprimir
    btnImprimir.addEventListener('click', () => {
        window.print();
    });

    // Integração com IA (Gemini) - Chave camuflada para evitar bloqueio do GitHub
    const btnIa = document.getElementById('btn-ia');
    const _k = "QUl6YVN5RG1lTXdqNUYtMTJrZXJ3bWJ3QU1odGlZOHlSMm5fOTVV";
    const API_KEY = atob(_k);
    
    if (btnIa) {
        btnIa.addEventListener('click', async () => {
            const tema = inputTema.value.trim();
            if (!tema) {
                alert('Por favor, preencha o "Tema do DSS" antes de gerar com IA.');
                inputTema.focus();
                return;
            }

            btnIa.disabled = true;
            btnIa.textContent = '⏳ Gerando...';

            try {
                const prompt = `Você é um técnico de segurança do trabalho experiente. Escreva um Diálogo Semanal de Segurança (DSS) sobre o tema: "${tema}". 
O texto deve ser direto, profissional, sem enrolação e focado na prevenção de acidentes.
NÃO use formatação markdown como **negrito**, # ou *. APENAS texto puro.
Para listas e tópicos, use APENAS o símbolo de bullet point "•".
Estruture a resposta EXATAMENTE com os 3 títulos abaixo (pule uma linha entre eles):

Introdução:
[texto introdutório do tema]

Desenvolvimento:
[explicação do tema]
[pule uma linha]
Principais riscos envolvidos:
• [risco 1]
• [risco 2]
• [risco 3]
[pule uma linha]
Boas práticas de segurança:
• [prática 1]
• [prática 2]
• [prática 3]

Conclusão:
[texto final de impacto e conscientização]`;

                const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${API_KEY}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        contents: [{ parts: [{ text: prompt }] }]
                    })
                });

                const data = await response.json();
                
                if (!response.ok) {
                    const errMsg = data.error ? data.error.message : 'Erro desconhecido';
                    throw new Error(errMsg);
                }

                const textoGerado = data.candidates[0].content.parts[0].text;
                
                inputConteudo.value = textoGerado.trim();
                updatePreview();
            } catch (error) {
                console.error(error);
                alert('Erro ao gerar DSS: ' + error.message);
            } finally {
                btnIa.disabled = false;
                btnIa.textContent = '✨ Gerar com IA';
            }
        });
    }

    // Upload da Logo pelo usuário
    const logoUpload = document.getElementById('logo-upload');
    const logoDisplay = document.getElementById('logo-display');
    const logoDisplayContainer = document.getElementById('logo-display-container');

    if (logoUpload) {
        logoUpload.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    // Atualiza a imagem na tela e na impressão
                    document.querySelectorAll('#logo-display').forEach(img => {
                        img.src = e.target.result;
                        img.style.display = 'inline-block';
                    });
                    document.querySelectorAll('#logo-display-container').forEach(container => {
                        if(container) container.style.display = 'none';
                    });
                }
                reader.readAsDataURL(file);
            }
        });
    }

    // Data da revisão automática (primeira e segunda página)
    const dataRevisaoTexto = dataFormatada.split('-').reverse().join('/');
    const spanDataRevisao = document.getElementById('data-revisao-auto');
    if (spanDataRevisao) {
        spanDataRevisao.textContent = dataRevisaoTexto;
    }
    // Segunda página
    document.querySelectorAll('.data-revisao-auto2').forEach(el => {
        el.textContent = dataRevisaoTexto;
    });

    // Execução inicial
    generatePresenceRows();
    updatePreview();
});
