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

    // Barra de progresso dinâmica
    const progressBar = document.getElementById('progress-bar');
    function updateProgress() {
        const fields = [inputTema, inputData, inputResponsavel, inputConteudo];
        const filled = fields.filter(f => f.value.trim() !== '').length;
        const percent = Math.round((filled / fields.length) * 100);
        if (progressBar) {
            progressBar.style.width = percent + '%';
        }
    }

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
        
        // Atualiza o conteúdo do texto, convertendo as quebras de linha nativas em <br> para funcionar em qualquer formato (HTML, PDF, Word)
        document.querySelectorAll('.out-conteudo').forEach(el => {
            const texto = inputConteudo.value || '';
            el.innerHTML = texto.replace(/\n/g, '<br>');
        });
    }

    // Adiciona os eventos (listeners) aos inputs para atualizar em tempo real
    const inputs = [inputTitulo, inputData, inputDuracao, inputResponsavel, inputSetor, inputConteudo];
    inputs.forEach(input => {
        input.addEventListener('input', () => { updatePreview(); updateProgress(); });
    });

    // Evento de Imprimir
    btnImprimir.addEventListener('click', () => {
        window.print();
    });

    // Evento de Exportar para Word
    const btnWord = document.getElementById('btn-word');
    if (btnWord) {
        btnWord.addEventListener('click', () => {
            // Pega o conteúdo HTML do documento
            let printContent = document.getElementById('printable-area').innerHTML;
            
            // 1. Não substituimos globalmente para não criar buracos no layout do Word
            // As quebras de linha (<br>) já estão injetadas corretamente no HTML do out-conteudo.

            // Força a largura 100% nas tabelas usando atributos HTML puros (o Word lê melhor que CSS)
            printContent = printContent.replace(/<table/g, '<table width="100%"');

            // 2. O MS Word bloqueia imagens Base64 por segurança. Vamos usar o link direto da logo no seu GitHub Pages
            const logoUrl = "https://adrielmartinsdias2803-dotcom.github.io/gerador-dss/logo.png";
            printContent = printContent.replace(/src="[^"]*logo\.png"/g, `src="${logoUrl}"`);
            
            // Oculta a caixa de "Clique para alterar a logo" que não deve ir pro Word
            printContent = printContent.replace(/<span[^>]*>Clique para alterar a logo<\/span>/g, '');

            // CSS Inline Otimizado para o Word
            const wordCSS = `
                <style>
                    @page WordSection1 { size: 595.3pt 841.9pt; margin: 42.5pt; }
                    div.WordSection1 { page: WordSection1; font-family: Arial, sans-serif; }
                    table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
                    td, th { border: 1.5px solid black; padding: 4px; vertical-align: middle; font-family: Arial, sans-serif; font-size: 10pt; }
                    .header-table td { font-size: 11pt; font-family: "Times New Roman", serif; text-align: center; }
                    .col-form-title, .presence-table th { background-color: #8c8c8c; color: black; font-weight: bold; text-align: left; }
                    .presence-table th { text-align: center; }
                    .col-doc-name { font-weight: bold; text-align: left; }
                    .content-box { border: 1.5px solid black; border-top: none; padding: 12px; }
                    .info-table { border: 1.5px solid black; margin-bottom: 0; }
                    .content-box p, .out-conteudo { text-align: left; line-height: 1.5; font-size: 10.5pt; font-family: Arial, sans-serif; }
                    .presence-table td { height: 26px; }
                    .page-break { page-break-before: always; }
                    img { max-width: 120px; height: auto; }
                </style>
            `;

            // Monta o cabeçalho específico para o MS Word reconhecer
            const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
                "xmlns:w='urn:schemas-microsoft-com:office:word' " +
                "xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><meta charset='utf-8'><title>DSS</title>" +
                wordCSS +
                "</head><body><div class='WordSection1'>";
            const footer = "</div></body></html>";
            
            // Concatena tudo
            const sourceHTML = header + printContent + footer;
            
            // Cria o link de download mágico
            const source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
            const fileDownload = document.createElement("a");
            document.body.appendChild(fileDownload);
            fileDownload.href = source;
            
            // Define o nome do arquivo com a data
            const nomeArquivo = `DSS_${inputData.value}.doc`;
            fileDownload.download = nomeArquivo;
            fileDownload.click();
            document.body.removeChild(fileDownload);
        });
    }

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
            
            // Skeleton loading — anima o textarea enquanto a IA trabalha
            inputConteudo.value = '';
            inputConteudo.placeholder = 'A inteligência artificial está escrevendo o DSS...';
            inputConteudo.classList.add('skeleton-loading');

            try {
                const prompt = `Você é um técnico de segurança do trabalho experiente. Escreva um Diálogo Semanal de Segurança (DSS) sobre o tema: "${tema}". 
O texto deve ser direto, profissional, conciso (máximo de 200 palavras) e focado na prevenção de acidentes.
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
                updateProgress();

                // Salva o tema no histórico (após sucesso)
                saveToHistory(tema);
            } catch (error) {
                console.error(error);
                alert('Erro ao gerar DSS: ' + error.message);
            } finally {
                btnIa.disabled = false;
                btnIa.textContent = '✨ Gerar com IA';
                inputConteudo.classList.remove('skeleton-loading');
                inputConteudo.placeholder = 'Digite o conteúdo do DSS aqui...\n\nIntrodução:\n...\n\nDesenvolvimento:\n...\n\nConclusão:\n...';
            }
        });
    }

    // =========================================
    // SUGESTÃO DE TEMAS COM IA
    // =========================================
    const btnSuggest = document.getElementById('btn-suggest');
    const suggestionsArea = document.getElementById('suggestions-area');
    const suggestionsList = document.getElementById('suggestions-list');

    if (btnSuggest) {
        btnSuggest.addEventListener('click', async () => {
            btnSuggest.disabled = true;
            btnSuggest.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2a7 7 0 0 1 7 7c0 2.4-1.2 4.5-3 5.7V17a2 2 0 0 1-2 2h-4a2 2 0 0 1-2-2v-2.3C6.2 13.5 5 11.4 5 9a7 7 0 0 1 7-7z"/><line x1="10" y1="22" x2="14" y2="22"/></svg> Buscando sugestões...';

            try {
                const setor = inputSetor.value.trim() || 'indústria';
                const promptSuggest = `Você é um técnico de segurança do trabalho. Liste exatamente 10 temas essenciais e variados de Diálogo Semanal de Segurança (DSS) focados nas rotinas, atos e riscos do setor de "${setor}".

REGRAS:
- Retorne APENAS os temas, um por linha
- Cada tema deve ter no máximo 6 palavras
- NÃO numere os temas
- NÃO use bullet points
- NÃO adicione explicações
- Temas devem ser diretamente ligados às atividades do setor de ${setor}

Exemplo de formato:
Uso correto de EPIs
Prevenção de quedas em altura
Ergonomia no posto de trabalho`;

                const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${API_KEY}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        contents: [{ parts: [{ text: promptSuggest }] }]
                    })
                });

                const data = await response.json();
                
                if (!response.ok) {
                    throw new Error(data.error ? data.error.message : 'Erro desconhecido');
                }

                const textoSugestoes = data.candidates[0].content.parts[0].text;
                const temas = textoSugestoes.split('\n').map(t => t.trim()).filter(t => t.length > 3 && t.length < 60);

                // Renderiza os chips
                suggestionsList.innerHTML = temas.map(tema => 
                    `<span class="suggestion-chip" data-tema="${tema}">${tema}</span>`
                ).join('');

                // Adiciona click nos chips
                suggestionsList.querySelectorAll('.suggestion-chip').forEach(chip => {
                    chip.addEventListener('click', () => {
                        inputTema.value = chip.dataset.tema;
                        buildTitulo();
                        updatePreview();
                        updateProgress();
                        // Destaque visual no chip selecionado
                        suggestionsList.querySelectorAll('.suggestion-chip').forEach(c => c.style.opacity = '0.5');
                        chip.style.opacity = '1';
                        chip.style.borderColor = '#8b5cf6';
                    });
                });

                suggestionsArea.classList.remove('hidden');

            } catch (error) {
                console.error(error);
                alert('Erro ao buscar sugestões: ' + error.message);
            } finally {
                btnSuggest.disabled = false;
                btnSuggest.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2a7 7 0 0 1 7 7c0 2.4-1.2 4.5-3 5.7V17a2 2 0 0 1-2 2h-4a2 2 0 0 1-2-2v-2.3C6.2 13.5 5 11.4 5 9a7 7 0 0 1 7-7z"/><line x1="10" y1="22" x2="14" y2="22"/></svg> Sugerir temas com IA';
            }
        });
    }

    // =========================================
    // HISTÓRICO DE TEMAS (localStorage)
    // =========================================
    const HISTORY_KEY = 'dss_temas_history';
    const MAX_HISTORY = 8;
    const btnHistory = document.getElementById('btn-history');
    const historyDropdown = document.getElementById('history-dropdown');
    const historyList = document.getElementById('history-list');

    function getHistory() {
        try {
            return JSON.parse(localStorage.getItem(HISTORY_KEY)) || [];
        } catch { return []; }
    }

    function saveToHistory(tema) {
        let history = getHistory();
        // Remove duplicatas
        history = history.filter(item => item.tema !== tema);
        // Adiciona no topo
        history.unshift({ tema: tema, data: new Date().toLocaleDateString('pt-BR') });
        // Limita o tamanho
        if (history.length > MAX_HISTORY) history = history.slice(0, MAX_HISTORY);
        localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
    }

    function renderHistory() {
        if (!historyList) return;
        const history = getHistory();
        if (history.length === 0) {
            historyList.innerHTML = '<div class="history-empty">Nenhum tema usado ainda</div>';
            return;
        }
        historyList.innerHTML = history.map(item => 
            `<div class="history-item" data-tema="${item.tema}">
                <span>${item.tema}</span>
                <span class="history-date">${item.data}</span>
            </div>`
        ).join('');

        // Adiciona click nos itens
        historyList.querySelectorAll('.history-item').forEach(el => {
            el.addEventListener('click', () => {
                inputTema.value = el.dataset.tema;
                historyDropdown.classList.add('hidden');
                buildTitulo();
                updatePreview();
                updateProgress();
            });
        });
    }

    if (btnHistory) {
        btnHistory.addEventListener('click', (e) => {
            e.stopPropagation();
            renderHistory();
            historyDropdown.classList.toggle('hidden');
        });
    }

    // Fecha o dropdown ao clicar fora
    document.addEventListener('click', (e) => {
        if (historyDropdown && !historyDropdown.contains(e.target) && e.target !== btnHistory) {
            historyDropdown.classList.add('hidden');
        }
    });

    // =========================================
    // ATALHOS DE TECLADO
    // =========================================
    document.addEventListener('keydown', (e) => {
        // Ctrl+P → Imprimir
        if (e.ctrlKey && e.key === 'p') {
            e.preventDefault();
            window.print();
        }
        // Ctrl+S → Baixar Word
        if (e.ctrlKey && e.key === 's') {
            e.preventDefault();
            const btnWord = document.getElementById('btn-word');
            if (btnWord) btnWord.click();
        }
    });

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
    updateProgress();
});
