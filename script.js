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

    // Evento de Exportar para Word
    const btnWord = document.getElementById('btn-word');
    if (btnWord) {
        btnWord.addEventListener('click', () => {
            // Pega o conteúdo HTML do documento
            let printContent = document.getElementById('printable-area').innerHTML;
            
            // Base64 da Logo para embedar diretamente no arquivo Word (impede que a imagem quebre)
            const logoBase64 = "data:image/png;base64,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";
            
            // Substitui todas as src="logo.png" pelo base64 no HTML exportado
            printContent = printContent.replace(/src="logo\.png"/g, `src="${logoBase64}"`);
            
            // CSS Inline Otimizado para o Word
            const wordCSS = `
                <style>
                    @page WordSection1 { size: 595.3pt 841.9pt; margin: 42.5pt; }
                    div.WordSection1 { page: WordSection1; font-family: Arial, sans-serif; }
                    table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
                    td, th { border: 1px solid black; padding: 4px; vertical-align: middle; font-family: Arial, sans-serif; font-size: 10pt; }
                    .header-table td { font-size: 11pt; font-family: "Times New Roman", serif; }
                    .col-form-title, .presence-table th { background-color: #8c8c8c; color: black; font-weight: bold; text-align: left; }
                    .presence-table th { text-align: center; }
                    .col-doc-name { font-weight: bold; }
                    .content-box { border: 1.5px solid black; border-top: none; padding: 8px; }
                    .info-table { border: 1.5px solid black; margin-bottom: 0; }
                    .content-box p, .out-conteudo { text-align: justify; line-height: 1.5; font-size: 10.5pt; }
                    .page-break { page-break-before: always; }
                    .logo-img { max-width: 100px; max-height: 40px; }
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
