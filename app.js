(() => {
    const COLUMNS = {
        local: 0,
        andar: 1,
        horaInicio: 4,
        horaFim: 8,
        status: 12,
        solucao: 13
    };

    let dadosProcessadosGlobais = null;
    let currentEditRowIndex = null;
    let currentFinalizeRowIndex = null;
    let cronometrosPorLinha = {};

    const origemInput = document.getElementById("origemFile");
    const btnProcessar = document.getElementById("btnProcessar");
    const btnExportar = document.getElementById("btnExportar");
    const statusBox = document.getElementById("status");

    const editModal = document.getElementById("editModal");
    const btnCancelModal = document.getElementById("btnCancelModal");
    const btnSaveModal = document.getElementById("btnSaveModal");

    const modalInputLocal = document.getElementById("modalInputLocal");
    const modalInputAndar = document.getElementById("modalInputAndar");
    const modalInputHoraInicio = document.getElementById("modalInputHoraInicio");
    const modalInputHoraFim = document.getElementById("modalInputHoraFim");
    const modalInputStatus = document.getElementById("modalInputStatus");
    const modalInputSolucao = document.getElementById("modalInputSolucao");

    const solutionModal = document.getElementById("solutionModal");
    const solutionInput = document.getElementById("solutionInput");
    const solutionError = document.getElementById("solutionError");
    const btnCancelSolutionModal = document.getElementById("btnCancelSolutionModal");
    const btnConfirmSolutionModal = document.getElementById("btnConfirmSolutionModal");
    const btnGithubLoad = document.getElementById("btnGithubLoad");
    const btnGithubSave = document.getElementById("btnGithubSave");
    const btnGithubConfig = document.getElementById("btnGithubConfig");
    const githubModal = document.getElementById("githubModal");
    const btnCancelGithubModal = document.getElementById("btnCancelGithubModal");
    const btnSaveGithubConfig = document.getElementById("btnSaveGithubConfig");
    const ghOwner = document.getElementById("ghOwner");
    const ghRepo = document.getElementById("ghRepo");
    const ghBranch = document.getElementById("ghBranch");
    const ghPath = document.getElementById("ghPath");
    const ghToken = document.getElementById("ghToken");

    const btnNovaAtividade = document.getElementById("btnNovaAtividade");
    const activityModal = document.getElementById("activityModal");
    const activityLocal = document.getElementById("activityLocal");
    const activityAndar = document.getElementById("activityAndar");
    const activitySetor = document.getElementById("activitySetor");
    const activityDescricao = document.getElementById("activityDescricao");
    const btnCancelActivityModal = document.getElementById("btnCancelActivityModal");
    const btnSaveActivityModal = document.getElementById("btnSaveActivityModal");

    const filterSetor = document.getElementById("filterSetor");
    const filterLocal = document.getElementById("filterLocal");
    const filterAndar = document.getElementById("filterAndar");
    const btnLimparFiltros = document.getElementById("btnLimparFiltros");

    const GITHUB_CFG_KEY = "github_json_config_v1";
    let githubFileSha = null;
    let githubConfig = carregarGithubConfig();

    btnProcessar.addEventListener("click", processarArquivo);
    btnExportar.addEventListener("click", exportarArquivo);
    origemInput.addEventListener("change", () => {
        if (origemInput.files && origemInput.files[0]) processarArquivo();
    });
    btnCancelModal.addEventListener("click", fecharModal);
    btnSaveModal.addEventListener("click", salvarEdicaoModal);
    btnCancelSolutionModal.addEventListener("click", fecharModalSolucao);
    btnConfirmSolutionModal.addEventListener("click", confirmarFinalizacaoComSolucao);
    if (btnGithubLoad) btnGithubLoad.addEventListener("click", carregarDadosDoGithub);
    if (btnGithubSave) btnGithubSave.addEventListener("click", salvarDadosNoGithub);
    if (btnGithubConfig) btnGithubConfig.addEventListener("click", abrirGithubModal);
    if (btnCancelGithubModal) btnCancelGithubModal.addEventListener("click", fecharGithubModal);
    if (btnSaveGithubConfig) btnSaveGithubConfig.addEventListener("click", salvarGithubConfig);
    if (btnNovaAtividade) btnNovaAtividade.addEventListener("click", abrirModalNovaAtividade);
    if (btnCancelActivityModal) btnCancelActivityModal.addEventListener("click", fecharModalNovaAtividade);
    if (btnSaveActivityModal) btnSaveActivityModal.addEventListener("click", salvarNovaAtividade);
    if (btnLimparFiltros) btnLimparFiltros.addEventListener("click", limparFiltros);
    if (filterSetor) filterSetor.addEventListener("input", rerenderComFiltros);
    if (filterLocal) filterLocal.addEventListener("input", rerenderComFiltros);
    if (filterAndar) filterAndar.addEventListener("input", rerenderComFiltros);
    setStatus("Aguardando arquivo para processar.", "status-loading");

    editModal.addEventListener("click", (event) => {
        if (event.target === editModal) fecharModal();
    });

    solutionModal.addEventListener("click", (event) => {
        if (event.target === solutionModal) fecharModalSolucao();
    });

    if (activityModal) {
        activityModal.addEventListener("click", (event) => {
            if (event.target === activityModal) fecharModalNovaAtividade();
        });
    }
    if (githubModal) {
        githubModal.addEventListener("click", (event) => {
            if (event.target === githubModal) fecharGithubModal();
        });
    }

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && currentEditRowIndex !== null) fecharModal();
        if (event.key === "Escape" && currentFinalizeRowIndex !== null) fecharModalSolucao();
        if (event.key === "Escape" && activityModal && activityModal.style.display === "flex") fecharModalNovaAtividade();
        if (event.key === "Escape" && githubModal && githubModal.style.display === "flex") fecharGithubModal();
    });

    async function processarArquivo() {
        if (!origemInput.files || !origemInput.files[0]) {
            setStatus("Erro: O arquivo de origem é obrigatório.", "status-error");
            return;
        }

        try {
            setStatus("Processando arquivo em memória...", "status-loading");
            btnExportar.disabled = true;

            const workbook = await readFileAsync(origemInput.files[0]);
            const worksheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[worksheetName];
            const origData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: false,
                defval: "",
                dateNF: "dd/mm/yyyy hh:mm:ss"
            });
            document.dispatchEvent(new CustomEvent("source-sheet-loaded", {
                detail: {
                    origData,
                    fileName: origemInput.files[0].name
                }
            }));

            const finalData = transformarDados(origData);
            const dataToSort = finalData.slice(1);
            dataToSort.sort((a, b) => {
                const cmpLocal = String(a[COLUMNS.local] || "").localeCompare(String(b[COLUMNS.local] || ""));
                if (cmpLocal !== 0) return cmpLocal;
                return String(b[COLUMNS.andar] || "").localeCompare(String(a[COLUMNS.andar] || ""));
            });

            limparCronometrosAtivos();
            dadosProcessadosGlobais = [finalData[0], ...dataToSort];
            renderHtmlCards(dadosProcessadosGlobais, "tableContainer");

            setStatus(`Operação concluída. ${dataToSort.length} registros processados.`, "status-success");
            btnExportar.disabled = false;
        } catch (err) {
            setStatus(`Erro crítico: ${err.message}`, "status-error");
        }
    }

    function transformarDados(origData) {
        const finalData = [];
        const origHeaders = origData[0] || [];

        finalData.push([
            "Local", "Andar", origHeaders[0] || "", origHeaders[1] || "",
            "Hora inicio", origHeaders[2] || "", origHeaders[3] || "",
            origHeaders[4] || "", "Hora fim", origHeaders[5] || "",
            origHeaders[6] || "", origHeaders[7] || "", "Status", "Solução"
        ]);

        for (let i = 1; i < origData.length; i++) {
            const row = origData[i];
            if (!row || row.length === 0) continue;

            finalData.push([
                "", "", row[0] || "", row[1] || "", "",
                row[2] || "", row[3] || "", row[4] || "", "",
                row[5] || "", row[6] || "", row[7] || "", "Aberto", ""
            ]);
        }

        return finalData;
    }

    function renderHtmlCards(aoa, containerId) {
        const container = document.getElementById(containerId);
        container.innerHTML = "";
        container.className = "cards-container";

        if (!aoa || aoa.length <= 1) return;

        const headers = aoa[0];
        const board = document.createElement("div");
        board.className = "status-board";

        const colAberto = criarColunaStatus("Aberto");
        const colAndamento = criarColunaStatus("Em andamento");
        const colConcluido = criarColunaStatus("Concluída");

        for (let i = 1; i < aoa.length; i++) {
            const rowData = aoa[i];
            if (!linhaPassaNosFiltros(rowData, headers)) continue;

            const card = document.createElement("article");
            card.className = "card";
            card.id = `card-${i}`;

            const valuesContainer = document.createElement("div");

            headers.forEach((headerName, colIndex) => {
                const row = document.createElement("div");
                row.className = "card-row";

                const label = document.createElement("span");
                label.className = "card-label";
                label.textContent = headerName || `Coluna ${colIndex + 1}`;

                const value = document.createElement("span");
                value.className = "card-value";
                value.id = `val-${i}-${colIndex}`;

                const cellData = rowData[colIndex];
                if (cellData !== undefined && cellData !== null && cellData !== "") value.textContent = cellData;

                row.appendChild(label);
                row.appendChild(value);
                valuesContainer.appendChild(row);
            });

            card.appendChild(valuesContainer);

            const timerDisplay = document.createElement("div");
            timerDisplay.className = "card-timer";
            timerDisplay.id = `timer-${i}`;
            timerDisplay.textContent = "Cronômetro: 00:00:00";
            card.appendChild(timerDisplay);

            const actionContainer = document.createElement("div");
            actionContainer.className = "card-actions";

            const btnTimerToggle = document.createElement("button");
            btnTimerToggle.className = "btn-card-timer";
            btnTimerToggle.id = `btn-timer-toggle-${i}`;
            btnTimerToggle.type = "button";
            btnTimerToggle.textContent = "Iniciar";
            btnTimerToggle.addEventListener("click", () => alternarCronometro(i));

            const btnPendencia = document.createElement("button");
            btnPendencia.className = "btn-card-pendencia";
            btnPendencia.id = `btn-pendencia-${i}`;
            btnPendencia.type = "button";
            btnPendencia.textContent = "Pendência";
            btnPendencia.addEventListener("click", () => registrarPendencia(i));

            const btnTimerFinalize = document.createElement("button");
            btnTimerFinalize.className = "btn-card-finalize";
            btnTimerFinalize.id = `btn-timer-finalize-${i}`;
            btnTimerFinalize.type = "button";
            btnTimerFinalize.textContent = "Finalizar";
            btnTimerFinalize.addEventListener("click", () => abrirModalSolucao(i));

            const btnConcluir = document.createElement("button");
            btnConcluir.className = "btn-card-concluir";
            btnConcluir.id = `btn-concluir-${i}`;
            btnConcluir.type = "button";
            btnConcluir.textContent = "Marcar Concluída";
            btnConcluir.addEventListener("click", () => executarConclusao(i));

            const btnCardEdit = document.createElement("button");
            btnCardEdit.className = "btn-card-edit";
            btnCardEdit.type = "button";
            btnCardEdit.textContent = "Editar";
            btnCardEdit.addEventListener("click", () => abrirModalEdicao(i));

            actionContainer.append(btnTimerToggle, btnPendencia, btnTimerFinalize, btnConcluir, btnCardEdit);
            card.appendChild(actionContainer);

            const categoria = obterCategoriaStatus(i);
            if (categoria === "concluida") colConcluido.cards.push(card);
            else if (categoria === "andamento") colAndamento.cards.push(card);
            else colAberto.cards.push(card);

            obterEstadoCronometro(i);
            atualizarCronometroDisplay(i);
            sincronizarEstadoCartao(i);
        }

        board.appendChild(montarColunaStatus(colAberto));
        board.appendChild(montarColunaStatus(colAndamento));
        board.appendChild(montarColunaStatus(colConcluido));
        container.appendChild(board);
    }

    function criarColunaStatus(titulo) {
        return { titulo, cards: [] };
    }

    function montarColunaStatus(coluna) {
        const wrap = document.createElement("section");
        wrap.className = "status-column";

        const title = document.createElement("h3");
        title.textContent = coluna.titulo;

        const count = document.createElement("span");
        count.className = "status-column-count";
        count.textContent = `(${coluna.cards.length})`;
        title.appendChild(count);
        wrap.appendChild(title);

        if (coluna.cards.length === 0) {
            const empty = document.createElement("p");
            empty.className = "status-column-empty";
            empty.textContent = "Sem atividades nesta coluna.";
            wrap.appendChild(empty);
            return wrap;
        }

        coluna.cards.forEach((card) => wrap.appendChild(card));
        return wrap;
    }

    function abrirModalEdicao(rowIndex) {
        currentEditRowIndex = rowIndex;
        const rowData = dadosProcessadosGlobais[rowIndex];

        modalInputLocal.value = rowData[COLUMNS.local] || "";
        modalInputAndar.value = rowData[COLUMNS.andar] || "";
        modalInputHoraInicio.value = rowData[COLUMNS.horaInicio] || "";
        modalInputHoraFim.value = rowData[COLUMNS.horaFim] || "";
        modalInputStatus.value = rowData[COLUMNS.status] || "";
        modalInputSolucao.value = rowData[COLUMNS.solucao] || "";

        editModal.style.display = "flex";
        modalInputLocal.focus();
    }

    function fecharModal() {
        editModal.style.display = "none";
        currentEditRowIndex = null;
    }

    function salvarEdicaoModal() {
        if (currentEditRowIndex === null) return;

        const i = currentEditRowIndex;
        const valLocal = modalInputLocal.value.trim();
        const valAndar = modalInputAndar.value.trim();
        const valInicio = modalInputHoraInicio.value.trim();
        const valFim = modalInputHoraFim.value.trim();
        const valStatus = modalInputStatus.value.trim();
        const valSolucao = modalInputSolucao.value.trim();
        const statusCanonico = canonizarStatus(valStatus);

        dadosProcessadosGlobais[i][COLUMNS.local] = valLocal;
        dadosProcessadosGlobais[i][COLUMNS.andar] = valAndar;
        dadosProcessadosGlobais[i][COLUMNS.horaInicio] = valInicio;
        dadosProcessadosGlobais[i][COLUMNS.horaFim] = valFim;
        dadosProcessadosGlobais[i][COLUMNS.status] = statusCanonico;
        dadosProcessadosGlobais[i][COLUMNS.solucao] = valSolucao;

        const updates = [
            { col: COLUMNS.local, val: valLocal },
            { col: COLUMNS.andar, val: valAndar },
            { col: COLUMNS.horaInicio, val: valInicio },
            { col: COLUMNS.horaFim, val: valFim },
            { col: COLUMNS.status, val: statusCanonico },
            { col: COLUMNS.solucao, val: valSolucao }
        ];

        updates.forEach((item) => {
            const node = document.getElementById(`val-${i}-${item.col}`);
            if (node) node.textContent = item.val;
        });

        recalcularCronometroDaLinha(i);
        sincronizarEstadoCartao(i);
        fecharModal();
        renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
    }

    function abrirModalSolucao(rowIndex) {
        if (!dadosProcessadosGlobais || !dadosProcessadosGlobais[rowIndex]) return;
        if (statusEhConcluida(dadosProcessadosGlobais[rowIndex][COLUMNS.status])) return;

        currentFinalizeRowIndex = rowIndex;
        solutionInput.value = String(dadosProcessadosGlobais[rowIndex][COLUMNS.solucao] || "").trim();
        solutionError.textContent = "";
        solutionModal.style.display = "flex";
        solutionInput.focus();
    }

    function fecharModalSolucao() {
        solutionModal.style.display = "none";
        currentFinalizeRowIndex = null;
        solutionError.textContent = "";
    }

    function confirmarFinalizacaoComSolucao() {
        if (currentFinalizeRowIndex === null) return;

        const rowIndex = currentFinalizeRowIndex;
        const textoSolucao = solutionInput.value.trim();
        if (!textoSolucao) {
            solutionError.textContent = "Informe a solução para finalizar.";
            return;
        }

        solutionError.textContent = "";
        dadosProcessadosGlobais[rowIndex][COLUMNS.solucao] = textoSolucao;

        const solucaoNode = document.getElementById(`val-${rowIndex}-${COLUMNS.solucao}`);
        if (solucaoNode) solucaoNode.textContent = textoSolucao;

        fecharModalSolucao();
        finalizarCronometro(rowIndex);
    }

    function registrarPendencia(rowIndex) {
        const estado = obterEstadoCronometro(rowIndex);
        if (!estado.rodando) {
            setStatus("Pendência só pode ser registrada com o cronômetro em andamento.", "status-error");
            return;
        }

        const texto = prompt("Descreva a pendência desta atividade:");
        if (!texto || !texto.trim()) return;

        const agora = new Date().toLocaleTimeString("pt-BR", { hour12: false });
        const atual = String(dadosProcessadosGlobais[rowIndex][COLUMNS.solucao] || "").trim();
        const pend = `[Pendência ${agora}] ${texto.trim()}`;
        const novoTexto = atual ? `${atual} | ${pend}` : pend;

        dadosProcessadosGlobais[rowIndex][COLUMNS.solucao] = novoTexto;
        const solucaoNode = document.getElementById(`val-${rowIndex}-${COLUMNS.solucao}`);
        if (solucaoNode) solucaoNode.textContent = novoTexto;
        setStatus("Pendência registrada sem interromper o cronômetro.", "status-loading");
    }

    function abrirModalNovaAtividade() {
        if (!activityModal) return;
        activityLocal.value = "";
        activityAndar.value = "";
        activitySetor.value = "";
        activityDescricao.value = "";
        activityModal.style.display = "flex";
        activityLocal.focus();
    }

    function fecharModalNovaAtividade() {
        if (!activityModal) return;
        activityModal.style.display = "none";
    }

    function obterCabecalhoPadraoAtividade() {
        return [
            "Local", "Andar", "Setor", "Solicitação",
            "Hora inicio", "Descrição", "Prioridade", "Responsável",
            "Hora fim", "Origem", "Equipe", "Observação", "Status", "Solução"
        ];
    }

    function salvarNovaAtividade() {
        const local = String(activityLocal.value || "").trim();
        const andar = String(activityAndar.value || "").trim();
        const setor = String(activitySetor.value || "").trim();
        const descricao = String(activityDescricao.value || "").trim();

        if (!local || !andar || !descricao) {
            setStatus("Nova atividade: Local, Andar e Descrição são obrigatórios.", "status-error");
            return;
        }

        if (!dadosProcessadosGlobais) {
            dadosProcessadosGlobais = [obterCabecalhoPadraoAtividade()];
        }

        const headers = dadosProcessadosGlobais[0];
        const novaLinha = Array(headers.length).fill("");
        novaLinha[COLUMNS.local] = local;
        novaLinha[COLUMNS.andar] = andar;
        novaLinha[COLUMNS.status] = "Aberto";
        novaLinha[COLUMNS.solucao] = "";

        const setorIdx = obterIndiceSetor(headers);
        if (setorIdx >= 0) novaLinha[setorIdx] = setor;

        const descIdx = obterIndiceDescricao(headers);
        if (descIdx >= 0) novaLinha[descIdx] = descricao;

        dadosProcessadosGlobais.push(novaLinha);
        obterEstadoCronometro(dadosProcessadosGlobais.length - 1);
        fecharModalNovaAtividade();
        renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
        btnExportar.disabled = false;
        setStatus("Nova atividade adicionada com sucesso.", "status-success");
    }

    function normalizarStatus(valor) {
        return String(valor || "")
            .trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase();
    }

    function statusEhConcluida(valor) {
        return normalizarStatus(valor) === "concluida";
    }

    function canonizarStatus(valor) {
        const norm = normalizarStatus(valor);
        if (!norm) return "Aberto";
        if (norm === "aberto") return "Aberto";
        if (norm.includes("andamento")) return "Em andamento";
        if (norm === "concluida") return "Concluída";
        return valor;
    }

    function normalizarTexto(valor) {
        return String(valor || "")
            .trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase();
    }

    function obterIndiceSetor(headers) {
        if (!headers || headers.length === 0) return -1;
        const idx = headers.findIndex((h) => normalizarTexto(h).includes("setor"));
        if (idx >= 0) return idx;
        return headers.length > 3 ? 3 : -1;
    }

    function obterIndiceDescricao(headers) {
        if (!headers || headers.length === 0) return -1;
        const idx = headers.findIndex((h) => {
            const t = normalizarTexto(h);
            return t.includes("descricao") || t.includes("atividade");
        });
        if (idx >= 0) return idx;
        return headers.length > 5 ? 5 : -1;
    }

    function linhaPassaNosFiltros(rowData, headers) {
        const termoSetor = normalizarTexto(filterSetor ? filterSetor.value : "");
        const termoLocal = normalizarTexto(filterLocal ? filterLocal.value : "");
        const termoAndar = normalizarTexto(filterAndar ? filterAndar.value : "");

        if (!termoSetor && !termoLocal && !termoAndar) return true;

        const setorIdx = obterIndiceSetor(headers);
        const valorSetor = setorIdx >= 0 ? normalizarTexto(rowData[setorIdx]) : "";
        const valorLocal = normalizarTexto(rowData[COLUMNS.local]);
        const valorAndar = normalizarTexto(rowData[COLUMNS.andar]);

        const okSetor = !termoSetor || valorSetor.includes(termoSetor);
        const okLocal = !termoLocal || valorLocal.includes(termoLocal);
        const okAndar = !termoAndar || valorAndar.includes(termoAndar);
        return okSetor && okLocal && okAndar;
    }

    function obterCategoriaStatus(rowIndex) {
        const row = dadosProcessadosGlobais && dadosProcessadosGlobais[rowIndex];
        if (!row) return "aberto";

        const statusNorm = normalizarStatus(row[COLUMNS.status]);
        const estado = obterEstadoCronometro(rowIndex);
        if (statusEhConcluida(row[COLUMNS.status])) return "concluida";
        if (estado.rodando || statusNorm.includes("andamento")) return "andamento";
        return "aberto";
    }

    function rerenderComFiltros() {
        if (!dadosProcessadosGlobais) return;
        renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
    }

    function limparFiltros() {
        if (filterSetor) filterSetor.value = "";
        if (filterLocal) filterLocal.value = "";
        if (filterAndar) filterAndar.value = "";
        rerenderComFiltros();
    }

    function formatarDataHoraBR(data) {
        const dia = String(data.getDate()).padStart(2, "0");
        const mes = String(data.getMonth() + 1).padStart(2, "0");
        const ano = data.getFullYear();
        const hora = String(data.getHours()).padStart(2, "0");
        const minuto = String(data.getMinutes()).padStart(2, "0");
        const segundo = String(data.getSeconds()).padStart(2, "0");
        return `${dia}/${mes}/${ano} ${hora}:${minuto}:${segundo}`;
    }

    function parseDataHoraFlexivel(valor) {
        const texto = String(valor || "").trim();
        if (!texto) return null;

        const br = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
        if (br) {
            const data = new Date(
                Number(br[3]),
                Number(br[2]) - 1,
                Number(br[1]),
                Number(br[4] || 0),
                Number(br[5] || 0),
                Number(br[6] || 0)
            );
            if (!Number.isNaN(data.getTime())) return data;
        }

        const apenasHora = texto.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
        if (apenasHora) {
            const agora = new Date();
            agora.setHours(Number(apenasHora[1]), Number(apenasHora[2]), Number(apenasHora[3] || 0), 0);
            if (!Number.isNaN(agora.getTime())) return agora;
        }

        const nativo = new Date(texto);
        if (!Number.isNaN(nativo.getTime())) return nativo;
        return null;
    }

    function formatarDuracao(ms) {
        const totalSegundos = Math.max(0, Math.floor(ms / 1000));
        const horas = Math.floor(totalSegundos / 3600);
        const minutos = Math.floor((totalSegundos % 3600) / 60);
        const segundos = totalSegundos % 60;
        return `${String(horas).padStart(2, "0")}:${String(minutos).padStart(2, "0")}:${String(segundos).padStart(2, "0")}`;
    }

    function calcularDuracaoLinha(rowIndex) {
        if (!dadosProcessadosGlobais || !dadosProcessadosGlobais[rowIndex]) return null;

        const horaInicio = parseDataHoraFlexivel(dadosProcessadosGlobais[rowIndex][COLUMNS.horaInicio]);
        const horaFim = parseDataHoraFlexivel(dadosProcessadosGlobais[rowIndex][COLUMNS.horaFim]);

        if (!horaInicio || !horaFim) return null;

        let diff = horaFim.getTime() - horaInicio.getTime();
        if (diff < 0) diff += 24 * 60 * 60 * 1000;
        return diff >= 0 ? diff : 0;
    }

    function obterEstadoCronometro(rowIndex) {
        if (!cronometrosPorLinha[rowIndex]) {
            const duracaoInicial = calcularDuracaoLinha(rowIndex);
            cronometrosPorLinha[rowIndex] = {
                rodando: false,
                inicioEpochMs: null,
                acumuladoMs: duracaoInicial !== null ? duracaoInicial : 0,
                intervalId: null
            };
        }
        return cronometrosPorLinha[rowIndex];
    }

    function limparCronometrosAtivos() {
        Object.values(cronometrosPorLinha).forEach((estado) => {
            if (estado && estado.intervalId) clearInterval(estado.intervalId);
        });
        cronometrosPorLinha = {};
    }

    function obterMsCronometro(rowIndex) {
        const estado = obterEstadoCronometro(rowIndex);
        let ms = estado.acumuladoMs;
        if (estado.rodando && estado.inicioEpochMs) ms += Date.now() - estado.inicioEpochMs;
        return Math.max(0, ms);
    }

    function atualizarCronometroDisplay(rowIndex) {
        const timerNode = document.getElementById(`timer-${rowIndex}`);
        if (!timerNode) return;

        const estado = obterEstadoCronometro(rowIndex);
        timerNode.textContent = `Cronômetro: ${formatarDuracao(obterMsCronometro(rowIndex))}`;
        timerNode.classList.toggle("card-timer-running", estado.rodando);
    }

    function alternarCronometro(rowIndex) {
        const estado = obterEstadoCronometro(rowIndex);
        if (estado.rodando) pausarCronometro(rowIndex);
        else iniciarCronometro(rowIndex);
    }

    function iniciarCronometro(rowIndex) {
        if (!dadosProcessadosGlobais || !dadosProcessadosGlobais[rowIndex]) return;
        if (statusEhConcluida(dadosProcessadosGlobais[rowIndex][COLUMNS.status])) return;

        const estado = obterEstadoCronometro(rowIndex);
        if (estado.rodando) return;

        const inicioAtual = String(dadosProcessadosGlobais[rowIndex][COLUMNS.horaInicio] || "").trim();
        if (!inicioAtual) {
            const agoraInicio = formatarDataHoraBR(new Date());
            dadosProcessadosGlobais[rowIndex][COLUMNS.horaInicio] = agoraInicio;
            const inicioNode = document.getElementById(`val-${rowIndex}-${COLUMNS.horaInicio}`);
            if (inicioNode) inicioNode.textContent = agoraInicio;
        }

        const fimNode = document.getElementById(`val-${rowIndex}-${COLUMNS.horaFim}`);
        if (fimNode && fimNode.textContent.trim() !== "") {
            dadosProcessadosGlobais[rowIndex][COLUMNS.horaFim] = "";
            fimNode.textContent = "";
        }

        dadosProcessadosGlobais[rowIndex][COLUMNS.status] = "Em andamento";
        const statusNode = document.getElementById(`val-${rowIndex}-${COLUMNS.status}`);
        if (statusNode) statusNode.textContent = "Em andamento";

        estado.rodando = true;
        estado.inicioEpochMs = Date.now();
        if (estado.intervalId) clearInterval(estado.intervalId);
        estado.intervalId = setInterval(() => atualizarCronometroDisplay(rowIndex), 1000);

        atualizarCronometroDisplay(rowIndex);
        sincronizarEstadoCartao(rowIndex);
        renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
    }

    function pausarCronometro(rowIndex) {
        const estado = obterEstadoCronometro(rowIndex);
        if (!estado.rodando) return;

        estado.acumuladoMs += Date.now() - estado.inicioEpochMs;
        estado.inicioEpochMs = null;
        estado.rodando = false;

        if (estado.intervalId) clearInterval(estado.intervalId);
        estado.intervalId = null;

        atualizarCronometroDisplay(rowIndex);
        sincronizarEstadoCartao(rowIndex);
    }

    function finalizarCronometro(rowIndex) {
        if (!dadosProcessadosGlobais || !dadosProcessadosGlobais[rowIndex]) return;

        const estado = obterEstadoCronometro(rowIndex);
        if (estado.rodando) pausarCronometro(rowIndex);

        const agoraFim = formatarDataHoraBR(new Date());
        dadosProcessadosGlobais[rowIndex][COLUMNS.horaFim] = agoraFim;
        const fimNode = document.getElementById(`val-${rowIndex}-${COLUMNS.horaFim}`);
        if (fimNode) fimNode.textContent = agoraFim;

        const inicioAtual = String(dadosProcessadosGlobais[rowIndex][COLUMNS.horaInicio] || "").trim();
        if (!inicioAtual) {
            dadosProcessadosGlobais[rowIndex][COLUMNS.horaInicio] = agoraFim;
            const inicioNode = document.getElementById(`val-${rowIndex}-${COLUMNS.horaInicio}`);
            if (inicioNode) inicioNode.textContent = agoraFim;
        }

        dadosProcessadosGlobais[rowIndex][COLUMNS.status] = "Concluída";
        const statusNode = document.getElementById(`val-${rowIndex}-${COLUMNS.status}`);
        if (statusNode) statusNode.textContent = "Concluída";

        const duracaoLinha = calcularDuracaoLinha(rowIndex);
        if (duracaoLinha !== null) estado.acumuladoMs = duracaoLinha;

        atualizarCronometroDisplay(rowIndex);
        sincronizarEstadoCartao(rowIndex);
        renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
    }

    function recalcularCronometroDaLinha(rowIndex) {
        const estado = obterEstadoCronometro(rowIndex);

        if (estado.intervalId) clearInterval(estado.intervalId);
        estado.intervalId = null;
        estado.rodando = false;
        estado.inicioEpochMs = null;

        const duracaoLinha = calcularDuracaoLinha(rowIndex);
        estado.acumuladoMs = duracaoLinha !== null ? duracaoLinha : 0;

        atualizarCronometroDisplay(rowIndex);
    }

    function sincronizarEstadoCartao(rowIndex) {
        const cardElement = document.getElementById(`card-${rowIndex}`);
        const btnConcluir = document.getElementById(`btn-concluir-${rowIndex}`);
        const btnTimerToggle = document.getElementById(`btn-timer-toggle-${rowIndex}`);
        const btnTimerFinalize = document.getElementById(`btn-timer-finalize-${rowIndex}`);
        const btnPendencia = document.getElementById(`btn-pendencia-${rowIndex}`);

        if (!dadosProcessadosGlobais || !dadosProcessadosGlobais[rowIndex]) return;

        const horaFimValor = String(dadosProcessadosGlobais[rowIndex][COLUMNS.horaFim] || "");
        const statusAtual = dadosProcessadosGlobais[rowIndex][COLUMNS.status];
        const concluida = statusEhConcluida(statusAtual);
        const estadoTimer = obterEstadoCronometro(rowIndex);

        if (btnTimerToggle) {
            btnTimerToggle.textContent = estadoTimer.rodando ? "Pausar" : "Iniciar";
            btnTimerToggle.style.display = concluida ? "none" : "inline-block";
        }

        if (btnTimerFinalize) {
            const temTempo = obterMsCronometro(rowIndex) > 0;
            const podeFinalizar = !concluida && (estadoTimer.rodando || temTempo);
            btnTimerFinalize.style.display = podeFinalizar ? "inline-block" : "none";
        }

        if (btnPendencia) {
            const podePendencia = !concluida && estadoTimer.rodando;
            btnPendencia.style.display = podePendencia ? "inline-block" : "none";
        }

        atualizarCronometroDisplay(rowIndex);

        if (concluida) {
            if (cardElement) cardElement.classList.add("card-completed");
            if (btnConcluir) btnConcluir.style.display = "none";
            return;
        }

        if (cardElement) cardElement.classList.remove("card-completed");
        if (btnConcluir) btnConcluir.style.display = estadoTimer.rodando ? "none" : "inline-block";
    }

    function executarConclusao(rowIndex) {
        const estado = obterEstadoCronometro(rowIndex);
        if (estado.rodando) pausarCronometro(rowIndex);

        const horaFimAtual = String(dadosProcessadosGlobais[rowIndex][COLUMNS.horaFim] || "").trim();
        if (!horaFimAtual) {
            const agoraFim = formatarDataHoraBR(new Date());
            dadosProcessadosGlobais[rowIndex][COLUMNS.horaFim] = agoraFim;
            const fimNode = document.getElementById(`val-${rowIndex}-${COLUMNS.horaFim}`);
            if (fimNode) fimNode.textContent = agoraFim;
        }

        dadosProcessadosGlobais[rowIndex][COLUMNS.status] = "Concluída";
        const statusNode = document.getElementById(`val-${rowIndex}-${COLUMNS.status}`);
        if (statusNode) statusNode.textContent = "Concluída";

        const duracaoLinha = calcularDuracaoLinha(rowIndex);
        if (duracaoLinha !== null) estado.acumuladoMs = duracaoLinha;

        atualizarCronometroDisplay(rowIndex);
        sincronizarEstadoCartao(rowIndex);
        renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
    }

    function exportarArquivo() {
        if (!dadosProcessadosGlobais) return;

        try {
            const worksheet = XLSX.utils.aoa_to_sheet(dadosProcessadosGlobais);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Relatório Processado");
            XLSX.writeFile(workbook, "Relatorio_Processado.xlsx");
        } catch (err) {
            setStatus(`Erro ao exportar: ${err.message}`, "status-error");
        }
    }

    function readFileAsync(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    resolve(XLSX.read(data, { type: "array" }));
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function carregarGithubConfig() {
        try {
            const raw = localStorage.getItem(GITHUB_CFG_KEY);
            if (!raw) {
                return { owner: "", repo: "", branch: "main", path: "data/atividades.json", token: "" };
            }
            const cfg = JSON.parse(raw);
            return {
                owner: cfg.owner || "",
                repo: cfg.repo || "",
                branch: cfg.branch || "main",
                path: cfg.path || "data/atividades.json",
                token: cfg.token || ""
            };
        } catch (_) {
            return { owner: "", repo: "", branch: "main", path: "data/atividades.json", token: "" };
        }
    }

    function salvarGithubConfigLocal() {
        localStorage.setItem(GITHUB_CFG_KEY, JSON.stringify(githubConfig));
    }

    function abrirGithubModal() {
        if (!githubModal) return;
        ghOwner.value = githubConfig.owner || "";
        ghRepo.value = githubConfig.repo || "";
        ghBranch.value = githubConfig.branch || "main";
        ghPath.value = githubConfig.path || "data/atividades.json";
        ghToken.value = githubConfig.token || "";
        githubModal.style.display = "flex";
    }

    function fecharGithubModal() {
        if (!githubModal) return;
        githubModal.style.display = "none";
    }

    function salvarGithubConfig() {
        githubConfig = {
            owner: String(ghOwner.value || "").trim(),
            repo: String(ghRepo.value || "").trim(),
            branch: String(ghBranch.value || "").trim() || "main",
            path: String(ghPath.value || "").trim() || "data/atividades.json",
            token: String(ghToken.value || "").trim()
        };
        salvarGithubConfigLocal();
        fecharGithubModal();
        setStatus("Configuração GitHub salva.", "status-success");
    }

    function validarGithubConfig() {
        if (!githubConfig.owner || !githubConfig.repo || !githubConfig.branch || !githubConfig.path || !githubConfig.token) {
            setStatus("Config GitHub incompleta. Preencha Owner/Repo/Branch/Path/Token.", "status-error");
            return false;
        }
        return true;
    }

    function githubContentsReadUrl() {
        const owner = encodeURIComponent(githubConfig.owner);
        const repo = encodeURIComponent(githubConfig.repo);
        const path = githubConfig.path.split("/").map((p) => encodeURIComponent(p)).join("/");
        const branch = encodeURIComponent(githubConfig.branch);
        return `https://api.github.com/repos/${owner}/${repo}/contents/${path}?ref=${branch}`;
    }

    function githubContentsWriteUrl() {
        const owner = encodeURIComponent(githubConfig.owner);
        const repo = encodeURIComponent(githubConfig.repo);
        const path = githubConfig.path.split("/").map((p) => encodeURIComponent(p)).join("/");
        return `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    }

    function toBase64Utf8(texto) {
        const bytes = new TextEncoder().encode(texto);
        let binary = "";
        for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
        return btoa(binary);
    }

    function fromBase64Utf8(base64) {
        const binary = atob(base64.replace(/\n/g, ""));
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return new TextDecoder().decode(bytes);
    }

    async function carregarDadosDoGithub() {
        if (!validarGithubConfig()) return;

        try {
            setStatus("Carregando JSON do GitHub...", "status-loading");
            const resp = await fetch(githubContentsReadUrl(), {
                headers: {
                    Authorization: `Bearer ${githubConfig.token}`,
                    Accept: "application/vnd.github+json"
                }
            });

            if (!resp.ok) {
                if (resp.status === 404) {
                    throw new Error("Arquivo JSON não encontrado no repositório. Salve uma vez para criar.");
                }
                const erro = await resp.text();
                throw new Error(`GitHub GET falhou (${resp.status}): ${erro}`);
            }

            const payload = await resp.json();
            githubFileSha = payload.sha || null;
            const raw = fromBase64Utf8(payload.content || "");
            const parsed = JSON.parse(raw);

            if (!parsed || !Array.isArray(parsed.dadosProcessadosGlobais) || parsed.dadosProcessadosGlobais.length === 0) {
                throw new Error("JSON não contém 'dadosProcessadosGlobais' válido.");
            }

            limparCronometrosAtivos();
            dadosProcessadosGlobais = parsed.dadosProcessadosGlobais;
            renderHtmlCards(dadosProcessadosGlobais, "tableContainer");
            btnExportar.disabled = dadosProcessadosGlobais.length <= 1;
            setStatus("Dados carregados do GitHub com sucesso.", "status-success");
        } catch (error) {
            setStatus(`Erro ao carregar do GitHub: ${error.message}`, "status-error");
        }
    }

    async function salvarDadosNoGithub() {
        if (!validarGithubConfig()) return;
        if (!dadosProcessadosGlobais || dadosProcessadosGlobais.length === 0) {
            setStatus("Sem dados para salvar no GitHub.", "status-error");
            return;
        }

        try {
            setStatus("Salvando JSON no GitHub...", "status-loading");

            if (!githubFileSha) {
                const preflight = await fetch(githubContentsReadUrl(), {
                    headers: {
                        Authorization: `Bearer ${githubConfig.token}`,
                        Accept: "application/vnd.github+json"
                    }
                });
                if (preflight.ok) {
                    const current = await preflight.json();
                    githubFileSha = current.sha || null;
                }
            }

            const documento = {
                schemaVersion: 1,
                updatedAt: new Date().toISOString(),
                dadosProcessadosGlobais
            };

            const body = {
                message: `Atualiza atividades em ${new Date().toLocaleString("pt-BR")}`,
                content: toBase64Utf8(JSON.stringify(documento, null, 2)),
                branch: githubConfig.branch
            };
            if (githubFileSha) body.sha = githubFileSha;

            const resp = await fetch(githubContentsWriteUrl(), {
                method: "PUT",
                headers: {
                    Authorization: `Bearer ${githubConfig.token}`,
                    Accept: "application/vnd.github+json",
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(body)
            });

            if (!resp.ok) {
                const erro = await resp.text();
                throw new Error(`GitHub PUT falhou (${resp.status}): ${erro}`);
            }

            const result = await resp.json();
            githubFileSha = result && result.content ? result.content.sha : githubFileSha;
            setStatus("Dados salvos no GitHub com sucesso.", "status-success");
        } catch (error) {
            setStatus(`Erro ao salvar no GitHub: ${error.message}`, "status-error");
        }
    }

    function setStatus(message, cssClass) {
        statusBox.textContent = message;
        statusBox.className = cssClass;
    }
})();

(() => {
    const USER_PADRAO = "HP";
    const loginOverlay = document.getElementById("loginOverlay");
    const loginForm = document.getElementById("loginForm");
    const loginUser = document.getElementById("loginUser");
    const loginError = document.getElementById("loginError");

    if (!loginOverlay || !loginForm || !loginUser || !loginError) return;

    loginForm.addEventListener("submit", (event) => {
        event.preventDefault();

        const usuario = String(loginUser.value || "").trim().toUpperCase();
        if (usuario !== USER_PADRAO) {
            loginError.textContent = `Usuário inválido. Digite "${USER_PADRAO}".`;
            loginUser.focus();
            loginUser.select();
            return;
        }

        loginError.textContent = "";
        loginOverlay.style.display = "none";
    });

    loginUser.value = "";
    loginUser.focus();
})();

(() => {
    const DEFAULT_CONFIG = { area: "Infra", matricula: "18881", colab: "Alisson" };
    const STORAGE_KEY = "os_processor_config";
    const cabecalho = ["AREA", "MATRICULA", "COLABORADOR", "DATA", "DESCRIÇÃO DA ATIVIDADE", "HORA INÍCIO", "HORA TÉRMINO", "SOLICITANTE", "SETOR", "OBSERVAÇÃO"];
    const origemInput = document.getElementById("origemFile");
    if (!origemInput) return;

    const osBtnSave = document.getElementById("osBtnSave");
    const osBtnConfig = document.getElementById("osBtnConfig");
    const osBtnClearLog = document.getElementById("osBtnClearLog");
    const osStatus = document.getElementById("osStatus");
    const osLog = document.getElementById("osLog");
    const osRowCount = document.getElementById("osRowCount");
    const osTableEmpty = document.getElementById("osTableEmpty");
    const osTableContainer = document.getElementById("osTableContainer");
    const osTableHeader = document.getElementById("osTableHeader");
    const osTableBody = document.getElementById("osTableBody");

    const osConfirmModal = document.getElementById("osConfirmModal");
    const osConfirmBody = document.getElementById("osConfirmBody");
    const osConfirmCancel = document.getElementById("osConfirmCancel");
    const osConfirmOk = document.getElementById("osConfirmOk");

    const osSettingsModal = document.getElementById("osSettingsModal");
    const osSettingsCancel = document.getElementById("osSettingsCancel");
    const osSettingsSave = document.getElementById("osSettingsSave");
    const osConfArea = document.getElementById("osConfArea");
    const osConfMatricula = document.getElementById("osConfMatricula");
    const osConfColab = document.getElementById("osConfColab");

    let appConfig = carregarConfig();
    let dadosConsolidados = [cabecalho];
    let onConfirmAction = null;

    document.addEventListener("source-sheet-loaded", (event) => {
        const detail = event.detail || {};
        if (!Array.isArray(detail.origData)) return;
        processarOrigemRaw(detail.origData, detail.fileName || "");
    });
    osBtnSave.addEventListener("click", onSaveClick);
    osBtnConfig.addEventListener("click", abrirModalConfig);
    osBtnClearLog.addEventListener("click", limparLog);

    osConfirmCancel.addEventListener("click", fecharConfirmacao);
    osConfirmOk.addEventListener("click", () => {
        const acao = onConfirmAction;
        fecharConfirmacao();
        if (typeof acao === "function") acao();
    });

    osSettingsCancel.addEventListener("click", fecharModalConfig);
    osSettingsSave.addEventListener("click", salvarConfig);

    osConfirmModal.addEventListener("click", (event) => {
        if (event.target === osConfirmModal) fecharConfirmacao();
    });
    osSettingsModal.addEventListener("click", (event) => {
        if (event.target === osSettingsModal) fecharModalConfig();
    });

    printLog("Ready.");
    setOsStatus("Aguardando", "status-loading");

    function carregarConfig() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return { ...DEFAULT_CONFIG };
            const parsed = JSON.parse(raw);
            return {
                area: parsed.area || DEFAULT_CONFIG.area,
                matricula: parsed.matricula || DEFAULT_CONFIG.matricula,
                colab: parsed.colab || DEFAULT_CONFIG.colab
            };
        } catch (_) {
            return { ...DEFAULT_CONFIG };
        }
    }

    function salvarConfigLocal() {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(appConfig));
    }

    function setOsStatus(msg, cssClass) {
        osStatus.textContent = msg;
        osStatus.className = cssClass;
    }

    function printLog(msg) {
        const hora = new Date().toLocaleTimeString("pt-BR", { hour12: false });
        osLog.textContent += `[${hora}] ${msg}\n`;
        osLog.scrollTop = osLog.scrollHeight;
    }

    function limparLog() {
        osLog.textContent = "";
        printLog("Log limpo.");
    }

    function formatDateBr(dateObj) {
        if (!dateObj || Number.isNaN(dateObj.getTime())) return "";
        const d = String(dateObj.getDate()).padStart(2, "0");
        const m = String(dateObj.getMonth() + 1).padStart(2, "0");
        const y = dateObj.getFullYear();
        return `${d}/${m}/${y}`;
    }

    function formatTimeOnly(dateObj) {
        if (!dateObj || Number.isNaN(dateObj.getTime())) return "";
        const h = String(dateObj.getHours()).padStart(2, "0");
        const m = String(dateObj.getMinutes()).padStart(2, "0");
        const s = String(dateObj.getSeconds()).padStart(2, "0");
        return `${h}:${m}:${s}`;
    }

    function corrigirDataBrasil(valor) {
        if (valor == null || valor === "") return null;

        if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
            return new Date(valor.getTime());
        }

        if (typeof valor === "number") {
            const dt = new Date(Math.round((valor - 25569) * 86400 * 1000));
            dt.setMinutes(dt.getMinutes() + dt.getTimezoneOffset());
            return dt;
        }

        const strValor = String(valor).trim();
        const partes = strValor.split(" ");
        const dataPart = partes[0];
        const timePart = partes[1] || "";
        const partesData = dataPart.split("/");

        if (partesData.length === 3) {
            const p1 = parseInt(partesData[0], 10);
            const p2 = parseInt(partesData[1], 10);
            let ano = parseInt(partesData[2], 10);
            if (ano < 100) ano += 2000;

            let dia;
            let mes;
            if (p1 > 12) {
                dia = p1;
                mes = p2;
            } else {
                mes = p1;
                dia = p2;
            }

            const dateFinal = new Date(ano, mes - 1, dia);
            if (timePart) {
                const t = timePart.split(":");
                if (t.length >= 2) {
                    dateFinal.setHours(parseInt(t[0], 10), parseInt(t[1], 10), t[2] ? parseInt(t[2], 10) : 0, 0);
                }
            }
            return dateFinal;
        }

        const nativo = new Date(strValor);
        return Number.isNaN(nativo.getTime()) ? null : nativo;
    }

    function renderTabelaOS() {
        osTableHeader.innerHTML = "";
        osTableBody.innerHTML = "";

        cabecalho.forEach((text) => {
            const th = document.createElement("th");
            th.textContent = text;
            osTableHeader.appendChild(th);
        });
        const thAcoes = document.createElement("th");
        thAcoes.textContent = "Ações";
        osTableHeader.appendChild(thAcoes);

        for (let i = 1; i < dadosConsolidados.length; i++) {
            const tr = document.createElement("tr");
            const row = dadosConsolidados[i];

            row.forEach((val, cellIdx) => {
                const td = document.createElement("td");
                const input = document.createElement("input");
                input.type = "text";
                input.className = "os-cell-input";
                input.value = val || "";
                input.addEventListener("change", (e) => {
                    dadosConsolidados[i][cellIdx] = e.target.value;
                });
                td.appendChild(input);
                tr.appendChild(td);
            });

            const tdDelete = document.createElement("td");
            tdDelete.style.textAlign = "center";
            const btnDelete = document.createElement("button");
            btnDelete.type = "button";
            btnDelete.className = "os-btn-delete";
            btnDelete.textContent = "Excluir";
            btnDelete.addEventListener("click", () => {
                dadosConsolidados.splice(i, 1);
                renderTabelaOS();
            });
            tdDelete.appendChild(btnDelete);
            tr.appendChild(tdDelete);
            osTableBody.appendChild(tr);
        }

        const count = Math.max(0, dadosConsolidados.length - 1);
        osRowCount.textContent = String(count);
        const temLinhas = count > 0;
        osTableContainer.hidden = !temLinhas;
        osTableEmpty.style.display = temLinhas ? "none" : "block";
    }

    function processarOrigemRaw(raw, fileName) {
        osBtnSave.disabled = true;
        setOsStatus("Processando dados...", "status-loading");
        printLog(`Lendo arquivo ${fileName || "(sem nome)"}`);

        try {
            dadosConsolidados = [cabecalho];
            for (let i = 1; i < raw.length; i++) {
                const r = raw[i];
                if (!r || !r[2]) continue;

                const dataObj = corrigirDataBrasil(r[51]);
                const horaInicioObj = corrigirDataBrasil(r[22]);
                const horaTerminoObj = corrigirDataBrasil(r[51]);

                dadosConsolidados.push([
                    appConfig.area,
                    appConfig.matricula,
                    appConfig.colab,
                    formatDateBr(dataObj),
                    `Problema: ${r[2] || ""}; Solução: ${r[7] || ""}`,
                    formatTimeOnly(horaInicioObj),
                    formatTimeOnly(horaTerminoObj),
                    r[4] || "",
                    r[1] || "",
                    r[32] || ""
                ]);
            }

            renderTabelaOS();
            osBtnSave.disabled = dadosConsolidados.length <= 1;
            setOsStatus("Revisão pronta", "status-success");
            printLog(`Importação concluída. ${dadosConsolidados.length - 1} linha(s) mapeada(s).`);
        } catch (err) {
            setOsStatus("Falha no processamento", "status-error");
            printLog(`Erro: ${err.message}`);
        }
    }

    function gerarArquivosSemanais() {
        if (dadosConsolidados.length <= 1) {
            printLog("Sem dados para exportar.");
            return;
        }

        const semanas = new Map();
        for (let i = 1; i < dadosConsolidados.length; i++) {
            const row = dadosConsolidados[i];
            const dataString = row[3];
            if (!dataString) continue;

            const [d, m, y] = String(dataString).split("/").map(Number);
            if (!d || !m || !y) continue;

            const dt = new Date(y, m - 1, d);
            if (Number.isNaN(dt.getTime())) continue;

            const dom = new Date(dt);
            dom.setDate(dt.getDate() - dt.getDay());
            const sab = new Date(dom);
            sab.setDate(dom.getDate() + 6);

            const key = `${formatDateBr(dom).replace(/\//g, "-")} ate ${formatDateBr(sab).replace(/\//g, "-")}`;
            if (!semanas.has(key)) semanas.set(key, [cabecalho]);
            semanas.get(key).push(row);
        }

        if (semanas.size === 0) {
            printLog("Nenhuma data válida encontrada para exportação.");
            setOsStatus("Sem datas válidas", "status-error");
            return;
        }

        semanas.forEach((linhas, periodo) => {
            const ws = XLSX.utils.aoa_to_sheet(linhas);
            ws["!cols"] = cabecalho.map((_, idx) => ({ wch: idx === 4 ? 65 : 15 }));
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "OS_Data");
            const colab = String(appConfig.colab || "colaborador").replace(/[<>:"/\\|?*]/g, "_");
            XLSX.writeFile(wb, `OS_Semana_${periodo}_${colab}.xlsx`);
        });

        setOsStatus("Exportação concluída", "status-success");
        printLog(`${semanas.size} arquivo(s) semanal(is) gerado(s).`);
    }

    function onSaveClick() {
        mostrarConfirmacao("Confirmar geração dos relatórios semanais com os dados revisados?", gerarArquivosSemanais);
    }

    function mostrarConfirmacao(texto, onConfirm) {
        osConfirmBody.textContent = texto;
        onConfirmAction = onConfirm;
        osConfirmModal.style.display = "flex";
    }

    function fecharConfirmacao() {
        onConfirmAction = null;
        osConfirmModal.style.display = "none";
    }

    function abrirModalConfig() {
        osConfArea.value = appConfig.area;
        osConfMatricula.value = appConfig.matricula;
        osConfColab.value = appConfig.colab;
        osSettingsModal.style.display = "flex";
    }

    function fecharModalConfig() {
        osSettingsModal.style.display = "none";
    }

    function salvarConfig() {
        appConfig = {
            area: osConfArea.value.trim() || DEFAULT_CONFIG.area,
            matricula: osConfMatricula.value.trim() || DEFAULT_CONFIG.matricula,
            colab: osConfColab.value.trim() || DEFAULT_CONFIG.colab
        };
        salvarConfigLocal();
        printLog("Configurações salvas.");
        fecharModalConfig();

        if (dadosConsolidados.length > 1) {
            mostrarConfirmacao("Aplicar novas configurações nas linhas já importadas?", () => {
                for (let i = 1; i < dadosConsolidados.length; i++) {
                    dadosConsolidados[i][0] = appConfig.area;
                    dadosConsolidados[i][1] = appConfig.matricula;
                    dadosConsolidados[i][2] = appConfig.colab;
                }
                renderTabelaOS();
                printLog("Configurações aplicadas aos dados atuais.");
            });
        }
    }
})();
