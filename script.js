// =======================
// CONFIGURAÇÃO BÁSICA
// =======================

// Prefixo para localStorage (evita conflito entre aplicações)
const APP_PREFIX = 'horarios_escola_';

// Dias e turnos
const DIAS_SEMANA = ['SEGUNDA', 'TERCA', 'QUARTA', 'QUINTA', 'SEXTA', 'SABADO'];
const TURNOS = ['MANHA', 'TARDE', 'NOITE'];

// Tempos de aula por turno
let TEMPOS_POR_TURNO = {
    MANHA: [
        { id: 1, inicio: '07:10', fim: '07:50', etiqueta: '1º Tempo', intervalo: false },
        { id: 2, inicio: '07:50', fim: '08:30', etiqueta: '2º Tempo', intervalo: false },
        { id: 3, inicio: '08:30', fim: '09:10', etiqueta: '3º Tempo', intervalo: false },
        { id: 4, inicio: '09:10', fim: '09:50', etiqueta: '4º Tempo', intervalo: false },
        { id: 5, inicio: '09:50', fim: '10:10', etiqueta: 'INTERVALO', intervalo: true },
        { id: 6, inicio: '10:10', fim: '10:50', etiqueta: '5º Tempo', intervalo: false },
        { id: 7, inicio: '10:50', fim: '11:30', etiqueta: '6º Tempo', intervalo: false },
        { id: 8, inicio: '11:30', fim: '12:10', etiqueta: '7º Tempo', intervalo: false }
    ],
    TARDE: [
        { id: 1, inicio: '13:30', fim: '14:10', etiqueta: '1º Tempo', intervalo: false },
        { id: 2, inicio: '14:10', fim: '14:50', etiqueta: '2º Tempo', intervalo: false },
        { id: 3, inicio: '14:50', fim: '15:30', etiqueta: '3º Tempo', intervalo: false },
        { id: 4, inicio: '15:30', fim: '16:10', etiqueta: '4º Tempo', intervalo: false },
        { id: 5, inicio: '16:10', fim: '16:30', etiqueta: 'INTERVALO', intervalo: true },
        { id: 6, inicio: '16:30', fim: '17:10', etiqueta: '5º Tempo', intervalo: false },
        { id: 7, inicio: '17:10', fim: '17:50', etiqueta: '6º Tempo', intervalo: false }
    ],
    NOITE: [
        { id: 1, inicio: '18:30', fim: '19:15', etiqueta: '1º Tempo', intervalo: false },
        { id: 2, inicio: '19:15', fim: '20:00', etiqueta: '2º Tempo', intervalo: false },
        { id: 3, inicio: '20:00', fim: '20:45', etiqueta: '3º Tempo', intervalo: false },
        { id: 4, inicio: '20:45', fim: '21:00', etiqueta: 'INTERVALO', intervalo: true },
        { id: 5, inicio: '21:00', fim: '21:45', etiqueta: '4º Tempo', intervalo: false },
        { id: 6, inicio: '21:45', fim: '22:30', etiqueta: '5º Tempo', intervalo: false }
    ]
};

function getTempos(turno) {
    return TEMPOS_POR_TURNO[turno] || TEMPOS_POR_TURNO.MANHA;
}

function siglaTurno(turno) {
    if (turno === 'MANHA') return 'M';
    if (turno === 'TARDE') return 'T';
    if (turno === 'NOITE') return 'N';
    return '';
}

function etiquetaTempoAbrev(tempo, turno) {
    if (tempo.intervalo) return 'INTERVALO';
    const sigla = siglaTurno(turno);
    return tempo.etiqueta.replace(' Tempo', ` Tempo-${sigla}`);
}

function etiquetaTempoMin(tempo) {
    if (tempo.intervalo) return 'INTERVALO';
    return tempo.etiqueta.replace(' Tempo', ' T');
}
// Dados principais
let cursos = ['Informática', 'Administração', 'Enfermagem'];
let professores = []; // {id, nome, disciplinas[], cor}
let turmas = [];      // {id, nome, curso, turno, descricao}
let aulas = [];       // {id, turno, dia, turmaId, tempoId, disciplina, professorId}

const ROLES = {
    ADMIN: 'ADMIN',
    COORD_TURNO: 'COORD_TURNO'
};

const USUARIOS_FIXOS = [
    {
        id: 'admin',
        nome: 'Administrador da Escola',
        login: 'admin',
        senha: 'admin',
        role: ROLES.ADMIN
    }
];

let usuarioLogado = null;

// Controle do editor visual / ajuste inteligente
let turnoAtualGrade = 'MANHA';
let modoVisualGrade = 'TURNO';
let cursoVisualGrade = '';
let pilhaUndo = []; // para desfazer ajustes inteligentes

// Controle do modal de edição de aula
let aulaEmEdicao = null;

let chartProfDias = null;
let chartProfTurnos = null;
let chartProfDiasMes = null;

const PALETA_GRAFICOS = ['#3498db', '#e74c3c', '#2ecc71', '#9b59b6', '#f1c40f', '#e67e22'];


// Cache para otimização
let cacheGrade = {};
let dadosModificados = false;

// =======================
// UTILITÁRIOS
// =======================

function gerarId() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}

function textoTurno(turno) {
    if (turno === 'MANHA') return 'Manhã';
    if (turno === 'TARDE') return 'Tarde';
    if (turno === 'NOITE') return 'Noite';
    return turno;
}

function textoDia(sigla) {
    switch (sigla) {
        case 'SEGUNDA': return 'Segunda';
        case 'TERCA':   return 'Terça';
        case 'QUARTA':  return 'Quarta';
        case 'QUINTA':  return 'Quinta';
        case 'SEXTA':   return 'Sexta';
        case 'SABADO':  return 'Sábado';
        default:        return sigla;
    }
}

function abreviar(str, tam) {
    if (!str) return '';
    return str.length > tam ? str.slice(0, tam - 1) + '…' : str;
}

function mostrarToast(mensagem, tipo = 'success') {
    // Remove toast anterior
    const toastAntigo = document.querySelector('.toast');
    if (toastAntigo) toastAntigo.remove();
    
    // Cria novo toast
    const toast = document.createElement('div');
    toast.className = `toast ${tipo}`;
    toast.textContent = mensagem;
    document.body.appendChild(toast);
    
    // Remove automaticamente após 3 segundos
    setTimeout(() => {
        if (toast.parentNode) {
            toast.style.animation = 'slideIn 0.3s ease reverse';
            setTimeout(() => toast.remove(), 300);
        }
    }, 3000);
}

// banner removido
// LocalStorage
function salvarLocal() {
    try {
        localStorage.setItem(`${APP_PREFIX}professores`, JSON.stringify(professores));
        localStorage.setItem(`${APP_PREFIX}turmas`, JSON.stringify(turmas));
        localStorage.setItem(`${APP_PREFIX}aulas`, JSON.stringify(aulas));
        localStorage.setItem(`${APP_PREFIX}cursos`, JSON.stringify(cursos));
        localStorage.setItem(`${APP_PREFIX}tempos_por_turno`, JSON.stringify(TEMPOS_POR_TURNO));
        dadosModificados = true;
        mostrarToast('Dados salvos com sucesso!');
    } catch (e) {
        console.warn('Não foi possível salvar no localStorage', e);
        mostrarToast('Erro ao salvar dados', 'error');
    }
}

function carregarLocal() {
    try {
        const p = localStorage.getItem(`${APP_PREFIX}professores`);
        const t = localStorage.getItem(`${APP_PREFIX}turmas`);
        const a = localStorage.getItem(`${APP_PREFIX}aulas`);
        const c = localStorage.getItem(`${APP_PREFIX}cursos`);
        const tempos = localStorage.getItem(`${APP_PREFIX}tempos_por_turno`);
        if (p) professores = JSON.parse(p);
        if (t) turmas = JSON.parse(t);
        if (a) aulas = JSON.parse(a);
        if (c) cursos = JSON.parse(c);
        if (tempos) {
            try { TEMPOS_POR_TURNO = JSON.parse(tempos); } catch {}
        }
    } catch (e) {
        console.warn('Não foi possível carregar do localStorage', e);
        mostrarToast('Erro ao carregar dados salvos', 'warning');
    }
}

function migrarEstruturaLocal() {
    let alterou = false;
    if (!TEMPOS_POR_TURNO || !TEMPOS_POR_TURNO.MANHA || !TEMPOS_POR_TURNO.TARDE || !TEMPOS_POR_TURNO.NOITE) {
        TEMPOS_POR_TURNO = TEMPOS_POR_TURNO || {};
        TEMPOS_POR_TURNO.MANHA = TEMPOS_POR_TURNO.MANHA || getTempos('MANHA');
        TEMPOS_POR_TURNO.TARDE = TEMPOS_POR_TURNO.TARDE || getTempos('TARDE');
        TEMPOS_POR_TURNO.NOITE = TEMPOS_POR_TURNO.NOITE || getTempos('NOITE');
        alterou = true;
    }
    ['MANHA','TARDE','NOITE'].forEach(turno => {
        const arr = TEMPOS_POR_TURNO[turno] || [];
        arr.forEach((x, i) => { if (x && typeof x.id !== 'number') { x.id = i + 1; alterou = true; } });
    });
    professores = Array.isArray(professores) ? professores : [];
    turmas = Array.isArray(turmas) ? turmas : [];
    aulas = Array.isArray(aulas) ? aulas : [];
    professores.forEach(p => {
        if (p && !Array.isArray(p.disciplinas)) {
            if (typeof p.disciplinas === 'string' && p.disciplinas.trim()) {
                p.disciplinas = p.disciplinas
                    .split(',')
                    .map(s => s.trim())
                    .filter(Boolean);
            } else {
                p.disciplinas = [];
            }
            alterou = true;
        }
        if (!Array.isArray(p.diasDisponiveis)) { p.diasDisponiveis = []; alterou = true; }
        if (!Array.isArray(p.turnosDisponiveis)) { p.turnosDisponiveis = []; alterou = true; }
        if (p.disponibilidadePorDia === undefined) { p.disponibilidadePorDia = null; alterou = true; }
        if (p.temposPorDiaTurno === undefined) { p.temposPorDiaTurno = null; alterou = true; }
        if (!Array.isArray(p.temposDisponiveis) || p.temposDisponiveis.length === 0) {
            p.temposDisponiveis = getTempos('MANHA').filter(t => !t.intervalo).map(t => t.id);
            alterou = true;
        }
    });
    turmas.forEach(t => {
        if (!t.turno || !['MANHA','TARDE','NOITE'].includes(t.turno)) { t.turno = 'MANHA'; alterou = true; }
    });
    aulas.forEach(a => {
        if (!a.turno || !['MANHA','TARDE','NOITE'].includes(a.turno)) { a.turno = 'MANHA'; alterou = true; }
        if (!DIAS_SEMANA.includes(a.dia)) { a.dia = 'SEGUNDA'; alterou = true; }
        if (typeof a.tempoId !== 'number') { a.tempoId = parseInt(a.tempoId, 10) || 1; alterou = true; }
    });
    if (alterou) salvarLocal();
}

// Verificar conflitos
function verificarConflitoProfessor(professorId, turno, dia, tempoId, aulaId = null) {
    return aulas.some(a => 
        a.id !== aulaId && // Ignora a própria aula se estiver editando
        a.professorId === professorId &&
        a.turno === turno &&
        a.dia === dia &&
        a.tempoId === tempoId
    );
}

function professorDisponivelNoHorario(prof, turno, dia, tempoId) {
    if (!prof) return true;
    const disp = prof.disponibilidadePorDia || null;
    const temposMap = prof.temposPorDiaTurno || null;
    const diasF = prof.diasDisponiveis || [];
    const turnosF = prof.turnosDisponiveis || [];
    const temposF = (prof.temposDisponiveis && prof.temposDisponiveis.length
        ? prof.temposDisponiveis
        : getTempos('MANHA').filter(t => !t.intervalo).map(t => t.id));

    let okTempo = temposF.length === 0 || temposF.includes(tempoId);
    if (temposMap && temposMap[dia] && temposMap[dia][turno]) {
        const lista = temposMap[dia][turno] || [];
        okTempo = lista.length === 0 ? okTempo : lista.includes(tempoId);
    }

    if (disp) {
        const turnosDia = disp[dia] || [];
        const okDiaTurno = turnosDia.length === 0 ? true : turnosDia.includes(turno);
        return okDiaTurno && okTempo;
    } else {
        const okDia = diasF.length === 0 || diasF.includes(dia);
        const okTurno = turnosF.length === 0 || turnosF.includes(turno);
        return okDia && okTurno && okTempo;
    }
}

// =======================
// NAVEGAÇÃO ENTRE SEÇÕES
// =======================

function autenticarUsuario(login, senha) {
    return USUARIOS_FIXOS.find(u => u.login === login && u.senha === senha) || null;
}

function aplicarPermissoesNaUI() {
    const menuBtns = Array.from(document.querySelectorAll('.menu-btn'));
    const btnRelatorios = menuBtns.find(b => b.textContent.includes('Relatórios'));
    const btnConsultas = menuBtns.find(b => b.textContent.includes('Consultas'));
    if (!usuarioLogado) {
        if (btnRelatorios) btnRelatorios.style.display = 'none';
        if (btnConsultas) btnConsultas.style.display = 'none';
        const blocoTempoVago = document.getElementById('profRelatorioMostrarVagos')?.closest('.form-group');
        const blocoGraficos = document.getElementById('profRelatorioMostrarGraficos')?.closest('.form-group');
        const charts = document.querySelector('.prof-charts');
        if (blocoTempoVago) blocoTempoVago.style.display = 'none';
        if (blocoGraficos) blocoGraficos.style.display = 'none';
        if (charts) charts.style.display = 'none';
        return;
    }
    if (btnRelatorios) btnRelatorios.style.display = '';
    const blocoTempoVago = document.getElementById('profRelatorioMostrarVagos')?.closest('.form-group');
    const blocoGraficos = document.getElementById('profRelatorioMostrarGraficos')?.closest('.form-group');
    const charts = document.querySelector('.prof-charts');
    const turnoRel = document.getElementById('turnoRelatorio');
    const turnoRelCurso = document.getElementById('turnoRelatorioCurso');
    if (usuarioLogado.role === ROLES.ADMIN) {
        if (btnConsultas) btnConsultas.style.display = '';
        if (blocoTempoVago) blocoTempoVago.style.display = '';
        if (blocoGraficos) blocoGraficos.style.display = '';
        if (charts) charts.style.display = '';
        if (turnoRel) {
            turnoRel.disabled = false;
        }
        if (turnoRelCurso) {
            turnoRelCurso.disabled = false;
        }
        return;
    }
    if (btnConsultas) btnConsultas.style.display = 'none';
    if (blocoTempoVago) blocoTempoVago.style.display = 'none';
    if (blocoGraficos) blocoGraficos.style.display = 'none';
    if (charts) charts.style.display = 'none';
    if (usuarioLogado.role === ROLES.COORD_TURNO && usuarioLogado.turno) {
        if (turnoRel) {
            turnoRel.value = usuarioLogado.turno;
            turnoRel.disabled = true;
        }
        if (turnoRelCurso) {
            turnoRelCurso.value = usuarioLogado.turno;
            turnoRelCurso.disabled = true;
        }
    }
}

function mostrarHintEditorTurno() {
    const text = document.getElementById('editorTurnoHintText');
    const chk = document.getElementById('chkMostrarInstrucoesEditor');
    if (!text || !chk) return;
    text.style.display = chk.checked ? '' : 'none';
}

function sairSistema() {
    usuarioLogado = null;
    try {
        localStorage.removeItem('usuarioAtual');
    } catch (e) {
    }
    const loginCard = document.getElementById('loginCard');
    const mainContent = document.querySelector('.main-content');
    const loginHint = document.getElementById('loginHint');
    const inputLogin = document.getElementById('loginUsuario');
    const inputSenha = document.getElementById('senhaUsuario');
    if (mainContent) {
        mainContent.style.display = 'none';
    }
    if (loginCard) {
        loginCard.style.display = '';
    }
    if (loginHint) {
        loginHint.textContent = '';
    }
    if (inputLogin) {
        inputLogin.value = '';
    }
    if (inputSenha) {
        inputSenha.value = '';
    }
    aplicarPermissoesNaUI();
}

function showSection(id) {
    document.querySelectorAll('.section').forEach(sec => sec.classList.remove('active'));
    document.getElementById(id).classList.add('active');

    document.querySelectorAll('.menu-btn').forEach(btn => btn.classList.remove('active'));
    const alvo = Array.from(document.querySelectorAll('.menu-btn'))
        .find(b => b.getAttribute('onclick')?.includes(`'${id}'`));
    if (alvo) alvo.classList.add('active');
    
    // Atualiza selects específicos quando a seção é aberta
    if (id === 'relatorios') {
        preencherSelectCursoRelatorio();
        preencherSelectProfessorRelatorio();
    }
    if (id === 'horarios') {
        mostrarHintEditorTurno();
    }
}

// =======================
// SELEÇÃO VISUAL (CHIPS)
// =======================

function selecionarChip(groupId, hiddenId, btn) {
    const group = document.getElementById(groupId);
    if (!group) return;
    group.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
    btn.classList.add('selecionado');

    const hidden = document.getElementById(hiddenId);
    if (hidden) hidden.value = btn.getAttribute('data-value');
}

// turno no formulário de horário
function selecionarChipTurnoHorario(btn) {
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        const valor = btn.getAttribute('data-value');
        if (usuarioLogado.turno && usuarioLogado.turno !== valor) {
            mostrarToast('Você só pode trabalhar no seu turno.', 'warning');
            return;
        }
    }
    selecionarChip('chips-turno-horario', 'turnoHorario', btn);
    atualizarChipsTurmasHorario();
    atualizarChipsTempos();
    turnoAtualGrade = btn.getAttribute('data-value');
    montarGradeTurno(turnoAtualGrade);
}

// turno na visualização da grade
function mudarTurnoVisual(btn) {
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        const valor = btn.getAttribute('data-value');
        if (usuarioLogado.turno && usuarioLogado.turno !== valor) {
            mostrarToast('Você só pode visualizar o seu turno.', 'warning');
            return;
        }
    }
    selecionarChip('chips-turno-visual', 'turnoGradeFake', btn);
    turnoAtualGrade = btn.getAttribute('data-value');
    montarGradeTurno(turnoAtualGrade);
}

function mudarModoVisual(btn) {
    const group = document.getElementById('chips-modo-visual');
    if (group) {
        group.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
        btn.classList.add('selecionado');
    }
    modoVisualGrade = btn.getAttribute('data-value') || 'TURNO';
    const filtroCurso = document.getElementById('filtro-curso-visual');
    if (filtroCurso) {
        filtroCurso.style.display = modoVisualGrade === 'CURSO' ? '' : 'none';
    }
    montarGradeTurno(turnoAtualGrade);
}

function mudarCursoVisual(select) {
    cursoVisualGrade = select.value || '';
    montarGradeTurno(turnoAtualGrade);
}

// Atualiza valor do hidden de tempos com base nos chips selecionados
function atualizarHiddenTempos() {
    const hidden = document.getElementById('tempoHorario');
    if (!hidden) return;

    const selecionados = Array
        .from(document.querySelectorAll('#chips-tempo .chip.selecionado'))
        .map(c => c.dataset.value);

    hidden.value = selecionados.join(',');
}

// tempos clicáveis (MULTI-SELEÇÃO)
function atualizarChipsTempos() {
    const container = document.getElementById('chips-tempo');
    if (!container) return;

    container.innerHTML = '';
    const hidden = document.getElementById('tempoHorario');
    if (hidden) hidden.value = '';

    const turnoSel = document.getElementById('turnoHorario')?.value || turnoAtualGrade || 'MANHA';
    const tempos = getTempos(turnoSel);
    tempos.forEach(t => {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'chip';
        chip.dataset.value = t.id;
        chip.textContent = t.intervalo ? 'INTERVALO' : etiquetaTempoAbrev(t, turnoSel);
        chip.onclick = () => {
            if (t.intervalo) {
                mostrarToast('Não é possível lançar aula no intervalo.', 'warning');
                return;
            }
            chip.classList.toggle('selecionado');
            atualizarHiddenTempos();
        };
        container.appendChild(chip);
    });
}

// =======================
// PROFESSORES
// =======================

function atualizarListaProfessores() {
    const corpo = document.getElementById('listaProfessores');
    const selProf = document.getElementById('professorHorario');
    const selProfRel = document.getElementById('professorRelatorio');
    const selProfModal = document.getElementById('modalProfessor');

    if (!corpo) return;
    corpo.innerHTML = '';
    if (selProf) selProf.innerHTML = '<option value="">Selecione...</option>';
    if (selProfRel) selProfRel.innerHTML = '<option value="">Selecione...</option>';
    if (selProfModal) selProfModal.innerHTML = '<option value="">Selecione...</option>';

    professores.forEach(p => {
        // tabela
        const tr = document.createElement('tr');

        const tdNome = document.createElement('td');
        tdNome.textContent = p.nome;
        tr.appendChild(tdNome);

        const tdDisc = document.createElement('td');
        tdDisc.textContent = p.disciplinas.join(', ');
        tr.appendChild(tdDisc);

        const tdCor = document.createElement('td');
        const box = document.createElement('div');
        box.style.width = '26px';
        box.style.height = '16px';
        box.style.borderRadius = '6px';
        box.style.background = p.cor;
        tdCor.appendChild(box);
        tr.appendChild(tdCor);

        const tdAcoes = document.createElement('td');
        tdAcoes.style.display = 'flex';
        tdAcoes.style.gap = '6px';

        const btnEd = document.createElement('button');
        btnEd.className = 'btn-secondary small';
        btnEd.textContent = 'Editar';
        btnEd.onclick = () => abrirModalProfessor(p.id);

        const btnDel = document.createElement('button');
        btnDel.className = 'btn-danger small';
        btnDel.textContent = 'Excluir';
        btnDel.onclick = () => excluirProfessor(p.id);

        tdAcoes.appendChild(btnEd);
        tdAcoes.appendChild(btnDel);
        tr.appendChild(tdAcoes);

        corpo.appendChild(tr);

        const trEdit = document.createElement('tr');
        trEdit.className = 'prof-edit-row';
        trEdit.id = `editProfRow-${p.id}`;
        trEdit.style.display = 'none';
        const tdEdit = document.createElement('td');
        tdEdit.colSpan = 4;
        const disciplinasGlobais = Array.from(new Set(professores.flatMap(px => px.disciplinas))).filter(Boolean);
        const optionsDisc = disciplinasGlobais.map(d => `<option value="${d}" ${p.disciplinas.includes(d) ? 'selected' : ''}>${d}</option>`).join('');
        tdEdit.innerHTML = `
            <div class="inline-editor">
                <div class="form-group">
                    <label>Nome</label>
                    <input type="text" id="editProfNome-${p.id}" value="${p.nome}">
                </div>
                <div class="form-group">
                    <label>Disciplinas</label>
                    <select id="editProfDiscs-${p.id}" multiple>${optionsDisc}</select>
                </div>
                <div class="form-group">
                    <label>Nova disciplina</label>
                    <input type="text" id="editProfDiscNova-${p.id}" placeholder="Adicionar nova">
                </div>
                <div class="form-group">
                    <label>Cor</label>
                    <input type="color" id="editProfCor-${p.id}" value="${p.cor}">
                </div>
                <div class="form-group">
                    <label>Dias disponíveis</label>
                    <div id="editProfDias-${p.id}" class="checks-grid">
                        ${['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA','SABADO'].map(d => 
                            `<label><input type="checkbox" value="${d}" ${ (p.diasDisponiveis||[]).includes(d) ? 'checked' : ''}> ${textoDia(d)}</label>`
                        ).join('')}
                    </div>
                </div>
                <div class="form-group">
                    <label>Turnos disponíveis</label>
                    <div id="editProfTurnos-${p.id}" class="checks-grid">
                        ${['MANHA','TARDE','NOITE'].map(t => 
                            `<label><input type="checkbox" value="${t}" ${ (p.turnosDisponiveis||[]).includes(t) ? 'checked' : ''}> ${textoTurno(t)}</label>`
                        ).join('')}
                    </div>
                </div>
                <div class="form-group">
                    <label>Tempos disponíveis</label>
                    <div id="editProfTempos-${p.id}" class="checks-grid">
                        ${getTempos('MANHA').filter(t=>!t.intervalo).map(t => {
                            const checked = (p.temposDisponiveis && p.temposDisponiveis.length ? p.temposDisponiveis : getTempos('MANHA').filter(x=>!x.intervalo).map(x=>x.id)).includes(t.id);
                            const lbl = etiquetaTempoAbrev(t, 'MANHA');
                            return `<label><input type="checkbox" value="${t.id}" ${checked ? 'checked' : ''}> ${lbl}</label>`;
                        }).join('')}
                    </div>
                </div>
                <div class="form-group">
                    <button type="button" class="btn-primary small" onclick="salvarEdicaoProfessorInline('${p.id}')">Salvar</button>
                    <button type="button" class="btn-secondary small" onclick="toggleEdicaoProfessorInline('${p.id}', true)">Cancelar</button>
                </div>
            </div>
        `;
        trEdit.appendChild(tdEdit);
        corpo.appendChild(trEdit);

        // selects
        if (selProf) {
            const opt1 = document.createElement('option');
            opt1.value = p.id;
            opt1.textContent = p.nome;
            selProf.appendChild(opt1);
        }

        if (selProfRel) {
            const opt2 = document.createElement('option');
            opt2.value = p.id;
            opt2.textContent = p.nome;
            selProfRel.appendChild(opt2);
        }

        if (selProfModal) {
            const opt3 = document.createElement('option');
            opt3.value = p.id;
            opt3.textContent = p.nome;
            selProfModal.appendChild(opt3);
        }
    });
}

function salvarProfessor(e) {
    e.preventDefault();
    const id = document.getElementById('professorId').value;
    const nome = document.getElementById('nomeProfessor').value.trim();
    const disciplinasStr = document.getElementById('disciplinasProfessor').value;
    const cor = document.getElementById('corProfessor').value || '#3498db';
    const dispDiaTurno = coletarDisponibilidade('dispProfessor');
    const temposDiaTurno = coletarTemposDisponibilidade('dispProfessor');
    const diasDisponiveis = Object.keys(dispDiaTurno).filter(d => (dispDiaTurno[d] || []).length > 0);
    const turnosDisponiveis = Array.from(new Set([].concat(...Object.values(dispDiaTurno))));
    const temposDisponiveis = (() => {
        const s = new Set();
        Object.values(temposDiaTurno || {}).forEach(turnoMap => {
            Object.values(turnoMap || {}).forEach(lista => {
                (lista || []).forEach(id => s.add(id));
            });
        });
        return Array.from(s);
    })();

    if (!nome) {
        mostrarToast('Informe o nome do professor.', 'warning');
        return;
    }

    const corLower = cor.toLowerCase();
    const conflitoCor = professores.find(p => p.cor && p.cor.toLowerCase() === corLower && p.id !== id);
    if (conflitoCor) {
        const msgCor =
            `A cor escolhida já está sendo usada por ${conflitoCor.nome}.\n` +
            'Deseja continuar usando a mesma cor para este professor?';
        const continuar = confirm(msgCor);
        if (!continuar) {
            return;
        }
    }

    const disciplinas = disciplinasStr
        ? disciplinasStr.split(',').map(s => s.trim()).filter(Boolean)
        : [];

    if (id) {
        const p = professores.find(x => x.id === id);
        if (p) {
            const novoProf = {
                ...p,
                nome,
                disciplinas,
                cor,
                diasDisponiveis,
                turnosDisponiveis,
                disponibilidadePorDia: dispDiaTurno,
                temposPorDiaTurno: temposDiaTurno,
                temposDisponiveis: temposDisponiveis.length
                    ? temposDisponiveis
                    : getTempos('MANHA').filter(t => !t.intervalo).map(t => t.id)
            };

            const aulasProfExistentes = aulas.filter(a => a.professorId === id);
            const aulasIndisponiveis = aulasProfExistentes.filter(a =>
                !professorDisponivelNoHorario(novoProf, a.turno, a.dia, a.tempoId)
            );

            if (aulasIndisponiveis.length > 0) {
                const detalhes = aulasIndisponiveis.slice(0, 5).map(a => {
                    const turma = turmas.find(t => t.id === a.turmaId);
                    const tempo = getTempos(a.turno).find(t => t.id === a.tempoId);
                    const descTempo = tempo ? tempo.etiqueta : `Tempo ${a.tempoId}`;
                    const descTurma = turma ? ` - Turma ${turma.nome}` : '';
                    return `${textoDia(a.dia)} / ${textoTurno(a.turno)} / ${descTempo}${descTurma}`;
                });
                const msgExtra = aulasIndisponiveis.length > detalhes.length
                    ? `\n... e mais ${aulasIndisponiveis.length - detalhes.length} aula(s).`
                    : '';
                const msg =
                    'Atenção: com a nova disponibilidade, este professor ficará indisponível ' +
                    `para ${aulasIndisponiveis.length} aula(s) já lançada(s):\n\n` +
                    detalhes.join('\n') +
                    msgExtra +
                    '\n\nDeseja salvar mesmo assim?';
                if (!confirm(msg)) {
                    return;
                }
            }

            p.nome = novoProf.nome;
            p.disciplinas = novoProf.disciplinas;
            p.cor = novoProf.cor;
            p.diasDisponiveis = novoProf.diasDisponiveis;
            p.turnosDisponiveis = novoProf.turnosDisponiveis;
            p.disponibilidadePorDia = novoProf.disponibilidadePorDia;
            p.temposPorDiaTurno = novoProf.temposPorDiaTurno;
            p.temposDisponiveis = novoProf.temposDisponiveis;
            mostrarToast('Professor atualizado com sucesso!');
        }
    } else {
        professores.push({
            id: gerarId(),
            nome,
            disciplinas,
            cor,
            diasDisponiveis,
            turnosDisponiveis,
            disponibilidadePorDia: dispDiaTurno,
            temposPorDiaTurno: temposDiaTurno,
            temposDisponiveis: temposDisponiveis.length ? temposDisponiveis : getTempos('MANHA').filter(t => !t.intervalo).map(t => t.id)
        });
        mostrarToast('Professor cadastrado com sucesso!');
    }

    salvarLocal();
    atualizarListaProfessores();
    document.getElementById('formProfessor').reset();
    document.getElementById('professorId').value = '';
}

function editarProfessor(id) {
    const p = professores.find(x => x.id === id);
    if (!p) return;
    document.getElementById('professorId').value = p.id;
    document.getElementById('nomeProfessor').value = p.nome;
    document.getElementById('disciplinasProfessor').value = p.disciplinas.join(', ');
    document.getElementById('corProfessor').value = p.cor;
    renderDisponibilidade('dispProfessor', p.disponibilidadePorDia || null, (p.diasDisponiveis || []), (p.turnosDisponiveis || []), p.temposPorDiaTurno || null);
}

function excluirProfessor(id) {
    if (!confirm('Excluir este professor? Todas as aulas dele serão removidas.')) return;
    const aulasProfessor = aulas.filter(a => a.professorId === id);
    if (aulasProfessor.length > 0) {
        if (!confirm(`Este professor tem ${aulasProfessor.length} aula(s) agendada(s). Deseja excluir mesmo assim?`)) return;
    }
    
    professores = professores.filter(p => p.id !== id);
    aulas = aulas.filter(a => a.professorId !== id);
    salvarLocal();
    atualizarListaProfessores();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast('Professor excluído com sucesso!');
}

function toggleEdicaoProfessorInline(id, fechar = false) {
    const row = document.getElementById(`editProfRow-${id}`);
    if (!row) {
        editarProfessor(id);
        return;
    }
    if (fechar) {
        row.style.display = 'none';
        return;
    }
    row.style.display = row.style.display === 'none' ? 'table-row' : 'none';
}

function salvarEdicaoProfessorInline(id) {
    const p = professores.find(x => x.id === id);
    if (!p) return;
    const nome = document.getElementById(`editProfNome-${id}`).value.trim();
    const sel = document.getElementById(`editProfDiscs-${id}`);
    const nova = document.getElementById(`editProfDiscNova-${id}`).value.trim();
    const cor = document.getElementById(`editProfCor-${id}`).value || '#3498db';
    const dias = Array.from(document.querySelectorAll(`#editProfDias-${id} input:checked`)).map(i => i.value);
    const turnos = Array.from(document.querySelectorAll(`#editProfTurnos-${id} input:checked`)).map(i => i.value);
    const tempos = Array.from(document.querySelectorAll(`#editProfTempos-${id} input:checked`)).map(i => parseInt(i.value, 10));
    if (!nome) {
        mostrarToast('Informe o nome do professor.', 'warning');
        return;
    }
    const corLower = cor.toLowerCase();
    const conflitoCor = professores.find(px => px.cor && px.cor.toLowerCase() === corLower && px.id !== id);
    if (conflitoCor) {
        const msgCor =
            `A cor escolhida já está sendo usada por ${conflitoCor.nome}.\n` +
            'Deseja continuar usando a mesma cor para este professor?';
        const continuar = confirm(msgCor);
        if (!continuar) {
            return;
        }
    }
    const selecionadas = sel ? Array.from(sel.selectedOptions).map(o => o.value) : [];
    const disciplinas = [...new Set([...selecionadas, ...(nova ? [nova] : [])])];
    p.nome = nome;
    p.disciplinas = disciplinas;
    p.cor = cor;
    p.diasDisponiveis = dias;
    p.turnosDisponiveis = turnos;
    p.temposDisponiveis = tempos.length ? tempos : getTempos('MANHA').filter(t => !t.intervalo).map(t => t.id);
    salvarLocal();
    atualizarListaProfessores();
    mostrarToast('Professor atualizado com sucesso!');
}

// =======================
// CURSOS / TURMAS
// =======================

function preencherSelectCursos() {
    const sel = document.getElementById('cursoTurma');
    if (!sel) return;
    sel.innerHTML = '';
    cursos.forEach(curso => {
        const opt = document.createElement('option');
        opt.value = curso;
        opt.textContent = curso;
        sel.appendChild(opt);
    });
}

function preencherSelectCursoVisual() {
    const sel = document.getElementById('cursoVisual');
    if (!sel) return;
    sel.innerHTML = '<option value="">Todos os cursos</option>';
    cursos.forEach(curso => {
        const opt = document.createElement('option');
        opt.value = curso;
        opt.textContent = curso;
        sel.appendChild(opt);
    });
    if (cursoVisualGrade && cursos.includes(cursoVisualGrade)) {
        sel.value = cursoVisualGrade;
    }
}

function abrirModalCurso() {
    document.getElementById('formNovoCurso').reset();
    document.getElementById('modalCurso').style.display = 'flex';
}

function fecharModalCurso() {
    document.getElementById('modalCurso').style.display = 'none';
}

function salvarNovoCurso(e) {
    e.preventDefault();
    const nome = document.getElementById('nomeNovoCurso').value.trim();
    if (!nome) {
        mostrarToast('Informe o nome do curso/área.', 'warning');
        return;
    }
    if (!cursos.includes(nome)) {
        cursos.push(nome);
        salvarLocal();
        preencherSelectCursos();
        preencherSelectCursoRelatorio();
        document.getElementById('cursoTurma').value = nome;
        mostrarToast('Curso cadastrado com sucesso!');
    }
    fecharModalCurso();
}

function excluirCursoSelecionado() {
    const sel = document.getElementById('cursoTurma');
    const cursoParaExcluir = sel.value;

    if (!cursoParaExcluir) {
        mostrarToast('Selecione um curso para excluir.', 'warning');
        return;
    }

    // Verifica se existem turmas vinculadas
    const turmasVinculadas = turmas.filter(t => t.curso === cursoParaExcluir);
    if (turmasVinculadas.length > 0) {
        mostrarToast(`Não é possível excluir: existem ${turmasVinculadas.length} turmas vinculadas a este curso.`, 'error');
        return;
    }

    if (!confirm(`Deseja realmente excluir o curso "${cursoParaExcluir}"?`)) {
        return;
    }

    // Remove do array
    cursos = cursos.filter(c => c !== cursoParaExcluir);
    
    // Salva e atualiza
    salvarLocal();
    preencherSelectCursos();
    preencherSelectCursoRelatorio();
    mostrarToast('Curso excluído com sucesso!');
}

function atualizarListaTurmas() {
    const corpo = document.getElementById('listaTurmas');
    const selTurmaLimpeza = document.getElementById('turmaLimpeza');
    if (!corpo) return;
    corpo.innerHTML = '';
    if (selTurmaLimpeza) selTurmaLimpeza.innerHTML = '<option value="">Selecione a turma...</option>';

    turmas.forEach(t => {
        const tr = document.createElement('tr');

        const tdNome = document.createElement('td');
        tdNome.textContent = t.nome;
        tr.appendChild(tdNome);

        const tdCurso = document.createElement('td');
        tdCurso.textContent = t.curso;
        tr.appendChild(tdCurso);

        const tdDesc = document.createElement('td');
        tdDesc.textContent = t.descricao || '';
        tr.appendChild(tdDesc);

        const tdTurno = document.createElement('td');
        tdTurno.textContent = textoTurno(t.turno);
        tr.appendChild(tdTurno);

        const tdAcoes = document.createElement('td');
        tdAcoes.style.display = 'flex';
        tdAcoes.style.gap = '6px';

        const btnEd = document.createElement('button');
        btnEd.className = 'btn-secondary small';
        btnEd.textContent = 'Editar';
        btnEd.onclick = () => toggleEdicaoTurmaInline(t.id);

        const btnDel = document.createElement('button');
        btnDel.className = 'btn-danger small';
        btnDel.textContent = 'Excluir';
        btnDel.onclick = () => excluirTurma(t.id);

        tdAcoes.appendChild(btnEd);
        tdAcoes.appendChild(btnDel);
        tr.appendChild(tdAcoes);

        corpo.appendChild(tr);
        
        if (selTurmaLimpeza) {
            const opt = document.createElement('option');
            opt.value = t.id;
            opt.textContent = `${t.nome} (${textoTurno(t.turno)})`;
            selTurmaLimpeza.appendChild(opt);
        }

        const trEdit = document.createElement('tr');
        trEdit.className = 'turma-edit-row';
        trEdit.id = `editRow-${t.id}`;
        trEdit.style.display = 'none';
        const tdEdit = document.createElement('td');
        tdEdit.colSpan = 5;
        const turnosOpts = TURNOS.map(tt => `<option value="${tt}" ${tt === t.turno ? 'selected' : ''}>${textoTurno(tt)}</option>`).join('');
        const cursosOpts = cursos.map(c => `<option value="${c}" ${c === t.curso ? 'selected' : ''}>${c}</option>`).join('');
        tdEdit.innerHTML = `
            <div class="inline-editor">
                <div class="form-group">
                    <label>Nome</label>
                    <input type="text" id="editNome-${t.id}" value="${t.nome}">
                </div>
                <div class="form-group">
                    <label>Curso / Área</label>
                    <select id="editCurso-${t.id}">${cursosOpts}</select>
                </div>
                <div class="form-group">
                    <label>Turno</label>
                    <select id="editTurno-${t.id}">${turnosOpts}</select>
                </div>
                <div class="form-group">
                    <label>Descrição</label>
                    <input type="text" id="editDesc-${t.id}" value="${t.descricao || ''}">
                </div>
                <div class="form-group">
                    <button type="button" class="btn-primary small" onclick="salvarEdicaoTurmaInline('${t.id}')">Salvar</button>
                    <button type="button" class="btn-secondary small" onclick="toggleEdicaoTurmaInline('${t.id}', true)">Cancelar</button>
                </div>
            </div>
        `;
        trEdit.appendChild(tdEdit);
        corpo.appendChild(trEdit);
    });

    atualizarChipsTurmasHorario();
    montarGradeTurno(turnoAtualGrade);
    preencherSelectCursoRelatorio();
}

function salvarTurma(e) {
    e.preventDefault();
    const id = document.getElementById('turmaId').value;
    const nome = document.getElementById('nomeTurma').value.trim();
    const curso = document.getElementById('cursoTurma').value;
    const turno = document.getElementById('turnoTurma').value;
    const descricao = document.getElementById('descricaoTurma').value.trim();

    if (!nome || !curso || !turno) {
        mostrarToast('Preencha nome, curso e turno.', 'warning');
        return;
    }

    if (id) {
        const t = turmas.find(x => x.id === id);
        if (t) {
            t.nome = nome;
            t.curso = curso;
            t.turno = turno;
            t.descricao = descricao;
            mostrarToast('Turma atualizada com sucesso!');
        }
    } else {
        turmas.push({
            id: gerarId(),
            nome,
            curso,
            turno,
            descricao
        });
        mostrarToast('Turma cadastrada com sucesso!');
    }

    salvarLocal();
    atualizarListaTurmas();
    document.getElementById('formTurma').reset();
    document.getElementById('turmaId').value = '';
    document.getElementById('turnoTurma').value = '';
    document.querySelectorAll('#chips-turno-turma .chip').forEach(c => c.classList.remove('selecionado'));
}

function editarTurma(id) {
    const t = turmas.find(x => x.id === id);
    if (!t) return;
    document.getElementById('turmaId').value = t.id;
    document.getElementById('nomeTurma').value = t.nome;
    document.getElementById('cursoTurma').value = t.curso;
    document.getElementById('turnoTurma').value = t.turno;
    document.getElementById('descricaoTurma').value = t.descricao || '';

    document.querySelectorAll('#chips-turno-turma .chip').forEach(c => {
        c.classList.toggle('selecionado', c.dataset.value === t.turno);
    });
}

function excluirTurma(id) {
    if (!confirm('Excluir esta turma e todos os seus horários?')) return;
    const aulasTurma = aulas.filter(a => a.turmaId === id);
    if (aulasTurma.length > 0) {
        if (!confirm(`Esta turma tem ${aulasTurma.length} aula(s) agendada(s). Deseja excluir mesmo assim?`)) return;
    }
    
    turmas = turmas.filter(t => t.id !== id);
    aulas = aulas.filter(a => a.turmaId !== id);
    salvarLocal();
    atualizarListaTurmas();
    mostrarToast('Turma excluída com sucesso!');
}

function toggleEdicaoTurmaInline(id, fechar = false) {
    const row = document.getElementById(`editRow-${id}`);
    if (!row) {
        editarTurma(id);
        return;
    }
    if (fechar) {
        row.style.display = 'none';
        return;
    }
    row.style.display = row.style.display === 'none' ? 'table-row' : 'none';
}

function salvarEdicaoTurmaInline(id) {
    const t = turmas.find(x => x.id === id);
    if (!t) return;
    const nome = document.getElementById(`editNome-${id}`).value.trim();
    const curso = document.getElementById(`editCurso-${id}`).value;
    const turno = document.getElementById(`editTurno-${id}`).value;
    const descricao = document.getElementById(`editDesc-${id}`).value.trim();
    if (!nome || !curso || !turno) {
        mostrarToast('Preencha nome, curso e turno.', 'warning');
        return;
    }
    t.nome = nome;
    t.curso = curso;
    t.turno = turno;
    t.descricao = descricao;
    salvarLocal();
    atualizarListaTurmas();
    mostrarToast('Turma atualizada com sucesso!');
}

// =======================
// HORÁRIOS (CADASTRO RÁPIDO)
// =======================

function atualizarChipsTurmasHorario() {
    const container = document.getElementById('chips-turma');
    if (!container) return;
    const turno = document.getElementById('turnoHorario').value;
    container.innerHTML = '';

    if (!turno) {
        container.innerHTML = '<span class="hint">Selecione um turno primeiro.</span>';
        return;
    }

    const turmasTurno = turmas.filter(t => t.turno === turno);
    if (!turmasTurno.length) {
        container.innerHTML = '<span class="hint">Não há turmas cadastradas neste turno.</span>';
        return;
    }

    turmasTurno.forEach(t => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'chip';
        btn.dataset.value = t.id;
        btn.textContent = t.nome;
        btn.onclick = () => selecionarChip('chips-turma', 'turmaHorario', btn);
        container.appendChild(btn);
    });
}

function salvarHorario(e) {
    e.preventDefault();
    const turno = document.getElementById('turnoHorario').value;
    const dia = document.getElementById('diaSemana').value;
    const turmaId = document.getElementById('turmaHorario').value;
    const temposStr = document.getElementById('tempoHorario').value;
    const disciplina = document.getElementById('disciplinaHorario').value.trim();
    const professorId = document.getElementById('professorHorario').value || null;

    if (!turno || !dia || !turmaId || !temposStr || !disciplina) {
        mostrarToast('Preencha turno, dia, turma, pelo menos um tempo e disciplina.', 'warning');
        return;
    }

    const tempoIds = temposStr
        .split(',')
        .map(v => parseInt(v, 10))
        .filter(n => !isNaN(n));

    if (!tempoIds.length) {
        mostrarToast('Selecione pelo menos um tempo.', 'warning');
        return;
    }

    const temposObjs = tempoIds
        .map(id => getTempos(turno).find(t => t.id === id))
        .filter(Boolean);

    if (temposObjs.some(t => t.intervalo)) {
        mostrarToast('Não é possível lançar aula no intervalo.', 'warning');
        return;
    }

    // Verifica disponibilidade e conflitos de professor
    if (professorId) {
        const prof = professores.find(p => p.id === professorId);
        if (!prof) {
            mostrarToast('Professor não encontrado.', 'error');
            return;
        }

        const indisponiveis = [];
        tempoIds.forEach(tempoId => {
            if (!professorDisponivelNoHorario(prof, turno, dia, tempoId)) {
                const tempo = getTempos(turno).find(t => t.id === tempoId);
                indisponiveis.push(tempo ? tempo.etiqueta : tempoId);
            }
        });
        if (indisponiveis.length > 0) {
            if (!confirm(`O professor não está disponível em ${dia} / ${textoTurno(turno)} nos tempos: ${indisponiveis.join(', ')}. Deseja lançar mesmo assim?`)) {
                return;
            }
        }

        const conflitos = [];
        tempoIds.forEach(tempoId => {
            if (verificarConflitoProfessor(professorId, turno, dia, tempoId)) {
                const tempo = getTempos(turno).find(t => t.id === tempoId);
                conflitos.push(tempo ? tempo.etiqueta : tempoId);
            }
        });
        if (conflitos.length > 0) {
            if (!confirm(`O professor já tem aula nos seguintes horários: ${conflitos.join(', ')}. Deseja continuar mesmo assim?`)) {
                return;
            }
        }
    }

    // Verifica conflitos em qualquer um dos tempos (mesma turma)
    const conflitosTurma = [];
    tempoIds.forEach(tempoId => {
        const ja = aulas.find(a =>
            a.turno === turno &&
            a.dia === dia &&
            a.turmaId === turmaId &&
            a.tempoId === tempoId
        );
        if (ja) conflitosTurma.push(tempoId);
    });

    if (conflitosTurma.length && !confirm('Já existe aula em algum dos tempos selecionados. Deseja substituir?')) {
        return;
    }

    // Remove aulas antigas nesses tempos (para essa turma/dia/turno)
    aulas = aulas.filter(a =>
        !(a.turno === turno &&
          a.dia === dia &&
          a.turmaId === turmaId &&
          tempoIds.includes(a.tempoId))
    );

    // Cria uma aula para cada tempo selecionado
    tempoIds.forEach(tempoId => {
        aulas.push({
            id: gerarId(),
            turno,
            dia,
            turmaId,
            tempoId,
            disciplina,
            professorId
        });
    });

    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast(`Aula(s) salva(s) com sucesso! (${tempoIds.length} tempo(s))`);

    document.getElementById('formHorario').reset();
    document.getElementById('horarioId').value = '';
    document.getElementById('turnoHorario').value = '';
    document.getElementById('diaSemana').value = '';
    document.getElementById('turmaHorario').value = '';
    document.getElementById('tempoHorario').value = '';
    document.querySelectorAll('#chips-turno-horario .chip, #chips-dia .chip, #chips-turma .chip, #chips-tempo .chip')
        .forEach(c => c.classList.remove('selecionado'));
}

// pega aula específica
function obterAula(turno, dia, turmaId, tempoId) {
    return aulas.find(a =>
        a.turno === turno &&
        a.dia === dia &&
        a.turmaId === turmaId &&
        a.tempoId === tempoId
    );
}

// =======================
// EDITOR VISUAL (GRADE) - ATUALIZADO
// =======================

function montarGradeTurno(turno) {
    const tabela = document.getElementById('tabelaGrade');
    if (!tabela) return;
    tabela.innerHTML = '';

    let turmasTurno = turmas.filter(t => t.turno === turno);
    if (modoVisualGrade === 'CURSO' && cursoVisualGrade) {
        turmasTurno = turmasTurno.filter(t => t.curso === cursoVisualGrade);
    }
    if (!turmasTurno.length) {
        const msg = modoVisualGrade === 'CURSO' && cursoVisualGrade
            ? 'Não há turmas deste curso neste turno.'
            : 'Não há turmas cadastradas neste turno.';
        tabela.innerHTML = `<tr><td class="hint" style="padding:8px;">${msg}</td></tr>`;
        return;
    }

    // Cabeçalho: DIA | HORÁRIO | TURMAS...
    const thead = document.createElement('thead');
    const trHead = document.createElement('tr');

    const thDia = document.createElement('th');
    thDia.textContent = 'DIA';
    thDia.className = 'dia-col';
    trHead.appendChild(thDia);

    const thHor = document.createElement('th');
    thHor.textContent = 'HORÁRIO';
    thHor.className = 'horario-col';
    trHead.appendChild(thHor);

    turmasTurno.forEach(turma => {
        const th = document.createElement('th');
        th.className = 'turma-col';

        const span = document.createElement('span');
        span.innerHTML =
            `<div class="turma-header-nome">${turma.nome}</div>` +
            (turma.descricao ? `<div class="turma-header-desc">${turma.descricao}</div>` : '');
        th.appendChild(span);

        trHead.appendChild(th);
    });

    thead.appendChild(trHead);
    tabela.appendChild(thead);

    const tbody = document.createElement('tbody');

    const chkOmitirSabado = document.getElementById('chkOmitirSabadoGrade');
    const omitirSabadoGrade = !!(chkOmitirSabado && chkOmitirSabado.checked);

    DIAS_SEMANA.forEach(dia => {
        if (omitirSabadoGrade && dia === 'SABADO') return;
        let primeiraLinhaDia = true;

        const temposTurno = getTempos(turno);
        temposTurno.forEach((tempo, idxTempo) => {
            const tr = document.createElement('tr');

            if (primeiraLinhaDia) {
                const tdDia = document.createElement('td');
                tdDia.className = 'celula-dia';
                tdDia.rowSpan = temposTurno.length;
                tdDia.textContent = textoDia(dia).toUpperCase();
                tr.appendChild(tdDia);
                primeiraLinhaDia = false;
            }

            const tdHor = document.createElement('td');
            tdHor.className = 'celula-horario';
            tdHor.innerHTML = tempo.inicio + '<br>' + tempo.fim;
            tr.appendChild(tdHor);

            if (tempo.intervalo) {
                const td = document.createElement('td');
                td.className = 'celula-intervalo-texto';
                td.colSpan = turmasTurno.length;
                td.textContent = 'INTERVALO';
                tr.appendChild(td);
            } else {
                turmasTurno.forEach(turma => {
                    const td = document.createElement('td');
                    
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (aula) {
                        td.className = 'celula-aula';

                        const prof = professores.find(p => p.id === aula.professorId);
                        const cor = prof ? prof.cor : '#3498db';

                        const bloco = document.createElement('div');
                        bloco.className = 'bloco-aula';
                        bloco.style.backgroundColor = cor;
                        bloco.style.color = '#fff';
                        bloco.style.padding = '4px 6px';
                        bloco.style.borderRadius = '4px';
                        bloco.style.fontSize = '0.7rem';
                        bloco.style.minHeight = '40px';
                        bloco.style.display = 'flex';
                        bloco.style.flexDirection = 'column';
                        bloco.style.justifyContent = 'center';

                        bloco.innerHTML =
                            '<strong>' + abreviar(aula.disciplina, 14) + '</strong>' +
                            '<small>' + (prof ? abreviar(prof.nome, 14) : 'Sem professor') + '</small>';

                        bloco.onclick = (ev) => {
                            ev.stopPropagation();
                            abrirEditorCelula(aula);
                        };

                        bloco.draggable = true;
                        bloco.addEventListener('dragstart', (ev) => {
                            const payload = JSON.stringify({ aulaId: aula.id });
                            ev.dataTransfer.setData('text/plain', payload);
                            ev.dataTransfer.setData('application/json', payload);
                            ev.dataTransfer.effectAllowed = 'copy';
                        });

                        td.appendChild(bloco);
                    } else {
                        td.className = 'celula-vazia';
                        td.textContent = 'Clique para adicionar aula';
                        
                        // Adiciona ícone visual para indicar que é clicável
                        const icon = document.createElement('div');
                        icon.style.fontSize = '0.9rem';
                        icon.style.opacity = '0.5';
                        icon.style.marginTop = '5px';
                        icon.innerHTML = '➕';
                        td.appendChild(icon);
                        
                        td.onclick = () => {
                            abrirEditorCelulaVazia(turno, dia, turma.id, tempo.id);
                        };
                    }

                    // Também permite editar clicando na célula em si (não só no bloco)
                    td.onclick = (ev) => {
                        if (ev.target === td || ev.target === td.firstChild) {
                            if (aula) {
                                abrirEditorCelula(aula);
                            } else {
                                abrirEditorCelulaVazia(turno, dia, turma.id, tempo.id);
                            }
                        }
                    };

                    td.dataset.turno = turno;
                    td.dataset.dia = dia;
                    td.dataset.turmaId = turma.id;
                    td.dataset.tempoId = tempo.id;
                    td.addEventListener('dragover', (ev) => {
                        if (!aula && !tempo.intervalo) {
                            ev.preventDefault();
                            td.classList.add('celula-drop-alvo');
                        }
                    });
                    td.addEventListener('dragenter', (ev) => {
                        if (!aula && !tempo.intervalo) {
                            ev.preventDefault();
                            td.classList.add('celula-drop-alvo');
                        }
                    });
                    td.addEventListener('dragleave', () => {
                        td.classList.remove('celula-drop-alvo');
                    });
                    td.addEventListener('drop', (ev) => {
                        ev.stopPropagation();
                        td.classList.remove('celula-drop-alvo');
                        if (aula || tempo.intervalo) return;
                        ev.preventDefault();
                        const data = ev.dataTransfer.getData('text/plain') || ev.dataTransfer.getData('application/json');
                        if (!data) return;
                        let payload;
                        try { payload = JSON.parse(data); } catch (e) { return; }
                        const src = aulas.find(a => a.id === payload.aulaId);
                        if (!src) return;
                        const profId = src.professorId || null;
                        if (profId) {
                            const prof = professores.find(p => p.id === profId);
                            if (!prof) {
                                mostrarToast('Professor não encontrado.', 'error');
                                return;
                            }
                            if (!professorDisponivelNoHorario(prof, turno, dia, tempo.id)) {
                                const continuarDisp = confirm('Este professor não está disponível neste dia/turno/tempo. Deseja continuar mesmo assim?');
                                if (!continuarDisp) return;
                            }
                            if (verificarConflitoProfessor(profId, turno, dia, tempo.id)) {
                                const continuar = confirm('Este professor já tem aula neste mesmo horário. Deseja continuar mesmo assim?');
                                if (!continuar) return;
                            }
                        }
                        const novaAula = {
                            id: gerarId(),
                            turno,
                            dia,
                            turmaId: turma.id,
                            tempoId: tempo.id,
                            disciplina: src.disciplina,
                            professorId: src.professorId || null
                        };
                        aulas.push(novaAula);
                        salvarLocal();
                        montarGradeTurno(turnoAtualGrade);
                        mostrarToast('Aula copiada para a célula selecionada.');
                    });

                    tr.appendChild(td);
                });
            }

            if (idxTempo === temposTurno.length - 1) {
                tr.classList.add('linha-dia');
            }

            tbody.appendChild(tr);
        });
    });

    tabela.appendChild(tbody);
    
    // Eventos são vinculados dinamicamente; não usar cache de innerHTML
}

// editor ao clicar em aula (abre MODAL)
function abrirEditorCelula(aula) {
    aulaEmEdicao = aula;

    const turma = turmas.find(t => t.id === aula.turmaId);
    const tempo = getTempos(aula.turno).find(t => t.id === aula.tempoId);

    document.getElementById('modalAulaId').value = aula.id;
    document.getElementById('modalTurma').value = turma ? turma.nome : '';
    document.getElementById('modalDia').value = textoDia(aula.dia);
    document.getElementById('modalHorario').value = tempo ? `${tempo.inicio} - ${tempo.fim}` : '';
    document.getElementById('modalDisciplina').value = aula.disciplina;
    document.getElementById('modalProfessor').value = aula.professorId || '';

    document.getElementById('modalAula').style.display = 'flex';
}

// editor em célula vazia: abre MODAL para criar nova aula
function abrirEditorCelulaVazia(turno, dia, turmaId, tempoId) {
    // Cria um objeto temporário para a nova aula
    const tempAula = {
        id: 'novo', // Identificador especial
        turno,
        dia,
        turmaId,
        tempoId,
        disciplina: '',
        professorId: ''
    };
    
    // Abre o mesmo modal de edição
    abrirEditorCelula(tempAula);
}

// Fechar modal de aula
function fecharModalAula() {
    document.getElementById('modalAula').style.display = 'none';
    aulaEmEdicao = null;
}

// Salvar edição do modal
function salvarEdicaoAula(e) {
    e.preventDefault();
    if (!aulaEmEdicao) {
        fecharModalAula();
        return;
    }

    const disc = document.getElementById('modalDisciplina').value.trim();
    const profId = document.getElementById('modalProfessor').value || null;

    if (!disc) {
        mostrarToast('Informe a disciplina.', 'warning');
        return;
    }

    // Verifica disponibilidade e conflito de professor
    if (profId) {
        const prof = professores.find(p => p.id === profId);
        if (!prof) {
            mostrarToast('Professor não encontrado.', 'error');
            return;
        }

        if (!professorDisponivelNoHorario(prof, aulaEmEdicao.turno, aulaEmEdicao.dia, aulaEmEdicao.tempoId)) {
            if (!confirm('Este professor não está disponível neste dia/turno/tempo. Deseja continuar mesmo assim?')) {
                return;
            }
        }

        const idParaIgnorar = aulaEmEdicao.id === 'novo' ? null : aulaEmEdicao.id;
        if (verificarConflitoProfessor(profId, aulaEmEdicao.turno, aulaEmEdicao.dia, aulaEmEdicao.tempoId, idParaIgnorar)) {
            if (!confirm('Este professor já tem aula neste mesmo horário. Deseja continuar mesmo assim?')) {
                return;
            }
        }
    }

    if (aulaEmEdicao.id === 'novo') {
        // CRIAR NOVA AULA
        const novaAula = {
            id: gerarId(),
            turno: aulaEmEdicao.turno,
            dia: aulaEmEdicao.dia,
            turmaId: aulaEmEdicao.turmaId,
            tempoId: aulaEmEdicao.tempoId,
            disciplina: disc,
            professorId: profId
        };
        aulas.push(novaAula);
        mostrarToast('Aula criada com sucesso!');
    } else {
        // EDITAR AULA EXISTENTE
        const aula = aulas.find(a => a.id === aulaEmEdicao.id);
        if (aula) {
            aula.disciplina = disc;
            aula.professorId = profId;
            mostrarToast('Aula atualizada com sucesso!');
        }
    }

    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    fecharModalAula();
}

// Excluir aula a partir do modal
function excluirAulaModal() {
    if (!aulaEmEdicao) {
        fecharModalAula();
        return;
    }
    if (!confirm('Remover esta aula?')) return;

    aulas = aulas.filter(a => a.id !== aulaEmEdicao.id);
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    fecharModalAula();
    mostrarToast('Aula removida com sucesso!');
}

function limparTodasAulas() {
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        mostrarToast('Apenas administradores podem apagar todos os horários.', 'warning');
        return;
    }
    if (!aulas.length) {
        mostrarToast('Não há horários para apagar.', 'warning');
        return;
    }
    if (!confirm('Apagar TODAS as aulas de todos os turnos, turmas e professores? Esta ação não pode ser desfeita.')) {
        return;
    }
    aulas = [];
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast('Todos os horários foram apagados.');
}

function limparAulasTurma() {
    const sel = document.getElementById('turmaLimpeza');
    if (!sel) return;
    const turmaId = sel.value;
    if (!turmaId) {
        mostrarToast('Selecione uma turma para apagar os horários.', 'warning');
        return;
    }
    const aulasTurma = aulas.filter(a => a.turmaId === turmaId);
    if (!aulasTurma.length) {
        mostrarToast('Esta turma não possui aulas lançadas.', 'warning');
        return;
    }
    const turma = turmas.find(t => t.id === turmaId);
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        if (turma && usuarioLogado.turno && turma.turno !== usuarioLogado.turno) {
            mostrarToast('Você só pode apagar horários das turmas do seu turno.', 'warning');
            return;
        }
    }
    const nomeTurma = turma ? turma.nome : '';
    const labelTurno = turma ? textoTurno(turma.turno) : '';
    if (!confirm(`Remover ${aulasTurma.length} aula(s) da turma ${nomeTurma} (${labelTurno})?`)) {
        return;
    }
    aulas = aulas.filter(a => a.turmaId !== turmaId);
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast('Horários da turma apagados com sucesso.');
}

function limparAulasDiaSemana() {
    const selDia = document.getElementById('diaLimpeza');
    const selTurno = document.getElementById('turnoLimpeza');
    if (!selDia || !selTurno) return;
    const dia = selDia.value;
    const turno = selTurno.value;
    if (!dia) {
        mostrarToast('Selecione um dia da semana.', 'warning');
        return;
    }
    const aulasDia = aulas.filter(a =>
        a.dia === dia &&
        (turno === 'TODOS' || a.turno === turno)
    );
    if (!aulasDia.length) {
        mostrarToast('Não há aulas neste dia/turno para apagar.', 'warning');
        return;
    }
    const labelDia = textoDia(dia);
    const labelTurno = turno === 'TODOS' ? 'todos os turnos' : textoTurno(turno);
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        if (turno === 'TODOS' || (usuarioLogado.turno && turno !== 'TODOS' && turno !== usuarioLogado.turno)) {
            mostrarToast('Você só pode apagar horários do seu turno.', 'warning');
            return;
        }
    }
    if (!confirm(`Remover ${aulasDia.length} aula(s) de ${labelDia} (${labelTurno})?`)) {
        return;
    }
    aulas = aulas.filter(a =>
        !(a.dia === dia && (turno === 'TODOS' || a.turno === turno))
    );
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast('Horários do dia/turno apagados com sucesso.');
}

// =======================
// AJUSTE INTELIGENTE
// =======================

function ajusteInteligente() {
    if (!confirm('O ajuste inteligente vai tentar compactar os horários (evitar buracos) neste turno. Continuar?')) {
        return;
    }

    const turno = turnoAtualGrade;
    const turmasTurno = turmas.filter(t => t.turno === turno);
    if (!turmasTurno.length) {
        mostrarToast('Não há turmas neste turno.', 'warning');
        return;
    }

    // salva estado para desfazer
    pilhaUndo.push(JSON.stringify(aulas));

    // só considera tempos não intervalos
    const temposValidos = getTempos(turnoAtualGrade).filter(t => !t.intervalo);

    turmasTurno.forEach(turma => {
        DIAS_SEMANA.forEach(dia => {
            // pega as aulas do dia/turno/turma
            const aulasDia = temposValidos.map(tempo =>
                obterAula(turno, dia, turma.id, tempo.id) || null
            );

            // compacta (remove nulls)
            const compactadas = aulasDia.filter(a => a !== null);

            // remove todas as aulas antigas desse dia/turma/turno
            aulas = aulas.filter(a =>
                !(a.turno === turno && a.dia === dia && a.turmaId === turma.id)
            );

            // reinsere nas primeiras posições de tempo, respeitando disponibilidade e conflitos de professor
            compactadas.forEach((aula, idx) => {
                const novoTempo = temposValidos[idx];
                if (novoTempo) {
                    let podeMover = true;
                    if (aula.professorId) {
                        const prof = professores.find(p => p.id === aula.professorId);
                        if (prof) {
                            if (!professorDisponivelNoHorario(prof, turno, dia, novoTempo.id)) {
                                podeMover = false;
                            } else {
                                const conflitoGlobal = aulas.some(a =>
                                    a.professorId === aula.professorId &&
                                    a.turno === turno &&
                                    a.dia === dia &&
                                    a.tempoId === novoTempo.id
                                );
                                if (conflitoGlobal) {
                                    podeMover = false;
                                }
                            }
                        }
                    }
                    if (podeMover) {
                        aula.tempoId = novoTempo.id;
                    }
                }
                aulas.push(aula);
            });
        });
    });

    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast('Ajuste inteligente aplicado! Use "Desfazer" se necessário.');

    const aulasTurno = aulas.filter(a => a.turno === turno && a.professorId);
    const mapaConflitos = {};
    aulasTurno.forEach(a => {
        const key = `${a.professorId}_${a.dia}_${a.tempoId}`;
        if (!mapaConflitos[key]) mapaConflitos[key] = [];
        mapaConflitos[key].push(a);
    });
    const conflitosExistentes = [];
    Object.keys(mapaConflitos).forEach(key => {
        const grupo = mapaConflitos[key];
        if (!grupo || grupo.length < 2) return;
        const turmasIds = Array.from(new Set(grupo.map(g => g.turmaId).filter(Boolean)));
        if (turmasIds.length < 2) return;
        const [profId, dia, tempoIdStr] = key.split('_');
        const tempoId = parseInt(tempoIdStr, 10);
        const prof = professores.find(p => p.id === profId);
        const tempo = getTempos(turno).find(t => t.id === tempoId);
        const nomesTurmas = turmasIds
            .map(id => {
                const t = turmas.find(tx => tx.id === id);
                return t ? t.nome : '';
            })
            .filter(Boolean);
        conflitosExistentes.push({
            profNome: prof ? prof.nome : profId,
            dia,
            tempoEtiqueta: tempo ? tempo.etiqueta : String(tempoId),
            turmas: nomesTurmas
        });
    });
    if (conflitosExistentes.length > 0) {
        let msg = 'O ajuste inteligente foi aplicado.\n\n';
        msg += 'Foram encontrados conflitos de professor já existentes (autorizados anteriormente):\n\n';
        conflitosExistentes.forEach(c => {
            msg += `- ${c.profNome}: ${textoDia(c.dia)} / ${textoTurno(turno)} / ${c.tempoEtiqueta} (turmas: ${c.turmas.join(', ')})\n`;
        });
        msg += '\nEsses conflitos foram mantidos porque já tinham sido lançados com confirmação do usuário.\n\n';
        msg += 'O ajuste inteligente apenas compacta os horários dentro do turno selecionado, evitando criar novos conflitos ou colocar o professor em tempos em que está indisponível.';
        alert(msg);
    }
}

function desfazerAjuste() {
    if (!pilhaUndo.length) {
        mostrarToast('Não há ajuste para desfazer.', 'warning');
        return;
    }
    const estado = pilhaUndo.pop();
    try {
        aulas = JSON.parse(estado);
        salvarLocal();
        montarGradeTurno(turnoAtualGrade);
        mostrarToast('Ajuste desfeito com sucesso!');
    } catch (e) {
        console.error('Erro ao desfazer ajuste:', e);
        mostrarToast('Erro ao desfazer ajuste.', 'error');
    }
}

// =======================
// RELATÓRIOS - ATUALIZADO
// =======================

function preencherSelectCursoRelatorio() {
    const sel = document.getElementById('cursoRelatorio');
    if (!sel) return;
    sel.innerHTML = '<option value="">Selecione...</option>';
    cursos.forEach(curso => {
        const opt = document.createElement('option');
        opt.value = curso;
        opt.textContent = curso;
        sel.appendChild(opt);
    });
}

function preencherSelectProfessorRelatorio() {
    const sel = document.getElementById('professorRelatorio');
    if (!sel) return;
    sel.innerHTML = '<option value="">Selecione...</option>';
    professores.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = p.nome;
        sel.appendChild(opt);
    });
}

document.addEventListener('DOMContentLoaded', () => {
    renderDisponibilidade('dispProfessor', null);
});

// =======================
// CONFIG TEMPOS POR TURNO
// =======================
function selecionarTurnoTempos(btn) {
    const group = document.getElementById('chips-turno-tempos');
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        const valor = btn.getAttribute('data-value');
        if (usuarioLogado.turno && usuarioLogado.turno !== valor) {
            mostrarToast('Você só pode configurar tempos do seu turno.', 'warning');
            return;
        }
    }
    if (group) {
        group.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
        btn.classList.add('selecionado');
    }
    renderTemposTurno();
}

function turnoTemposSelecionado() {
    const sel = document.querySelector('#chips-turno-tempos .chip.selecionado');
    return sel ? sel.getAttribute('data-value') : 'MANHA';
}

function renderTemposTurno() {
    const tbody = document.getElementById('temposConfigBody');
    if (!tbody) return;
    const turno = turnoTemposSelecionado();
    const tempos = getTempos(turno);
    tbody.innerHTML = '';
    tempos.forEach((t, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${t.id}</td>
            <td><input type="text" value="${t.etiqueta}"></td>
            <td><input type="time" value="${t.inicio}"></td>
            <td><input type="time" value="${t.fim}"></td>
            <td class="centered"><input type="checkbox" ${t.intervalo ? 'checked' : ''}></td>
            <td><button class="btn-danger small" type="button">Excluir</button></td>
        `;
        const inputs = tr.querySelectorAll('input');
        inputs[0].addEventListener('change', (e) => { t.etiqueta = e.target.value.trim() || t.etiqueta; });
        inputs[1].addEventListener('change', (e) => { t.inicio = e.target.value || t.inicio; });
        inputs[2].addEventListener('change', (e) => { t.fim = e.target.value || t.fim; });
        inputs[3].addEventListener('change', (e) => { t.intervalo = !!e.target.checked; });
        tr.querySelector('button.btn-danger').addEventListener('click', () => {
            TEMPOS_POR_TURNO[turno] = TEMPOS_POR_TURNO[turno].filter(x => x.id !== t.id);
            // Reindexa IDs
            TEMPOS_POR_TURNO[turno].forEach((x, i) => x.id = i + 1);
            renderTemposTurno();
        });
        tbody.appendChild(tr);
    });
}

function adicionarTempoTurno() {
    const turno = turnoTemposSelecionado();
    const arr = TEMPOS_POR_TURNO[turno];
    const nextId = (arr[arr.length - 1]?.id || 0) + 1;
    arr.push({ id: nextId, inicio: '00:00', fim: '00:00', etiqueta: `${nextId}º Tempo`, intervalo: false });
    renderTemposTurno();
}

function salvarTemposTurno() {
    salvarLocal();
    mostrarToast('Tempos do turno salvos.');
    // Atualiza chips e grade se estiverem nesse turno
    atualizarChipsTempos();
    montarGradeTurno(turnoAtualGrade);
}

document.addEventListener('DOMContentLoaded', renderTemposTurno);

function selecionarTurnoTemposModal(btn) {
    const group = document.getElementById('chips-turno-tempos-modal');
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        const valor = btn.getAttribute('data-value');
        if (usuarioLogado.turno && usuarioLogado.turno !== valor) {
            mostrarToast('Você só pode configurar tempos do seu turno.', 'warning');
            return;
        }
    }
    if (group) {
        group.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
        btn.classList.add('selecionado');
    }
    renderTemposTurnoModal();
}

function turnoTemposSelecionadoModal() {
    const sel = document.querySelector('#chips-turno-tempos-modal .chip.selecionado');
    return sel ? sel.getAttribute('data-value') : 'MANHA';
}

function renderTemposTurnoModal() {
    const tbody = document.getElementById('temposConfigBodyModal');
    if (!tbody) return;
    const turno = turnoTemposSelecionadoModal();
    const tempos = getTempos(turno);
    tbody.innerHTML = '';
    tempos.forEach((t) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${t.id}</td>
            <td><input type="text" value="${t.etiqueta}"></td>
            <td><input type="time" value="${t.inicio}"></td>
            <td><input type="time" value="${t.fim}"></td>
            <td class="centered"><input type="checkbox" ${t.intervalo ? 'checked' : ''}></td>
            <td><button class="btn-danger small" type="button">Excluir</button></td>
        `;
        const inputs = tr.querySelectorAll('input');
        inputs[0].addEventListener('change', (e) => { t.etiqueta = e.target.value.trim() || t.etiqueta; });
        inputs[1].addEventListener('change', (e) => { t.inicio = e.target.value || t.inicio; });
        inputs[2].addEventListener('change', (e) => { t.fim = e.target.value || t.fim; });
        inputs[3].addEventListener('change', (e) => { t.intervalo = !!e.target.checked; });
        tr.querySelector('button.btn-danger').addEventListener('click', () => {
            TEMPOS_POR_TURNO[turno] = TEMPOS_POR_TURNO[turno].filter(x => x.id !== t.id);
            TEMPOS_POR_TURNO[turno].forEach((x, i) => x.id = i + 1);
            renderTemposTurnoModal();
        });
        tbody.appendChild(tr);
    });
}

function adicionarTempoTurnoModal() {
    const turno = turnoTemposSelecionadoModal();
    const arr = TEMPOS_POR_TURNO[turno];
    const nextId = (arr[arr.length - 1]?.id || 0) + 1;
    arr.push({ id: nextId, inicio: '00:00', fim: '00:00', etiqueta: `${nextId}º Tempo`, intervalo: false });
    renderTemposTurnoModal();
}

function salvarTemposTurnoModal() {
    salvarLocal();
    mostrarToast('Tempos do turno salvos.');
    atualizarChipsTempos();
    montarGradeTurno(turnoAtualGrade);
}

function abrirModalTempos() {
    const modal = document.getElementById('modalTempos');
    if (!modal) return;
    renderTemposTurnoModal();
    modal.style.display = 'flex';
}

function fecharModalTempos() {
    const modal = document.getElementById('modalTempos');
    if (modal) modal.style.display = 'none';
}

document.addEventListener('DOMContentLoaded', () => {
    const modal = document.getElementById('modalTempos');
    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) fecharModalTempos();
        });
    }
});

document.addEventListener('DOMContentLoaded', () => {
    const cont = document.getElementById('modalDispProfessor');
    if (cont) {
        cont.addEventListener('change', () => {
            const disp = coletarDisponibilidade('modalDispProfessor');
            renderTemposDisponibilidade('modalDispTemposProfessor', disp, null);
        });
    }
});
function gerarRelatorioTurnoPDF() {
    const { jsPDF } = window.jspdf;
    const turno = document.getElementById('turnoRelatorio').value;
    const layout = document.getElementById('layoutRelatorio').value;
    const chkOmitirSabado = document.getElementById('omitirSabadoRelatorios');
    const omitirSabado = !!(chkOmitirSabado && chkOmitirSabado.checked);

    const turmasTurno = turmas.filter(t => t.turno === turno);
    if (!turmasTurno.length) {
        mostrarToast('Não há turmas neste turno.', 'warning');
        return;
    }

    const colCount = 2 + turmasTurno.map(t => t.nome).length;
    const format = colCount > 12 ? 'a3' : 'a4';
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format });

    const titulo = `HORÁRIO - TURNO DA ${textoTurno(turno).toUpperCase()} — Gerado em: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`;
    doc.setFont(undefined, 'bold');
    doc.setFontSize(8);
    doc.text(titulo, doc.internal.pageSize.getWidth() / 2, 6, { align: 'center' });

    const head = [['DIA', 'HORÁRIO', ...turmasTurno.map(t => t.nome)]];
    const body = [];

    DIAS_SEMANA.forEach(dia => {
        if (omitirSabado && dia === 'SABADO') return;
        const temposTurno = getTempos(turno);
        temposTurno.forEach((tempo, idx) => {
            const row = [];
            
            // Coluna DIA com RowSpan para evitar repetição
            if (idx === 0) {
                row.push({ 
                    content: textoDia(dia), 
                    rowSpan: temposTurno.length,
                    styles: { valign: 'middle', halign: 'center' } 
                });
            }
            
            row.push(`${tempo.inicio} - ${tempo.fim}`);

            if (tempo.intervalo) {
                turmasTurno.forEach(() => row.push('INTERVALO'));
            } else {
                turmasTurno.forEach(turma => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) {
                        row.push('-');
                    } else {
                        const prof = professores.find(p => p.id === aula.professorId);
                        const textoCelula = `${aula.disciplina}\n${prof ? prof.nome : ''}`;
                        if (layout === 'colorido' && prof && prof.cor) {
                            const rgb = hexToRgb(prof.cor);
                            if (rgb) {
                                row.push({
                                    content: textoCelula,
                                    styles: { fillColor: [rgb.r, rgb.g, rgb.b], textColor: 255, valign: 'middle', halign: 'center' }
                                });
                            } else {
                                row.push(textoCelula);
                            }
                        } else {
                            row.push(textoCelula);
                        }
                    }
                });
            }

            body.push(row);
        });
    });

    const bodyFont = colCount > 12 ? 8 : 7;
    doc.autoTable({
        head,
        body,
        startY: 8,
        theme: 'grid',
        styles: {
            fontSize: bodyFont,
            cellPadding: 1,
            valign: 'middle',
            halign: 'center',
            overflow: 'linebreak',
            cellWidth: 'wrap',
            minCellHeight: 6
        },
        margin: { top: 6, right: 6, bottom: 6, left: 6 },
        headStyles: {
            fillColor: [52, 152, 219],
            textColor: 255,
            fontStyle: 'bold',
            fontSize: 8,
            cellPadding: 1,
            halign: 'center',
            valign: 'middle'
        },
        columnStyles: {
            0: { cellWidth: 10, cellPadding: 0 },
            1: { cellWidth: 20 }
        },
        didParseCell: function(data) {
            if (data.section === 'body' && data.column.index === 0) {
                data.cell.text = [];
            }
        },
        didDrawCell: function(data) {
            if (data.section === 'body' && data.column.index === 0) {
                const raw = data.row.raw[0];
                const text = raw && raw.content ? raw.content : data.cell.text[0];
                if (text) {
                    const prevSize = doc.getFontSize();
                    doc.setFontSize(7);
                    const xCenter = data.cell.x + data.cell.width / 2;
                    const x = xCenter + 5;
                    const yCenter = data.cell.y + data.cell.height / 2;
                    const ptToMm = 0.3528;
                    const offset = (doc.getFontSize() * ptToMm) * 0.5;
                    const y = yCenter + offset;
                    doc.setTextColor(0);
                    doc.text(text, x, y, { align: 'center', angle: 90 });
                    doc.setFontSize(prevSize);
                }
            }
        }
    });

    if (layout === 'colorido') {
        let y = doc.lastAutoTable.finalY + 8;
        doc.setFontSize(9);
        doc.setFont(undefined, 'bold');
        doc.text('LEGENDA DE PROFESSORES:', 14, y);
        y += 6;
        doc.setFont(undefined, 'normal');
        doc.setFontSize(8);
        
        let x = 14;
        professores.forEach((p, i) => {
            if (y > doc.internal.pageSize.height - 20) {
                doc.addPage();
                y = 20;
                x = 14;
            }
            
            const rgb = hexToRgb(p.cor || '#3498db') || { r: 52, g: 152, b: 219 };
            doc.setFillColor(rgb.r, rgb.g, rgb.b);
            doc.rect(x, y - 3, 4, 4, 'F');
            const discs = (p.disciplinas && p.disciplinas.length) ? p.disciplinas.join(', ') : '-';
            const legenda = ` ${p.nome} — ${discs}`;
            doc.text(legenda, x + 6, y);
            
            x += 45;
            if (x > doc.internal.pageSize.width - 50) {
                x = 14;
                y += 6;
            }
        });
    }

    doc.save(`horario_turno_${turno.toLowerCase()}_${new Date().toISOString().slice(0,10)}.pdf`);
    mostrarToast('PDF gerado com sucesso!');
}

function gerarRelatorioTurnoXLS() {
    if (typeof XLSX === 'undefined') {
        mostrarToast('Biblioteca XLSX não carregada. Verifique sua conexão.', 'error');
        return;
    }

    const turno = document.getElementById('turnoRelatorio').value;
    const layout = document.getElementById('layoutRelatorio').value;
    const turmasTurno = turmas.filter(t => t.turno === turno);
    if (!turmasTurno.length) {
        mostrarToast('Não há turmas neste turno.', 'warning');
        return;
    }

    const chkOmitirSabado = document.getElementById('omitirSabadoRelatorioCurso');
    const omitirSabado = !!(chkOmitirSabado && chkOmitirSabado.checked);

    const sheetData = [];
    sheetData.push(['DIA', 'HORÁRIO', ...turmasTurno.map(t => t.nome)]);

    DIAS_SEMANA.forEach(dia => {
        if (omitirSabado && dia === 'SABADO') return;
        const temposTurno = getTempos(turno);
        temposTurno.forEach(tempo => {
            const row = [];
            row.push(textoDia(dia));
            row.push(`${tempo.inicio} - ${tempo.fim}`);

            if (tempo.intervalo) {
                turmasTurno.forEach(() => row.push('INTERVALO'));
            } else {
                turmasTurno.forEach(turma => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) {
                        row.push('');
                    } else {
                        const prof = professores.find(p => p.id === aula.professorId);
                        row.push(
                            `${aula.disciplina}${prof ? ' - ' + prof.nome : ''}`
                        );
                    }
                });
            }

            sheetData.push(row);
        });
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    
    const colWidths = turmasTurno.map(() => ({ wch: 18 }));
    ws['!cols'] = [{ wch: 6 }, { wch: 12 }, ...colWidths];
    
    const rowsPerDia = getTempos(turno).length;
    const merges = [];
    DIAS_SEMANA.forEach((d, idxDia) => {
        const r0 = 1 + idxDia * rowsPerDia;
        const r1 = r0 + rowsPerDia - 1;
        merges.push({ s: { r: r0, c: 0 }, e: { r: r1, c: 0 } });
        const addr = XLSX.utils.encode_cell({ r: r0, c: 0 });
        if (ws[addr]) {
            ws[addr].s = Object.assign({}, ws[addr].s || {}, { alignment: { horizontal: 'center', vertical: 'center', textRotation: 90 } });
        }
    });
    ws['!merges'] = merges;
    const headDiaAddr = XLSX.utils.encode_cell({ r: 0, c: 0 });
    if (ws[headDiaAddr]) {
        ws[headDiaAddr].s = Object.assign({}, ws[headDiaAddr].s || {}, { alignment: { horizontal: 'center', vertical: 'center', textRotation: 90 } });
    }
    
    if (layout === 'colorido') {
        const toArgb = (hex) => {
            const m = /^#?([a-fA-F0-9]{6})$/.exec(hex || '');
            const h = m ? m[1] : '3498db';
            const r = parseInt(h.slice(0, 2), 16);
            const g = parseInt(h.slice(2, 4), 16);
            const b = parseInt(h.slice(4, 6), 16);
            const s = (n) => n.toString(16).padStart(2, '0').toUpperCase();
            return 'FF' + s(r) + s(g) + s(b);
        };
    DIAS_SEMANA.forEach((dia, idxDia) => {
        const temposTurno2 = getTempos(turno);
        temposTurno2.forEach((tempo, idxTempo) => {
            if (tempo.intervalo) return;
            turmasTurno.forEach((turma, idxTurma) => {
                const aula = obterAula(turno, dia, turma.id, tempo.id);
                if (!aula) return;
                const prof = professores.find(p => p.id === aula.professorId);
                if (!prof || !prof.cor) return;
                const r = 1 + idxDia * rowsPerDia + idxTempo;
                const c = 2 + idxTurma;
                const addr = XLSX.utils.encode_cell({ r, c });
                if (ws[addr]) {
                    ws[addr].s = Object.assign({}, ws[addr].s || {}, {
                        fill: { fgColor: { rgb: toArgb(prof.cor).slice(2) } },
                        font: { color: { rgb: 'FFFFFFFF' } },
                        alignment: { horizontal: 'center', vertical: 'center' }
                    });
                }
            });
        });
    });
    }
    
    XLSX.utils.book_append_sheet(wb, ws, 'Turno');
    XLSX.writeFile(wb, `horario_turno_${turno.toLowerCase()}_${new Date().toISOString().slice(0,10)}.xlsx`);
    mostrarToast('Excel gerado com sucesso!');
}

function gerarRelatorioCursoPDF() {
    const { jsPDF } = window.jspdf;
    const curso = document.getElementById('cursoRelatorio').value;
    const turno = document.getElementById('turnoRelatorioCurso').value;
    const layoutSel = document.getElementById('layoutRelatorioCurso');
    const layout = layoutSel ? layoutSel.value : 'compacto';

    if (!curso) {
        mostrarToast('Selecione um curso.', 'warning');
        return;
    }

    const turmasCurso = turmas.filter(t => t.curso === curso && t.turno === turno);
    if (!turmasCurso.length) {
        mostrarToast(`Não há turmas do curso ${curso} no turno da ${textoTurno(turno).toLowerCase()}.`, 'warning');
        return;
    }

    const colCount = 2 + turmasCurso.map(t => t.nome).length;
    const format = colCount > 12 ? 'a3' : 'a4';
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format });

    const titulo = `HORÁRIO - CURSO ${curso.toUpperCase()} - TURNO DA ${textoTurno(turno).toUpperCase()} — Gerado em: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`;
    doc.setFont(undefined, 'bold');
    doc.setFontSize(8);
    doc.text(titulo, doc.internal.pageSize.getWidth() / 2, 6, { align: 'center' });

    const head = [['DIA', 'HORÁRIO', ...turmasCurso.map(t => t.nome)]];
    const body = [];

    const chkOmitirSabado = document.getElementById('omitirSabadoRelatorioCurso');
    const omitirSabado = !!(chkOmitirSabado && chkOmitirSabado.checked);

    DIAS_SEMANA.forEach(dia => {
        if (omitirSabado && dia === 'SABADO') return;
        const temposTurno = getTempos(turno);
        temposTurno.forEach((tempo, idx) => {
            const row = [];
            
            // Coluna DIA com RowSpan
            if (idx === 0) {
                row.push({ 
                    content: textoDia(dia), 
                    rowSpan: temposTurno.length,
                    styles: { valign: 'middle', halign: 'center' } 
                });
            }

            // Remove a palavra INTERVALO da coluna de horário
            row.push(`${tempo.inicio} - ${tempo.fim}`);

            if (tempo.intervalo) {
                turmasCurso.forEach(() => row.push('INTERVALO'));
            } else {
                turmasCurso.forEach(turma => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) {
                        row.push('-');
                    } else {
                        const prof = professores.find(p => p.id === aula.professorId);
                        const textoCelula = `${aula.disciplina}\n${prof ? prof.nome : ''}`;
                        if (layout === 'colorido' && prof && prof.cor) {
                            const rgb = hexToRgb(prof.cor);
                            if (rgb) {
                                row.push({
                                    content: textoCelula,
                                    styles: { fillColor: [rgb.r, rgb.g, rgb.b], textColor: 255, valign: 'middle', halign: 'center' }
                                });
                            } else {
                                row.push(textoCelula);
                            }
                        } else {
                            row.push(textoCelula);
                        }
                    }
                });
            }

            body.push(row);
        });
    });

    const bodyFont = colCount > 12 ? 8 : 7;
    doc.autoTable({
        head,
        body,
        startY: 8,
        theme: 'grid',
        styles: {
            fontSize: bodyFont,
            cellPadding: 1,
            valign: 'middle',
            halign: 'center'
        },
        margin: { top: 6, right: 6, bottom: 6, left: 6 },
        headStyles: {
            fillColor: [231, 76, 60],
            textColor: 255,
            fontStyle: 'bold',
            fontSize: 8,
            cellPadding: 1,
            halign: 'center',
            valign: 'middle'
        },
        columnStyles: {
            0: { cellWidth: 10, cellPadding: 0 },
            1: { cellWidth: 20 }
        },
        didParseCell: function(data) {
            if (data.section === 'body' && data.column.index === 0) {
                data.cell.text = [];
            }
        },
        didDrawCell: function(data) {
            if (data.section === 'body' && data.column.index === 0) {
                const raw = data.row.raw[0];
                const text = raw && raw.content ? raw.content : data.cell.text[0];
                if (text) {
                    const prevSize = doc.getFontSize();
                    doc.setFontSize(7);
                    const xCenter = data.cell.x + data.cell.width / 2;
                    const x = xCenter + 5;
                    const yCenter = data.cell.y + data.cell.height / 2;
                    const ptToMm = 0.3528;
                    const offset = (doc.getFontSize() * ptToMm) * 0.5;
                    const y = yCenter + offset;
                    doc.setTextColor(0);
                    doc.text(text, x, y, { align: 'center', angle: 90 });
                    doc.setFontSize(prevSize);
                }
            }
        }
    });

    doc.save(`horario_curso_${curso.toLowerCase().replace(/\s+/g, '_')}_${turno.toLowerCase()}.pdf`);
    mostrarToast('PDF do curso gerado com sucesso!');
}

function gerarRelatorioCursoXLS() {
    if (typeof XLSX === 'undefined') {
        mostrarToast('Biblioteca XLSX não carregada.', 'error');
        return;
    }

    const curso = document.getElementById('cursoRelatorio').value;
    const turno = document.getElementById('turnoRelatorioCurso').value;
    const layoutSel = document.getElementById('layoutRelatorioCurso');
    const layout = layoutSel ? layoutSel.value : 'compacto';

    if (!curso) {
        mostrarToast('Selecione um curso.', 'warning');
        return;
    }

    const turmasCurso = turmas.filter(t => t.curso === curso && t.turno === turno);
    if (!turmasCurso.length) {
        mostrarToast(`Não há turmas do curso ${curso} no turno selecionado.`, 'warning');
        return;
    }

    const sheetData = [];
    sheetData.push(['DIA', 'HORÁRIO', ...turmasCurso.map(t => t.nome)]);

    DIAS_SEMANA.forEach(dia => {
        const temposTurno = getTempos(turno);
        temposTurno.forEach(tempo => {
            const row = [];
            row.push(textoDia(dia));
            row.push(`${tempo.inicio} - ${tempo.fim}`);

            if (tempo.intervalo) {
                turmasCurso.forEach(() => row.push('INTERVALO'));
            } else {
                turmasCurso.forEach(turma => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) {
                        row.push('');
                    } else {
                        const prof = professores.find(p => p.id === aula.professorId);
                        row.push(`${aula.disciplina}${prof ? ' - ' + prof.nome : ''}`);
                    }
                });
            }

            sheetData.push(row);
        });
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    
    const colWidths = turmasCurso.map(() => ({ wch: 18 }));
    ws['!cols'] = [{ wch: 6 }, { wch: 12 }, ...colWidths];
    
    const rowsPerDia = getTempos(turno).length;
    const merges = [];
    DIAS_SEMANA.forEach((d, idxDia) => {
        const r0 = 1 + idxDia * rowsPerDia;
        const r1 = r0 + rowsPerDia - 1;
        merges.push({ s: { r: r0, c: 0 }, e: { r: r1, c: 0 } });
        const addr = XLSX.utils.encode_cell({ r: r0, c: 0 });
        if (ws[addr]) {
            ws[addr].s = Object.assign({}, ws[addr].s || {}, { alignment: { horizontal: 'center', vertical: 'center', textRotation: 90 } });
        }
    });
    ws['!merges'] = merges;
    const headDiaAddr = XLSX.utils.encode_cell({ r: 0, c: 0 });
    if (ws[headDiaAddr]) {
        ws[headDiaAddr].s = Object.assign({}, ws[headDiaAddr].s || {}, { alignment: { horizontal: 'center', vertical: 'center', textRotation: 90 } });
    }
    
    if (layout === 'colorido') {
        const toArgb = (hex) => {
            const m = /^#?([a-fA-F0-9]{6})$/.exec(hex || '');
            const h = m ? m[1] : '3498db';
            const r = parseInt(h.slice(0, 2), 16);
            const g = parseInt(h.slice(2, 4), 16);
            const b = parseInt(h.slice(4, 6), 16);
            const s = (n) => n.toString(16).padStart(2, '0').toUpperCase();
            return 'FF' + s(r) + s(g) + s(b);
        };
        DIAS_SEMANA.forEach((dia, idxDia) => {
            const temposTurno2 = getTempos(turno);
            temposTurno2.forEach((tempo, idxTempo) => {
                if (tempo.intervalo) return;
                turmasCurso.forEach((turma, idxTurma) => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) return;
                    const prof = professores.find(p => p.id === aula.professorId);
                    if (!prof || !prof.cor) return;
                    const r = 1 + idxDia * rowsPerDia + idxTempo;
                    const c = 2 + idxTurma;
                    const addr = XLSX.utils.encode_cell({ r, c });
                    if (ws[addr]) {
                        ws[addr].s = Object.assign({}, ws[addr].s || {}, {
                            fill: { fgColor: { rgb: toArgb(prof.cor).slice(2) } },
                            font: { color: { rgb: 'FFFFFFFF' } },
                            alignment: { horizontal: 'center', vertical: 'center' }
                        });
                    }
                });
            });
        });
    }
    
    XLSX.utils.book_append_sheet(wb, ws, 'Curso');
    XLSX.writeFile(wb, `horario_curso_${curso.toLowerCase().replace(/\s+/g, '_')}_${turno.toLowerCase()}_${new Date().toISOString().slice(0,10)}.xlsx`);
    mostrarToast('Excel do curso gerado com sucesso!');
}

function gerarTabelaProfessor() {
    const profId = document.getElementById('professorRelatorio').value;
    const corpo = document.querySelector('#tabelaProfessor tbody');
    const tituloEl = document.getElementById('tabelaProfessorTitulo');
    const chkGraficos = document.getElementById('profRelatorioMostrarGraficos');
    const chkVagos = document.getElementById('profRelatorioMostrarVagos');
    const hintTempoVago = document.getElementById('profTempoVagoHint');
    const containerGraficos = document.querySelector('.prof-charts');
    if (!corpo) return;
    corpo.innerHTML = '';
    limparGraficosProfessor();

    if (!profId) {
        if (tituloEl) tituloEl.textContent = '';
        mostrarToast('Selecione um professor.', 'warning');
        return;
    }

    const prof = professores.find(p => p.id === profId);
    if (!prof) {
        if (tituloEl) tituloEl.textContent = '';
        mostrarToast('Professor não encontrado.', 'error');
        return;
    }

    let aulasProf = aulas.filter(a => a.professorId === profId);
    let incluirVagos = !chkVagos || chkVagos.checked;
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        incluirVagos = false;
    }

    if (hintTempoVago) {
        hintTempoVago.style.display = incluirVagos ? '' : 'none';
    }

    const agoraTabela = new Date();
    const dataTabela = agoraTabela.toLocaleDateString('pt-BR');
    const horaTabela = agoraTabela.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    const legendaTabela = `Pré-visualização: horário do professor ${prof.nome} — Gerado em: ${dataTabela} às ${horaTabela}`;

    if (!aulasProf.length && !incluirVagos) {
        if (tituloEl) tituloEl.textContent = legendaTabela;
        corpo.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;">Nenhum horário lançado para este professor.</td></tr>';
        return;
    }

    if (tituloEl) {
        tituloEl.textContent = legendaTabela;
    }

    const carga = calcularCargaHorariaProfessor(aulasProf);

    const disciplinasUsadas = [...new Set(
        aulasProf
            .map(a => a.disciplina)
            .filter(Boolean)
    )];

    const grupos = {};
    aulasProf.forEach(a => {
        const key = `${a.dia}_${a.turno}`;
        if (!grupos[key]) grupos[key] = [];
        grupos[key].push(a);
    });

    if (incluirVagos) {
        const temposPorDiaTurno = {};
        aulasProf.forEach(a => {
            const key = `${a.dia}_${a.turno}`;
            if (!temposPorDiaTurno[key]) temposPorDiaTurno[key] = new Set();
            temposPorDiaTurno[key].add(a.tempoId);
        });

        Object.keys(temposPorDiaTurno).forEach(key => {
            const [dia, turno] = key.split('_');
            const temposTurno = getTempos(turno).filter(t => !t.intervalo);
            const usados = temposPorDiaTurno[key];
            if (!usados || !usados.size) return;
            const usadosOrdenados = Array.from(usados).sort((a, b) => a - b);
            const minId = usadosOrdenados[0];
            const maxId = usadosOrdenados[usadosOrdenados.length - 1];
            temposTurno.forEach(t => {
                if (t.id >= minId && t.id <= maxId && !usados.has(t.id)) {
                    const profObj = professores.find(p => p.id === profId);
                    if (!profObj || !professorDisponivelNoHorario(profObj, turno, dia, t.id)) {
                        return;
                    }
                    grupos[key] = grupos[key] || [];
                    grupos[key].push({
                        id: `vago_${dia}_${turno}_${t.id}`,
                        turno,
                        dia,
                        turmaId: null,
                        tempoId: t.id,
                        disciplina: 'Tempo vago na escola',
                        professorId: profId,
                        _vago: true
                    });
                }
            });
        });
    }

    // Ordena as chaves: por dia da semana, depois por turno
    const chavesOrdenadas = Object.keys(grupos).sort((a, b) => {
        const [diaA, turnoA] = a.split('_');
        const [diaB, turnoB] = b.split('_');
        
        const idxDiaA = DIAS_SEMANA.indexOf(diaA);
        const idxDiaB = DIAS_SEMANA.indexOf(diaB);
        if (idxDiaA !== idxDiaB) return idxDiaA - idxDiaB;
        
        return TURNOS.indexOf(turnoA) - TURNOS.indexOf(turnoB);
    });

    let totalVagos = 0;

    chavesOrdenadas.forEach((chave, grupoIndex) => {
        const [dia, turno] = chave.split('_');
        const aulasGrupo = (grupos[chave] || []).sort((a, b) => a.tempoId - b.tempoId);
        aulasGrupo.forEach((a, idx) => {
            if (a._vago) totalVagos++;
            const turma = turmas.find(t => t.id === a.turmaId);
            const tempo = getTempos(turno).find(t => t.id === a.tempoId);
            
            const tr = document.createElement('tr');
            
            // Dia (apenas na primeira linha do grupo)
            const tdDia = document.createElement('td');
            if (idx === 0) {
                tdDia.textContent = textoDia(dia);
                tdDia.rowSpan = aulasGrupo.length;
                tr.appendChild(tdDia);
            }
            
            // Turno (apenas na primeira linha do grupo)
            const tdTurno = document.createElement('td');
            if (idx === 0) {
                tdTurno.textContent = textoTurno(turno);
                tdTurno.rowSpan = aulasGrupo.length;
                tr.appendChild(tdTurno);
            }
            
            // Horário
            const tdHorario = document.createElement('td');
            tdHorario.textContent = tempo ? `${tempo.inicio} - ${tempo.fim}` : '';
            tr.appendChild(tdHorario);
            
            // Tempo
            const tdTempo = document.createElement('td');
            tdTempo.textContent = tempo ? tempo.etiqueta : '';
            tr.appendChild(tdTempo);
            
            // Turma
            const turmaNome = turma ? turma.nome : '';
            const isVago = !!a._vago;
            if (!isVago) {
                const prevTurmaNome = idx === 0 ? null : (turmas.find(t => t.id === aulasGrupo[idx - 1].turmaId)?.nome || '');
                const isInicioRunTurma = idx === 0 || turmaNome !== prevTurmaNome;
                if (isInicioRunTurma) {
                    let spanTurma = 1;
                    for (let j = idx + 1; j < aulasGrupo.length; j++) {
                        const proxA = aulasGrupo[j];
                        const proxTurma = turmas.find(t => t.id === proxA.turmaId);
                        const proxNome = proxTurma ? proxTurma.nome : '';
                        if (proxNome === turmaNome && !proxA._vago) spanTurma++;
                        else break;
                    }
                    const tdTurma = document.createElement('td');
                    tdTurma.textContent = turmaNome;
                    tdTurma.rowSpan = spanTurma;
                    tdTurma.style.verticalAlign = 'middle';
                    tr.appendChild(tdTurma);
                }
            } else {
                const tdTurma = document.createElement('td');
                tdTurma.textContent = 'Escola';
                tr.appendChild(tdTurma);
            }
            
            // Disciplina
            const discNome = a.disciplina || '';
            const prevDisc = idx === 0 ? null : (aulasGrupo[idx - 1].disciplina || '');
            const isInicioRunDisc = idx === 0 || discNome !== prevDisc || a._vago || !!aulasGrupo[idx - 1]._vago;
            if (isInicioRunDisc) {
                let spanDisc = 1;
                for (let j = idx + 1; j < aulasGrupo.length; j++) {
                    const prox = aulasGrupo[j];
                    const proxDisc = prox.disciplina || '';
                    if (proxDisc === discNome && !prox._vago && !a._vago) spanDisc++;
                    else break;
                }
                const tdDisciplina = document.createElement('td');
                tdDisciplina.textContent = discNome;
                if (a._vago) tdDisciplina.style.fontStyle = 'italic';
                tdDisciplina.rowSpan = spanDisc;
                tdDisciplina.style.verticalAlign = 'middle';
                tr.appendChild(tdDisciplina);
            }
            
            corpo.appendChild(tr);
        });
        
        // Adiciona linha de separação entre grupos (exceto no último)
        if (grupoIndex < chavesOrdenadas.length - 1) {
            const trSeparador = document.createElement('tr');
            trSeparador.innerHTML = '<td colspan="6" style="background:#f5f5f5;height:1px;padding:0;"></td>';
            corpo.appendChild(trSeparador);
        }
    });

    const trResumo = document.createElement('tr');
    const tdLabel = document.createElement('td');
    tdLabel.colSpan = 3;
    tdLabel.textContent = 'Total de tempos na semana';
    const tdValor = document.createElement('td');
    tdValor.colSpan = 3;
    tdValor.textContent = `${carga.temposSemana} tempos (${carga.horasSemana}h semanais, aproximadamente ${carga.horasMes}h mensais)`;
    trResumo.appendChild(tdLabel);
    trResumo.appendChild(tdValor);
    trResumo.style.fontWeight = '600';
    corpo.appendChild(trResumo);

    if (totalVagos > 0 && incluirVagos) {
        const horasVagoSemana = totalVagos;
        const horasVagoMes = horasVagoSemana * 5;
        const trVagos = document.createElement('tr');
        const tdLabelV = document.createElement('td');
        tdLabelV.colSpan = 3;
        tdLabelV.textContent = 'Tempo vago na escola (entre aulas)';
        const tdValorV = document.createElement('td');
        tdValorV.colSpan = 3;
        tdValorV.textContent = `${totalVagos} tempos (${horasVagoSemana}h semanais, aproximadamente ${horasVagoMes}h mensais)`;
        trVagos.appendChild(tdLabelV);
        trVagos.appendChild(tdValorV);
        trVagos.style.fontStyle = 'italic';
        corpo.appendChild(trVagos);
    }

    let mostrarGraficos = !chkGraficos || chkGraficos.checked;
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        mostrarGraficos = false;
    }
    if (containerGraficos) {
        containerGraficos.style.display = mostrarGraficos ? '' : 'none';
    }
    if (mostrarGraficos) {
        atualizarGraficosProfessor(aulasProf);
    }
}

function preverImpressaoProfessor() {
    const profId = document.getElementById('professorRelatorio').value;
    if (!profId) {
        mostrarToast('Selecione um professor e gere a tabela antes de imprimir.', 'warning');
        return;
    }
    const corpo = document.querySelector('#tabelaProfessor tbody');
    if (!corpo || !corpo.children.length) {
        mostrarToast('Gere a tabela do professor antes de pré-visualizar a impressão.', 'warning');
        return;
    }
    const prof = professores.find(p => p.id === profId);
    const tituloEl = document.querySelector('#relatorios h2');
    let tituloOriginal = null;
    if (prof && tituloEl) {
        tituloOriginal = tituloEl.textContent;
        tituloEl.textContent = `Relatório do Professor ${prof.nome}`;
    }
    window.print();
    if (tituloEl && tituloOriginal !== null) {
        tituloEl.textContent = tituloOriginal;
    }
}

function calcularCargaHorariaProfessor(aulasProf) {
    const temposSemana = aulasProf.length;
    const horasSemana = temposSemana;
    const horasMes = temposSemana * 5;
    return { temposSemana, horasSemana, horasMes };
}

function limparGraficosProfessor() {
    if (chartProfDias) {
        chartProfDias.destroy();
        chartProfDias = null;
    }
    if (chartProfTurnos) {
        chartProfTurnos.destroy();
        chartProfTurnos = null;
    }
    if (chartProfDiasMes) {
        chartProfDiasMes.destroy();
        chartProfDiasMes = null;
    }
}

function atualizarGraficosProfessor(aulasProf) {
    if (typeof Chart === 'undefined') return;
    const canvasDias = document.getElementById('chartProfDiasSemana');
    const canvasTurnos = document.getElementById('chartProfTurnos');
    const canvasDiasMes = document.getElementById('chartProfDiasMes');
    if (!canvasDias || !canvasTurnos || !canvasDiasMes) return;
    const ctxDias = canvasDias.getContext('2d');
    const ctxTurnos = canvasTurnos.getContext('2d');
    if (!ctxDias || !ctxTurnos) return;

    const byDia = {};
    const byTurno = {};
    aulasProf.forEach(a => {
        byDia[a.dia] = (byDia[a.dia] || 0) + 1;
        byTurno[a.turno] = (byTurno[a.turno] || 0) + 1;
    });

    const diasKeys = DIAS_SEMANA.filter(d => byDia[d]);
    const diasLabels = diasKeys.map(d => textoDia(d));
    const diasValues = diasKeys.map(d => byDia[d]);
    const diasMesValues = diasValues.map(v => v * 5);

    const turnosKeys = TURNOS.filter(t => byTurno[t]);
    const turnosLabels = turnosKeys.map(t => textoTurno(t));
    const turnosValues = turnosKeys.map(t => byTurno[t]);

    const resumoDiasEl = document.getElementById('chartProfDiasResumo');
    const resumoTurnosEl = document.getElementById('chartProfTurnosResumo');
    const resumoMesEl = document.getElementById('chartProfDiasMesResumo');
    const cargaTotal = calcularCargaHorariaProfessor(aulasProf);

    if (chartProfDias) chartProfDias.destroy();
    chartProfDias = new Chart(ctxDias, {
        type: 'bar',
        data: {
            labels: diasLabels,
            datasets: [{
                label: 'Tempos por dia',
                data: diasValues,
                backgroundColor: PALETA_GRAFICOS.slice(0, diasLabels.length)
            }]
        },
        options: {
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const tempos = context.parsed.y;
                            const horasSemana = tempos;
                            const horasMes = tempos * 5;
                            return `${context.label}: ${tempos} tempos (${horasSemana}h/sem, ${horasMes}h/mês)`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        stepSize: 1
                    },
                    title: {
                        display: true,
                        text: 'Tempos na semana'
                    }
                }
            }
        }
    });

    if (chartProfTurnos) chartProfTurnos.destroy();
    chartProfTurnos = new Chart(ctxTurnos, {
        type: 'bar',
        data: {
            labels: turnosLabels,
            datasets: [{
                label: 'Tempos por turno',
                data: turnosValues,
                backgroundColor: PALETA_GRAFICOS.slice(0, turnosLabels.length)
            }]
        },
        options: {
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const tempos = context.parsed.y;
                            const horasSemana = tempos;
                            const horasMes = tempos * 5;
                            return `${context.label}: ${tempos} tempos (${horasSemana}h/sem, ${horasMes}h/mês)`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        stepSize: 1
                    },
                    title: {
                        display: true,
                        text: 'Tempos na semana'
                    }
                }
            }
        }
    });

    const ctxDiasMes = canvasDiasMes.getContext('2d');
    if (!ctxDiasMes) return;
    if (chartProfDiasMes) chartProfDiasMes.destroy();
    chartProfDiasMes = new Chart(ctxDiasMes, {
        type: 'bar',
        data: {
            labels: diasLabels,
            datasets: [{
                label: 'Horas por mês',
                data: diasMesValues,
                backgroundColor: PALETA_GRAFICOS.slice(0, diasLabels.length)
            }]
        },
        options: {
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const horasMes = context.parsed.y;
                            const temposSemana = horasMes / 5;
                            return `${context.label}: ${horasMes}h/mês (${temposSemana} tempos/sem)`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    },
                    title: {
                        display: true,
                        text: 'Horas por mês'
                    }
                }
            }
        }
    });

    if (resumoDiasEl) {
        resumoDiasEl.textContent =
            `Total: ${cargaTotal.temposSemana} tempos na semana ` +
            `(${cargaTotal.horasSemana}h semanais, aproximadamente ${cargaTotal.horasMes}h mensais).`;
    }
    if (resumoTurnosEl) {
        resumoTurnosEl.textContent =
            `Total: ${cargaTotal.temposSemana} tempos na semana ` +
            `(${cargaTotal.horasSemana}h semanais, aproximadamente ${cargaTotal.horasMes}h mensais).`;
    }
    if (resumoMesEl) {
        resumoMesEl.textContent =
            `Total: ${cargaTotal.horasMes}h por mês ` +
            `(${cargaTotal.temposSemana} tempos por semana).`;
    }
}

function gerarRelatorioProfessorPDF() {
    const { jsPDF } = window.jspdf;
    const profId = document.getElementById('professorRelatorio').value;
    if (!profId) {
        mostrarToast('Selecione um professor.', 'warning');
        return;
    }

    const prof = professores.find(p => p.id === profId);
    if (!prof) {
        mostrarToast('Professor não encontrado.', 'error');
        return;
    }

    const chkVagos = document.getElementById('profRelatorioMostrarVagos');
    let incluirVagos = !chkVagos || chkVagos.checked;
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        incluirVagos = false;
    }

    let aulasProf = aulas.filter(a => a.professorId === profId);
    if (!aulasProf.length && !incluirVagos) {
        mostrarToast('Este professor não possui aulas lançadas.', 'warning');
        return;
    }

    const carga = calcularCargaHorariaProfessor(aulasProf);
    const chkGraficos = document.getElementById('profRelatorioMostrarGraficos');
    let incluirGraficos = !chkGraficos || chkGraficos.checked;
    if (usuarioLogado && usuarioLogado.role === ROLES.COORD_TURNO) {
        incluirGraficos = false;
    }

    const grupos = {};
    aulasProf.forEach(a => {
        const key = `${a.dia}_${a.turno}`;
        if (!grupos[key]) grupos[key] = [];
        grupos[key].push(a);
    });

    if (incluirVagos) {
        const temposPorDiaTurno = {};
        aulasProf.forEach(a => {
            const key = `${a.dia}_${a.turno}`;
            if (!temposPorDiaTurno[key]) temposPorDiaTurno[key] = new Set();
            temposPorDiaTurno[key].add(a.tempoId);
        });

        Object.keys(temposPorDiaTurno).forEach(key => {
            const [dia, turno] = key.split('_');
            const temposTurno = getTempos(turno).filter(t => !t.intervalo);
            const usados = temposPorDiaTurno[key];
            if (!usados || !usados.size) return;
            const usadosOrdenados = Array.from(usados).sort((a, b) => a - b);
            const minId = usadosOrdenados[0];
            const maxId = usadosOrdenados[usadosOrdenados.length - 1];
            temposTurno.forEach(t => {
                if (t.id >= minId && t.id <= maxId && !usados.has(t.id)) {
                    const profObj = prof;
                    if (profObj && !professorDisponivelNoHorario(profObj, turno, dia, t.id)) {
                        return;
                    }
                    grupos[key] = grupos[key] || [];
                    grupos[key].push({
                        id: `vago_${dia}_${turno}_${t.id}`,
                        turno,
                        dia,
                        turmaId: null,
                        tempoId: t.id,
                        disciplina: 'Tempo vago na escola',
                        professorId: profId,
                        _vago: true
                    });
                }
            });
        });
    }

    const chavesOrdenadas = Object.keys(grupos).sort((a, b) => {
        const [diaA, turnoA] = a.split('_');
        const [diaB, turnoB] = b.split('_');
        const idxDiaA = DIAS_SEMANA.indexOf(diaA);
        const idxDiaB = DIAS_SEMANA.indexOf(diaB);
        if (idxDiaA !== idxDiaB) return idxDiaA - idxDiaB;
        return TURNOS.indexOf(turnoA) - TURNOS.indexOf(turnoB);
    });

    const agoraPDF = new Date();
    const dataPDF = agoraPDF.toLocaleDateString('pt-BR');
    const horaPDF = agoraPDF.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    const titulo = `HORÁRIO DO PROFESSOR - ${prof.nome.toUpperCase()}`;
    doc.setFontSize(10);
    doc.setFont(undefined, 'bold');
    doc.text(titulo, doc.internal.pageSize.getWidth() / 2, 26, {
        align: 'center',
        maxWidth: doc.internal.pageSize.getWidth() - 20
    });

    const head = [['Dia', 'Turno', 'Horário', 'Tempo', 'Turma', 'Disciplina']];
    const body = [];
    const sepRows = new Set();
    let bodyRowIndex = 0;

    let totalVagos = 0;

    chavesOrdenadas.forEach((chave, groupIndex) => {
        const [dia, turno] = chave.split('_');
        const aulasGrupo = (grupos[chave] || []).sort((a, b) => a.tempoId - b.tempoId);
        const rowsGrupo = aulasGrupo.length;
        aulasGrupo.forEach((a, idx) => {
            if (a._vago) totalVagos++;
            const turma = turmas.find(t => t.id === a.turmaId);
            const tempo = getTempos(turno).find(t => t.id === a.tempoId);

            const turmaNome = turma ? turma.nome : '';
            const isVago = !!a._vago;
            let turmaCell;
            if (isVago) {
                const prevIsVago = idx === 0 ? false : !!aulasGrupo[idx - 1]._vago;
                const isInicioRunEscola = idx === 0 || !prevIsVago;
                if (isInicioRunEscola) {
                    let spanEscola = 1;
                    for (let j = idx + 1; j < aulasGrupo.length; j++) {
                        const proxA = aulasGrupo[j];
                        if (proxA._vago) spanEscola++;
                        else break;
                    }
                    turmaCell = { content: 'Escola', rowSpan: spanEscola, styles: { valign: 'middle' } };
                } else {
                    turmaCell = '';
                }
            } else {
                const prevTurmaNome = idx === 0 ? null : (turmas.find(t => t.id === aulasGrupo[idx - 1].turmaId)?.nome || '');
                const isInicioRunTurma = idx === 0 || turmaNome !== prevTurmaNome;
                if (isInicioRunTurma) {
                    let spanTurma = 1;
                    for (let j = idx + 1; j < aulasGrupo.length; j++) {
                        const proxA = aulasGrupo[j];
                        const proxTurma = turmas.find(t => t.id === proxA.turmaId);
                        const proxNome = proxTurma ? proxTurma.nome : '';
                        if (proxNome === turmaNome && !proxA._vago) spanTurma++;
                        else break;
                    }
                    turmaCell = { content: turmaNome, rowSpan: spanTurma, styles: { valign: 'middle' } };
                } else {
                    turmaCell = '';
                }
            }

            const discNome = a.disciplina || '';
            const prevAula = idx === 0 ? null : aulasGrupo[idx - 1];
            const prevDisc = prevAula ? (prevAula.disciplina || '') : null;
            const prevIsVago = prevAula ? !!prevAula._vago : false;
            let isInicioRunDisc;
            if (isVago) {
                isInicioRunDisc = idx === 0 || !prevIsVago;
            } else {
                isInicioRunDisc = idx === 0 || discNome !== prevDisc || prevIsVago;
            }
            let discCell;
            if (isInicioRunDisc) {
                let spanDisc = 1;
                for (let j = idx + 1; j < aulasGrupo.length; j++) {
                    const prox = aulasGrupo[j];
                    if (isVago) {
                        if (prox._vago) spanDisc++;
                        else break;
                    } else {
                        const proxDisc = prox.disciplina || '';
                        if (proxDisc === discNome && !prox._vago) spanDisc++;
                        else break;
                    }
                }
                const discStyles = { valign: 'middle' };
                if (a._vago) discStyles.fontStyle = 'italic';
                discCell = { content: discNome, rowSpan: spanDisc, styles: discStyles };
            } else {
                discCell = '';
            }
            
            const diaCell = idx === 0
                ? { content: textoDia(dia), _rowsGrupo: rowsGrupo, styles: { valign: 'middle' } }
                : '';

            const turnoCell = idx === 0
                ? { content: textoTurno(turno), _rowsGrupo: rowsGrupo, styles: { valign: 'middle' } }
                : '';

            const row = [
                diaCell,
                turnoCell,
                tempo ? `${tempo.inicio} - ${tempo.fim}` : '',
                tempo ? tempo.etiqueta : '',
                turmaCell,
                discCell
            ];
            
            body.push(row);
            bodyRowIndex++;
        });
        const nextKey = chavesOrdenadas[groupIndex + 1];
        const nextDia = nextKey ? nextKey.split('_')[0] : null;
        if (!nextKey || nextDia !== dia) {
            sepRows.add(bodyRowIndex - 1);
        }
    });

    const disciplinasUsadas = [...new Set(aulasProf.map(a => a.disciplina).filter(Boolean))];

    const disciplinasLabel =
        (disciplinasUsadas.length
            ? disciplinasUsadas.join(', ')
            : (Array.isArray(prof.disciplinas) && prof.disciplinas.length
                ? prof.disciplinas.join(', ')
                : 'Não informado'));

    doc.setFontSize(8);
    doc.setFont(undefined, 'normal');
    doc.text(
        `Disciplinas: ${disciplinasLabel}`,
        14,
        30
    );

    doc.autoTable({
        head,
        body,
        startY: 34,
        theme: 'grid',
        styles: {
            fontSize: 9,
            cellPadding: 2,
            valign: 'middle',
            halign: 'center',
            minCellHeight: 12
        },
        headStyles: {
            fillColor: [52, 152, 219],
            textColor: 255,
            fontStyle: 'bold',
            fontSize: 7,
            cellPadding: 1,
            minCellHeight: 6,
            halign: 'center',
            valign: 'middle'
        },
        columnStyles: {
            0: { cellWidth: 12, cellPadding: 0 },  // Dia
            1: { cellWidth: 12, cellPadding: 0 },  // Turno
            2: { cellWidth: 32, cellPadding: 1 },  // Horário
            3: { cellWidth: 26, cellPadding: 1 },  // Tempo
            4: { cellWidth: 30 },                  // Turma
            5: { cellWidth: 40, cellPadding: 1 }   // Disciplina (mais estreita, mas ainda confortável)
        },
        didParseCell: function(data) {
            if (!data || !data.cell) return;

            const colIdx = data.column.index;

            if (data.section === 'head' &&
                (colIdx === 0 || colIdx === 1) &&
                data.cell.styles) {
                data.cell.styles.fontSize = 6;
            }

            if (data.section === 'body' && data.cell.styles) {
                if (colIdx === 0) {
                    data.cell.styles.fontSize = 8;
                }
                if (colIdx === 5) {
                    data.cell.styles.fontSize = 8;
                }
            }
        },
        didDrawCell: function(data) {
            return;
        }
    });

    const finalY = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY : 28;
    doc.setFontSize(8);
    doc.setFont(undefined, 'normal');
    doc.text(
        `Total de tempos na semana: ${carga.temposSemana} tempos (${carga.horasSemana}h semanais, aproximadamente ${carga.horasMes}h mensais)`,
        14,
        finalY + 5
    );
    if (incluirVagos && totalVagos > 0) {
        const horasVagoSemana = totalVagos;
        const horasVagoMes = horasVagoSemana * 5;
        doc.setFont(undefined, 'italic');
        doc.text(
            `Tempo vago na escola (entre aulas): ${totalVagos} tempos (${horasVagoSemana}h semanais, aproximadamente ${horasVagoMes}h mensais).`,
            14,
            finalY + 9
        );
        doc.setFont(undefined, 'normal');
    }
    doc.text(
        `Gerado em: ${dataPDF} às ${horaPDF}`,
        14,
        finalY + 13
    );

    if (incluirGraficos) {
        adicionarGraficosProfessorNoPDF(doc, aulasProf);
    }

    doc.save(`horario_prof_${prof.nome.replace(/\s+/g, '_').toLowerCase()}_${new Date().toISOString().slice(0,10)}.pdf`);
    mostrarToast('PDF do professor gerado com sucesso!');
}

function criarImagemGraficoBarra(labels, values, titulo, eixoY) {
    if (typeof Chart === 'undefined') return null;
    if (!labels || !labels.length) return null;
    try {
        const canvas = document.createElement('canvas');
        canvas.width = 500;
        canvas.height = 260;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
            canvas.remove();
            return null;
        }
        const cores = ['#3498db', '#e67e22', '#2ecc71', '#9b59b6', '#f1c40f', '#e74c3c'];
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: titulo,
                    data: values,
                    backgroundColor: cores.slice(0, labels.length)
                }]
            },
            options: {
                responsive: false,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: { display: true, text: titulo }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { precision: 0 },
                        title: eixoY ? { display: true, text: eixoY } : undefined
                    }
                }
            }
        });
        const dataUrl = canvas.toDataURL('image/png');
        chart.destroy();
        canvas.remove();
        return dataUrl;
    } catch (e) {
        console.error('Erro ao gerar imagem de gráfico:', e);
        return null;
    }
}

function desenharGraficoBarrasPDF(doc, x, y, width, height, labels, values, titulo, eixoY) {
    if (!labels || !labels.length) return;
    if (!values || !values.length) return;
    const maxVal = Math.max(...values) || 1;
    const top = y + 8;
    const bottom = y + height - 10;
    const left = x + 8;
    const right = x + width - 8;
    const chartHeight = bottom - top;
    const chartWidth = right - left;
    const barWidth = chartWidth / labels.length * 0.7;
    const gap = chartWidth / labels.length * 0.3;

    const coresRgb = PALETA_GRAFICOS.map(hex => {
        const c = hexToRgb(hex);
        if (c) return [c.r, c.g, c.b];
        return [52, 152, 219];
    });

    doc.setFontSize(10);
    doc.setFont(undefined, 'bold');
    doc.text(titulo, x + width / 2, y + 4, { align: 'center' });

    doc.setFontSize(8);
    doc.setFont(undefined, 'normal');
    doc.line(left, top, left, bottom);
    doc.line(left, bottom, right, bottom);

    labels.forEach((label, idx) => {
        const val = values[idx];
        const ratio = val / maxVal;
        const h = ratio * chartHeight;
        const bx = left + idx * (barWidth + gap) + gap / 2;
        const by = bottom - h;
        const cor = coresRgb[idx % coresRgb.length];
        doc.setFillColor(cor[0], cor[1], cor[2]);
        doc.rect(bx, by, barWidth, h, 'F');
        const lx = bx + barWidth / 2;
        doc.text(String(val), lx, by - 1, { align: 'center' });
        doc.text(label, lx, bottom + 4, { align: 'center' });
    });

    if (eixoY) {
        doc.text(eixoY, left - 6, top - 2, { align: 'left' });
    }
}

function adicionarGraficosProfessorNoPDF(doc, aulasProf) {
    try {
        if (!aulasProf || !aulasProf.length) return;

        const byDia = {};
        const byTurno = {};
        aulasProf.forEach(a => {
            byDia[a.dia] = (byDia[a.dia] || 0) + 1;
            byTurno[a.turno] = (byTurno[a.turno] || 0) + 1;
        });

        const diasKeys = DIAS_SEMANA.filter(d => byDia[d]);
        const diasLabels = diasKeys.map(d => textoDia(d));
        const diasValues = diasKeys.map(d => byDia[d]);
        const diasMesValues = diasValues.map(v => v * 5);

        const turnosKeys = TURNOS.filter(t => byTurno[t]);
        const turnosLabels = turnosKeys.map(t => textoTurno(t));
        const turnosValues = turnosKeys.map(t => byTurno[t]);

        doc.addPage();

        const resumoHead = [['Dia', 'Tempos/semana', 'Horas/mês']];
        const resumoBody = diasLabels.map((label, idx) => [
            label,
            diasValues[idx],
            diasMesValues[idx]
        ]);

        if (resumoBody.length) {
            doc.autoTable({
                head: resumoHead,
                body: resumoBody,
                startY: 20,
                theme: 'grid',
                styles: {
                    fontSize: 9,
                    cellPadding: 2,
                    valign: 'middle'
                },
                headStyles: {
                    fillColor: [52, 152, 219],
                    textColor: 255,
                    fontStyle: 'bold'
                }
            });
        }

        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 10;
        const fullWidth = pageWidth - margin * 2;
        const colWidth = (pageWidth - margin * 3) / 2;
        const chartHeight = 70;
        let y = doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY + 8 : 20;

        const maxPageY = doc.internal.pageSize.getHeight() - 10;

        if (diasLabels.length || turnosLabels.length) {
            if (y + chartHeight > maxPageY) {
                doc.addPage();
                y = 20;
            }

            if (diasLabels.length) {
                desenharGraficoBarrasPDF(
                    doc,
                    margin,
                    y,
                    colWidth,
                    chartHeight,
                    diasLabels,
                    diasValues,
                    'Tempos por dia da semana',
                    'Tempos na semana'
                );
            }

            if (turnosLabels.length) {
                desenharGraficoBarrasPDF(
                    doc,
                    margin + colWidth + margin,
                    y,
                    colWidth,
                    chartHeight,
                    turnosLabels,
                    turnosValues,
                    'Tempos por turno',
                    'Tempos na semana'
                );
            }

            y += chartHeight + 8;
        }

        if (diasLabels.length) {
            if (y + chartHeight > maxPageY) {
                doc.addPage();
                y = 20;
            }
            desenharGraficoBarrasPDF(
                doc,
                margin,
                y,
                fullWidth,
                chartHeight,
                diasLabels,
                diasMesValues,
                'Horas por mês por dia da semana',
                'Horas por mês'
            );
        }
    } catch (e) {
        console.error('Erro ao adicionar gráficos no PDF do professor:', e);
    }
}

// Converte cor hex para RGB
function hexToRgb(hex) {
    if (!hex) return null;
    const res = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return res ? {
        r: parseInt(res[1], 16),
        g: parseInt(res[2], 16),
        b: parseInt(res[3], 16)
    } : null;
}

// =======================
// CONSULTAS - ATUALIZADO
// =======================

function gerarConsulta() {
    const tipo = document.getElementById('tipoConsulta').value;
    const thead = document.querySelector('#tabelaConsulta thead');
    const tbody = document.querySelector('#tabelaConsulta tbody');
    if (!thead || !tbody) return;
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (tipo === 'professores') {
        thead.innerHTML = `
            <tr>
                <th>Professor</th>
                <th>Disciplinas</th>
                <th>Total de Aulas</th>
                <th>Turnos Atuantes</th>
            </tr>
        `;
        
        professores.forEach(p => {
            const aulasProfessor = aulas.filter(a => a.professorId === p.id);
            const turnos = [...new Set(aulasProfessor.map(a => a.turno))];
            const disciplinasUsadas = [...new Set(
                aulasProfessor
                    .map(a => a.disciplina)
                    .filter(Boolean)
            )];
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${p.nome}</strong></td>
                <td>${disciplinasUsadas.join(', ') || '-'}</td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;">${aulasProfessor.length}</span></td>
                <td>${turnos.map(t => textoTurno(t)).join(', ') || '-'}</td>
            `;
            tbody.appendChild(tr);
        });
    } 
    else if (tipo === 'turmas') {
        thead.innerHTML = `
            <tr>
                <th>Curso / Área</th>
                <th>Turno</th>
                <th>Turma</th>
                <th>Descrição</th>
                <th>Total de Aulas</th>
            </tr>
        `;
        
        // Agrupa turmas por curso
        const turmasPorCurso = {};
        turmas.forEach(t => {
            if (!turmasPorCurso[t.curso]) {
                turmasPorCurso[t.curso] = [];
            }
            turmasPorCurso[t.curso].push(t);
        });
        
        Object.keys(turmasPorCurso).forEach(curso => {
            const turmasCurso = turmasPorCurso[curso];
            
            // Cabeçalho do grupo de curso
            const trHeader = document.createElement('tr');
            trHeader.className = 'curso-group-header';
            trHeader.innerHTML = `
                <td colspan="5" style="background-color:#e0e0e0; font-weight:bold; padding:8px;">
                    ${curso} <span style="font-weight:normal; font-size:0.9em;">(${turmasCurso.length} turmas)</span>
                </td>
            `;
            tbody.appendChild(trHeader);

            // Agrupa por turno dentro do curso
            const turmasPorTurno = {};
            turmasCurso.forEach(t => {
                if (!turmasPorTurno[t.turno]) turmasPorTurno[t.turno] = [];
                turmasPorTurno[t.turno].push(t);
            });

            Object.keys(turmasPorTurno).forEach(turno => {
                const turmasDoTurno = turmasPorTurno[turno];
                
                turmasDoTurno.forEach((t, idx) => {
                    const totalAulas = aulas.filter(a => a.turmaId === t.id).length;
                    const tr = document.createElement('tr');
                    
                    // Coluna Curso (vazia pois já tem cabeçalho, ou merged se preferir layout tabela pura)
                    // Aqui mantemos vazia para indentação visual sob o cabeçalho
                    const tdCurso = document.createElement('td');
                    tdCurso.style.border = 'none'; // visual mais limpo
                    tr.appendChild(tdCurso);

                    // Coluna Turno (apenas na primeira do grupo)
                    const tdTurno = document.createElement('td');
                    if (idx === 0) {
                        tdTurno.textContent = textoTurno(turno);
                        tdTurno.rowSpan = turmasDoTurno.length;
                        tdTurno.style.verticalAlign = 'middle';
                        tdTurno.style.fontWeight = 'bold';
                        tr.appendChild(tdTurno);
                    } else if (idx === 0) { 
                        // nunca entra aqui, mas garante logica
                    }

                    // Turma
                    const tdTurma = document.createElement('td');
                    tdTurma.innerHTML = `<strong>${t.nome}</strong>`;
                    tr.appendChild(tdTurma);

                    // Descrição
                    const tdDesc = document.createElement('td');
                    tdDesc.textContent = t.descricao || '-';
                    tr.appendChild(tdDesc);

                    // Total
                    const tdTotal = document.createElement('td');
                    tdTotal.innerHTML = `<span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${totalAulas}</span>`;
                    tr.appendChild(tdTotal);

                    tbody.appendChild(tr);
                });
            });
            
            // Espaço entre cursos
            const trEspaco = document.createElement('tr');
            trEspaco.innerHTML = '<td colspan="5" style="height:10px; border:none;"></td>';
            tbody.appendChild(trEspaco);
        });
    } 
    else if (tipo === 'turnos') {
        thead.innerHTML = `
            <tr>
                <th>Turno</th>
                <th>Quantidade de Turmas</th>
                <th>Total de Aulas</th>
                <th>Cursos Presentes</th>
            </tr>
        `;
        
        TURNOS.forEach(turno => {
            const turmasTurno = turmas.filter(t => t.turno === turno);
            const idsTurmas = turmasTurno.map(t => t.id);
            const totalAulas = aulas.filter(a => idsTurmas.includes(a.turmaId)).length;
            const cursosTurno = [...new Set(turmasTurno.map(t => t.curso))];
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${textoTurno(turno)}</strong></td>
                <td><span style="background:#fff8e1;padding:2px 8px;border-radius:12px;">${turmasTurno.length}</span></td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;">${totalAulas}</span></td>
                <td>${cursosTurno.join(', ') || '-'}</td>
            `;
            tbody.appendChild(tr);
        });
    } 
    else if (tipo === 'cursos') {
        thead.innerHTML = `
            <tr>
                <th>Curso / Área</th>
                <th>Turmas Matutinas</th>
                <th>Turmas Vespertinas</th>
                <th>Turmas Noturnas</th>
                <th>Total de Turmas</th>
                <th>Total de Aulas</th>
            </tr>
        `;
        
        cursos.forEach(curso => {
            const turmasCurso = turmas.filter(t => t.curso === curso);
            const turmasManha = turmasCurso.filter(t => t.turno === 'MANHA').length;
            const turmasTarde = turmasCurso.filter(t => t.turno === 'TARDE').length;
            const turmasNoite = turmasCurso.filter(t => t.turno === 'NOITE').length;
            const idsTurmas = turmasCurso.map(t => t.id);
            const totalAulas = aulas.filter(a => idsTurmas.includes(a.turmaId)).length;
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${curso}</strong></td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;">${turmasManha}</span></td>
                <td><span style="background:#fff8e1;padding:2px 8px;border-radius:12px;">${turmasTarde}</span></td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;background:#2c3e50;color:white;">${turmasNoite}</span></td>
                <td><span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${turmasCurso.length}</span></td>
                <td><span style="background:#d4edda;padding:2px 8px;border-radius:12px;font-weight:bold;">${totalAulas}</span></td>
            `;
            tbody.appendChild(tr);
        });
    }
    
    mostrarToast('Consulta gerada com sucesso!');
}

// =======================
// BACKUP E RESTAURAÇÃO
// =======================

function exportarBackup() {
    const dados = {
        professores,
        turmas,
        aulas,
        cursos,
        exportadoEm: new Date().toISOString(),
        versao: '1.0'
    };
    
    const blob = new Blob([JSON.stringify(dados, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `backup_horarios_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    mostrarToast('Backup exportado com sucesso!');
}

function importarBackup(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const dados = JSON.parse(e.target.result);
            
            // Valida estrutura básica
            if (!dados.professores || !dados.turmas || !dados.aulas) {
                mostrarToast('Arquivo de backup inválido.', 'error');
                return;
            }
            
            if (!confirm(`Importar backup com:\n• ${dados.professores.length} professor(es)\n• ${dados.turmas.length} turma(s)\n• ${dados.aulas.length} aula(s)\n\nIsso substituirá todos os dados atuais. Continuar?`)) {
                return;
            }
            
            professores = dados.professores;
            turmas = dados.turmas;
            aulas = dados.aulas;
            cursos = dados.cursos || cursos;
            
            salvarLocal();
            location.reload();
            
        } catch (error) {
            console.error('Erro ao importar backup:', error);
            mostrarToast('Erro ao importar backup. Arquivo corrompido.', 'error');
        }
    };
    
    reader.readAsText(file);
    
    // Limpa o input para permitir importar o mesmo arquivo novamente
    event.target.value = '';
}

// =======================
// INICIALIZAÇÃO
// =======================

document.addEventListener('DOMContentLoaded', () => {
    carregarLocal();
    migrarEstruturaLocal();

    // formulários
    const loginCard = document.getElementById('loginCard');
    const formLogin = document.getElementById('formLogin');
    const loginHint = document.getElementById('loginHint');
    const mainContent = document.querySelector('.main-content');

    if (formLogin && mainContent && loginCard) {
        mainContent.style.display = 'none';
        formLogin.addEventListener('submit', (e) => {
            e.preventDefault();
            const login = document.getElementById('loginUsuario').value.trim();
            const senha = document.getElementById('senhaUsuario').value.trim();
            const user = autenticarUsuario(login, senha);
            if (!user) {
                if (loginHint) loginHint.textContent = 'Usuário ou senha inválidos.';
                return;
            }
            usuarioLogado = user;
            localStorage.setItem('usuarioAtual', JSON.stringify({
                id: user.id,
                role: user.role,
                turno: user.turno || null,
                nome: user.nome
            }));
            loginCard.style.display = 'none';
            mainContent.style.display = 'flex';
            if (loginHint) loginHint.textContent = '';
            aplicarPermissoesNaUI();
            showSection('horarios');
        });
        const salvo = localStorage.getItem('usuarioAtual');
        if (salvo) {
            try {
                const u = JSON.parse(salvo);
                const match = USUARIOS_FIXOS.find(x => x.id === u.id);
                if (match) {
                    usuarioLogado = match;
                    loginCard.style.display = 'none';
                    mainContent.style.display = 'flex';
                    aplicarPermissoesNaUI();
                }
            } catch (e) {
            }
        }
    }

    const formProfessor = document.getElementById('formProfessor');
    if (formProfessor) formProfessor.addEventListener('submit', salvarProfessor);

    const formTurma = document.getElementById('formTurma');
    if (formTurma) formTurma.addEventListener('submit', salvarTurma);

    const formHorario = document.getElementById('formHorario');
    if (formHorario) formHorario.addEventListener('submit', salvarHorario);

    const formNovoCurso = document.getElementById('formNovoCurso');
    if (formNovoCurso) formNovoCurso.addEventListener('submit', salvarNovoCurso);

    // modal de aula
    const formModal = document.getElementById('formModalAula');
    if (formModal) formModal.addEventListener('submit', salvarEdicaoAula);

    const btnExcluirModal = document.getElementById('btnExcluirAulaModal');
    if (btnExcluirModal) btnExcluirModal.addEventListener('click', excluirAulaModal);

    const modalAula = document.getElementById('modalAula');
    if (modalAula) {
        modalAula.addEventListener('click', (e) => {
            if (e.target === modalAula) fecharModalAula();
        });
    }

    // modal curso
    const modalCurso = document.getElementById('modalCurso');
    if (modalCurso) {
        modalCurso.addEventListener('click', (e) => {
            if (e.target === modalCurso) fecharModalCurso();
        });
    }

    // inicial
    preencherSelectCursos();
    preencherSelectCursoVisual();
    atualizarListaProfessores();
    atualizarListaTurmas();
    atualizarChipsTempos();
    montarGradeTurno(turnoAtualGrade);
    preencherSelectCursoRelatorio();
    preencherSelectProfessorRelatorio();
    const chkVagosRelProf = document.getElementById('profRelatorioMostrarVagos');
    const hintTempoVago = document.getElementById('profTempoVagoHint');
    if (chkVagosRelProf && hintTempoVago) {
        const syncTempoVagoHint = () => {
            hintTempoVago.style.display = chkVagosRelProf.checked ? '' : 'none';
        };
        chkVagosRelProf.addEventListener('change', syncTempoVagoHint);
        syncTempoVagoHint();
    }
    const chkMostrarInstrucoesEditor = document.getElementById('chkMostrarInstrucoesEditor');
    if (chkMostrarInstrucoesEditor) {
        chkMostrarInstrucoesEditor.addEventListener('change', () => mostrarHintEditorTurno());
        mostrarHintEditorTurno();
    }
    
    // Auto-preencher disciplina ao selecionar professor no Editor de Horários
    const selectProfHorario = document.getElementById('professorHorario');
    if (selectProfHorario) {
        selectProfHorario.addEventListener('change', function() {
            const profId = this.value;
            const disciplinaInput = document.getElementById('disciplinaHorario');
            if (profId && disciplinaInput) {
                const prof = professores.find(p => p.id === profId);
                if (prof && prof.disciplinas && prof.disciplinas.length > 0) {
                    // Pega a primeira disciplina cadastrada (ou todas unidas)
                    // Como é um array, vamos pegar a primeira para preencher o campo
                    disciplinaInput.value = prof.disciplinas[0]; 
                }
            }
        });
    }

    // Auto-preencher disciplina no Modal de Edição de Aula
    const selectProfModal = document.getElementById('modalProfessor');
    if (selectProfModal) {
        selectProfModal.addEventListener('change', function() {
            const profId = this.value;
            const disciplinaInput = document.getElementById('modalDisciplina');
            if (profId && disciplinaInput) {
                const prof = professores.find(p => p.id === profId);
                if (prof && prof.disciplinas && prof.disciplinas.length > 0) {
                    disciplinaInput.value = prof.disciplinas[0];
                }
            }
        });
    }
    
    mostrarToast('Sistema carregado com sucesso!', 'success');
});

function abrirModalProfessor(id) {
    const p = professores.find(px => px.id === id);
    if (!p) return;
    const modal = document.getElementById('modalProfessorPerfil');
    document.getElementById('modalProfessorId').value = p.id;
    document.getElementById('modalProfNome').value = p.nome;
    document.getElementById('modalProfDisciplinas').value = p.disciplinas.join(', ');
    document.getElementById('modalProfCor').value = p.cor || '#3498db';
    const diasSel = p.diasDisponiveis || [];
    const turnosSel = p.turnosDisponiveis || [];
    const dispCont = document.getElementById('modalDispProfessor');
    if (dispCont) {
        renderDisponibilidade('modalDispProfessor', p.disponibilidadePorDia || null, diasSel, turnosSel, p.temposPorDiaTurno || null);
    }
    modal.style.display = 'flex';
}

function fecharModalProfessor() {
    const modal = document.getElementById('modalProfessorPerfil');
    if (modal) modal.style.display = 'none';
}

function salvarModalProfessor(e) {
    e.preventDefault();
    const id = document.getElementById('modalProfessorId').value;
    const p = professores.find(px => px.id === id);
    if (!p) return;
    const nome = document.getElementById('modalProfNome').value.trim();
    const discsStr = document.getElementById('modalProfDisciplinas').value;
    const cor = document.getElementById('modalProfCor').value || '#3498db';
    if (!nome) {
        mostrarToast('Informe o nome do professor.', 'warning');
        return;
    }
    const corLower = cor.toLowerCase();
    const conflitoCor = professores.find(px => px.cor && px.cor.toLowerCase() === corLower && px.id !== id);
    if (conflitoCor) {
        const msgCor =
            `A cor escolhida já está sendo usada por ${conflitoCor.nome}.\n` +
            'Deseja continuar usando a mesma cor para este professor?';
        const continuar = confirm(msgCor);
        if (!continuar) {
            return;
        }
    }
    p.nome = nome;
    p.disciplinas = discsStr ? discsStr.split(',').map(s => s.trim()).filter(Boolean) : [];
    p.cor = cor;
    const dispDiaTurno = coletarDisponibilidade('modalDispProfessor');
    const temposDiaTurno = coletarTemposDisponibilidade('modalDispProfessor');
    const dias = Object.keys(dispDiaTurno).filter(d => (dispDiaTurno[d] || []).length > 0);
    const turnos = Array.from(new Set([].concat(...Object.values(dispDiaTurno))));
    const tempos = (() => {
        const s = new Set();
        Object.values(temposDiaTurno || {}).forEach(turnoMap => {
            Object.values(turnoMap || {}).forEach(lista => {
                (lista || []).forEach(id => s.add(id));
            });
        });
        return Array.from(s);
    })();

    const novoProf = {
        ...p,
        nome,
        disciplinas: p.disciplinas,
        cor,
        diasDisponiveis: dias,
        turnosDisponiveis: turnos,
        disponibilidadePorDia: dispDiaTurno,
        temposPorDiaTurno: temposDiaTurno,
        temposDisponiveis: tempos.length
            ? tempos
            : getTempos('MANHA').filter(t => !t.intervalo).map(t => t.id)
    };

    const aulasProfExistentes = aulas.filter(a => a.professorId === p.id);
    const aulasIndisponiveis = aulasProfExistentes.filter(a =>
        !professorDisponivelNoHorario(novoProf, a.turno, a.dia, a.tempoId)
    );

    if (aulasIndisponiveis.length > 0) {
        const detalhes = aulasIndisponiveis.slice(0, 5).map(a => {
            const turma = turmas.find(t => t.id === a.turmaId);
            const tempo = getTempos(a.turno).find(t => t.id === a.tempoId);
            const descTempo = tempo ? tempo.etiqueta : `Tempo ${a.tempoId}`;
            const descTurma = turma ? ` - Turma ${turma.nome}` : '';
            return `${textoDia(a.dia)} / ${textoTurno(a.turno)} / ${descTempo}${descTurma}`;
        });
        const msgExtra = aulasIndisponiveis.length > detalhes.length
            ? `\n... e mais ${aulasIndisponiveis.length - detalhes.length} aula(s).`
            : '';
        const msg =
            'Atenção: com a nova disponibilidade, este professor ficará indisponível ' +
            `para ${aulasIndisponiveis.length} aula(s) já lançada(s):\n\n` +
            detalhes.join('\n') +
            msgExtra +
            '\n\nDeseja salvar mesmo assim?';
        if (!confirm(msg)) {
            return;
        }
    }

    p.nome = novoProf.nome;
    p.disciplinas = novoProf.disciplinas;
    p.cor = novoProf.cor;
    p.diasDisponiveis = novoProf.diasDisponiveis;
    p.turnosDisponiveis = novoProf.turnosDisponiveis;
    p.disponibilidadePorDia = novoProf.disponibilidadePorDia;
    p.temposPorDiaTurno = novoProf.temposPorDiaTurno;
    p.temposDisponiveis = novoProf.temposDisponiveis;
    salvarLocal();
    atualizarListaProfessores();
    fecharModalProfessor();
    mostrarToast('Professor atualizado com sucesso!');
}

document.addEventListener('DOMContentLoaded', () => {
    const formModalProf = document.getElementById('formModalProfessor');
    if (formModalProf) {
        formModalProf.addEventListener('submit', salvarModalProfessor);
    }
    const modalProf = document.getElementById('modalProfessorPerfil');
    if (modalProf) {
        modalProf.addEventListener('click', (e) => {
            if (e.target === modalProf) fecharModalProfessor();
        });
    }
});

function renderDisponibilidade(containerId, dados = null, diasLegacy = [], turnosLegacy = [], temposMap = null) {
    const cont = document.getElementById(containerId);
    if (!cont) return;
    const dias = ['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA','SABADO'];
    const turnos = ['MANHA','TARDE','NOITE'];
    cont.innerHTML = '';
    dias.forEach(d => {
        const day = document.createElement('div');
        day.className = 'disp-day';
        day.innerHTML = `
            <div class="day-label">${textoDia(d)}</div>
            <div class="turnos-row">
                ${turnos.map(t =>
                    `<label>
                        <input type="checkbox" data-dia="${d}" data-turno="${t}" data-tipo="turno">
                        ${textoTurno(t)}
                    </label>`
                ).join('')}
            </div>
            <div class="tempos-row" id="tempos-${containerId}-${d}"></div>
        `;
        cont.appendChild(day);
    });
    // Apply data
    const map = dados || {};
    dias.forEach(d => {
        const arr = map[d] || null;
        turnos.forEach(t => {
            const chk = cont.querySelector(`#${containerId} input[data-dia="${d}"][data-turno="${t}"][data-tipo="turno"]`) || cont.querySelector(`input[data-dia="${d}"][data-turno="${t}"][data-tipo="turno"]`);
            if (!chk) return;
            if (arr) {
                chk.checked = arr.includes(t);
            } else {
                const okDia = diasLegacy && diasLegacy.includes ? diasLegacy.includes(d) : false;
                const okTurno = turnosLegacy && turnosLegacy.includes ? turnosLegacy.includes(t) : false;
                chk.checked = okDia && okTurno;
            }
        });
        renderTemposDia(containerId, d, dados, temposMap);
    });

    cont.addEventListener('change', (e) => {
        const target = e.target;
        if (target && target.matches('input[type="checkbox"][data-dia][data-turno][data-tipo="turno"]')) {
            const dia = target.getAttribute('data-dia');
            renderTemposDia(containerId, dia, coletarDisponibilidade(containerId), null);
        }
    });
}

function coletarDisponibilidade(containerId) {
    const cont = document.getElementById(containerId);
    const out = {};
    if (!cont) return out;
    const dias = ['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA','SABADO'];
    dias.forEach(d => {
        const checks = Array.from(cont.querySelectorAll(`input[data-dia="${d}"][data-tipo="turno"]:checked`));
        out[d] = checks.map(i => i.getAttribute('data-turno'));
    });
    return out;
}

function renderTemposDia(containerId, dia, disp = null, temposMap = null) {
    const cont = document.getElementById(`tempos-${containerId}-${dia}`);
    if (!cont) return;
    const turnos = ['MANHA','TARDE','NOITE'];
    cont.innerHTML = '';
    turnos.forEach(t => {
        const ativo = disp && disp[dia] ? disp[dia].includes(t) : false;
        if (!ativo) return;
        const lista = temposMap && temposMap[dia] && temposMap[dia][t] ? temposMap[dia][t] : getTempos(t).filter(x=>!x.intervalo).map(x=>x.id);
        const row = document.createElement('div');
        row.className = 'turno-tempos-row';
        const lblTurno = document.createElement('div');
        lblTurno.className = 'turno-tempos-label';
        lblTurno.textContent = textoTurno(t);
        const wrapper = document.createElement('div');
        wrapper.className = 'filters-table turno-tempos-table';
        getTempos(t).filter(x=>!x.intervalo).forEach(tp => {
            const label = document.createElement('label');
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.value = tp.id;
            chk.setAttribute('data-dia', dia);
            chk.setAttribute('data-turno', t);
            chk.setAttribute('data-tipo', 'tempo');
            chk.checked = lista.includes(tp.id);
            label.appendChild(chk);
            label.appendChild(document.createTextNode(` ${etiquetaTempoMin(tp)}`));
            wrapper.appendChild(label);
        });
        row.appendChild(lblTurno);
        row.appendChild(wrapper);
        cont.appendChild(row);
    });
}

function coletarTemposDisponibilidade(containerId) {
    const cont = document.getElementById(containerId);
    const out = {};
    if (!cont) return out;
    const dias = ['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA','SABADO'];
    const turnos = ['MANHA','TARDE','NOITE'];
    dias.forEach(d => {
        out[d] = {};
        turnos.forEach(t => {
            const checks = Array.from(document.querySelectorAll(`#${containerId} input[data-dia="${d}"][data-turno="${t}"][data-tipo="tempo"]:checked`));
            out[d][t] = checks.map(i => parseInt(i.value, 10)).filter(n => !isNaN(n));
        });
    });
    return out;
}
