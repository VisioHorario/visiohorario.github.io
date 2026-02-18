// =======================
// CONFIGURAÇÃO BÁSICA
// =======================

// Prefixo para localStorage (evita conflito entre aplicações)
const APP_PREFIX = 'horarios_escola_';

// Dias e turnos
const DIAS_SEMANA = ['SEGUNDA', 'TERCA', 'QUARTA', 'QUINTA', 'SEXTA', 'SABADO'];

function diasInsercaoRapida() {
    return DIAS_SEMANA.filter(d => d !== 'SABADO');
}
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

let ORDEM_TURMAS_POR_TURNO = {};

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
let cursos = ['Informática', 'Administração', 'Enfermagem'];
let professores = [];
let turmas = [];
let aulas = [];

const CURSO_TEMPO_INTEGRAL = 'Tempo Integral';

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

let turnoAtualGrade = 'MANHA';
let modoVisualGrade = 'TURNO';
let cursoVisualGrade = '';
let pilhaUndo = [];
let pilhaRedo = [];

// Controle do modal de edição de aula
let aulaEmEdicao = null;

let chartProfDias = null;
let chartProfTurnos = null;
let chartProfDiasMes = null;

const PALETA_GRAFICOS = ['#3498db', '#e74c3c', '#2ecc71', '#9b59b6', '#f1c40f', '#e67e22'];


// Cache para otimização
let cacheGrade = {};
let dadosModificados = false;
let melhorPercentualAjustePorTurno = {};
let gradeCelulasSelecionadas = [];
let clipboardAulas = [];
let clipboardModo = null;
let menuContextoGrade = null;
let menuContextoTimeout = null;

function limparSelecaoGrade() {
    gradeCelulasSelecionadas.forEach(td => {
        td.style.outline = '';
        td.style.outlineOffset = '';
    });
    gradeCelulasSelecionadas = [];
}

function alternarSelecaoCelula(td) {
    if (!td) return;
    const idx = gradeCelulasSelecionadas.indexOf(td);
    if (idx >= 0) {
        gradeCelulasSelecionadas.splice(idx, 1);
        td.style.outline = '';
        td.style.outlineOffset = '';
    } else {
        gradeCelulasSelecionadas.push(td);
        td.style.outline = '2px solid #2980b9';
        td.style.outlineOffset = '-2px';
    }
}

function obterAulasDasCelulas(celulas) {
    if (!celulas || !celulas.length) return [];
    const lista = [];
    celulas.forEach(td => {
        const turno = td.dataset.turno;
        const dia = td.dataset.dia;
        const turmaId = td.dataset.turmaId;
        const tempoId = parseInt(td.dataset.tempoId, 10);
        const aula = obterAula(turno, dia, turmaId, tempoId);
        if (aula) {
            lista.push(aula);
        }
    });
    return lista;
}

function fecharMenuContextoGrade() {
    if (menuContextoGrade) {
        menuContextoGrade.style.display = 'none';
    }
    if (menuContextoTimeout) {
        clearTimeout(menuContextoTimeout);
        menuContextoTimeout = null;
    }
}

function abrirMenuContextoGrade(ev, td) {
    if (!td) return;
    if (!ev.ctrlKey && gradeCelulasSelecionadas.length === 0) {
        limparSelecaoGrade();
        alternarSelecaoCelula(td);
    } else if (!ev.ctrlKey && gradeCelulasSelecionadas.indexOf(td) === -1) {
        limparSelecaoGrade();
        alternarSelecaoCelula(td);
    }
    if (!menuContextoGrade) {
        menuContextoGrade = document.createElement('div');
        menuContextoGrade.style.position = 'fixed';
        menuContextoGrade.style.zIndex = '9999';
        menuContextoGrade.style.background = '#ffffff';
        menuContextoGrade.style.border = '1px solid #ccc';
        menuContextoGrade.style.boxShadow = '0 2px 6px rgba(0,0,0,0.15)';
        menuContextoGrade.style.minWidth = '140px';
        menuContextoGrade.style.fontSize = '0.85rem';
        document.body.appendChild(menuContextoGrade);
    }
    const destinos = gradeCelulasSelecionadas.length ? [...gradeCelulasSelecionadas] : [td];
    const aulasSelecionadas = obterAulasDasCelulas(destinos);
    const podeCopiar = aulasSelecionadas.length > 0;
    const podeRecortar = aulasSelecionadas.length > 0;
    const podeExcluir = aulasSelecionadas.length > 0;
    const podeColar = clipboardAulas && clipboardAulas.length > 0;
    const criarItem = (rotulo, habilitado, acao) => {
        const item = document.createElement('div');
        item.textContent = rotulo;
        item.style.padding = '6px 10px';
        item.style.cursor = habilitado ? 'pointer' : 'default';
        item.style.color = habilitado ? '#333' : '#aaa';
        item.onmouseenter = () => {
            if (habilitado) item.style.background = '#f0f0f0';
        };
        item.onmouseleave = () => {
            item.style.background = 'transparent';
        };
        if (habilitado) {
            item.onclick = () => {
                fecharMenuContextoGrade();
                acao();
            };
        }
        menuContextoGrade.appendChild(item);
    };
    menuContextoGrade.innerHTML = '';
    criarItem('Copiar', podeCopiar, () => {
        clipboardModo = 'copy';
        clipboardAulas = aulasSelecionadas.map(a => ({ ...a }));
    });
    criarItem('Recortar', podeRecortar, () => {
        if (!aulasSelecionadas.length) return;
        clipboardModo = 'cut';
        clipboardAulas = aulasSelecionadas.map(a => ({ ...a }));
    });
    criarItem('Excluir aula', podeExcluir, () => {
        if (!aulasSelecionadas.length) return;
        if (!confirm('Deseja excluir a(s) aula(s) selecionada(s)?')) return;
        capturarEstadoAulas();
        const aulasParaExcluir = [];
        aulasSelecionadas.forEach(sel => {
            const mesmas = aulas.filter(a =>
                a.turno === sel.turno &&
                a.dia === sel.dia &&
                a.turmaId === sel.turmaId &&
                a.tempoId === sel.tempoId
            );
            mesmas.forEach(a => {
                if (!aulasParaExcluir.some(x => x.id === a.id)) {
                    aulasParaExcluir.push(a);
                }
            });
        });
        const ids = new Set(aulasParaExcluir.map(a => a.id));
        aulas = aulas.filter(a => !ids.has(a.id));
        salvarLocal();
        montarGradeTurno(turnoAtualGrade);
    });
    criarItem('Colar', podeColar, () => {
        if (!clipboardAulas || !clipboardAulas.length) return;
        const destinosPaste = gradeCelulasSelecionadas.length ? [...gradeCelulasSelecionadas] : [td];
        const backupAulas = JSON.stringify(aulas);
        const backupClipboardModo = clipboardModo;
        const backupClipboardAulas = JSON.stringify(clipboardAulas);
        let operacaoFalhou = false;
        capturarEstadoAulas();
        let aulasUsadasCut = [];
        try {
            let avisouPortMat = false;
            let idx = 0;
            for (let i = 0; i < destinosPaste.length && idx < clipboardAulas.length; i++) {
                const cel = destinosPaste[i];
                const base = clipboardAulas[idx];
                idx++;
                const turno = cel.dataset.turno;
                const dia = cel.dataset.dia;
                const turmaId = cel.dataset.turmaId;
                const tempoId = parseInt(cel.dataset.tempoId, 10);
                const profId = base.professorId || null;
                if (profId) {
                    const prof = professores.find(p => p.id === profId);
                    if (!prof) {
                        mostrarToast('Professor não encontrado.', 'error');
                        operacaoFalhou = true;
                        break;
                    }
                    if (!professorDisponivelNoHorario(prof, turno, dia, tempoId)) {
                        const continuarDisp = confirm('Professor não está disponível neste dia/turno/tempo. Deseja continuar mesmo assim?');
                        if (!continuarDisp) {
                            operacaoFalhou = true;
                            break;
                        }
                    }
                    const existenteProf = aulas.find(a =>
                        a.id !== base.id &&
                        a.professorId === profId &&
                        a.turno === turno &&
                        a.dia === dia &&
                        a.tempoId === tempoId &&
                        a.turmaId !== turmaId
                    );
                    if (existenteProf) {
                        const continuar = confirm('Professor já tem aula neste mesmo horário em outra turma. Deseja continuar mesmo assim?');
                        if (!continuar) {
                            operacaoFalhou = true;
                            break;
                        }
                    }
                }
                const grupoDisc = grupoDisciplinaPortMat(base.disciplina);
                if (grupoDisc && !avisouPortMat) {
                    const aulasMesmoDiaMesmaDisciplina = aulas.filter(a =>
                        a.turmaId === turmaId &&
                        a.dia === dia &&
                        grupoDisciplinaPortMat(a.disciplina) === grupoDisc &&
                        !(a.turno === turno && a.dia === dia && a.turmaId === turmaId && a.tempoId === tempoId)
                    );
                    const totalFinalDiscDia = aulasMesmoDiaMesmaDisciplina.length + 1;
                    const limite = 2;
                    if (totalFinalDiscDia > limite) {
                        mostrarToast('Boas práticas pedagógicas sugerem não concentrar mais de 2 tempos de Português ou Matemática no mesmo dia para a mesma turma.', 'warning');
                        avisouPortMat = true;
                    }
                }
                const existente = obterAula(turno, dia, turmaId, tempoId);
                if (existente) {
                    aulas = aulas.filter(a => a.id !== existente.id);
                }
                if (clipboardModo === 'cut') {
                    aulasUsadasCut.push(base);
                }
                const nova = {
                    id: gerarId(),
                    turno,
                    dia,
                    turmaId,
                    tempoId,
                    disciplina: base.disciplina,
                    professorId: base.professorId || null
                };
                aulas.push(nova);
            }
        } catch (e) {
            operacaoFalhou = true;
        }
        if (operacaoFalhou) {
            aulas = JSON.parse(backupAulas);
            clipboardModo = backupClipboardModo;
            clipboardAulas = JSON.parse(backupClipboardAulas);
            salvarLocal();
            montarGradeTurno(turnoAtualGrade);
            mostrarToast('Não foi possível colar as aulas recortadas neste(s) horário(s). As aulas foram devolvidas à célula de origem.', 'warning');
            return;
        }
        if (clipboardModo === 'cut') {
            const idsOrigem = new Set(aulasUsadasCut.map(a => a.id));
            aulas = aulas.filter(a => !idsOrigem.has(a.id));
            clipboardAulas = [];
            clipboardModo = null;
        }
        salvarLocal();
        montarGradeTurno(turnoAtualGrade);
    });
    menuContextoGrade.style.left = ev.clientX + 'px';
    menuContextoGrade.style.top = ev.clientY + 'px';
    menuContextoGrade.style.display = 'block';
    if (menuContextoTimeout) {
        clearTimeout(menuContextoTimeout);
    }
    menuContextoTimeout = setTimeout(() => {
        if (menuContextoGrade) {
            menuContextoGrade.style.display = 'none';
        }
        menuContextoTimeout = null;
    }, 5000);
}

// =======================
// UTILITÁRIOS
// =======================

function atualizarIndicadorDesfazerTurno() {
    const btn = document.getElementById('btnDesfazerTurno');
    if (!btn) return;
    const turno = turnoAtualGrade;
    let qtd = 0;
    pilhaUndo.forEach(item => {
        try {
            const entry = JSON.parse(item);
            if (entry && entry.turno === turno) {
                qtd++;
            }
        } catch (e) {
        }
    });
    const base = '↩️ Desfazer';
    btn.textContent = qtd > 0 ? `${base} (${qtd})` : base;
}

function atualizarIndicadorRefazerTurno() {
    const btn = document.getElementById('btnRefazerTurno');
    if (!btn) return;
    const turno = turnoAtualGrade;
    let qtd = 0;
    pilhaRedo.forEach(item => {
        try {
            const entry = JSON.parse(item);
            if (entry && entry.turno === turno) {
                qtd++;
            }
        } catch (e) {
        }
    });
    const base = '↪️ Refazer';
    btn.textContent = qtd > 0 ? `${base} (${qtd})` : base;
}

function capturarEstadoAulas() {
    const turno = turnoAtualGrade;
    const aulasTurno = aulas.filter(a => a.turno === turno);
    pilhaUndo.push(JSON.stringify({ turno, aulas: aulasTurno }));
    if (pilhaUndo.length > 20) {
        pilhaUndo.shift();
    }
    atualizarIndicadorDesfazerTurno();
    pilhaRedo = [];
    atualizarIndicadorRefazerTurno();
}

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

function textoTipoTurma(tipo) {
    if (tipo === 'INTEGRADO') return 'Integrado ao Médio';
    if (tipo === 'SUBSEQUENTE') return 'Subsequente ao Médio';
    if (tipo === 'MEDIO') return 'Médio';
    if (tipo === 'FUNDAMENTAL1') return 'Fundamental 1';
    if (tipo === 'FUNDAMENTAL') return 'Fundamental 2';
    return '';
}

function textoAnoTurma(ano) {
    if (ano === '1ANO') return '1º Ano';
    if (ano === '2ANO') return '2º Ano';
    if (ano === '3ANO') return '3º Ano';
    if (ano === '4ANO') return '4º Ano';
    if (ano === '5ANO') return '5º Ano';
    if (ano === '6ANO') return '6º Ano';
    if (ano === '7ANO') return '7º Ano';
    if (ano === '8ANO') return '8º Ano';
    if (ano === '9ANO') return '9º Ano';
    return '';
}

function textoFaseTurma(fase) {
    if (fase === 'I') return 'I Fase';
    if (fase === 'II') return 'II Fase';
    if (fase === 'III') return 'III Fase';
    if (fase === 'IV') return 'IV Fase';
    if (fase === 'V') return 'V Fase';
    if (fase === 'VI') return 'VI Fase';
    return '';
}

function compararTurmasPorCursoETipo(a, b) {
    const tipoA = a.tipoTurma || '';
    const tipoB = b.tipoTurma || '';
    const subA = tipoA === 'SUBSEQUENTE';
    const subB = tipoB === 'SUBSEQUENTE';
    if (subA !== subB) return subA ? 1 : -1;
    const ca = (a.curso || '').toString();
    const cb = (b.curso || '').toString();
    const cmpCurso = ca.localeCompare(cb, 'pt-BR', { sensitivity: 'base' });
    if (cmpCurso !== 0) return cmpCurso;
    return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
}

function atualizarOpcoesFaseTurma(idSelectAno, idSelectFase) {
    const selAno = document.getElementById(idSelectAno);
    const selFase = document.getElementById(idSelectFase);
    if (!selAno || !selFase) return;
    const ano = selAno.value;
    let fases = [];
    let tipo = '';
    if (idSelectAno === 'anoTurma') {
        const selTipo = document.getElementById('tipoTurma');
        tipo = selTipo ? selTipo.value : '';
    } else if (idSelectAno === 'modalAnoTurma') {
        const selTipo = document.getElementById('modalTipoTurma');
        tipo = selTipo ? selTipo.value : '';
    }
    if (tipo !== 'FUNDAMENTAL1' && tipo !== 'FUNDAMENTAL') {
        if (ano === '1ANO') {
            fases = ['I', 'II'];
        } else if (ano === '2ANO') {
            fases = ['III', 'IV'];
        } else if (ano === '3ANO') {
            fases = ['V', 'VI'];
        }
    }
    selFase.innerHTML = '<option value="">Fase...</option>';
    fases.forEach(f => {
        const opt = document.createElement('option');
        opt.value = f;
        opt.textContent = textoFaseTurma(f);
        selFase.appendChild(opt);
    });
}

function atualizarOpcoesAnoTurma(idSelectTipo, idSelectAno, idSelectFase) {
    const selTipo = document.getElementById(idSelectTipo);
    const selAno = document.getElementById(idSelectAno);
    const selFase = idSelectFase ? document.getElementById(idSelectFase) : null;
    if (!selTipo || !selAno) return;
    const tipo = selTipo.value;
    let anos = [];
    if (tipo === 'FUNDAMENTAL1') {
        anos = ['1ANO', '2ANO', '3ANO', '4ANO', '5ANO'];
    } else if (tipo === 'FUNDAMENTAL') {
        anos = ['6ANO', '7ANO', '8ANO', '9ANO'];
    } else if (tipo) {
        anos = ['1ANO', '2ANO', '3ANO'];
    }
    const valorAtual = selAno.value;
    selAno.innerHTML = '<option value="">Ano...</option>';
    anos.forEach(a => {
        const opt = document.createElement('option');
        opt.value = a;
        opt.textContent = textoAnoTurma(a);
        selAno.appendChild(opt);
    });
    if (anos.includes(valorAtual)) {
        selAno.value = valorAtual;
    } else {
        selAno.value = '';
    }
    if (selFase) {
        selFase.innerHTML = '<option value="">Fase...</option>';
    }
}

function abreviar(str, tam) {
    if (!str) return '';
    return str.length > tam ? str.slice(0, tam - 1) + '…' : str;
}

let apelidosDisciplinas = {};

function apelidarDisciplina(nome) {
    if (!nome) return '';
    const key = nome.trim();
    if (apelidosDisciplinas && Object.prototype.hasOwnProperty.call(apelidosDisciplinas, key)) {
        const ap = apelidosDisciplinas[key];
        if (ap && ap.trim()) return ap.trim();
    }
    return nome;
}

function normalizarPalavrasDisciplina(nome) {
    const original = (nome || '').trim();
    if (!original) return [];
    const partes = original.split(/\s+/).map(p => p.replace(/[.,;:]/g, ''));
    const stop = ['DE', 'DA', 'DO', 'DAS', 'DOS', 'E', 'EM', 'COM', 'PARA'];
    const filtradas = partes.filter(p => p && !stop.includes(p.toUpperCase()));
    return filtradas.length ? filtradas : partes.filter(Boolean);
}

function capitalizarParteDisciplina(str) {
    if (!str) return '';
    const lower = str.toLowerCase();
    return lower.charAt(0).toUpperCase() + lower.slice(1);
}

function sugerirApelidoDisciplina(nome) {
    const original = (nome || '').trim();
    if (!original) return '';
    const palavras = normalizarPalavrasDisciplina(original);
    const totalChars = original.length;
    if (palavras.length === 1) {
        return original;
    }
    if (palavras.length <= 2 && totalChars <= 25) {
        const p1 = capitalizarParteDisciplina(palavras[0].slice(0, 4));
        const p2 = palavras[1] ? capitalizarParteDisciplina(palavras[1].slice(0, 2)) : '';
        return (p1 + (p2 ? ' ' + p2 : '')).trim();
    }
    const iniciais = palavras.map(p => p.charAt(0).toUpperCase()).join(' ');
    return iniciais;
}

function grupoDisciplinaPortMat(nome) {
    const n = (nome || '').toUpperCase().trim();
    if (!n) return null;
    if (
        n.includes('PORTUG') ||
        n === 'PORT AP' ||
        n === 'PORT APLICADO' ||
        n === 'PORT APLICADA'
    ) {
        return 'PORTUG';
    }
    if (
        n.includes('MATEM') ||
        n === 'MAT AP' ||
        n === 'MAT APLICADO' ||
        n === 'MAT APLICADA'
    ) {
        return 'MATEM';
    }
    return null;
}

function calcularQualidadeAjusteTurno(turno) {
    const aulasTurno = aulas.filter(a => a.turno === turno);
    if (!aulasTurno.length) {
        return { percentual: 100, totalAvaliado: 0, totalOk: 0 };
    }

    let totalAvaliado = 0;
    let totalOk = 0;

    const mapaFamilias = {};
    aulasTurno.forEach(a => {
        const tipo = grupoDisciplinaPortMat(a.disciplina);
        if (!tipo) return;
        const chave = `${a.turmaId || ''}|${a.professorId || ''}|${tipo}`;
        if (!mapaFamilias[chave]) {
            mapaFamilias[chave] = [];
        }
        mapaFamilias[chave].push(a);
    });
    const familias = Object.values(mapaFamilias);
    familias.forEach(lista => {
        if (!lista || lista.length !== 4) return;
        totalAvaliado++;
        const porDia = {};
        lista.forEach(a => {
            if (!porDia[a.dia]) porDia[a.dia] = [];
            porDia[a.dia].push(a.tempoId);
        });
        const dias = Object.keys(porDia);
        if (dias.length !== 2) return;
        let ok = true;
        for (const d of dias) {
            const tempos = porDia[d].slice().sort((x, y) => x - y);
            if (tempos.length !== 2) {
                ok = false;
                break;
            }
            if (Math.abs(tempos[0] - tempos[1]) !== 1) {
                ok = false;
                break;
            }
        }
        if (ok) {
            totalOk++;
        }
    });

    const mapaPares = {};
    aulasTurno.forEach(a => {
        const turmaId = a.turmaId || '';
        const profId = a.professorId || '';
        const disc = (a.disciplina || '').toUpperCase();
        if (!turmaId || !disc) return;
        const chave = `${turmaId}|${profId}|${disc}`;
        if (!mapaPares[chave]) {
            mapaPares[chave] = [];
        }
        mapaPares[chave].push(a);
    });
    Object.values(mapaPares).forEach(lista => {
        if (!lista || lista.length !== 2) return;
        totalAvaliado++;
        const a1 = lista[0];
        const a2 = lista[1];
        if (a1.dia !== a2.dia) return;
        const diff = Math.abs(a1.tempoId - a2.tempoId);
        if (diff === 1) {
            totalOk++;
        }
    });

    if (!totalAvaliado) {
        return { percentual: 100, totalAvaliado: 0, totalOk: 0 };
    }
    const percentual = Math.round((totalOk / totalAvaliado) * 100);
    return { percentual, totalAvaliado, totalOk };
}

function preencherSelectDisciplinasApelido() {
    const sel = document.getElementById('disciplinaApelidoSelect');
    if (!sel) return;
    const disciplinas = [...new Set(
        [
            ...professores
                .flatMap(p => Array.isArray(p.disciplinas) ? p.disciplinas : []),
            ...aulas
                .map(a => a.disciplina)
        ].filter(Boolean)
    )].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const atual = sel.value;
    sel.innerHTML = '<option value="">Selecione...</option>' +
        disciplinas.map(d => `<option value="${d}">${d}</option>`).join('');
    if (atual) sel.value = atual;
    const inputApelido = document.getElementById('apelidoDisciplinaInput');
    const chkRemover = document.getElementById('removerApelidoDisciplinaCheckbox');
    if (inputApelido) {
        inputApelido.disabled = false;
        const discSel = sel.value;
        let valor = '';
        if (discSel && apelidosDisciplinas[discSel]) {
            valor = apelidosDisciplinas[discSel];
        } else if (discSel) {
            valor = sugerirApelidoDisciplina(discSel);
        }
        inputApelido.value = valor;
    }
    if (chkRemover) {
        chkRemover.checked = false;
    }
}

function salvarApelidoDisciplina(ev) {
    ev.preventDefault();
    const sel = document.getElementById('disciplinaApelidoSelect');
    const input = document.getElementById('apelidoDisciplinaInput');
    if (!sel || !input) return;
    const disciplina = sel.value.trim();
    const apelido = input.value.trim();
    if (!disciplina) {
        mostrarToast('Selecione uma disciplina para apelidar.', 'warning');
        return;
    }
    if (!apelido) {
        delete apelidosDisciplinas[disciplina];
        mostrarToast('Apelido removido para esta disciplina.');
    } else {
        apelidosDisciplinas[disciplina] = apelido;
        mostrarToast('Apelido salvo para a disciplina.');
    }
    salvarLocal();
}

function toggleRemoverApelidoDisciplina(chk) {
    const input = document.getElementById('apelidoDisciplinaInput');
    if (!input) return;
    if (chk && chk.checked) {
        input.value = '';
        input.disabled = true;
    } else {
        input.disabled = false;
        const sel = document.getElementById('disciplinaApelidoSelect');
        if (sel && sel.value && apelidosDisciplinas[sel.value]) {
            input.value = apelidosDisciplinas[sel.value];
        }
    }
}

function mostrarToast(mensagem, tipo = 'success') {
    const toastAntigo = document.querySelector('.toast');
    if (toastAntigo && toastAntigo.parentNode) {
        toastAntigo.parentNode.removeChild(toastAntigo);
    }
    const toast = document.createElement('div');
    toast.className = `toast ${tipo}`;
    const spanMsg = document.createElement('span');
    spanMsg.textContent = mensagem;
    toast.appendChild(spanMsg);
    const btnFechar = document.createElement('button');
    btnFechar.type = 'button';
    btnFechar.className = 'toast-close';
    btnFechar.textContent = '×';
    btnFechar.onclick = () => {
        if (toast.parentNode) {
            toast.style.animation = 'slideIn 0.3s ease reverse';
            setTimeout(() => toast.remove(), 300);
        }
    };
    toast.appendChild(btnFechar);
    document.body.appendChild(toast);
    setTimeout(() => {
        if (toast.parentNode) {
            toast.style.animation = 'slideIn 0.3s ease reverse';
            setTimeout(() => toast.remove(), 300);
        }
    }, 10000);
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
        localStorage.setItem(`${APP_PREFIX}apelidos_disciplinas`, JSON.stringify(apelidosDisciplinas));
        localStorage.setItem(`${APP_PREFIX}ordem_turmas_turno`, JSON.stringify(ORDEM_TURMAS_POR_TURNO));
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
        const ap = localStorage.getItem(`${APP_PREFIX}apelidos_disciplinas`);
        const ordemTurmas = localStorage.getItem(`${APP_PREFIX}ordem_turmas_turno`);
        if (p) professores = JSON.parse(p);
        if (t) turmas = JSON.parse(t);
        if (a) aulas = JSON.parse(a);
        if (c) cursos = JSON.parse(c);
        if (!cursos.includes(CURSO_TEMPO_INTEGRAL)) {
            cursos.push(CURSO_TEMPO_INTEGRAL);
        }
        if (tempos) {
            try { TEMPOS_POR_TURNO = JSON.parse(tempos); } catch {}
        }
        if (ap) {
            try { apelidosDisciplinas = JSON.parse(ap) || {}; } catch { apelidosDisciplinas = {}; }
        }
        if (ordemTurmas) {
            try { ORDEM_TURMAS_POR_TURNO = JSON.parse(ordemTurmas) || {}; } catch { ORDEM_TURMAS_POR_TURNO = {}; }
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
        if (!p.baseTipo) {
            p.baseTipo = 'TECNICA';
            alterou = true;
        }
    });
    turmas.forEach(t => {
        if (!t.turno || !['MANHA','TARDE','NOITE'].includes(t.turno)) { t.turno = 'MANHA'; alterou = true; }
        if (!t.tipoTurma) { t.tipoTurma = ''; alterou = true; }
        if (!t.anoTurma) { t.anoTurma = ''; alterou = true; }
        if (!t.faseTurma) { t.faseTurma = ''; alterou = true; }
    });
    aulas.forEach(a => {
        if (!a.turno || !['MANHA','TARDE','NOITE'].includes(a.turno)) { a.turno = 'MANHA'; alterou = true; }
        if (!DIAS_SEMANA.includes(a.dia)) { a.dia = 'SEGUNDA'; alterou = true; }
        if (typeof a.tempoId !== 'number') { a.tempoId = parseInt(a.tempoId, 10) || 1; alterou = true; }
    });
    if (alterou) salvarLocal();
}

function prioridadeProfessorAlocacao(professorId) {
    if (!professorId) return 500;
    
    const prof = professores.find(p => p.id === professorId);
    if (!prof) return 500;
    
    let prioridade = 0;
    
    // 1. Prioridade por tipo de professor (Base Comum > Técnico)
    const baseTipo = prof.baseTipo || 'TECNICA';
    if (baseTipo === 'COMUM') prioridade -= 100;
    
    // 2. Priorizar professores com menos aulas alocadas
    const aulasAtuais = aulas.filter(a => a.professorId === professorId).length;
    prioridade -= aulasAtuais * 10;
    
    // 3. Considerar disponibilidade do professor
    const turnoAtual = turnoAtualGrade;
    const diasDisponiveis = DIAS_SEMANA.filter(dia => 
        professorDisponivelNoDia(professorId, turnoAtual, dia)
    );
    prioridade -= diasDisponiveis.length * 5;
    
    // 4. Priorizar professores com carga horária mais equilibrada
    const cargaPorDia = calcularCargaHorariaPorDia(professorId, turnoAtual);
    const desvioPadrao = calcularDesvioPadraoCarga(cargaPorDia);
    prioridade -= Math.round(desvioPadrao * 3);
    
    return prioridade;
}

// Função auxiliar para calcular carga horária por dia
function calcularCargaHorariaPorDia(professorId, turno) {
    const cargaPorDia = {};
    DIAS_SEMANA.forEach(dia => {
        cargaPorDia[dia] = aulas.filter(a => 
            a.professorId === professorId && 
            a.turno === turno && 
            a.dia === dia
        ).length;
    });
    return cargaPorDia;
}

// Função auxiliar para calcular desvio padrão
function calcularDesvioPadraoCarga(cargaPorDia) {
    const valores = Object.values(cargaPorDia);
    if (valores.length === 0) return 0;
    
    const media = valores.reduce((sum, val) => sum + val, 0) / valores.length;
    const variancia = valores.reduce((sum, val) => sum + Math.pow(val - media, 2), 0) / valores.length;
    return Math.sqrt(variancia);
}

// Função auxiliar para verificar disponibilidade por dia
function professorDisponivelNoDia(professorId, turno, dia) {
    const prof = professores.find(p => p.id === professorId);
    if (!prof || !prof.disponibilidade) return true;
    
    const dispTurno = prof.disponibilidade[turno];
    if (!dispTurno) return false;
    
    return dispTurno[dia] !== false;
}

// ===========================================
// MELHORIA 5: OTIMIZAÇÃO DE ALOCAÇÃO DE 2-TEMPOS
// ===========================================

// Função para otimizar a alocação de disciplinas com 2 tempos
function otimizarAlocacaoDoisTempos(turno) {
    const turmasTurno = turmas.filter(t => t.turno === turno);
    
    turmasTurno.forEach(turma => {
        // Identificar disciplinas com exatamente 2 tempos
        const aulasTurma = aulas.filter(a => 
            a.turno === turno && 
            a.turmaId === turma.id &&
            a.disciplina
        );
        
        const contagemDisciplinas = {};
        aulasTurma.forEach(a => {
            const disc = (a.disciplina || '').toUpperCase();
            contagemDisciplinas[disc] = (contagemDisciplinas[disc] || 0) + 1;
        });
        
        // Focar nas disciplinas com exatamente 2 tempos
        const disciplinasDoisTempos = Object.keys(contagemDisciplinas).filter(
            disc => contagemDisciplinas[disc] === 2
        );
        
        disciplinasDoisTempos.forEach(discNome => {
            const aulasDisc = aulasTurma.filter(a => 
                (a.disciplina || '').toUpperCase() === discNome
            );
            
            if (aulasDisc.length === 2) {
                const professorId = aulasDisc[0].professorId;
                if (professorId) {
                    melhorarDistribuicaoDoisTempos(aulasDisc, turno, turma.id, professorId);
                }
            }
        });
    });
}

// Função para melhorar a distribuição de 2 tempos
function melhorarDistribuicaoDoisTempos(aulasDisc, turno, turmaId, professorId) {
    const diasAtuais = Array.from(new Set(aulasDisc.map(a => a.dia)));
    
    // Se já estão no mesmo dia, verificar se estão consecutivos
    if (diasAtuais.length === 1) {
        const dia = diasAtuais[0];
        const tempos = aulasDisc.map(a => a.tempoId).sort((a, b) => a - b);
        
        // Se não estão consecutivos, tentar juntá-los
        if (tempos[1] !== tempos[0] + 1) {
            tentarJuntarTemposConsecutivos(aulasDisc, turno, dia, turmaId, professorId);
        }
    } else {
        // Se estão em dias diferentes, avaliar qual é melhor
        const melhorDia = encontrarMelhorDiaParaDoisTempos(turno, turmaId, professorId);
        if (melhorDia && melhorDia !== diasAtuais[0] && melhorDia !== diasAtuais[1]) {
            tentarMoverParaMelhorDia(aulasDisc, turno, turmaId, professorId, melhorDia);
        }
    }
}

// Encontrar o melhor dia para alocar 2 tempos consecutivos
function encontrarMelhorDiaParaDoisTempos(turno, turmaId, professorId) {
    const dias = DIAS_SEMANA;
    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
    
    let melhorDia = null;
    let melhorPontuacao = -1;
    
    dias.forEach(dia => {
        // Verificar disponibilidade do professor
        if (!professorDisponivelNoDia(professorId, turno, dia)) return;
        
        // Verificar se há tempos consecutivos disponíveis
        const temposOcupados = new Set(
            aulas.filter(a => 
                a.turno === turno && 
                a.dia === dia && 
                a.turmaId === turmaId
            ).map(a => a.tempoId)
        );
        
        let pontuacao = 0;
        
        // Procurar por pares consecutivos disponíveis
        for (let i = 0; i < temposTurno.length - 1; i++) {
            const t1 = temposTurno[i].id;
            const t2 = temposTurno[i + 1].id;
            
            if (!temposOcupados.has(t1) && !temposOcupados.has(t2)) {
                // Verificar disponibilidade do professor nos dois tempos
                if (professorDisponivelNoHorario(
                    professores.find(p => p.id === professorId), 
                    turno, dia, t1
                ) && professorDisponivelNoHorario(
                    professores.find(p => p.id === professorId), 
                    turno, dia, t2
                )) {
                    pontuacao += 10; // Par perfeito encontrado
                    
                    // Bonus se for no início ou meio do dia (evitar fim do dia)
                    if (i < 2) pontuacao += 3; // Início do dia
                    else if (i < temposTurno.length - 3) pontuacao += 2; // Meio do dia
                    
                    if (pontuacao > melhorPontuacao) {
                        melhorPontuacao = pontuacao;
                        melhorDia = dia;
                    }
                }
            }
        }
    });
    
    return melhorDia;
}

// Tentar juntar tempos em consecutivos no mesmo dia
function tentarJuntarTemposConsecutivos(aulasDisc, turno, dia, turmaId, professorId) {
    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
    const temposOcupados = new Set(
        aulas.filter(a => 
            a.turno === turno && 
            a.dia === dia && 
            a.turmaId === turmaId &&
            !aulasDisc.includes(a)
        ).map(a => a.tempoId)
    );
    
    // Procurar por pares consecutivos disponíveis
    for (let i = 0; i < temposTurno.length - 1; i++) {
        const t1 = temposTurno[i].id;
        const t2 = temposTurno[i + 1].id;
        
        if (!temposOcupados.has(t1) && !temposOcupados.has(t2)) {
            // Verificar disponibilidade do professor
            const prof = professores.find(p => p.id === professorId);
            if (prof && 
                professorDisponivelNoHorario(prof, turno, dia, t1) &&
                professorDisponivelNoHorario(prof, turno, dia, t2) &&
                !verificarConflitoProfessor(professorId, turno, dia, t1) &&
                !verificarConflitoProfessor(professorId, turno, dia, t2)
            ) {
                // Aplicar a mudança
                aulasDisc[0].dia = dia;
                aulasDisc[0].tempoId = t1;
                aulasDisc[1].dia = dia;
                aulasDisc[1].tempoId = t2;
                
                salvarLocal();
                return true;
            }
        }
    }
    
    return false;
}

// Tentar mover para o melhor dia identificado
function tentarMoverParaMelhorDia(aulasDisc, turno, turmaId, professorId, melhorDia) {
    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
    const temposOcupados = new Set(
        aulas.filter(a => 
            a.turno === turno && 
            a.dia === melhorDia && 
            a.turmaId === turmaId
        ).map(a => a.tempoId)
    );
    
    // Procurar por pares consecutivos disponíveis no melhor dia
    for (let i = 0; i < temposTurno.length - 1; i++) {
        const t1 = temposTurno[i].id;
        const t2 = temposTurno[i + 1].id;
        
        if (!temposOcupados.has(t1) && !temposOcupados.has(t2)) {
            const prof = professores.find(p => p.id === professorId);
            if (prof && 
                professorDisponivelNoHorario(prof, turno, melhorDia, t1) &&
                professorDisponivelNoHorario(prof, turno, melhorDia, t2) &&
                !verificarConflitoProfessor(professorId, turno, melhorDia, t1) &&
                !verificarConflitoProfessor(professorId, turno, melhorDia, t2)
            ) {
                // Aplicar a mudança
                aulasDisc[0].dia = melhorDia;
                aulasDisc[0].tempoId = t1;
                aulasDisc[1].dia = melhorDia;
                aulasDisc[1].tempoId = t2;
                
                salvarLocal();
                return true;
            }
        }
    }
    
    return false;
}

// Integrar a otimização no ajuste inteligente principal
function ajusteInteligenteCompleto() {
    const turno = turnoAtualGrade;
    
    if (!confirm('O ajuste inteligente completo vai otimizar horários, priorizar alocações e melhorar distribuição de 2-tempos. Continuar?')) {
        return;
    }
    
    capturarEstadoAulas();
    
    // Executar o ajuste inteligente original
    ajusteInteligente();
    
    // Aplicar otimização específica para 2-tempos
    otimizarAlocacaoDoisTempos(turno);
    
    montarGradeTurno(turnoAtualGrade);
    mostrarToast('Ajuste inteligente completo aplicado com otimizações!');
}

// Verificar conflitos
function verificarConflitoProfessor(professorId, turno, dia, tempoId, aulaId = null) {
    return aulas.some(a => 
        a.id !== aulaId &&
        a.professorId === professorId &&
        a.turno === turno &&
        a.dia === dia &&
        a.tempoId === tempoId
    );
}

function confirmarDisponibilidadeProfessorDetalhada(prof, turno, dia, tempoId) {
    if (professorDisponivelNoHorario(prof, turno, dia, tempoId)) {
        return true;
    }
    const profId = prof.id;
    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
    const diasForaDisp = new Set();
    aulas.forEach(a => {
        if (a.professorId === profId && a.turno === turno) {
            if (!professorDisponivelNoHorario(prof, a.turno, a.dia, a.tempoId)) {
                diasForaDisp.add(a.dia);
            }
        }
    });
    const diasComDisp = DIAS_SEMANA.filter(diaSigla =>
        temposTurno.some(t => professorDisponivelNoHorario(prof, turno, diaSigla, t.id))
    );
    let msg = 'Este professor não está disponível neste dia/turno/tempo.\n\n';
    if (diasComDisp.length && temposTurno.length) {
        const pad = function(text, width) {
            text = String(text);
            if (text.length >= width) return text;
            return text + ' '.repeat(width - text.length);
        };
        msg += 'Disponibilidades deste professor no turno ' + turno + ':\n';
        const colTempoWidth = 6;
        const colDiaWidth = 12;
        const headerTempo = pad('Tp', colTempoWidth);
        const headerDias = diasComDisp.map(d => {
            let nome = textoDia(d).slice(0, 3);
            if (diasForaDisp.has(d)) {
                nome = '*' + nome + '*';
            }
            return pad(nome, colDiaWidth);
        });
        msg += headerTempo + ' | ' + headerDias.join(' | ') + '\n';
        msg += ''.padEnd(headerTempo.length + 3 + headerDias.length * (colDiaWidth + 3) - 3, '-') + '\n';
        temposTurno.forEach((t, idxTempo) => {
            const rotuloTempo = pad((idxTempo + 1) + 'º', colTempoWidth);
            const cols = diasComDisp.map(diaSigla =>
                professorDisponivelNoHorario(prof, turno, diaSigla, t.id)
                    ? pad(t.inicio + '-' + t.fim, colDiaWidth)
                    : pad('', colDiaWidth)
            );
            msg += rotuloTempo + ' | ' + cols.join(' | ') + '\n';
        });
    } else {
        msg += 'Não há qualquer tempo disponível para este professor neste turno.\n';
    }
    msg += '\nDeseja continuar mesmo assim?';
    return confirm(msg);
}

function confirmarConflitoProfessorDetalhado(profId, turno, dia, tempoId, turmaId, aulaIdIgnorar = null) {
    if (!verificarConflitoProfessor(profId, turno, dia, tempoId, aulaIdIgnorar)) {
        return true;
    }
    const conflitos = aulas.filter(a =>
        a.professorId === profId &&
        a.turno === turno &&
        a.dia === dia &&
        a.tempoId === tempoId &&
        a.id !== aulaIdIgnorar
    );
    const conflitosReais = conflitos.filter(a => a.turmaId !== turmaId);
    if (!conflitosReais.length) {
        return true;
    }
    let msg = 'Este professor já tem aula neste mesmo horário:\n\n';
    conflitosReais.forEach(a => {
        const turmaConf = turmas.find(t => t.id === a.turmaId);
        const nomeTurma = turmaConf ? turmaConf.nome : (a.turmaId || '');
        const nomeDisc = a.disciplina || '';
        msg += '- ' + nomeTurma + (nomeDisc ? ' — ' + nomeDisc : '') + '\n';
    });
    msg += '\nDeseja continuar mesmo assim?';
    return confirm(msg);
}

function verificarConflitosTurno() {
    const turno = turnoAtualGrade;
    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
    if (!temposTurno.length) {
        mostrarToast('Não há tempos configurados para este turno.', 'warning');
        return;
    }
    const aulasTurno = aulas.filter(a => a.turno === turno);
    if (!aulasTurno.length) {
        mostrarToast('Não há aulas lançadas neste turno para verificar conflitos.', 'warning');
        return;
    }

    const conflitos = [];
    const mapaProfDiaTempo = {};

    aulasTurno.forEach(aula => {
        const dia = aula.dia;
        const tempoId = aula.tempoId;
        const turmaId = aula.turmaId;
        const profId = aula.professorId || null;

        const tempoValido = temposTurno.some(t => t.id === tempoId);
        if (!tempoValido) {
            conflitos.push({
                tipo: 'TEMPO_INVALIDO',
                aula
            });
        }

        if (profId) {
            const prof = professores.find(p => p.id === profId);
            if (prof && !professorDisponivelNoHorario(prof, turno, dia, tempoId)) {
                conflitos.push({
                    tipo: 'INDISPONIBILIDADE_PROFESSOR',
                    aula
                });
            }

            const chaveProf = `${profId}|${dia}|${tempoId}`;
            if (!mapaProfDiaTempo[chaveProf]) {
                mapaProfDiaTempo[chaveProf] = [];
            }
            mapaProfDiaTempo[chaveProf].push(aula);
        }
    });

    Object.keys(mapaProfDiaTempo).forEach(chave => {
        const lista = mapaProfDiaTempo[chave];
        if (lista.length > 1) {
            conflitos.push({
                tipo: 'CONFLITO_PROFESSOR',
                aulas: lista
            });
        }
    });

    if (!conflitos.length) {
        mostrarToast('Nenhum conflito encontrado neste turno.', 'success');
        return;
    }

    const descricoes = [];
    let conflitosProfessor = 0;
    let conflitosDisponibilidade = 0;
    let conflitosTempo = 0;

    conflitos.forEach(c => {
        if (c.tipo === 'CONFLITO_PROFESSOR') {
            conflitosProfessor++;
            const aulasLista = c.aulas || [];
            aulasLista.forEach(a => {
                const prof = a.professorId ? professores.find(p => p.id === a.professorId) : null;
                const turma = turmas.find(t => t.id === a.turmaId);
                const tempo = getTempos(a.turno).find(t => t.id === a.tempoId);
                descricoes.push(`Professor em conflito: ${prof ? prof.nome : 'sem nome'} - ${a.disciplina || ''} - ${turma ? turma.nome : ''} - ${a.dia} ${textoTurno(a.turno)} ${tempo ? tempo.etiqueta : 'Tempo ' + a.tempoId}`);
            });
        } else if (c.tipo === 'INDISPONIBILIDADE_PROFESSOR') {
            conflitosDisponibilidade++;
            const a = c.aula;
            const prof = a.professorId ? professores.find(p => p.id === a.professorId) : null;
            const turma = turmas.find(t => t.id === a.turmaId);
            const tempo = getTempos(a.turno).find(t => t.id === a.tempoId);
            descricoes.push(`Professor indisponível: ${prof ? prof.nome : 'sem nome'} - ${a.disciplina || ''} - ${turma ? turma.nome : ''} - ${a.dia} ${textoTurno(a.turno)} ${tempo ? tempo.etiqueta : 'Tempo ' + a.tempoId}`);
        } else if (c.tipo === 'TEMPO_INVALIDO') {
            conflitosTempo++;
            const a = c.aula;
            const turma = turmas.find(t => t.id === a.turmaId);
            descricoes.push(`Tempo inválido na configuração: ${a.disciplina || ''} - ${turma ? turma.nome : ''} - ${a.dia} ${textoTurno(a.turno)} Tempo ${a.tempoId}`);
        }
    });

    const resumo = [];
    if (conflitosProfessor) resumo.push(`${conflitosProfessor} conflito(s) de professor em duas turmas no mesmo horário`);
    if (conflitosDisponibilidade) resumo.push(`${conflitosDisponibilidade} aula(s) em horário de indisponibilidade do professor`);
    if (conflitosTempo) resumo.push(`${conflitosTempo} aula(s) usando tempo inexistente neste turno`);

    const msgResumo = resumo.join(' | ');
    const msgDetalhada = descricoes.slice(0, 10).join('\n');
    let msg = `Conflitos encontrados neste turno:\n${msgResumo}`;
    if (msgDetalhada) {
        msg += `\n\nDetalhes (até 10 registros):\n${msgDetalhada}`;
    }
    alert(msg);
}

// Nova função para detecção proativa de conflitos com informações detalhadas
function verificarConflitosPotenciais(professorId, turno, dia, tempoId, aulaId = null) {
    const conflitos = [];
    
    // 1. Verificar conflitos com outras aulas do mesmo professor
    const conflitosProfessor = aulas.filter(a => 
        a.id !== aulaId &&
        a.professorId === professorId &&
        a.turno === turno &&
        a.dia === dia &&
        a.tempoId === tempoId
    );
    
    if (conflitosProfessor.length > 0) {
        conflitosProfessor.forEach(aulaConflito => {
            const turma = turmas.find(t => t.id === aulaConflito.turmaId);
            const nomeTurma = turma ? turma.nome : 'Turma desconhecida';
            conflitos.push({
                tipo: 'PROFESSOR',
                mensagem: `Conflito com aula em ${nomeTurma} - ${aulaConflito.disciplina || 'Sem disciplina'}`,
                aula: aulaConflito
            });
        });
    }
    
    // 2. Verificar disponibilidade do professor
    const prof = professores.find(p => p.id === professorId);
    if (prof) {
        if (!professorDisponivelNoHorario(prof, turno, dia, tempoId)) {
            conflitos.push({
                tipo: 'DISPONIBILIDADE',
                mensagem: 'Professor não disponível neste horário',
                detalhes: `Verifique a disponibilidade de ${prof.nome} nas configurações`
            });
        }
    }
    
    // 3. Verificar limite de tempos por dia para o professor
    const aulasProfDia = aulas.filter(a => 
        a.professorId === professorId &&
        a.turno === turno &&
        a.dia === dia &&
        a.id !== aulaId
    );
    
    if (aulasProfDia.length >= 4) {
        conflitos.push({
            tipo: 'LIMITE_TEMPOS',
            mensagem: 'Limite de 4 tempos por dia atingido para este professor',
            detalhes: `Professor já tem ${aulasProfDia.length} aulas neste dia`
        });
    }
    
    // 4. Verificar se a turma já tem aula neste horário
    if (aulaId) {
        const aulaAtual = aulas.find(a => a.id === aulaId);
        if (aulaAtual) {
            const conflitoTurma = aulas.some(a => 
                a.id !== aulaId &&
                a.turmaId === aulaAtual.turmaId &&
                a.turno === turno &&
                a.dia === dia &&
                a.tempoId === tempoId
            );
            
            if (conflitoTurma) {
                conflitos.push({
                    tipo: 'TURMA',
                    mensagem: 'Turma já possui aula neste horário',
                    detalhes: 'Cada turma só pode ter uma aula por tempo'
                });
            }
        }
    }
    
    return conflitos;
}

// Função para exibir conflitos de forma amigável
function exibirConflitosDetalhados(conflitos) {
    if (conflitos.length === 0) return null;
    
    let mensagem = 'Conflitos detectados:\n\n';
    
    conflitos.forEach((conflito, index) => {
        mensagem += `${index + 1}. ${conflito.mensagem}\n`;
        if (conflito.detalhes) {
            mensagem += `   → ${conflito.detalhes}\n`;
        }
        mensagem += '\n';
    });
    
    mensagem += 'Deseja continuar mesmo assim?';
    return confirm(mensagem);
}

function verificarConflitosProfessoresGlobais() {
    if (!aulas || !aulas.length) {
        alert('Não há aulas lançadas para verificar.');
        return;
    }

    const conflitosPorChave = new Map();

    aulas.forEach(aula => {
        if (!aula.professorId) return;
        const chave = `${aula.professorId}||${aula.turno}||${aula.dia}||${aula.tempoId}`;
        if (!conflitosPorChave.has(chave)) {
            conflitosPorChave.set(chave, []);
        }
        conflitosPorChave.get(chave).push(aula);
    });

    const linhas = [];

    conflitosPorChave.forEach((lista, chave) => {
        if (lista.length <= 1) return;
        const [profId, turno, dia, tempoId] = chave.split('||');
        const prof = professores.find(p => p.id === profId);
        const nomeProf = prof && prof.nome ? prof.nome : profId;
        const tempo = getTempos(turno).find(t => t.id === Number(tempoId));
        const etiquetaTempo = tempo ? tempo.etiqueta : `Tempo ${tempoId}`;
        const turmasConflito = lista.map(a => {
            const turma = turmas.find(t => t.id === a.turmaId);
            return turma && turma.nome ? turma.nome : a.turmaId;
        });
        linhas.push(
            `${nomeProf} — ${textoDia(dia)} / ${textoTurno(turno)} / ${etiquetaTempo}: ` +
            turmasConflito.join(', ')
        );
    });

    if (!linhas.length) {
        alert('Nenhum conflito de professor encontrado entre turmas.');
        return;
    }

    const mensagem =
        'Foram encontrados conflitos de professor (mesmo horário em mais de uma turma):\n\n' +
        linhas.join('\n');
    alert(mensagem);
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
        const turnosDia = disp[dia];
        if (!turnosDia || !Array.isArray(turnosDia) || turnosDia.length === 0) {
            return false;
        }
        const okDiaTurno = turnosDia.includes(turno);
        return okDiaTurno && okTempo;
    } else {
        const okDia = diasF.length === 0 || diasF.includes(dia);
        const okTurno = turnosF.length === 0 || turnosF.includes(turno);
        return okDia && okTurno && okTempo;
    }
}

function diagnosticoDisponibilidadeProfessorTurno(professorId, turno, quantidadeSolicitada) {
    const prof = professores.find(p => p.id === professorId);
    if (!prof) return null;
    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
    if (!temposTurno.length) return null;
    const dias = DIAS_SEMANA || [];
    let capacidade = 0;
    let ocupados = 0;
    dias.forEach(dia => {
        temposTurno.forEach(t => {
            if (!professorDisponivelNoHorario(prof, turno, dia, t.id)) return;
            capacidade++;
            const ja = aulas.some(a =>
                a.professorId === professorId &&
                a.turno === turno &&
                a.dia === dia &&
                a.tempoId === t.id
            );
            if (ja) {
                ocupados++;
            }
        });
    });
    const livres = capacidade - ocupados;
    const faltam = quantidadeSolicitada > livres ? (quantidadeSolicitada - livres) : 0;
    return { capacidade, ocupados, livres, faltam, professor: prof };
}

function gerarRelatorioFalhasInsercaoRapidaPlanilha(gruposFalha, ignoradas, gruposAplicados) {
    if (!gruposFalha || !gruposFalha.length) return;
    if (!window.jspdf || !window.jspdf.jsPDF) {
        return;
    }
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const titulo = 'Relatório de falhas da inserção rápida via planilha';
    doc.setFontSize(12);
    doc.setFont(undefined, 'bold');
    doc.text(titulo, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });
    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');
    const resumo = [
        `Data: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`,
        `Solicitações aplicadas: ${gruposAplicados}`,
        `Linhas ignoradas na leitura do arquivo: ${ignoradas}`,
        `Total de grupos com falha de alocação: ${gruposFalha.length}`
    ];
    let y = 25;
    resumo.forEach(l => {
        doc.text(l, 10, y);
        y += 5;
    });
    const head = [['Turma', 'Turno', 'Disciplina', 'Professor', 'Qtd. não alocada', 'Situação / Sugestão']];
    const body = [];
    gruposFalha.forEach(g => {
        const turma = turmas.find(t => t.id === g.turmaId);
        const nomeTurma = turma && turma.nome ? turma.nome : String(g.turmaId);
        const turnoTxt = textoTurno(g.turno);
        const prof = g.professorId ? professores.find(p => p.id === g.professorId) : null;
        const nomeProf = prof && prof.nome ? prof.nome : '-';
        const qtdNaoAlocada = g.quantidade || 0;
        let situacao = 'Não foi possível alocar as aulas solicitadas.';
        if (g.professorId) {
            const diag = diagnosticoDisponibilidadeProfessorTurno(g.professorId, g.turno, g.quantidade);
            if (diag) {
                if (diag.capacidade === 0) {
                    situacao += ` O professor não possui qualquer disponibilidade registrada neste turno; seriam necessários pelo menos ${g.quantidade} tempo(s) livres neste turno.`;
                } else if (diag.faltam > 0) {
                    situacao += ` Disponibilidade insuficiente: com a configuração atual o professor tem ${diag.livres} tempo(s) livre(s) no turno; para este pedido seriam necessários pelo menos ${diag.faltam} tempo(s) livres adicionais.`;
                } else {
                    situacao += ' Verifique conflitos com outras turmas ou limites de tempos por dia.';
                }
            }
        }
        situacao += ` Quantidade de tempos deste pedido que não pôde ser alocada: ${qtdNaoAlocada}.`;
        body.push([nomeTurma, turnoTxt, g.disciplinaUpper || g.disciplina || '-', nomeProf, String(qtdNaoAlocada), situacao]);
    });
    doc.autoTable({
        head,
        body,
        startY: y + 2,
        theme: 'grid',
        styles: {
            fontSize: 8,
            cellPadding: 2,
            valign: 'top'
        },
        headStyles: {
            fillColor: [230, 126, 34],
            textColor: 255
        },
        columnStyles: {
            0: { cellWidth: 22 },
            1: { cellWidth: 18 },
            2: { cellWidth: 30 },
            3: { cellWidth: 30 },
            4: { cellWidth: 16 },
            5: { cellWidth: 72 }
        }
    });
    const nomeArquivo = `falhas_insercao_planilha_${new Date().toISOString().slice(0,10)}.pdf`;
    doc.save(nomeArquivo);
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
    if (groupId === 'chips-turno-turma') {
        const selCurso = document.getElementById('cursoTurma');
        const cursoVal = selCurso ? selCurso.value : '';
        if (cursoVal === CURSO_TEMPO_INTEGRAL) {
            btn.classList.toggle('selecionado');
            const selecionados = Array.from(group.querySelectorAll('.chip.selecionado'))
                .map(c => c.getAttribute('data-value'));
            const hidden = document.getElementById(hiddenId);
            if (hidden) hidden.value = selecionados.join(',');
            return;
        }
    }
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
    atualizarIndicadorDesfazerTurno();
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
    atualizarIndicadorDesfazerTurno();
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

function toggleInsercaoRapida() {
    const box = document.getElementById('editorQuickInsert');
    if (!box) return;
    const btn = document.getElementById('btnToggleInsercaoRapida');
    const visivel = box.style.display !== 'none';
    const novoVisivel = !visivel;
    box.style.display = novoVisivel ? 'block' : 'none';
    if (btn) {
        if (novoVisivel) {
            btn.className = 'btn-primary';
            btn.textContent = '💡 Inserção rápida (ativa)';
        } else {
            btn.className = 'btn-secondary';
            btn.textContent = '💡 Inserção rápida';
        }
    }
}

function baixarModeloExcelInsercaoRapida() {
    if (!turmas || !turmas.length) {
        mostrarToast('Não há turmas cadastradas para gerar o modelo.', 'warning');
        return;
    }
    const cabecalho = ['TURMA', 'TURNO', 'PROFESSOR', 'DISCIPLINA', 'QTDE_TEMPOS'];
    const linhas = [cabecalho.join(';')];
    const ordenadas = [...turmas].sort((a, b) => {
        const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
        const ta = ordemTurno[a.turno] || 99;
        const tb = ordemTurno[b.turno] || 99;
        if (ta !== tb) return ta - tb;
        return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
    });
    ordenadas.forEach(t => {
        const nomeTurma = (t.nome || '').replace(/;/g, ',');
        const turno = t.turno || '';
        for (let i = 0; i < 6; i++) {
            linhas.push([nomeTurma, turno, '', '', ''].join(';'));
        }
    });
    const conteudo = linhas.join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'modelo_insercao_rapida.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function importarExcelInsercaoRapida(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        let text = '';
        const result = e.target.result;
        if (result instanceof ArrayBuffer) {
            try {
                const decUtf8 = new TextDecoder('utf-8', { fatal: false });
                text = decUtf8.decode(result);
                if (text.indexOf('\uFFFD') !== -1) {
                    const decLatin = new TextDecoder('iso-8859-1', { fatal: false });
                    text = decLatin.decode(result);
                }
            } catch (err) {
                const decLatin = new TextDecoder('iso-8859-1', { fatal: false });
                text = decLatin.decode(result);
            }
        } else {
            text = result || '';
        }
        if (text.charCodeAt(0) === 0xFEFF) {
            text = text.substring(1);
        }
        const linhas = text.split(/\r?\n/).map(l => l.trim()).filter(l => l);
        if (linhas.length <= 1) {
            mostrarToast('Arquivo vazio ou sem dados além do cabeçalho.', 'warning');
            return;
        }

        let ignoradas = 0;
        let linhasValidas = 0;
        const grupos = new Map();

        for (let i = 1; i < linhas.length; i++) {
            const linha = linhas[i];
            const cols = linha.split(';');
            if (cols.length < 5) {
                ignoradas++;
                continue;
            }
            const turmaNome = (cols[0] || '').trim();
            const turnoPlanilha = (cols[1] || '').trim();
            const professorNome = (cols[2] || '').trim();
            const disciplina = (cols[3] || '').trim();
            const qtdStr = (cols[4] || '').trim();
            if (!turmaNome || !disciplina || !qtdStr) {
                ignoradas++;
                continue;
            }
            if (!Object.prototype.hasOwnProperty.call(apelidosDisciplinas, disciplina)) {
                const sugestao = sugerirApelidoDisciplina(disciplina);
                if (sugestao) {
                    apelidosDisciplinas[disciplina] = sugestao;
                }
            }
            const turma = turmas.find(t =>
                (t.nome || '').toUpperCase() === turmaNome.toUpperCase()
            );
            if (!turma) {
                ignoradas++;
                continue;
            }
            let quantidade = parseInt(qtdStr, 10);
            if (isNaN(quantidade) || quantidade < 1) {
                ignoradas++;
                continue;
            }
            let professorId = null;
            if (professorNome) {
                const prof = professores.find(p =>
                    (p.nome || '').toUpperCase() === professorNome.toUpperCase()
                );
                if (prof) {
                    professorId = prof.id;
                }
            }

            const disciplinaUpper = (disciplina || '').toUpperCase();
            const chaveGrupo = `${turma.id}|${disciplinaUpper}|${professorId || ''}`;
            if (!grupos.has(chaveGrupo)) {
                grupos.set(chaveGrupo, {
                    turno: turma.turno,
                    turmaId: turma.id,
                    disciplina,
                    disciplinaUpper,
                    professorId,
                    quantidade: 0
                });
            }
            const grupo = grupos.get(chaveGrupo);
            grupo.quantidade += quantidade;
            linhasValidas++;
        }

        let gruposAplicados = 0;
        const gruposFalha = [];
        const listaGrupos = Array.from(grupos.values());
        listaGrupos.sort((a, b) => {
            const pa = prioridadeProfessorAlocacao(a.professorId || null);
            const pb = prioridadeProfessorAlocacao(b.professorId || null);
            if (pa !== pb) return pa - pb;
            return (b.quantidade || 0) - (a.quantidade || 0);
        });
        listaGrupos.forEach(grupo => {
            if (grupo.quantidade > 0) {
                const ok = ajusteInteligenteInsercaoRapida(
                    grupo.turno,
                    grupo.turmaId,
                    grupo.disciplina,
                    grupo.professorId,
                    grupo.quantidade
                );
                if (ok) {
                    harmonizarFamiliasPortMatInsercaoRapida(grupo.turno, grupo.turmaId, grupo.professorId || null);
                    gruposAplicados++;
                } else {
                    gruposFalha.push(grupo);
                }
            }
        });

        if (gruposAplicados === 0) {
            gerarRelatorioFalhasInsercaoRapidaPlanilha(gruposFalha, ignoradas, gruposAplicados);
            const msg = 'Nenhum pedido da planilha pôde ser aplicado. Foi gerado um PDF com os problemas encontrados e sugestões de disponibilidade por professor.';
            mostrarToast(msg, 'warning');
        } else {
            let msg = `Inserção rápida via planilha concluída. Solicitações aplicadas: ${gruposAplicados}. Linhas ignoradas: ${ignoradas}.`;
            if (gruposFalha.length) {
                gerarRelatorioFalhasInsercaoRapidaPlanilha(gruposFalha, ignoradas, gruposAplicados);
                msg += ' Alguns pedidos não puderam ser alocados. Foi gerado um PDF com detalhes dos problemas e sugestões por professor.';
            }
            mostrarToast(msg, gruposFalha.length ? 'warning' : 'success');
        }
        event.target.value = '';
    };
    reader.readAsArrayBuffer(file);
}

function obterDadosInsercaoRapida() {
    const turnoSel = document.getElementById('rapidoTurno');
    const turmaSel = document.getElementById('rapidoTurma');
    const profSel = document.getElementById('rapidoProfessor');
    const discInput = document.getElementById('rapidoDisciplina');
    const qtdInput = document.getElementById('rapidoQuantidade');
    let turmaIds = [];
    if (turmaSel) {
        const opts = Array.from(turmaSel.options || []);
        turmaIds = opts.filter(o => o.selected && o.value).map(o => o.value);
    }
    return {
        turno: turnoSel ? turnoSel.value : '',
        turmaId: turmaIds.length === 1 ? turmaIds[0] : '',
        turmaIds,
        professorId: profSel ? profSel.value : '',
        disciplina: discInput ? discInput.value.trim() : '',
        quantidade: qtdInput ? qtdInput.value : '1'
    };
}

function tentarAplicarInsercaoRapida(turno, dia, turmaId, tempoIds, disciplina, professorId) {
    if (!disciplina) {
        return false;
    }
    const temposTurno = getTempos(turno);
    let realocacoes = [];
    if (professorId) {
        const prof = professores.find(p => p.id === professorId);
        if (!prof) {
            return false;
        }
        const indisponiveis = [];
        tempoIds.forEach(id => {
            if (!professorDisponivelNoHorario(prof, turno, dia, id)) {
                const tempo = temposTurno.find(t => t.id === id);
                indisponiveis.push(tempo ? tempo.etiqueta : id);
            }
        });
        if (indisponiveis.length > 0) {
            return false;
        }
        const conflitos = [];
        tempoIds.forEach(id => {
            if (verificarConflitoProfessor(professorId, turno, dia, id)) {
                const tempo = temposTurno.find(t => t.id === id);
                conflitos.push(tempo ? tempo.etiqueta : id);
            }
        });
        if (conflitos.length > 0) {
            return false;
        }
    }
    const conflitosTurma = [];
    tempoIds.forEach(id => {
        const ja = aulas.find(a =>
            a.turno === turno &&
            a.dia === dia &&
            a.turmaId === turmaId &&
            a.tempoId === id
        );
        if (ja) conflitosTurma.push(id);
    });
    if (conflitosTurma.length) {
        const dias = diasInsercaoRapida();
        const slotsProibidos = new Set();
        tempoIds.forEach(id => slotsProibidos.add(`${dia}|${id}`));
        const aulasBloqueando = [];
        tempoIds.forEach(id => {
            const aula = aulas.find(a =>
                a.turno === turno &&
                a.dia === dia &&
                a.turmaId === turmaId &&
                a.tempoId === id
            );
            if (aula && !aulasBloqueando.some(x => x.id === aula.id)) {
                aulasBloqueando.push(aula);
            }
        });
        for (let i = 0; i < aulasBloqueando.length; i++) {
            const aula = aulasBloqueando[i];
            const sugestao = sugerirNovoSlotParaAulaRealloc(aula, turno, dias, temposTurno, slotsProibidos);
            if (!sugestao) {
                return false;
            }
            realocacoes.push({
                aulaId: aula.id,
                novoDia: sugestao.dia,
                novoTempoId: sugestao.tempoId
            });
            slotsProibidos.add(`${sugestao.dia}|${sugestao.tempoId}`);
        }
    }
    const grupoDisc = grupoDisciplinaPortMat(disciplina);
    const disciplinaSujeitaRegra = !!grupoDisc;
    if (disciplinaSujeitaRegra) {
        const aulasMesmoDiaMesmaDisciplina = aulas.filter(a =>
            a.turmaId === turmaId &&
            a.dia === dia &&
            grupoDisciplinaPortMat(a.disciplina) === grupoDisc &&
            !tempoIds.includes(a.tempoId)
        );
        const totalFinalDiscDia = aulasMesmoDiaMesmaDisciplina.length + tempoIds.length;
        const limite = 2;
        if (totalFinalDiscDia > limite) {
            if (!professorId) {
                return false;
            }
            const profs = new Set(aulasMesmoDiaMesmaDisciplina.map(a => a.professorId || null));
            profs.add(professorId);
            if (profs.size > 1) {
                return false;
            }
        }
    }
    capturarEstadoAulas();
    realocacoes.forEach(r => {
        const aula = aulas.find(a => a.id === r.aulaId);
        if (aula) {
            aula.dia = r.novoDia;
            aula.tempoId = r.novoTempoId;
        }
    });
    aulas = aulas.filter(a =>
        !(a.turno === turno &&
          a.dia === dia &&
          a.turmaId === turmaId &&
          tempoIds.includes(a.tempoId))
    );
    tempoIds.forEach(id => {
        aulas.push({
            id: gerarId(),
            turno,
            dia,
            turmaId,
            tempoId: id,
            disciplina,
            professorId
        });
    });
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast(`Aula(s) lançada(s) pela inserção rápida. (${tempoIds.length} tempo(s))`);
    return true;
}

function sugerirNovoSlotParaAulaRealloc(aula, turno, dias, temposTurno, slotsProibidos) {
    const profId = aula.professorId || null;
    let prof = null;
    if (profId) {
        prof = professores.find(p => p.id === profId);
        if (!prof) return null;
    }
    const grupoDisc = grupoDisciplinaPortMat(aula.disciplina);
    const disciplinaSujeitaRegra = !!grupoDisc;
    for (let d = 0; d < dias.length; d++) {
        const dia = dias[d];
        for (let i = 0; i < temposTurno.length; i++) {
            const t = temposTurno[i];
            if (t.intervalo) continue;
            const chave = `${dia}|${t.id}`;
            if (slotsProibidos.has(chave)) continue;
            const ocupadoTurma = aulas.some(x =>
                x.id !== aula.id &&
                x.turno === turno &&
                x.dia === dia &&
                x.turmaId === aula.turmaId &&
                x.tempoId === t.id
            );
            if (ocupadoTurma) continue;
            if (prof) {
                if (!professorDisponivelNoHorario(prof, turno, dia, t.id)) continue;
                if (verificarConflitoProfessor(profId, turno, dia, t.id)) continue;
                const maxPorTurmaDia = 4;
                const aulasProfTurmaDia = aulas.filter(a =>
                    a.id !== aula.id &&
                    a.turno === turno &&
                    a.dia === dia &&
                    a.turmaId === aula.turmaId &&
                    a.professorId === profId
                );
                const totalFinalProfTurmaDia = aulasProfTurmaDia.length + 1;
                if (totalFinalProfTurmaDia > maxPorTurmaDia) continue;
            }
        if (disciplinaSujeitaRegra) {
                const aulasDiaMesmaDisc = aulas.filter(a =>
                    a.id !== aula.id &&
                    a.turno === turno &&
                    a.dia === dia &&
                    a.turmaId === aula.turmaId &&
                    grupoDisciplinaPortMat(a.disciplina) === grupoDisc
                );
                const total = aulasDiaMesmaDisc.length + 1;
                if (total > 2) continue;
            }
            return { dia, tempoId: t.id };
        }
    }
    return null;
}

function encontrarBlocoDoisTemposDiaPortMat(turno, dia, turmaId, disciplina, professorId, temposTurno, temposValidos, dias) {
    const grupoDisc = grupoDisciplinaPortMat(disciplina);
    if (!grupoDisc) return null;
    let prof = null;
    if (professorId) {
        prof = professores.find(p => p.id === professorId);
        if (!prof) return null;
    }
    const aulasDiaDisc = aulas.filter(a =>
        a.turmaId === turmaId &&
        a.dia === dia &&
        grupoDisciplinaPortMat(a.disciplina) === grupoDisc
    );
    if (aulasDiaDisc.length > 1) {
        return null;
    }
    for (let i = 0; i < temposValidos.length - 1; i++) {
        const t1 = temposValidos[i];
        const t2 = temposValidos[i + 1];
        const idx1 = temposTurno.findIndex(tt => tt.id === t1.id);
        const idx2 = temposTurno.findIndex(tt => tt.id === t2.id);
        if (idx1 === -1 || idx2 === -1) continue;
        if (idx2 <= idx1) continue;
        if (prof) {
            if (!professorDisponivelNoHorario(prof, turno, dia, t1.id)) continue;
            if (!professorDisponivelNoHorario(prof, turno, dia, t2.id)) continue;
            if (verificarConflitoProfessor(professorId, turno, dia, t1.id)) continue;
            if (verificarConflitoProfessor(professorId, turno, dia, t2.id)) continue;
        }
        const slotsProibidos = new Set();
        slotsProibidos.add(`${dia}|${t1.id}`);
        slotsProibidos.add(`${dia}|${t2.id}`);
        const aulasBloqueando = [];
        [t1.id, t2.id].forEach(idTempo => {
            const aTurma = aulas.find(a =>
                a.turno === turno &&
                a.dia === dia &&
                a.turmaId === turmaId &&
                a.tempoId === idTempo
            );
            if (aTurma) aulasBloqueando.push(aTurma);
        });
        if (!aulasBloqueando.length) {
            const totalFinal = aulasDiaDisc.length + 2;
            if (totalFinal > 2) {
                continue;
            }
            return {
                dia,
                tempos: [t1.id, t2.id],
                realocacoes: []
            };
        }
        const realocacoes = [];
        let falhou = false;
        for (let k = 0; k < aulasBloqueando.length; k++) {
            const aula = aulasBloqueando[k];
            const sugestao = sugerirNovoSlotParaAulaRealloc(aula, turno, dias, temposTurno, slotsProibidos);
            if (!sugestao) {
                falhou = true;
                break;
            }
            realocacoes.push({
                aulaId: aula.id,
                novoDia: sugestao.dia,
                novoTempoId: sugestao.tempoId
            });
            slotsProibidos.add(`${sugestao.dia}|${sugestao.tempoId}`);
        }
        if (falhou) continue;
        const totalFinal = aulasDiaDisc.length + 2;
        if (totalFinal > 2) {
            continue;
        }
        return {
            dia,
            tempos: [t1.id, t2.id],
            realocacoes
        };
    }
    return null;
}

function aplicarInsercaoRapidaEspecialPortMat(turno, turmaId, disciplina, professorId, quantidade) {
    const nomeDisciplinaNorm = (disciplina || '').toUpperCase();
    const temposTurno = getTempos(turno);
    const temposValidos = temposTurno.filter(t => !t.intervalo);
    if (!temposValidos.length) {
        mostrarToast('Não há tempos configurados para este turno.', 'warning');
        return false;
    }
    const dias = diasInsercaoRapida();
    const blocosEscolhidos = [];
    const realocacoesAcumuladas = [];
    for (let d = 0; d < dias.length; d++) {
        const dia = dias[d];
        const resultado = encontrarBlocoDoisTemposDiaPortMat(
            turno,
            dia,
            turmaId,
            disciplina,
            professorId,
            temposTurno,
            temposValidos,
            dias
        );
        if (resultado) {
            blocosEscolhidos.push({ dia: resultado.dia, tempos: resultado.tempos });
            realocacoesAcumuladas.push(...resultado.realocacoes);
        }
        if (blocosEscolhidos.length === 2) break;
    }
    if (blocosEscolhidos.length < 2) {
        return false;
    }
    capturarEstadoAulas();
    realocacoesAcumuladas.forEach(r => {
        const aula = aulas.find(a => a.id === r.aulaId);
        if (aula) {
            aula.dia = r.novoDia;
            aula.tempoId = r.novoTempoId;
        }
    });
    blocosEscolhidos.forEach(b => {
        b.tempos.forEach(idTempo => {
            aulas = aulas.filter(a =>
                !(a.turno === turno &&
                  a.dia === b.dia &&
                  a.turmaId === turmaId &&
                  a.tempoId === idTempo)
            );
        });
    });
    blocosEscolhidos.forEach(b => {
        b.tempos.forEach(idTempo => {
            aulas.push({
                id: gerarId(),
                turno,
                dia: b.dia,
                turmaId,
                tempoId: idTempo,
                disciplina,
                professorId: professorId || null
            });
        });
    });
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    const b1 = blocosEscolhidos[0];
    const b2 = blocosEscolhidos[1];
    mostrarToast(`Foram lançados 4 tempos (2 em ${textoDia(b1.dia)}, 2 em ${textoDia(b2.dia)}).`);
    return true;
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
    const selProfRapido = document.getElementById('rapidoProfessor');

    if (!corpo) return;
    corpo.innerHTML = '';
    if (selProf) selProf.innerHTML = '<option value="">Selecione...</option>';
    if (selProfRel) selProfRel.innerHTML = '<option value="">Selecione...</option>';
    if (selProfModal) selProfModal.innerHTML = '<option value="">Selecione...</option>';
    if (selProfRapido) selProfRapido.innerHTML = '<option value="">Selecione...</option>';

    const ordenados = [...professores].sort((a, b) => {
        const pesoBase = (x) => x.baseTipo === 'COMUM' ? 2 : 1;
        const pa = pesoBase(a);
        const pb = pesoBase(b);
        if (pa !== pb) return pa - pb;
        return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
    });

    ordenados.forEach(p => {
        // tabela
        const tr = document.createElement('tr');

        const tdNome = document.createElement('td');
        tdNome.textContent = p.nome;
        tr.appendChild(tdNome);

        const tdDisc = document.createElement('td');
        tdDisc.textContent = p.disciplinas.join(', ');
        tr.appendChild(tdDisc);

        const tdBase = document.createElement('td');
        const base = p.baseTipo === 'COMUM' ? 'Base Comum' : 'Base Técnica';
        tdBase.textContent = base;
        tr.appendChild(tdBase);

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
        tdEdit.colSpan = 5;
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

        if (selProfRapido) {
            const opt4 = document.createElement('option');
            opt4.value = p.id;
            opt4.textContent = p.nome;
            selProfRapido.appendChild(opt4);
        }
    });
}

function salvarProfessor(e) {
    e.preventDefault();
    const id = document.getElementById('professorId').value;
    const nome = document.getElementById('nomeProfessor').value.trim();
    const disciplinasStr = document.getElementById('disciplinasProfessor').value;
    const apelidoProfInput = document.getElementById('apelidoProfessorDisciplina');
    const cor = document.getElementById('corProfessor').value || '#3498db';
    const baseProfessor = document.getElementById('baseProfessor').value || 'TECNICA';
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

    if (apelidoProfInput) {
        const apelido = apelidoProfInput.value.trim();
        const primeiraDisciplina = disciplinas.length ? disciplinas[0] : '';
        if (primeiraDisciplina && apelido) {
            apelidosDisciplinas[primeiraDisciplina] = apelido;
        }
    }

    if (id) {
        const p = professores.find(x => x.id === id);
        if (p) {
            const novoProf = {
                ...p,
                nome,
                disciplinas,
                cor,
                baseTipo: baseProfessor,
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
            p.baseTipo = novoProf.baseTipo;
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
            baseTipo: baseProfessor,
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
    const form = document.getElementById('formProfessor');
    if (form) form.reset();
    document.getElementById('professorId').value = '';
    if (apelidoProfInput) {
        apelidoProfInput.value = '';
        delete apelidoProfInput.dataset.manual;
    }
}

function editarProfessor(id) {
    const p = professores.find(x => x.id === id);
    if (!p) return;
    document.getElementById('professorId').value = p.id;
    document.getElementById('nomeProfessor').value = p.nome;
    document.getElementById('disciplinasProfessor').value = p.disciplinas.join(', ');
    const apelidoProfInput = document.getElementById('apelidoProfessorDisciplina');
    if (apelidoProfInput) {
        const primeiraDisciplina = p.disciplinas && p.disciplinas.length ? p.disciplinas[0] : '';
        if (primeiraDisciplina && apelidosDisciplinas[primeiraDisciplina]) {
            apelidoProfInput.value = apelidosDisciplinas[primeiraDisciplina];
            apelidoProfInput.dataset.manual = '1';
        } else {
            apelidoProfInput.value = '';
            delete apelidoProfInput.dataset.manual;
        }
    }
    document.getElementById('corProfessor').value = p.cor;
    const baseValor = p.baseTipo || 'TECNICA';
    const hiddenBase = document.getElementById('baseProfessor');
    if (hiddenBase) hiddenBase.value = baseValor;
    const groupBase = document.getElementById('chips-base-prof');
    if (groupBase) {
        const btnSelecionado = Array.from(groupBase.querySelectorAll('.chip'))
            .find(b => b.getAttribute('data-value') === baseValor);
        groupBase.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
        if (btnSelecionado) {
            btnSelecionado.classList.add('selecionado');
        }
    }
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
    const optNa = document.createElement('option');
    optNa.value = 'Não se aplica';
    optNa.textContent = 'Não se aplica';
    sel.appendChild(optNa);
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
    const optNa = document.createElement('option');
    optNa.value = 'Não se aplica';
    optNa.textContent = 'Não se aplica';
    sel.appendChild(optNa);
    cursos.forEach(curso => {
        const opt = document.createElement('option');
        opt.value = curso;
        opt.textContent = curso;
        sel.appendChild(opt);
    });
    if (cursoVisualGrade && (cursos.includes(cursoVisualGrade) || cursoVisualGrade === 'Não se aplica')) {
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

    if (!cursoParaExcluir || cursoParaExcluir === 'Não se aplica') {
        mostrarToast('Selecione um curso para excluir.', 'warning');
        return;
    }

    if (cursoParaExcluir === CURSO_TEMPO_INTEGRAL) {
        mostrarToast('O curso "Tempo Integral" é padrão do sistema e não pode ser excluído.', 'warning');
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

function preencherSelectTurmaConsulta() {
    const sel = document.getElementById('filtroTurmaConsulta');
    if (!sel) return;
    sel.innerHTML = '<option value="">Todas as turmas</option>';
    const ordenadas = [...turmas].sort((a, b) => {
        const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
        const ta = ordemTurno[a.turno] || 99;
        const tb = ordemTurno[b.turno] || 99;
        if (ta !== tb) return ta - tb;
        const ca = (a.curso || '').toString();
        const cb = (b.curso || '').toString();
        const cmpCurso = ca.localeCompare(cb, 'pt-BR', { sensitivity: 'base' });
        if (cmpCurso !== 0) return cmpCurso;
        return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
    });
    ordenadas.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = `${t.nome} (${textoTurno(t.turno)})`;
        sel.appendChild(opt);
    });
}

function atualizarSelectTurmaConsulta() {
    const sel = document.getElementById('filtroTurmaConsulta');
    if (!sel) return;
    sel.innerHTML = '<option value="">Todas as turmas</option>';
    const ordenadas = [...turmas].sort((a, b) => {
        const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
        const ta = ordemTurno[a.turno] || 99;
        const tb = ordemTurno[b.turno] || 99;
        if (ta !== tb) return ta - tb;
        const ca = (a.curso || '').toString();
        const cb = (b.curso || '').toString();
        const cmpCurso = ca.localeCompare(cb, 'pt-BR', { sensitivity: 'base' });
        if (cmpCurso !== 0) return cmpCurso;
        return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
    });
    ordenadas.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = `${t.nome} (${textoTurno(t.turno)})`;
        sel.appendChild(opt);
    });
}

function atualizarSelectTurmasInsercaoRapida() {
    const sel = document.getElementById('rapidoTurma');
    if (!sel) return;
    const turnoSel = document.getElementById('rapidoTurno')?.value || '';
    sel.innerHTML = '';
    let lista = [...turmas];
    if (turnoSel) {
        lista = lista.filter(t => t.turno === turnoSel);
    }
    lista.sort(compararTurmasPorCursoETipo);
    lista.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = `${t.nome} (${textoTurno(t.turno)})`;
        sel.appendChild(opt);
    });
}

function aplicarInsercaoRapidaAuto() {
    const dados = obterDadosInsercaoRapida();
    const turno = dados.turno || turnoAtualGrade;
    const disciplina = dados.disciplina;
    const professorId = dados.professorId || null;
    let quantidade = parseInt(dados.quantidade || '1', 10);
    const turmaIds = Array.isArray(dados.turmaIds) && dados.turmaIds.length ? dados.turmaIds : (dados.turmaId ? [dados.turmaId] : []);
    if (!turmaIds.length) {
        mostrarToast('Selecione pelo menos uma turma na inserção rápida.', 'warning');
        return;
    }
    let turmasSucesso = 0;
    const turmasFalha = [];
    turmaIds.forEach(idTurma => {
        const ok = ajusteInteligenteInsercaoRapida(turno, idTurma, disciplina, professorId, quantidade);
        if (ok) {
            harmonizarFamiliasPortMatInsercaoRapida(turno, idTurma, professorId);
            turmasSucesso++;
        } else {
            const turma = turmas.find(t => t.id === idTurma);
            const nome = turma && turma.nome ? turma.nome : String(idTurma);
            turmasFalha.push(nome);
        }
    });
    if (turmasFalha.length) {
        const nomes = turmasFalha.join(', ');
        let msg = turmasSucesso > 0
            ? `Inserção rápida aplicada em ${turmasSucesso} turma(s). Não foi possível alocar nas turmas: ${nomes}.`
            : `Não foi possível alocar automaticamente nas turmas selecionadas: ${nomes}.`;
        if (professorId) {
            const diag = diagnosticoDisponibilidadeProfessorTurno(professorId, turno, quantidade);
            if (diag && diag.capacidade === 0) {
                msg += ` O professor selecionado não possui qualquer disponibilidade registrada neste turno. Para este pedido seriam necessários pelo menos ${quantidade} tempo(s) disponíveis no turno ${textoTurno(turno)}.`;
            } else if (diag && diag.faltam > 0) {
                msg += ` Possível causa: disponibilidade insuficiente do professor neste turno. Com a configuração atual, ele tem ${diag.livres} tempo(s) livre(s) no turno ${textoTurno(turno)}; para atender este pedido seriam necessários pelo menos ${diag.faltam} tempo(s) livres adicionais.`;
            }
        }
        mostrarToast(msg, 'warning');
    }
}

function ajusteInteligenteInsercaoRapida(turno, turmaId, disciplina, professorId, quantidade) {
    if (!turno || !turmaId || !disciplina) {
        mostrarToast('Informe turno, turma e disciplina na inserção rápida.', 'warning');
        return false;
    }
    if (isNaN(quantidade) || quantidade < 1) quantidade = 1;
    const quantidadeOriginal = quantidade;
    const grupoDisc = grupoDisciplinaPortMat(disciplina);
    const disciplinaEspecial = !!grupoDisc;
    if (disciplinaEspecial && quantidade >= 4) {
        const okEsp = aplicarInsercaoRapidaEspecialPortMat(turno, turmaId, disciplina, professorId, quantidade);
        if (!okEsp) {
            mostrarToast('Não foi possível alocar automaticamente 4 tempos (2+2) para esta disciplina com as regras atuais.', 'warning');
        }
        return !!okEsp;
    }
    const temposTurno = getTempos(turno);
    const temposValidos = temposTurno.filter(t => !t.intervalo);
    if (!temposValidos.length) {
        mostrarToast('Não há tempos configurados para este turno.', 'warning');
        return false;
    }
    const dias = diasInsercaoRapida();
    for (let d = 0; d < dias.length; d++) {
        const dia = dias[d];
        for (let i = 0; i < temposTurno.length; i++) {
            const tempoInicial = temposTurno[i];
            if (tempoInicial.intervalo) continue;
            const tempoIds = [];
            for (let j = i; j < temposTurno.length && tempoIds.length < quantidade; j++) {
                const t = temposTurno[j];
                if (t.intervalo) continue;
                tempoIds.push(t.id);
            }
            if (tempoIds.length !== quantidade) continue;
            const ok = tentarAplicarInsercaoRapida(turno, dia, turmaId, tempoIds, disciplina, professorId);
            if (ok) {
                return true;
            }
        }
    }

    let restantes = quantidadeOriginal;
    let inseridos = 0;

    if (!disciplinaEspecial && restantes >= 2) {
        while (restantes >= 2) {
            let alocouBloco = false;
            for (let d = 0; d < dias.length && !alocouBloco; d++) {
                const dia = dias[d];
                for (let i = 0; i < temposTurno.length && !alocouBloco; i++) {
                    const tempoInicial = temposTurno[i];
                    if (tempoInicial.intervalo) continue;
                    const tempoIds = [];
                    for (let j = i; j < temposTurno.length && tempoIds.length < 2; j++) {
                        const t = temposTurno[j];
                        if (t.intervalo) continue;
                        tempoIds.push(t.id);
                    }
                    if (tempoIds.length !== 2) continue;
                    const ok = tentarAplicarInsercaoRapida(turno, dia, turmaId, tempoIds, disciplina, professorId);
                    if (ok) {
                        restantes -= 2;
                        inseridos += 2;
                        alocouBloco = true;
                    }
                }
            }
            if (!alocouBloco) {
                break;
            }
        }
    }

    if (restantes > 0) {
        if (quantidadeOriginal === 1) {
            for (let d = 0; d < dias.length && restantes > 0; d++) {
                const dia = dias[d];
                for (let i = 0; i < temposValidos.length && restantes > 0; i++) {
                    const t = temposValidos[i];
                    const ok = tentarAplicarInsercaoRapida(turno, dia, turmaId, [t.id], disciplina, professorId);
                    if (ok) {
                        restantes--;
                        inseridos++;
                    }
                }
            }
        }
        for (let d = 0; d < dias.length && restantes > 0; d++) {
            const dia = dias[d];
            for (let i = 0; i < temposValidos.length && restantes > 0; i++) {
                const t = temposValidos[i];
                const ok = tentarAplicarInsercaoRapida(turno, dia, turmaId, [t.id], disciplina, professorId);
                if (ok) {
                    restantes--;
                    inseridos++;
                }
            }
        }
    }

    if (inseridos === 0) {
        mostrarToast('Não foi possível alocar automaticamente as aulas com as regras atuais.', 'warning');
        return false;
    }
    if (inseridos < quantidadeOriginal) {
        mostrarToast(`Foram lançados ${inseridos} de ${quantidadeOriginal} tempos solicitados; não foi possível alocar todos com as regras atuais.`, 'warning');
    }
    return true;
}

function atualizarListaTurmas() {
    const corpo = document.getElementById('listaTurmas');
    const selTurmaLimpeza = document.getElementById('turmaLimpeza');
    if (!corpo) return;
    corpo.innerHTML = '';
    if (selTurmaLimpeza) selTurmaLimpeza.innerHTML = '<option value="">Selecione a turma...</option>';

    const ordenadas = [...turmas].sort((a, b) => {
        const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
        const ta = ordemTurno[a.turno] || 99;
        const tb = ordemTurno[b.turno] || 99;
        if (ta !== tb) return ta - tb;
        const ca = (a.curso || '').toString();
        const cb = (b.curso || '').toString();
        const cmpCurso = ca.localeCompare(cb, 'pt-BR', { sensitivity: 'base' });
        if (cmpCurso !== 0) return cmpCurso;
        return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
    });

    ordenadas.forEach(t => {
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

        const tdInfo = document.createElement('td');
        const tipoTxt = textoTipoTurma(t.tipoTurma);
        const anoTxt = textoAnoTurma(t.anoTurma);
        const faseTxt = textoFaseTurma(t.faseTurma);
        const partes = [];
        if (tipoTxt) partes.push(tipoTxt);
        if (anoTxt) partes.push(anoTxt);
        if (faseTxt) partes.push(faseTxt);
        tdInfo.textContent = partes.length ? partes.join(' — ') : '-';
        tr.appendChild(tdInfo);

        const tdAcoes = document.createElement('td');
        tdAcoes.style.display = 'flex';
        tdAcoes.style.gap = '6px';

        const btnEd = document.createElement('button');
        btnEd.className = 'btn-secondary small';
        btnEd.textContent = 'Editar';
        btnEd.onclick = () => abrirModalTurma(t.id);

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
        tdEdit.colSpan = 6;
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
    preencherSelectTurmaConsulta();
}

function salvarTurma(e) {
    e.preventDefault();
    const id = document.getElementById('turmaId').value;
    const nome = document.getElementById('nomeTurma').value.trim();
    const curso = document.getElementById('cursoTurma').value;
    const turnoRaw = document.getElementById('turnoTurma').value;
    const turnosSelecionados = (turnoRaw || '').split(',').map(t => t.trim()).filter(t => t);
    const tipo = document.getElementById('tipoTurma').value;
    const ano = document.getElementById('anoTurma').value;
    const fase = document.getElementById('faseTurma').value;
    const descricao = document.getElementById('descricaoTurma').value.trim();

    if (!nome || !curso || !turnosSelecionados.length || !tipo || !ano) {
        mostrarToast('Preencha nome, curso, turno, tipo e ano.', 'warning');
        return;
    }

    if (id) {
        const t = turmas.find(x => x.id === id);
        if (t) {
            t.nome = nome;
            t.curso = curso;
            t.turno = turnosSelecionados[0];
            t.tipoTurma = tipo;
            t.anoTurma = ano;
            t.faseTurma = fase;
            t.descricao = descricao;
            mostrarToast('Turma atualizada com sucesso!');
        }
    } else {
        turnosSelecionados.forEach(turno => {
            turmas.push({
                id: gerarId(),
                nome,
                curso,
                turno,
                tipoTurma: tipo,
                anoTurma: ano,
                faseTurma: fase,
                descricao
            });
        });
        if (turnosSelecionados.length > 1) {
            const nomesTurnos = turnosSelecionados.map(t => textoTurno(t)).join(' e ');
            mostrarToast(`Turma cadastrada nos turnos ${nomesTurnos} com sucesso!`);
        } else {
            mostrarToast('Turma cadastrada com sucesso!');
        }
    }

    salvarLocal();
    atualizarListaTurmas();
    document.getElementById('formTurma').reset();
    document.getElementById('turmaId').value = '';
    document.getElementById('turnoTurma').value = '';
    document.getElementById('tipoTurma').value = '';
    document.getElementById('anoTurma').value = '';
    const faseSel = document.getElementById('faseTurma');
    if (faseSel) {
        faseSel.innerHTML = '<option value="">Fase...</option>';
    }
    document.querySelectorAll('#chips-turno-turma .chip').forEach(c => c.classList.remove('selecionado'));
}

function editarTurma(id) {
    const t = turmas.find(x => x.id === id);
    if (!t) return;
    document.getElementById('turmaId').value = t.id;
    document.getElementById('nomeTurma').value = t.nome;
    document.getElementById('cursoTurma').value = t.curso;
    document.getElementById('turnoTurma').value = t.turno;
    const tipoSel = document.getElementById('tipoTurma');
    if (tipoSel) {
        tipoSel.value = t.tipoTurma || '';
        atualizarOpcoesAnoTurma('tipoTurma', 'anoTurma', 'faseTurma');
    }
    const anoSel = document.getElementById('anoTurma');
    if (anoSel) {
        anoSel.value = t.anoTurma || '';
        atualizarOpcoesFaseTurma('anoTurma', 'faseTurma');
        const faseSel = document.getElementById('faseTurma');
        if (faseSel) {
            faseSel.value = t.faseTurma || '';
        }
    }
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

function abrirModalTurma(id) {
    const t = turmas.find(x => x.id === id);
    if (!t) return;
    const modal = document.getElementById('modalTurma');
    if (!modal) return;
    const inpId = document.getElementById('modalTurmaId');
    const inpNome = document.getElementById('modalTurmaNome');
    const selCurso = document.getElementById('modalTurmaCurso');
    const hiddenTurno = document.getElementById('modalTurmaTurno');
    const selTipo = document.getElementById('modalTipoTurma');
    const selAno = document.getElementById('modalAnoTurma');
    const selFase = document.getElementById('modalFaseTurma');
    const inpDesc = document.getElementById('modalTurmaDescricao');
    if (!inpId || !inpNome || !selCurso || !hiddenTurno || !inpDesc || !selTipo || !selAno || !selFase) return;
    inpId.value = t.id;
    inpNome.value = t.nome;
    selCurso.innerHTML = '';
    const optNa = document.createElement('option');
    optNa.value = 'Não se aplica';
    optNa.textContent = 'Não se aplica';
    selCurso.appendChild(optNa);
    cursos.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        selCurso.appendChild(opt);
    });
    selCurso.value = t.curso;
    hiddenTurno.value = t.turno;
    const groupTurno = document.getElementById('chips-modal-turno-turma');
    if (groupTurno) {
        const btnSel = Array.from(groupTurno.querySelectorAll('.chip'))
            .find(b => b.getAttribute('data-value') === t.turno);
        groupTurno.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
        if (btnSel) btnSel.classList.add('selecionado');
    }
    selTipo.value = t.tipoTurma || '';
    atualizarOpcoesAnoTurma('modalTipoTurma', 'modalAnoTurma', 'modalFaseTurma');
    selAno.value = t.anoTurma || '';
    atualizarOpcoesFaseTurma('modalAnoTurma', 'modalFaseTurma');
    selFase.value = t.faseTurma || '';
    inpDesc.value = t.descricao || '';
    modal.style.display = 'flex';
}

function fecharModalTurma() {
    const modal = document.getElementById('modalTurma');
    if (modal) modal.style.display = 'none';
}

function salvarModalTurma(e) {
    e.preventDefault();
    const id = document.getElementById('modalTurmaId').value;
    const t = turmas.find(x => x.id === id);
    if (!t) return;
    const nome = document.getElementById('modalTurmaNome').value.trim();
    const curso = document.getElementById('modalTurmaCurso').value;
    const turno = document.getElementById('modalTurmaTurno').value;
    const tipo = document.getElementById('modalTipoTurma').value;
    const ano = document.getElementById('modalAnoTurma').value;
    const fase = document.getElementById('modalFaseTurma').value;
    const descricao = document.getElementById('modalTurmaDescricao').value.trim();
    if (!nome || !curso || !turno || !tipo || !ano) {
        mostrarToast('Preencha nome, curso, turno, tipo e ano.', 'warning');
        return;
    }
    t.nome = nome;
    t.curso = curso;
    t.turno = turno;
    t.tipoTurma = tipo;
    t.anoTurma = ano;
    t.faseTurma = fase;
    t.descricao = descricao;
    salvarLocal();
    atualizarListaTurmas();
    fecharModalTurma();
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
    const ordenadas = [...turmasTurno].sort(compararTurmasPorCursoETipo);
    const gruposPorCurso = {};
    ordenadas.forEach(t => {
        const curso = t.curso || 'Sem curso';
        if (!gruposPorCurso[curso]) gruposPorCurso[curso] = [];
        gruposPorCurso[curso].push(t);
    });
    Object.keys(gruposPorCurso).forEach(curso => {
        const col = document.createElement('div');
        col.className = 'chips-curso-column';
        gruposPorCurso[curso].forEach(t => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'chip';
            btn.dataset.value = t.id;
            btn.textContent = t.nome;
            btn.onclick = () => selecionarChip('chips-turma', 'turmaHorario', btn);
            col.appendChild(btn);
        });
        container.appendChild(col);
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
            const msgInd = `Professor não está disponível em ${dia} / ${textoTurno(turno)} nos tempos: ${indisponiveis.join(', ')}. Deseja continuar mesmo assim?`;
            if (!confirm(msgInd)) {
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
            const msgConf = `Professor já tem aula nos seguintes horários: ${conflitos.join(', ')}. Deseja continuar mesmo assim?`;
            if (!confirm(msgConf)) {
                return;
            }
        }
    }

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

    const grupoDisc = grupoDisciplinaPortMat(disciplina);
    const disciplinaSujeitaRegra = !!grupoDisc;
    if (disciplinaSujeitaRegra) {
        const aulasMesmoDiaMesmaDisciplina = aulas.filter(a =>
            a.turmaId === turmaId &&
            a.dia === dia &&
            grupoDisciplinaPortMat(a.disciplina) === grupoDisc &&
            !tempoIds.includes(a.tempoId)
        );
        const totalFinalDiscDia = aulasMesmoDiaMesmaDisciplina.length + tempoIds.length;
        const limite = 2;
        if (totalFinalDiscDia > limite) {
            mostrarToast(`Não é permitido lançar mais de ${limite} tempos da mesma disciplina no mesmo dia para esta turma.`, 'warning');
            return;
        }
    }

    capturarEstadoAulas();
    aulas = aulas.filter(a =>
        !(a.turno === turno &&
          a.dia === dia &&
          a.turmaId === turmaId &&
          tempoIds.includes(a.tempoId))
    );

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
    turmasTurno = [...turmasTurno].sort(compararTurmasPorCursoETipo);

    const ordemSalva = ORDEM_TURMAS_POR_TURNO[turno];
    if (Array.isArray(ordemSalva) && ordemSalva.length) {
        const mapaPos = new Map();
        ordemSalva.forEach((id, idx) => {
            if (!mapaPos.has(id)) mapaPos.set(id, idx);
        });
        turmasTurno.sort((a, b) => {
            const ia = mapaPos.has(a.id) ? mapaPos.get(a.id) : Number.MAX_SAFE_INTEGER;
            const ib = mapaPos.has(b.id) ? mapaPos.get(b.id) : Number.MAX_SAFE_INTEGER;
            if (ia !== ib) return ia - ib;
            return 0;
        });
    }
    if (!turmasTurno.length) {
        const msg = modoVisualGrade === 'CURSO' && cursoVisualGrade
            ? 'Não há turmas deste curso neste turno.'
            : 'Não há turmas cadastradas neste turno.';
        tabela.innerHTML = `<tr><td class="hint" style="padding:8px;">${msg}</td></tr>`;
        return;
    }

    const mapaContagemDisciplinaTurma = {};
    aulas.forEach(a => {
        if (a.turno !== turno) return;
        if (!a.turmaId || !a.disciplina) return;
        const profKey = a.professorId || '';
        const chave = a.turmaId + '||' + a.disciplina + '||' + profKey;
        mapaContagemDisciplinaTurma[chave] = (mapaContagemDisciplinaTurma[chave] || 0) + 1;
    });

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

    turmasTurno.forEach((turma, idxTurma) => {
        const th = document.createElement('th');
        th.className = 'turma-col';
        th.setAttribute('draggable', 'true');
        th.dataset.turmaId = turma.id;
        th.dataset.colIndex = String(idxTurma + 2);

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

        if (dia !== 'SEGUNDA') {
            const trSep = document.createElement('tr');
            trSep.className = 'linha-dia';
            const tdSep = document.createElement('td');
            tdSep.colSpan = 2 + turmasTurno.length;
            tdSep.textContent = textoDia(dia).toUpperCase();
            trSep.appendChild(tdSep);
            tbody.appendChild(trSep);
        }

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
                        bloco.style.position = 'relative';

                        bloco.innerHTML =
                            '<strong>' + abreviar(apelidarDisciplina(aula.disciplina), 14) + '</strong>' +
                            '<small>' + (prof ? abreviar(prof.nome, 14) : 'Sem professor') + '</small>';

                        const chaveCont = turma.id + '||' + (aula.disciplina || '') + '||' + (aula.professorId || '');
                        const totalDisciplinaTurma = mapaContagemDisciplinaTurma[chaveCont] || 0;
                        if (totalDisciplinaTurma > 0) {
                            const badge = document.createElement('div');
                            badge.textContent = totalDisciplinaTurma;
                            badge.style.position = 'absolute';
                            badge.style.top = '2px';
                            badge.style.right = '4px';
                            badge.style.fontSize = '0.55rem';
                            badge.style.color = '#333';
                            badge.style.backgroundColor = 'rgba(255,255,255,0.8)';
                            badge.style.padding = '0 2px';
                            badge.style.borderRadius = '2px';
                            badge.style.pointerEvents = 'none';
                            bloco.appendChild(badge);
                        }

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
                        const icon = document.createElement('div');
                        icon.style.fontSize = '0.9rem';
                        icon.style.opacity = '0.5';
                        icon.style.marginTop = '5px';
                        icon.innerHTML = '➕';
                        td.appendChild(icon);
                    }

                    td.onclick = (ev) => {
                        if (ev.ctrlKey) {
                            alternarSelecaoCelula(td);
                            return;
                        }
                        if (aula) {
                            if (ev.target === td || ev.target === td.firstChild) {
                                abrirEditorCelula(aula);
                            }
                        } else {
                            abrirEditorCelulaVazia(turno, dia, turma.id, tempo.id);
                        }
                    };

                    td.dataset.turno = turno;
                    td.dataset.dia = dia;
                    td.dataset.turmaId = turma.id;
                    td.dataset.tempoId = tempo.id;
                    td.oncontextmenu = (ev) => {
                        ev.preventDefault();
                        abrirMenuContextoGrade(ev, td);
                    };
                    td.addEventListener('dragover', (ev) => {
                        if (!tempo.intervalo) {
                            ev.preventDefault();
                            td.classList.add('celula-drop-alvo');
                        }
                    });
                    td.addEventListener('dragenter', (ev) => {
                        if (!tempo.intervalo) {
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
                        if (tempo.intervalo) return;
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
                            if (!confirmarDisponibilidadeProfessorDetalhada(prof, turno, dia, tempo.id)) {
                                return;
                            }
                            if (!confirmarConflitoProfessorDetalhado(profId, turno, dia, tempo.id, turma.id)) {
                                return;
                            }
                        }
                        const grupoDisc = grupoDisciplinaPortMat(src.disciplina);
                        if (grupoDisc) {
                            const aulasMesmoDiaMesmaDisciplina = aulas.filter(a =>
                                a.turmaId === turma.id &&
                                a.dia === dia &&
                                grupoDisciplinaPortMat(a.disciplina) === grupoDisc &&
                                !(a.turno === turno && a.dia === dia && a.turmaId === turma.id && a.tempoId === tempo.id)
                            );
                            const totalFinalDiscDia = aulasMesmoDiaMesmaDisciplina.length + 1;
                            const limite = 2;
                            if (totalFinalDiscDia > limite) {
                                mostrarToast('Boas práticas pedagógicas sugerem não concentrar mais de 2 tempos de Português ou Matemática no mesmo dia para a mesma turma.', 'warning');
                            }
                        }
                        capturarEstadoAulas();
                        aulas = aulas.filter(a =>
                            !(a.turno === turno &&
                              a.dia === dia &&
                              a.turmaId === turma.id &&
                              a.tempoId === tempo.id)
                        );
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

            tbody.appendChild(tr);
        });
    });

    tabela.appendChild(tbody);

    const thTurmaCols = Array.from(thead.querySelectorAll('th.turma-col'));
    let dragSrcColIndex = null;

    thTurmaCols.forEach(th => {
        th.addEventListener('dragstart', (ev) => {
            dragSrcColIndex = Array.from(th.parentElement.children).indexOf(th);
            ev.dataTransfer.effectAllowed = 'move';
        });
        th.addEventListener('dragover', (ev) => {
            ev.preventDefault();
            ev.dataTransfer.dropEffect = 'move';
        });
        th.addEventListener('drop', (ev) => {
            ev.preventDefault();
            const destIndex = Array.from(th.parentElement.children).indexOf(th);
            if (dragSrcColIndex === null || dragSrcColIndex === destIndex) return;

            const table = th.closest('table');
            if (!table) return;

            const rows = Array.from(table.rows);
            rows.forEach(row => {
                if (dragSrcColIndex < row.cells.length && destIndex < row.cells.length) {
                    const srcCell = row.cells[dragSrcColIndex];
                    const destCell = row.cells[destIndex];
                    if (srcCell && destCell && srcCell !== destCell) {
                        const refNode = dragSrcColIndex < destIndex ? destCell.nextSibling : destCell;
                        row.insertBefore(srcCell, refNode);
                    }
                }
            });

            const headerRow = thead.querySelector('tr');
            if (headerRow) {
                const cols = Array.from(headerRow.querySelectorAll('th.turma-col'));
                const novaOrdemIds = cols
                    .map(c => c.dataset.turmaId)
                    .filter(Boolean);
                ORDEM_TURMAS_POR_TURNO[turno] = novaOrdemIds;
                try {
                    localStorage.setItem(`${APP_PREFIX}ordem_turmas_turno`, JSON.stringify(ORDEM_TURMAS_POR_TURNO));
                } catch (e) {}
            }

            dragSrcColIndex = null;
        });
    });

    atualizarIndicadorDesfazerTurno();
    atualizarIndicadorRefazerTurno();
}

// editor ao clicar em aula (abre MODAL)
function abrirEditorCelula(aula) {
    aulaEmEdicao = aula;

    const turma = turmas.find(t => t.id === aula.turmaId);
    const tempo = getTempos(aula.turno).find(t => t.id === aula.tempoId);

    document.getElementById('modalAulaId').value = aula.id;
    document.getElementById('modalAulaTurma').value = turma ? turma.nome : '';
    document.getElementById('modalDia').value = textoDia(aula.dia);
    document.getElementById('modalHorario').value = tempo ? `${tempo.inicio} - ${tempo.fim}` : '';
    document.getElementById('modalDisciplina').value = aula.disciplina;
    const campoApelidoModal = document.getElementById('modalApelidoDisciplina');
    if (campoApelidoModal) {
        const nomeDisc = aula.disciplina || '';
        let valor = '';
        if (nomeDisc && Object.prototype.hasOwnProperty.call(apelidosDisciplinas, nomeDisc)) {
            valor = apelidosDisciplinas[nomeDisc] || '';
        } else if (nomeDisc.length > 10) {
            valor = sugerirApelidoDisciplina(nomeDisc);
        }
        campoApelidoModal.value = valor;
        delete campoApelidoModal.dataset.manual;
    }
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
    const campoApelidoModal = document.getElementById('modalApelidoDisciplina');
    const profId = document.getElementById('modalProfessor').value || null;

    if (!disc) {
        mostrarToast('Informe a disciplina.', 'warning');
        return;
    }

    if (campoApelidoModal) {
        const apelido = campoApelidoModal.value.trim();
        if (disc && apelido) {
            apelidosDisciplinas[disc] = apelido;
        }
    }

    if (profId) {
        const prof = professores.find(p => p.id === profId);
        if (!prof) {
            mostrarToast('Professor não encontrado.', 'error');
            return;
        }
        if (!confirmarDisponibilidadeProfessorDetalhada(prof, aulaEmEdicao.turno, aulaEmEdicao.dia, aulaEmEdicao.tempoId)) {
            return;
        }
        const idParaIgnorar = aulaEmEdicao.id === 'novo' ? null : aulaEmEdicao.id;
        if (!confirmarConflitoProfessorDetalhado(profId, aulaEmEdicao.turno, aulaEmEdicao.dia, aulaEmEdicao.tempoId, aulaEmEdicao.turmaId, idParaIgnorar)) {
            return;
        }
    }

    const grupoDisc = grupoDisciplinaPortMat(disc);
    if (grupoDisc) {
        const aulasMesmoDiaMesmaDisciplina = aulas.filter(a =>
            a.turmaId === aulaEmEdicao.turmaId &&
            a.dia === aulaEmEdicao.dia &&
            grupoDisciplinaPortMat(a.disciplina) === grupoDisc &&
            a.id !== (aulaEmEdicao.id === 'novo' ? null : aulaEmEdicao.id)
        );
        const totalFinalDiscDia = aulasMesmoDiaMesmaDisciplina.length + 1;
        const limite = 2;
        if (totalFinalDiscDia > limite) {
            mostrarToast('Boas práticas pedagógicas sugerem não concentrar mais de 2 tempos de Português ou Matemática no mesmo dia para a mesma turma.', 'warning');
        }
    }

    if (aulaEmEdicao.id === 'novo') {
        capturarEstadoAulas();
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
        const aula = aulas.find(a => a.id === aulaEmEdicao.id);
        if (aula) {
            capturarEstadoAulas();
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
    capturarEstadoAulas();
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
    capturarEstadoAulas();
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
    capturarEstadoAulas();
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
    capturarEstadoAulas();
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

function redistribuirTurmaTurnoPorInsercaoRapida(turma, turno) {
    const aulasTurmaTurno = aulas.filter(a =>
        a.turno === turno &&
        a.turmaId === turma.id &&
        a.disciplina
    );
    if (!aulasTurmaTurno.length) return;
    const grupos = {};
    aulasTurmaTurno.forEach(a => {
        const discKey = (a.disciplina || '').toUpperCase();
        const profKey = a.professorId || '';
        const key = profKey + '|' + discKey;
        if (!grupos[key]) grupos[key] = [];
        grupos[key].push(a);
    });
    const mapaFamiliasEspeciais = {};
    Object.keys(grupos).forEach(key => {
        const partes = key.split('|');
        const profKey = partes[0] || '';
        const discKey = partes.slice(1).join('|') || '';
        const tipoFamilia = grupoDisciplinaPortMat(discKey);
        if (!tipoFamilia) return;
        const familyKey = profKey + '||' + tipoFamilia;
        if (!mapaFamiliasEspeciais[familyKey]) {
            mapaFamiliasEspeciais[familyKey] = {
                professorId: profKey || null,
                tipo: tipoFamilia,
                keys: [],
                aulas: [],
                total: 0
            };
        }
        mapaFamiliasEspeciais[familyKey].keys.push(key);
        const listaGrupo = grupos[key] || [];
        mapaFamiliasEspeciais[familyKey].aulas.push(...listaGrupo);
        mapaFamiliasEspeciais[familyKey].total += listaGrupo.length;
    });
    const familiasParaTratar = Object.values(mapaFamiliasEspeciais).filter(f => f.total === 4 && f.keys.length > 1);
    familiasParaTratar.forEach(f => {
        f.keys.forEach(k => {
            delete grupos[k];
        });
    });
    const chavesGrupos = Object.keys(grupos);
    if (!chavesGrupos.length && !familiasParaTratar.length) return;
    chavesGrupos.sort((k1, k2) => {
        const profId1 = (k1.split('|')[0] || '') || null;
        const profId2 = (k2.split('|')[0] || '') || null;
        const p1 = prioridadeProfessorAlocacao(profId1);
        const p2 = prioridadeProfessorAlocacao(profId2);
        if (p1 !== p2) return p1 - p2;
        const g1 = grupos[k1] || [];
        const g2 = grupos[k2] || [];
        return g2.length - g1.length;
    });
    aulas = aulas.filter(a =>
        !(a.turno === turno &&
          a.turmaId === turma.id &&
          a.disciplina)
    );
    familiasParaTratar.forEach(f => {
        const grupoAulas = f.aulas || [];
        if (!grupoAulas.length) return;
        const base = grupoAulas.find(a => grupoDisciplinaPortMat(a.disciplina) === f.tipo) || grupoAulas[0];
        const disciplinaBase = base.disciplina;
        const professorId = base.professorId || null;
        const quantidade = grupoAulas.length;
        const ok = ajusteInteligenteInsercaoRapida(turno, turma.id, disciplinaBase, professorId, quantidade);
        if (!ok) {
            grupoAulas.forEach(a => aulas.push(a));
            return;
        }
        const contPorDisciplina = {};
        grupoAulas.forEach(a => {
            const nome = a.disciplina || '';
            contPorDisciplina[nome] = (contPorDisciplina[nome] || 0) + 1;
        });
        Object.keys(contPorDisciplina).forEach(nome => {
            if (nome === disciplinaBase) return;
            let restantes = contPorDisciplina[nome];
            if (restantes <= 0) return;
            const candidatas = aulas.filter(a =>
                a.turno === turno &&
                a.turmaId === turma.id &&
                (a.professorId || null) === professorId &&
                (a.disciplina || '') === disciplinaBase
            );
            for (let i = 0; i < candidatas.length && restantes > 0; i++) {
                candidatas[i].disciplina = nome;
                restantes--;
            }
        });
    });
    chavesGrupos.forEach(key => {
        const grupo = grupos[key];
        if (!grupo || !grupo.length) return;
        const base = grupo[0];
        const disciplina = base.disciplina;
        const professorId = base.professorId || null;
        const quantidade = grupo.length;
        const ok = ajusteInteligenteInsercaoRapida(turno, turma.id, disciplina, professorId, quantidade);
        if (!ok) {
            grupo.forEach(a => aulas.push(a));
        }
    });
}

function harmonizarFamiliasPortMatInsercaoRapida(turno, turmaId, professorId) {
    if (!professorId) return;
    const aulasTurmaTurno = aulas.filter(a =>
        a.turno === turno &&
        a.turmaId === turmaId &&
        a.disciplina &&
        a.professorId
    );
    if (!aulasTurmaTurno.length) return;
    const grupos = {};
    aulasTurmaTurno.forEach(a => {
        const discKey = (a.disciplina || '').toUpperCase();
        const profKey = a.professorId || '';
        if (profKey !== professorId) return;
        const key = profKey + '|' + discKey;
        if (!grupos[key]) grupos[key] = [];
        grupos[key].push(a);
    });
    const mapaFamiliasEspeciais = {};
    Object.keys(grupos).forEach(key => {
        const partes = key.split('|');
        const profKey = partes[0] || '';
        const discKey = partes.slice(1).join('|') || '';
        const tipoFamilia = grupoDisciplinaPortMat(discKey);
        if (!tipoFamilia) return;
        const familyKey = profKey + '||' + tipoFamilia;
        if (!mapaFamiliasEspeciais[familyKey]) {
            mapaFamiliasEspeciais[familyKey] = {
                professorId: profKey || null,
                tipo: tipoFamilia,
                aulas: [],
                total: 0
            };
        }
        const listaGrupo = grupos[key] || [];
        mapaFamiliasEspeciais[familyKey].aulas.push(...listaGrupo);
        mapaFamiliasEspeciais[familyKey].total += listaGrupo.length;
    });
    const familiasParaTratar = Object.values(mapaFamiliasEspeciais).filter(f => f.total === 4);
    if (!familiasParaTratar.length) return;
    familiasParaTratar.forEach(f => {
        const grupoAulas = f.aulas || [];
        if (!grupoAulas.length) return;
        const base = grupoAulas.find(a => grupoDisciplinaPortMat(a.disciplina) === f.tipo) || grupoAulas[0];
        const disciplinaBase = base.disciplina;
        const profId = base.professorId || null;
        const quantidade = grupoAulas.length;
        const contPorDisciplina = {};
        grupoAulas.forEach(a => {
            const nome = a.disciplina || '';
            contPorDisciplina[nome] = (contPorDisciplina[nome] || 0) + 1;
        });
        const backup = JSON.stringify(aulas);
        const idsFamilia = new Set(grupoAulas.map(a => a.id));
        aulas = aulas.filter(a => !idsFamilia.has(a.id));
        const ok = ajusteInteligenteInsercaoRapida(turno, turmaId, disciplinaBase, profId, quantidade);
        if (!ok) {
            aulas = JSON.parse(backup);
            return;
        }
        const aulasAntes = JSON.parse(backup);
        const idsAntes = new Set(aulasAntes.map(a => a.id));
        const novos = aulas.filter(a =>
            !idsAntes.has(a.id) &&
            a.turno === turno &&
            a.turmaId === turmaId &&
            (a.professorId || null) === profId &&
            (a.disciplina || '') === disciplinaBase
        );
        if (!novos.length) {
            aulas = JSON.parse(backup);
            return;
        }
        novos.sort((a, b) => {
            if (a.dia === b.dia) return a.tempoId - b.tempoId;
            return a.dia < b.dia ? -1 : 1;
        });
        let idx = 0;
        Object.keys(contPorDisciplina).forEach(nome => {
            let qtd = contPorDisciplina[nome];
            while (qtd > 0 && idx < novos.length) {
                novos[idx].disciplina = nome;
                idx++;
                qtd--;
            }
        });
    });
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
}

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

    const snapshotInicial = JSON.stringify(aulas);
    const infoInicial = calcularQualidadeAjusteTurno(turno);
    const qualidadeInicial = infoInicial.percentual;
    let melhorAnterior = typeof melhorPercentualAjustePorTurno[turno] === 'number'
        ? melhorPercentualAjustePorTurno[turno]
        : qualidadeInicial;
    if (melhorAnterior < qualidadeInicial) {
        melhorAnterior = qualidadeInicial;
    }

    let melhorSnapshot = snapshotInicial;
    let melhorQualidade = qualidadeInicial;
    let melhorInfo = infoInicial;

    capturarEstadoAulas();

    const temposValidos = getTempos(turnoAtualGrade).filter(t => !t.intervalo);

    for (let tentativa = 0; tentativa < 10; tentativa++) {
        turmasTurno.forEach(turma => {
            DIAS_SEMANA.forEach(dia => {
                const aulasDia = temposValidos.map(tempo =>
                    obterAula(turno, dia, turma.id, tempo.id) || null
                );
                const compactadas = aulasDia.filter(a => a !== null);
                if (!compactadas.length) {
                    return;
                }

                const gruposDisciplina = {};
                compactadas.forEach(a => {
                    const chaveDisc = (a.disciplina || '') + '|' + (a.professorId || '');
                    if (!gruposDisciplina[chaveDisc]) {
                        gruposDisciplina[chaveDisc] = [];
                    }
                    gruposDisciplina[chaveDisc].push(a);
                });
                const ordemProcessamento = [];
                const processados = new Set();
                Object.values(gruposDisciplina).forEach(lista => {
                    if (lista.length >= 2) {
                        lista.forEach(a => {
                            ordemProcessamento.push(a);
                            if (a.id != null) {
                                processados.add(a.id);
                            }
                        });
                    }
                });
                Object.values(gruposDisciplina).forEach(lista => {
                    if (lista.length === 1) {
                        const a = lista[0];
                        if (a && a.id != null && !processados.has(a.id)) {
                            ordemProcessamento.push(a);
                            processados.add(a.id);
                        } else if (a && a.id == null) {
                            ordemProcessamento.push(a);
                        }
                    }
                });

                const aulasOutrasTurmas = aulas.filter(a =>
                    !(a.turno === turno && a.dia === dia && a.turmaId === turma.id)
                );

                const novasAulasDia = [];
                const temposUsados = new Set();

                ordemProcessamento.forEach(aula => {
                    const profId = aula.professorId || null;
                    const tempoOriginalId = aula.tempoId;
                    let tempoEscolhidoId = null;

                    for (let i = 0; i < temposValidos.length; i++) {
                        const tempo = temposValidos[i];
                        if (temposUsados.has(tempo.id)) continue;

                        let podeUsar = true;

                        if (profId) {
                            const prof = professores.find(p => p.id === profId);
                            if (!prof) {
                                podeUsar = false;
                            } else {
                                if (!professorDisponivelNoHorario(prof, turno, dia, tempo.id)) {
                                    podeUsar = false;
                                }
                            }

                            if (podeUsar) {
                                const conflitoProfessor = aulasOutrasTurmas.some(a =>
                                    a.professorId === profId &&
                                    a.turno === turno &&
                                    a.dia === dia &&
                                    a.tempoId === tempo.id
                                );
                                if (conflitoProfessor) {
                                    podeUsar = false;
                                }
                            }
                        }

                        if (podeUsar) {
                            tempoEscolhidoId = tempo.id;
                            break;
                        }
                    }

                    if (!tempoEscolhidoId) {
                        const tempoOriginalExiste = temposValidos.some(t => t.id === tempoOriginalId);
                        if (tempoOriginalExiste && !temposUsados.has(tempoOriginalId)) {
                            let podeManter = true;
                            if (profId) {
                                const prof = professores.find(p => p.id === profId);
                                if (!prof || !professorDisponivelNoHorario(prof, turno, dia, tempoOriginalId)) {
                                    podeManter = false;
                                } else {
                                    const conflitoOriginal = aulasOutrasTurmas.some(a =>
                                        a.professorId === profId &&
                                        a.turno === turno &&
                                        a.dia === dia &&
                                        a.tempoId === tempoOriginalId
                                    );
                                    if (conflitoOriginal) {
                                        podeManter = false;
                                    }
                                }
                            }
                            if (podeManter) {
                                tempoEscolhidoId = tempoOriginalId;
                            }
                        }
                    }

                    if (!tempoEscolhidoId) {
                        tempoEscolhidoId = tempoOriginalId;
                    }

                    temposUsados.add(tempoEscolhidoId);
                    aula.tempoId = tempoEscolhidoId;
                    novasAulasDia.push(aula);
                });

                aulas = aulas.filter(a =>
                    !(a.turno === turno && a.dia === dia && a.turmaId === turma.id)
                );
                novasAulasDia.forEach(aula => {
                    aulas.push(aula);
                });
            });
        });

        turmasTurno.forEach(turma => {
            const profsTurma = new Set(
                aulas
                    .filter(a =>
                        a.turno === turno &&
                        a.turmaId === turma.id &&
                        a.professorId
                    )
                    .map(a => a.professorId)
            );
            profsTurma.forEach(profId => {
                harmonizarFamiliasPortMatInsercaoRapida(turno, turma.id, profId);
            });
        });

        const infoDepois = calcularQualidadeAjusteTurno(turno);
        const qualidadeDepois = infoDepois.percentual;
        if (qualidadeDepois > melhorQualidade) {
            melhorQualidade = qualidadeDepois;
            melhorInfo = infoDepois;
            melhorSnapshot = JSON.stringify(aulas);
        }
    }

    if (melhorQualidade < melhorAnterior) {
        aulas = JSON.parse(snapshotInicial);
        salvarLocal();
        montarGradeTurno(turnoAtualGrade);
        mostrarToast(
            `Ajuste inteligente não melhorou a alocação (mantido em ${melhorAnterior}%). Os horários anteriores foram preservados.`,
            'warning'
        );
        return;
    }

    const houveAlteracao = melhorSnapshot !== snapshotInicial;

    if (!houveAlteracao) {
        aulas = JSON.parse(snapshotInicial);
        melhorPercentualAjustePorTurno[turno] = melhorQualidade;
        salvarLocal();
        montarGradeTurno(turnoAtualGrade);
        mostrarToast(
            `Ajuste inteligente executado, mas não foi necessário alterar os horários deste turno (já estão na melhor configuração encontrada: ${melhorQualidade}%).`,
            'info'
        );
        return;
    }

    aulas = JSON.parse(melhorSnapshot);
    melhorPercentualAjustePorTurno[turno] = melhorQualidade;
    salvarLocal();
    montarGradeTurno(turnoAtualGrade);
    mostrarToast(
        `Ajuste inteligente aplicado com sucesso: ${melhorQualidade}% dos grupos de aulas estão bem alocados (${melhorInfo.totalOk} de ${melhorInfo.totalAvaliado}). Use "Desfazer" se necessário.`
    );
}

function desfazerAjuste() {
    const turno = turnoAtualGrade;
    if (!pilhaUndo.length) {
        mostrarToast('Não há ação para desfazer neste turno.', 'warning');
        return;
    }
    let idx = pilhaUndo.length - 1;
    while (idx >= 0) {
        try {
            const entry = JSON.parse(pilhaUndo[idx]);
            if (entry && entry.turno === turno) {
                const aulasAtuaisTurno = aulas.filter(a => a.turno === turno);
                pilhaRedo.push(JSON.stringify({ turno, aulas: aulasAtuaisTurno }));
                if (pilhaRedo.length > 20) {
                    pilhaRedo.shift();
                }
                pilhaUndo.splice(idx, 1);
                aulas = aulas.filter(a => a.turno !== turno);
                if (Array.isArray(entry.aulas)) {
                    entry.aulas.forEach(a => aulas.push(a));
                }
                salvarLocal();
                montarGradeTurno(turnoAtualGrade);
                mostrarToast('Ação desfeita com sucesso neste turno.');
                atualizarIndicadorDesfazerTurno();
                atualizarIndicadorRefazerTurno();
                return;
            }
        } catch (e) {
        }
        idx--;
    }
    mostrarToast('Não há ação anterior registrada para este turno.', 'warning');
}

function refazerAjuste() {
    const turno = turnoAtualGrade;
    if (!pilhaRedo.length) {
        mostrarToast('Não há ação para refazer neste turno.', 'warning');
        return;
    }
    let idx = pilhaRedo.length - 1;
    while (idx >= 0) {
        try {
            const entry = JSON.parse(pilhaRedo[idx]);
            if (entry && entry.turno === turno) {
                const aulasAtuaisTurno = aulas.filter(a => a.turno === turno);
                pilhaUndo.push(JSON.stringify({ turno, aulas: aulasAtuaisTurno }));
                if (pilhaUndo.length > 20) {
                    pilhaUndo.shift();
                }
                pilhaRedo.splice(idx, 1);
                aulas = aulas.filter(a => a.turno !== turno);
                if (Array.isArray(entry.aulas)) {
                    entry.aulas.forEach(a => aulas.push(a));
                }
                salvarLocal();
                montarGradeTurno(turnoAtualGrade);
                mostrarToast('Ação refeita com sucesso neste turno.');
                atualizarIndicadorDesfazerTurno();
                atualizarIndicadorRefazerTurno();
                return;
            }
        } catch (e) {
        }
        idx--;
    }
    mostrarToast('Não há ação para refazer neste turno.', 'warning');
}

// =======================
// RELATÓRIOS - ATUALIZADO
// =======================

function preencherSelectCursoRelatorio() {
    const sel = document.getElementById('cursoRelatorio');
    const selFiltroCurso = document.getElementById('filtroCursoConsulta');
    if (sel) {
        sel.innerHTML = '';
        cursos.forEach(curso => {
            const opt = document.createElement('option');
            opt.value = curso;
            opt.textContent = curso;
            sel.appendChild(opt);
        });
    }
    if (selFiltroCurso) {
        selFiltroCurso.innerHTML = '<option value="">Todos os cursos / áreas</option>';
        cursos.forEach(curso => {
            const opt = document.createElement('option');
            opt.value = curso;
            opt.textContent = curso;
            selFiltroCurso.appendChild(opt);
        });
    }
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
    renderTemposTurno();
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

    const campoDiscProf = document.getElementById('disciplinasProfessor');
    const campoApelidoProf = document.getElementById('apelidoProfessorDisciplina');
    if (campoDiscProf && campoApelidoProf) {
        campoDiscProf.addEventListener('input', () => {
            const texto = campoDiscProf.value || '';
            const primeira = texto.split(',')[0].trim();
            if (primeira.length > 10 && !campoApelidoProf.dataset.manual) {
                const sugestao = sugerirApelidoDisciplina(primeira);
                if (sugestao) {
                    campoApelidoProf.value = sugestao;
                }
            }
        });
        campoApelidoProf.addEventListener('input', () => {
            campoApelidoProf.dataset.manual = '1';
        });
    }

    const modalDisc = document.getElementById('modalDisciplina');
    const modalApelido = document.getElementById('modalApelidoDisciplina');
    if (modalDisc && modalApelido) {
        modalDisc.addEventListener('input', () => {
            const nome = modalDisc.value.trim();
            if (nome.length > 10 && !modalApelido.dataset.manual) {
                const sugestao = sugerirApelidoDisciplina(nome);
                if (sugestao) {
                    modalApelido.value = sugestao;
                }
            } else if (!nome && !modalApelido.dataset.manual) {
                modalApelido.value = '';
            }
        });
        modalApelido.addEventListener('input', () => {
            modalApelido.dataset.manual = '1';
        });
    }
});
function gerarRelatorioTurnoPDF() {
    const { jsPDF } = window.jspdf;
    const turno = document.getElementById('turnoRelatorio').value;
    const layout = document.getElementById('layoutRelatorio').value;
    const chkOmitirSabado = document.getElementById('omitirSabadoRelatorios');
    const omitirSabado = !!(chkOmitirSabado && chkOmitirSabado.checked);

    let turmasTurno = turmas.filter(t => t.turno === turno).sort(compararTurmasPorCursoETipo);
    const ordemSalva = ORDEM_TURMAS_POR_TURNO[turno];
    if (Array.isArray(ordemSalva) && ordemSalva.length) {
        const mapaPos = new Map();
        ordemSalva.forEach((id, idx) => {
            if (!mapaPos.has(id)) mapaPos.set(id, idx);
        });
        turmasTurno.sort((a, b) => {
            const ia = mapaPos.has(a.id) ? mapaPos.get(a.id) : Number.MAX_SAFE_INTEGER;
            const ib = mapaPos.has(b.id) ? mapaPos.get(b.id) : Number.MAX_SAFE_INTEGER;
            if (ia !== ib) return ia - ib;
            return 0;
        });
    }
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
    const diasUsados = DIAS_SEMANA.filter(d => !(omitirSabado && d === 'SABADO'));

    diasUsados.forEach((dia, idxDia) => {
        const temposTurno = getTempos(turno);
        temposTurno.forEach((tempo, idx) => {
            const row = [];
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
                        const nomeDisciplina = apelidarDisciplina(aula.disciplina || '');
                        const nomeProfessor = prof ? prof.nome : '';
                        const textoCelula = nomeProfessor ? `${nomeDisciplina}\n${nomeProfessor}` : nomeDisciplina;
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

        if (idxDia < diasUsados.length - 1) {
            body.push([{
                content: '',
                colSpan: colCount,
                styles: {
                    fillColor: [60, 60, 60],
                    textColor: 255,
                    cellPadding: 0,
                    minCellHeight: 1
                }
            }]);
        }
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
            1: { cellWidth: 20, fontStyle: 'bold' }
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
    let turmasTurno = turmas.filter(t => t.turno === turno).sort(compararTurmasPorCursoETipo);
    const ordemSalva = ORDEM_TURMAS_POR_TURNO[turno];
    if (Array.isArray(ordemSalva) && ordemSalva.length) {
        const mapaPos = new Map();
        ordemSalva.forEach((id, idx) => {
            if (!mapaPos.has(id)) mapaPos.set(id, idx);
        });
        turmasTurno.sort((a, b) => {
            const ia = mapaPos.has(a.id) ? mapaPos.get(a.id) : Number.MAX_SAFE_INTEGER;
            const ib = mapaPos.has(b.id) ? mapaPos.get(b.id) : Number.MAX_SAFE_INTEGER;
            if (ia !== ib) return ia - ib;
            return 0;
        });
    }
    if (!turmasTurno.length) {
        mostrarToast('Não há turmas neste turno.', 'warning');
        return;
    }

    const chkOmitirSabado = document.getElementById('omitirSabadoRelatorios');
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
                        const nomeDisc = apelidarDisciplina(aula.disciplina || '');
                        row.push(
                            `${nomeDisc}${prof ? ' - ' + prof.nome : ''}`
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
    let idxDiaPlanilha = 0;
    DIAS_SEMANA.forEach((d) => {
        if (omitirSabado && d === 'SABADO') return;
        const r0 = 1 + idxDiaPlanilha * rowsPerDia;
        const r1 = r0 + rowsPerDia - 1;
        merges.push({ s: { r: r0, c: 0 }, e: { r: r1, c: 0 } });
        const addr = XLSX.utils.encode_cell({ r: r0, c: 0 });
        if (ws[addr]) {
            ws[addr].s = Object.assign({}, ws[addr].s || {}, { alignment: { horizontal: 'center', vertical: 'center', textRotation: 90 } });
        }
        idxDiaPlanilha++;
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
        let idxDiaPlanilha2 = 0;
        DIAS_SEMANA.forEach((dia) => {
            if (omitirSabado && dia === 'SABADO') return;
            const temposTurno2 = getTempos(turno);
            temposTurno2.forEach((tempo, idxTempo) => {
                if (tempo.intervalo) return;
                turmasTurno.forEach((turma, idxTurma) => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) return;
                    const prof = professores.find(p => p.id === aula.professorId);
                    if (!prof || !prof.cor) return;
                    const r = 1 + idxDiaPlanilha2 * rowsPerDia + idxTempo;
                    const c = 2 + idxTurma;
                    const addr = XLSX.utils.encode_cell({ r, c });
                    if (ws[addr]) {
                        ws[addr].s = Object.assign({}, ws[addr].s || {}, {
                            fill: { patternType: 'solid', fgColor: { rgb: toArgb(prof.cor) } },
                            font: { color: { rgb: 'FFFFFFFF' } },
                            alignment: { horizontal: 'center', vertical: 'center' }
                        });
                    }
                });
            });
            idxDiaPlanilha2++;
        });
    }

    const totalCols = 2 + turmasTurno.length;
    let idxDiaBorda = 0;
    DIAS_SEMANA.forEach((d) => {
        if (omitirSabado && d === 'SABADO') return;
        const r1 = 1 + idxDiaBorda * rowsPerDia + rowsPerDia - 1;
        for (let c = 0; c < totalCols; c++) {
            const addr = XLSX.utils.encode_cell({ r: r1, c });
            if (ws[addr]) {
                const prev = ws[addr].s || {};
                const prevBorder = prev.border || {};
                ws[addr].s = Object.assign({}, prev, {
                    border: Object.assign({}, prevBorder, {
                        bottom: { style: 'medium', color: { rgb: 'FF000000' } }
                    })
                });
            }
        }
        idxDiaBorda++;
    });
    
    XLSX.utils.book_append_sheet(wb, ws, 'Turno');
    XLSX.writeFile(wb, `horario_turno_${turno.toLowerCase()}_${new Date().toISOString().slice(0,10)}.xlsx`);
    mostrarToast('Excel gerado com sucesso!');
}

function gerarRelatorioCursoPDF() {
    const { jsPDF } = window.jspdf;
    const selCursos = document.getElementById('cursoRelatorio');
    const cursosSelecionados = selCursos
        ? Array.from(selCursos.selectedOptions).map(o => o.value).filter(v => v)
        : [];
    const turno = document.getElementById('turnoRelatorioCurso').value;
    const layoutSel = document.getElementById('layoutRelatorioCurso');
    const layout = layoutSel ? layoutSel.value : 'compacto';

    if (!cursosSelecionados.length) {
        mostrarToast('Selecione ao menos um curso.', 'warning');
        return;
    }

    let turmasSelecionadas = turmas
        .filter(t => t.turno === turno && cursosSelecionados.includes(t.curso))
        .sort(compararTurmasPorCursoETipo);

    const ordemSalva = ORDEM_TURMAS_POR_TURNO[turno];
    if (Array.isArray(ordemSalva) && ordemSalva.length) {
        const mapaPos = new Map();
        ordemSalva.forEach((id, idx) => {
            if (!mapaPos.has(id)) mapaPos.set(id, idx);
        });
        turmasSelecionadas.sort((a, b) => {
            const ia = mapaPos.has(a.id) ? mapaPos.get(a.id) : Number.MAX_SAFE_INTEGER;
            const ib = mapaPos.has(b.id) ? mapaPos.get(b.id) : Number.MAX_SAFE_INTEGER;
            if (ia !== ib) return ia - ib;
            return 0;
        });
    }

    if (!turmasSelecionadas.length) {
        mostrarToast('Não há turmas nos cursos/turno selecionados.', 'warning');
        return;
    }

    const colCount = 2 + turmasSelecionadas.map(t => t.nome).length;
    const format = colCount > 12 ? 'a3' : 'a4';
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format });

    const listaCursosTitulo = cursosSelecionados.join(' • ');
    const titulo = `HORÁRIO - CURSOS (${listaCursosTitulo}) - TURNO DA ${textoTurno(turno).toUpperCase()} — Gerado em: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`;
    doc.setFont(undefined, 'bold');
    doc.setFontSize(8);
    doc.text(titulo, doc.internal.pageSize.getWidth() / 2, 6, { align: 'center' });

    const head = [['DIA', 'HORÁRIO', ...turmasSelecionadas.map(t => t.nome)]];
    const body = [];
    const chkOmitirSabado = document.getElementById('omitirSabadoRelatorioCurso');
    const omitirSabado = !!(chkOmitirSabado && chkOmitirSabado.checked);
    const diasUsados = DIAS_SEMANA.filter(d => !(omitirSabado && d === 'SABADO'));

    diasUsados.forEach((dia, idxDia) => {
        const temposTurno = getTempos(turno);
        temposTurno.forEach((tempo, idx) => {
            const row = [];
            if (idx === 0) {
                row.push({
                    content: textoDia(dia),
                    rowSpan: temposTurno.length,
                    styles: { valign: 'middle', halign: 'center' }
                });
            }
            row.push(`${tempo.inicio} - ${tempo.fim}`);
            if (tempo.intervalo) {
                turmasSelecionadas.forEach(() => row.push('INTERVALO'));
            } else {
                turmasSelecionadas.forEach(turma => {
                    const aula = obterAula(turno, dia, turma.id, tempo.id);
                    if (!aula) {
                        row.push('-');
                    } else {
                        const prof = professores.find(p => p.id === aula.professorId);
                        const nomeDisciplina = apelidarDisciplina(aula.disciplina || '');
                        const nomeProfessor = prof ? prof.nome : '';
                        const textoCelula = nomeProfessor ? `${nomeDisciplina}\n${nomeProfessor}` : nomeDisciplina;
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

        if (idxDia < diasUsados.length - 1) {
            body.push([{
                content: '',
                colSpan: colCount,
                styles: {
                    fillColor: [60, 60, 60],
                    textColor: 255,
                    cellPadding: 0,
                    minCellHeight: 1
                }
            }]);
        }
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
            1: { cellWidth: 20, fontStyle: 'bold' }
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

    doc.save(`horario_cursos_${turno.toLowerCase()}.pdf`);
    mostrarToast('PDF de cursos gerado com sucesso!');
}

function gerarRelatorioCursoXLS() {
    if (typeof XLSX === 'undefined') {
        mostrarToast('Biblioteca XLSX não carregada.', 'error');
        return;
    }

    const selCursos = document.getElementById('cursoRelatorio');
    const cursosSelecionados = selCursos
        ? Array.from(selCursos.selectedOptions).map(o => o.value).filter(v => v)
        : [];
    const turno = document.getElementById('turnoRelatorioCurso').value;
    const layoutSel = document.getElementById('layoutRelatorioCurso');
    const layout = layoutSel ? layoutSel.value : 'compacto';

    if (!cursosSelecionados.length) {
        mostrarToast('Selecione ao menos um curso.', 'warning');
        return;
    }

    const wb = XLSX.utils.book_new();

    cursosSelecionados.forEach(curso => {
        let turmasCurso = turmas.filter(t => t.curso === curso && t.turno === turno);
        const ordemSalva = ORDEM_TURMAS_POR_TURNO[turno];
        if (Array.isArray(ordemSalva) && ordemSalva.length) {
            const mapaPos = new Map();
            ordemSalva.forEach((id, idx) => {
                if (!mapaPos.has(id)) mapaPos.set(id, idx);
            });
            turmasCurso.sort((a, b) => {
                const ia = mapaPos.has(a.id) ? mapaPos.get(a.id) : Number.MAX_SAFE_INTEGER;
                const ib = mapaPos.has(b.id) ? mapaPos.get(b.id) : Number.MAX_SAFE_INTEGER;
                if (ia !== ib) return ia - ib;
                return 0;
            });
        }
        if (!turmasCurso.length) {
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
                        const nomeDisc = apelidarDisciplina(aula.disciplina || '');
                        row.push(`${nomeDisc}${prof ? ' - ' + prof.nome : ''}`);
                    }
                    });
                }

                sheetData.push(row);
            });
        });

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
                                fill: { patternType: 'solid', fgColor: { rgb: toArgb(prof.cor) } },
                                font: { color: { rgb: 'FFFFFFFF' } },
                                alignment: { horizontal: 'center', vertical: 'center' }
                            });
                        }
                    });
                });
            });
        }

        const totalCols = 2 + turmasCurso.length;
        DIAS_SEMANA.forEach((d, idxDia) => {
            const r1 = 1 + idxDia * rowsPerDia + rowsPerDia - 1;
            for (let c = 0; c < totalCols; c++) {
                const addr = XLSX.utils.encode_cell({ r: r1, c });
                if (ws[addr]) {
                    const prev = ws[addr].s || {};
                    const prevBorder = prev.border || {};
                    ws[addr].s = Object.assign({}, prev, {
                        border: Object.assign({}, prevBorder, {
                            bottom: { style: 'medium', color: { rgb: 'FF000000' } }
                        })
                    });
                }
            }
        });

        XLSX.utils.book_append_sheet(wb, ws, curso.substring(0, 25));
    });

    XLSX.writeFile(wb, `horario_cursos_${turno.toLowerCase()}_${new Date().toISOString().slice(0,10)}.xlsx`);
    mostrarToast('Excel de cursos gerado com sucesso!');
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

    const baseLabel = prof.baseTipo === 'COMUM' ? 'Base Comum' : 'Base Técnica';

    if (!aulasProf.length && !incluirVagos) {
        if (tituloEl) tituloEl.textContent = `${legendaTabela} — ${baseLabel}`;
        corpo.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:20px;">Nenhum horário lançado para este professor.</td></tr>';
        return;
    }

    if (tituloEl) {
        tituloEl.textContent = `${legendaTabela} — ${baseLabel}`;
    }

    const carga = calcularCargaHorariaProfessor(aulasProf);

    const disciplinasUsadas = [...new Set(
        aulasProf
            .map(a => a.disciplina)
            .filter(Boolean)
    )];

    const cargaPorDisciplinaTurnoSemana = {};
    aulasProf.forEach(a => {
        if (!a || a._vago) return;
        const disc = a.disciplina || '';
        const turno = a.turno || '';
        if (!disc || !turno) return;
        const key = `${disc}__${turno}`;
        if (!cargaPorDisciplinaTurnoSemana[key]) cargaPorDisciplinaTurnoSemana[key] = 0;
        cargaPorDisciplinaTurnoSemana[key] += 1;
    });
    const cargaPorDisciplinaTurnoMes = {};
    Object.keys(cargaPorDisciplinaTurnoSemana).forEach(key => {
        const tempos = cargaPorDisciplinaTurnoSemana[key];
        cargaPorDisciplinaTurnoMes[key] = tempos * 5;
    });

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

    // Ordena as chaves: por turno (Manhã, Tarde, Noite) e depois por dia da semana
    const chavesOrdenadas = Object.keys(grupos).sort((a, b) => {
        const [diaA, turnoA] = a.split('_');
        const [diaB, turnoB] = b.split('_');
        const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
        const ta = ordemTurno[turnoA] || 99;
        const tb = ordemTurno[turnoB] || 99;
        if (ta !== tb) return ta - tb;
        const idxDiaA = DIAS_SEMANA.indexOf(diaA);
        const idxDiaB = DIAS_SEMANA.indexOf(diaB);
        return idxDiaA - idxDiaB;
    });

    const linhasPorTurnoCH = {};
    chavesOrdenadas.forEach(key => {
        const [, turno] = key.split('_');
        const qtd = (grupos[key] || []).length;
        if (!qtd) return;
        linhasPorTurnoCH[turno] = (linhasPorTurnoCH[turno] || 0) + qtd;
    });
    const turnosCHImpressos = new Set();

    let totalVagos = 0;

    chavesOrdenadas.forEach((chave, grupoIndex) => {
        const [dia, turno] = chave.split('_');
        const aulasGrupo = (grupos[chave] || []).slice().sort((a, b) => {
            return (a.tempoId || 0) - (b.tempoId || 0);
        });
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
            
            const turmaNome = turma ? turma.nome : '';
            const isVago = !!a._vago;
            const tipoTexto = turma ? (textoTipoTurma(turma.tipoTurma) || '-') : '';
            const anoTexto = turma ? (textoAnoTurma(turma.anoTurma) || '-') : '';
            const faseTexto = turma ? (textoFaseTurma(turma.faseTurma) || '-') : '';
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
                    const tdTipo = document.createElement('td');
                    tdTipo.textContent = tipoTexto;
                    tdTipo.rowSpan = spanTurma;
                    tdTipo.style.verticalAlign = 'middle';
                    tr.appendChild(tdTipo);
                    const tdAno = document.createElement('td');
                    tdAno.textContent = anoTexto;
                    tdAno.rowSpan = spanTurma;
                    tdAno.style.verticalAlign = 'middle';
                    tr.appendChild(tdAno);
                    const tdFase = document.createElement('td');
                    tdFase.textContent = faseTexto;
                    tdFase.rowSpan = spanTurma;
                    tdFase.style.verticalAlign = 'middle';
                    tr.appendChild(tdFase);
                }
            } else {
                const tdTurma = document.createElement('td');
                tdTurma.textContent = 'Escola';
                tr.appendChild(tdTurma);
                const tdTipo = document.createElement('td');
                tdTipo.textContent = '';
                tr.appendChild(tdTipo);
                const tdAno = document.createElement('td');
                tdAno.textContent = '';
                tr.appendChild(tdAno);
                const tdFase = document.createElement('td');
                tdFase.textContent = '';
                tr.appendChild(tdFase);
            }
            
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
                tdDisciplina.textContent = apelidarDisciplina(discNome);
                if (a._vago) tdDisciplina.style.fontStyle = 'italic';
                tdDisciplina.rowSpan = spanDisc;
                tdDisciplina.style.verticalAlign = 'middle';
                tdDisciplina.style.whiteSpace = 'normal';
                tdDisciplina.style.wordBreak = 'break-word';
                tr.appendChild(tdDisciplina);
            }
            if (a._vago || !discNome || !turno) {
                const tdHoraAula = document.createElement('td');
                tdHoraAula.textContent = '';
                tr.appendChild(tdHoraAula);
            } else if (isInicioRunDisc) {
                let spanHora = 1;
                for (let j = idx + 1; j < aulasGrupo.length; j++) {
                    const prox = aulasGrupo[j];
                    const proxDisc = prox.disciplina || '';
                    if (!prox._vago && proxDisc === discNome) {
                        spanHora++;
                    } else {
                        break;
                    }
                }
                const tdHoraAula = document.createElement('td');
                const keyDiscTurno = `${discNome}__${turno}`;
                const horasMesDisc = cargaPorDisciplinaTurnoMes[keyDiscTurno] || 0;
                tdHoraAula.textContent = horasMesDisc ? `${horasMesDisc}h` : '';
                tdHoraAula.style.textAlign = 'center';
                tdHoraAula.rowSpan = spanHora;
                tdHoraAula.style.verticalAlign = 'middle';
                tr.appendChild(tdHoraAula);
            }

            if (!turnosCHImpressos.has(turno)) {
                const tdCHTurno = document.createElement('td');
                const infoTurno = (carga && carga.porTurno && carga.porTurno[turno]) ? carga.porTurno[turno] : null;
                const horasMesTurno = infoTurno ? infoTurno.horasMes : 0;
                tdCHTurno.textContent = horasMesTurno ? `${horasMesTurno}h` : '';
                tdCHTurno.style.textAlign = 'center';
                tdCHTurno.rowSpan = linhasPorTurnoCH[turno] || aulasGrupo.length;
                tdCHTurno.style.verticalAlign = 'middle';
                tr.appendChild(tdCHTurno);
                turnosCHImpressos.add(turno);
            }
            
            corpo.appendChild(tr);
        });
        
        if (grupoIndex < chavesOrdenadas.length - 1) {
            const trSeparador = document.createElement('tr');
            trSeparador.innerHTML = '<td colspan="11" style="background:#f5f5f5;height:1px;padding:0;"></td>';
            corpo.appendChild(trSeparador);
        }
    });

    const resumoCargaEl = document.getElementById('profResumoCarga');
    const resumoVagosEl = document.getElementById('profResumoVagos');
    const partesTurno = [];
    const porTurno = (carga && carga.porTurno) ? carga.porTurno : {};
    ['MANHA', 'TARDE', 'NOITE'].forEach(turno => {
        const info = porTurno[turno];
        if (info && info.temposSemana) {
            partesTurno.push(
                `${textoTurno(turno)}: ${info.temposSemana} tempos (${info.horasSemana}h/sem, ${info.horasMes}h/mês)`
            );
        }
    });
    let resumoTexto = '';
    if (partesTurno.length) {
        resumoTexto = `Por turno: ${partesTurno.join(' | ')}. `;
    }
    resumoTexto += `Total: ${carga.temposSemana} tempos (${carga.horasSemana}h semanais, aproximadamente ${carga.horasMes}h mensais).`;
    if (resumoCargaEl) {
        resumoCargaEl.textContent = `Carga horária (por turno e total)  ${resumoTexto}`;
    }

    if (totalVagos > 0 && incluirVagos) {
        const horasVagoSemana = totalVagos;
        const horasVagoMes = horasVagoSemana * 5;
        if (resumoVagosEl) {
            resumoVagosEl.textContent =
                `Tempo vago na escola (entre aulas)  ${totalVagos} tempos ` +
                `(${horasVagoSemana}h semanais, aproximadamente ${horasVagoMes}h mensais)`;
        }
    } else if (resumoVagosEl) {
        resumoVagosEl.textContent = '';
    }

    const trTotalCHTurno = document.createElement('tr');
    for (let i = 0; i < 10; i++) {
        const tdVazio = document.createElement('td');
        tdVazio.textContent = '';
        trTotalCHTurno.appendChild(tdVazio);
    }
    const tdTotalCH = document.createElement('td');
    const totalHorasTurnos = carga && typeof carga.horasMes === 'number' ? carga.horasMes : 0;
    tdTotalCH.textContent = totalHorasTurnos ? `${totalHorasTurnos}h` : '';
    tdTotalCH.style.fontWeight = '600';
    trTotalCHTurno.appendChild(tdTotalCH);
    corpo.appendChild(trTotalCHTurno);

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
    document.body.classList.add('print-relatorios');
    window.print();
    document.body.classList.remove('print-relatorios');
    if (tituloEl && tituloOriginal !== null) {
        tituloEl.textContent = tituloOriginal;
    }
}

function preverImpressaoConsulta() {
    const tipo = document.getElementById('tipoConsulta')?.value || '';
    const tabela = document.getElementById('tabelaConsulta');
    if (!tabela) return;
    const temConteudo = tabela.querySelector('tbody tr');
    if (!temConteudo) {
        mostrarToast('Gere a consulta antes de pré-visualizar a impressão.', 'warning');
        return;
    }
    const tituloEl = document.querySelector('#consultas h2');
    const tituloResultadoEl = document.getElementById('tituloResultadoConsulta');
    const tituloOriginal = tituloEl ? tituloEl.textContent : null;
    const resultadoOriginal = tituloResultadoEl ? tituloResultadoEl.textContent : null;
    if (tituloEl) {
        let desc = '';
        if (tipo === 'professores') desc = 'por Professores';
        else if (tipo === 'turmas') desc = 'por Turmas';
        else if (tipo === 'turnos') desc = 'por Turnos';
        else if (tipo === 'cursos') desc = 'por Cursos / Áreas';
        tituloEl.textContent = desc ? `Consultas Avançadas ${desc}` : 'Consultas Avançadas';
    }
    if (tituloResultadoEl) {
        let descTipo = '';
        if (tipo === 'professores') descTipo = 'Professores';
        else if (tipo === 'turmas') descTipo = 'Turmas';
        else if (tipo === 'turnos') descTipo = 'Turnos';
        else if (tipo === 'cursos') descTipo = 'Cursos / Áreas';
        else if (tipo === 'disponibilidadeProfessores') descTipo = 'Tempos disponíveis dos Professores';
        else if (tipo === 'disponibilidadeProfessoresInformada') descTipo = 'Disponibilidade dos Professores Informada';
        else if (tipo === 'tempoProfessoresSala') descTipo = 'Tempo dos Professores em Sala';
        tituloResultadoEl.textContent = descTipo ? `Resultado - ${descTipo}` : 'Resultado';
    }
    document.body.classList.add('print-consultas');
    window.print();
    document.body.classList.remove('print-consultas');
    if (tituloEl && tituloOriginal !== null) {
        tituloEl.textContent = tituloOriginal;
    }
    if (tituloResultadoEl && resultadoOriginal !== null) {
        tituloResultadoEl.textContent = resultadoOriginal;
    }
}

function onChangeTipoConsulta() {
    const tipo = document.getElementById('tipoConsulta')?.value || '';
    const grupoBase = document.getElementById('filtroBaseProfGroup');
    const selBase = document.getElementById('filtroBaseProfessor');
    const grupoTurma = document.getElementById('filtroTurmaGroup');
    const grupoTurno = document.getElementById('filtroTurnoGroup');
    const grupoCurso = document.getElementById('filtroCursoGroup');
    const grupoOpcoes = document.getElementById('filtroOpcoesConsultaGroup');
    if (grupoBase) {
        grupoBase.style.display = (tipo === 'professores' || tipo === 'disponibilidadeProfessores' || tipo === 'disponibilidadeProfessoresInformada' || tipo === 'tempoProfessoresSala') ? '' : 'none';
    }
    if (tipo !== 'professores' && tipo !== 'disponibilidadeProfessores' && tipo !== 'disponibilidadeProfessoresInformada' && tipo !== 'tempoProfessoresSala' && selBase) {
        selBase.value = '';
    }
    if (grupoTurma) {
        grupoTurma.style.display = tipo === 'turmas' ? '' : 'none';
    }
    if (grupoTurno) {
        grupoTurno.style.display = (tipo === 'turnos' || tipo === 'disponibilidadeProfessores' || tipo === 'disponibilidadeProfessoresInformada' || tipo === 'tempoProfessoresSala') ? '' : 'none';
    }
    if (grupoCurso) {
        grupoCurso.style.display = tipo === 'cursos' ? '' : 'none';
    }
    if (grupoOpcoes) {
        grupoOpcoes.style.display = tipo === 'professores' ? '' : 'none';
    }
}

function calcularCargaHorariaProfessor(aulasProf) {
    const lista = Array.isArray(aulasProf) ? aulasProf : [];
    const temposSemana = lista.length;
    const horasSemana = temposSemana;
    const horasMes = temposSemana * 5;

    const temposPorTurno = {};
    lista.forEach(a => {
        if (!a || !a.turno) return;
        if (!temposPorTurno[a.turno]) temposPorTurno[a.turno] = 0;
        temposPorTurno[a.turno] += 1;
    });

    const porTurno = {};
    Object.keys(temposPorTurno).forEach(turno => {
        const tempos = temposPorTurno[turno];
        porTurno[turno] = {
            temposSemana: tempos,
            horasSemana: tempos,
            horasMes: tempos * 5
        };
    });

    return { temposSemana, horasSemana, horasMes, porTurno };
}

function calcularTemposVagosProfessor(aulasProf, prof) {
    if (!aulasProf || !aulasProf.length) return { vagosSemana: 0, vagosMes: 0 };
    const temposPorDiaTurno = {};
    aulasProf.forEach(a => {
        const key = `${a.dia}_${a.turno}`;
        if (!temposPorDiaTurno[key]) temposPorDiaTurno[key] = new Set();
        temposPorDiaTurno[key].add(a.tempoId);
    });
    let totalVagos = 0;
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
                if (prof && !professorDisponivelNoHorario(prof, turno, dia, t.id)) return;
                totalVagos++;
            }
        });
    });
    return { vagosSemana: totalVagos, vagosMes: totalVagos * 5 };
}

function contarDiasDisponiveisProfessor(prof) {
    if (!prof) return 0;
    const disp = prof.disponibilidadePorDia || null;
    if (disp && typeof disp === 'object') {
        let cont = 0;
        DIAS_SEMANA.forEach(dia => {
            const valor = disp[dia];
            if (Array.isArray(valor)) {
                if (valor.length > 0) cont++;
            } else if (valor) {
                cont++;
            }
        });
        return cont;
    }
    const diasArr = Array.isArray(prof.diasDisponiveis) ? prof.diasDisponiveis : [];
    const diasValidos = diasArr.filter(d => DIAS_SEMANA.includes(d));
    return new Set(diasValidos).size;
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
        const partesTurno = [];
        const porTurno = (cargaTotal && cargaTotal.porTurno) ? cargaTotal.porTurno : {};
        ['MANHA', 'TARDE', 'NOITE'].forEach(turno => {
            const info = porTurno[turno];
            if (info && info.temposSemana) {
                partesTurno.push(
                    `${textoTurno(turno)}: ${info.temposSemana} tempos (${info.horasSemana}h/sem, ${info.horasMes}h/mês)`
                );
            }
        });
        let textoTurnos = '';
        if (partesTurno.length) {
            textoTurnos = `Por turno: ${partesTurno.join(' | ')}. `;
        }
        textoTurnos +=
            `Total: ${cargaTotal.temposSemana} tempos na semana ` +
            `(${cargaTotal.horasSemana}h semanais, aproximadamente ${cargaTotal.horasMes}h mensais).`;
        resumoTurnosEl.textContent = textoTurnos;
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
        const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
        const ta = ordemTurno[turnoA] || 99;
        const tb = ordemTurno[turnoB] || 99;
        if (ta !== tb) return ta - tb;
        const idxDiaA = DIAS_SEMANA.indexOf(diaA);
        const idxDiaB = DIAS_SEMANA.indexOf(diaB);
        return idxDiaA - idxDiaB;
    });

    const agoraPDF = new Date();
    const dataPDF = agoraPDF.toLocaleDateString('pt-BR');
    const horaPDF = agoraPDF.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    const baseLabelProf = prof.baseTipo === 'COMUM' ? 'Base Comum' : 'Base Técnica';
    const titulo = `HORÁRIO DO PROFESSOR - ${prof.nome.toUpperCase()} (${baseLabelProf})`;
    doc.setFontSize(10);
    doc.setFont(undefined, 'bold');
    doc.text(titulo, doc.internal.pageSize.getWidth() / 2, 26, {
        align: 'center',
        maxWidth: doc.internal.pageSize.getWidth() - 20
    });

    const head = [['Dia', 'Turno', 'Horário', 'Tempo', 'Turma', 'Tipo', 'Ano', 'Fase', 'Disciplina', 'C.H./Disciplina', 'C.H./Turno']];
    const body = [];
    const sepRows = new Set();
    let bodyRowIndex = 0;

    const cargaPorDisciplinaTurnoSemana = {};
    aulasProf.forEach(a => {
        if (!a || a._vago) return;
        const disc = a.disciplina || '';
        const turnoAula = a.turno || '';
        if (!disc || !turnoAula) return;
        const key = `${disc}__${turnoAula}`;
        if (!cargaPorDisciplinaTurnoSemana[key]) cargaPorDisciplinaTurnoSemana[key] = 0;
        cargaPorDisciplinaTurnoSemana[key] += 1;
    });
    const cargaPorDisciplinaTurnoMes = {};
    Object.keys(cargaPorDisciplinaTurnoSemana).forEach(key => {
        const tempos = cargaPorDisciplinaTurnoSemana[key];
        cargaPorDisciplinaTurnoMes[key] = tempos * 5;
    });

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
            const tipoTexto = turma ? (textoTipoTurma(turma.tipoTurma) || '-') : '';
            const anoTexto = turma ? (textoAnoTurma(turma.anoTurma) || '-') : '';
            const faseTexto = turma ? (textoFaseTurma(turma.faseTurma) || '-') : '';
            const isVago = !!a._vago;
            let turmaCell;
            let tipoCell;
            let anoCell;
            let faseCell;
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
                    tipoCell = { content: '', rowSpan: spanEscola, styles: { valign: 'middle' } };
                    anoCell = { content: '', rowSpan: spanEscola, styles: { valign: 'middle' } };
                    faseCell = { content: '', rowSpan: spanEscola, styles: { valign: 'middle' } };
                } else {
                    turmaCell = '';
                    tipoCell = '';
                    anoCell = '';
                    faseCell = '';
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
                    tipoCell = { content: tipoTexto, rowSpan: spanTurma, styles: { valign: 'middle' } };
                    anoCell = { content: anoTexto, rowSpan: spanTurma, styles: { valign: 'middle' } };
                    faseCell = { content: faseTexto, rowSpan: spanTurma, styles: { valign: 'middle' } };
                } else {
                    turmaCell = '';
                    tipoCell = '';
                    anoCell = '';
                    faseCell = '';
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
                discCell = { content: apelidarDisciplina(discNome), rowSpan: spanDisc, styles: discStyles };
            } else {
                discCell = '';
            }
            
            const diaCell = idx === 0
                ? { content: textoDia(dia), _rowsGrupo: rowsGrupo, styles: { valign: 'middle' } }
                : '';

            const turnoCell = idx === 0
                ? { content: textoTurno(turno), _rowsGrupo: rowsGrupo, styles: { valign: 'middle' } }
                : '';

            const keyDiscTurno = (!isVago && discNome && turno) ? `${discNome}__${turno}` : null;
            const horasMesDisc = keyDiscTurno ? (cargaPorDisciplinaTurnoMes[keyDiscTurno] || 0) : 0;
            const infoTurno = (carga && carga.porTurno && carga.porTurno[turno]) ? carga.porTurno[turno] : null;
            const horasMesTurno = infoTurno ? infoTurno.horasMes : 0;

            const row = [
                diaCell,
                turnoCell,
                tempo ? `${tempo.inicio} - ${tempo.fim}` : '',
                tempo ? tempo.etiqueta : '',
                turmaCell,
                tipoCell,
                anoCell,
                faseCell,
                discCell,
                (isVago || !discNome || !horasMesDisc) ? '' : `${horasMesDisc}h`,
                horasMesTurno ? `${horasMesTurno}h` : ''
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
            ? disciplinasUsadas.map(d => apelidarDisciplina(d)).join(', ')
            : (Array.isArray(prof.disciplinas) && prof.disciplinas.length
                ? prof.disciplinas.map(d => apelidarDisciplina(d)).join(', ')
                : 'Não informado'));

    const baseLabelPdf = prof.baseTipo === 'COMUM' ? 'Base Comum' : 'Base Técnica';

    doc.setFontSize(8);
    doc.setFont(undefined, 'normal');
    doc.text(`Base: ${baseLabelPdf}`, 14, 30);
    doc.text(`Disciplinas: ${disciplinasLabel}`, 14, 34);

    doc.autoTable({
        head,
        body,
        startY: 36,
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
            0: { cellWidth: 8, cellPadding: 0 },
            1: { cellWidth: 10, cellPadding: 0 },
            2: { cellWidth: 24, cellPadding: 1 },
            3: { cellWidth: 16, cellPadding: 1 },
            4: { cellWidth: 22 },
            5: { cellWidth: 18 },
            6: { cellWidth: 12 },
            7: { cellWidth: 12 },
            8: { cellWidth: 26, cellPadding: 1 },
            9: { cellWidth: 12, cellPadding: 1 }
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
                if (colIdx === 8 || colIdx === 9) {
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

    let yResumo = finalY + 5;
    const porTurnoResumo = (carga && carga.porTurno) ? carga.porTurno : {};
    const linhasTurnos = [];
    ['MANHA', 'TARDE', 'NOITE'].forEach(turno => {
        const info = porTurnoResumo[turno];
        if (info && info.temposSemana) {
            linhasTurnos.push(
                `${textoTurno(turno)}: ${info.temposSemana} tempos (${info.horasSemana}h semanais, aproximadamente ${info.horasMes}h mensais)`
            );
        }
    });

    if (linhasTurnos.length) {
        doc.text('Carga horária por turno:', 14, yResumo);
        yResumo += 4;
        linhasTurnos.forEach(linha => {
            doc.text(linha, 16, yResumo);
            yResumo += 4;
        });
    }

    const textoTotal =
        `Carga horária total: ${carga.temposSemana} tempos ` +
        `(${carga.horasSemana}h semanais, aproximadamente ${carga.horasMes}h mensais).`;
    doc.text(textoTotal, 14, yResumo);
    yResumo += 4;

    if (incluirVagos && totalVagos > 0) {
        const horasVagoSemana = totalVagos;
        const horasVagoMes = horasVagoSemana * 5;
        doc.setFont(undefined, 'italic');
        doc.text(
            `Tempo vago na escola (entre aulas): ${totalVagos} tempos (${horasVagoSemana}h semanais, aproximadamente ${horasVagoMes}h mensais).`,
            14,
            yResumo
        );
        doc.setFont(undefined, 'normal');
        yResumo += 4;
    }

    doc.text(
        `Gerado em: ${dataPDF} às ${horaPDF}`,
        14,
        yResumo
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
    const filtroBase = document.getElementById('filtroBaseProfessor')?.value || '';
    const filtroTurmaId = document.getElementById('filtroTurmaConsulta')?.value || '';
    const filtroTurno = document.getElementById('filtroTurnoConsulta')?.value || '';
    const filtroCurso = document.getElementById('filtroCursoConsulta')?.value || '';
    const thead = document.querySelector('#tabelaConsulta thead');
    const tbody = document.querySelector('#tabelaConsulta tbody');
    const tituloResultadoEl = document.getElementById('tituloResultadoConsulta');
    if (!thead || !tbody) return;
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (tituloResultadoEl) {
        let descTipo = '';
        if (tipo === 'professores') descTipo = 'Professores';
        else if (tipo === 'turmas') descTipo = 'Turmas';
        else if (tipo === 'turnos') descTipo = 'Turnos';
        else if (tipo === 'cursos') descTipo = 'Cursos / Áreas';
        else if (tipo === 'disponibilidadeProfessores') descTipo = 'Tempos disponíveis dos Professores';
        else if (tipo === 'disponibilidadeProfessoresInformada') descTipo = 'Disponibilidade dos Professores Informada';
        else if (tipo === 'tempoProfessoresSala') descTipo = 'Tempo dos Professores em Sala';
        tituloResultadoEl.textContent = descTipo ? `Resultado - ${descTipo}` : 'Resultado';
    }

    if (tipo === 'professores') {
        thead.innerHTML = `
            <tr>
                <th>Professor</th>
                <th>Disciplinas</th>
                <th>CH por disciplina</th>
                <th>Total de Aulas</th>
                <th>Turnos Atuantes</th>
                <th>CH mensal (h)</th>
                <th class="col-ch-janelas">CH em janelas (h/mês)</th>
                <th>Dias atuantes</th>
                <th>Dias disponíveis</th>
                <th>Nº turmas</th>
                <th>Nº disciplinas</th>
            </tr>
        `;
        
        const ordenados = [...professores].sort((a, b) => {
            const pesoBase = (x) => x.baseTipo === 'COMUM' ? 2 : 1;
            const pa = pesoBase(a);
            const pb = pesoBase(b);
            if (pa !== pb) return pa - pb;
            return (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' });
        });

            ordenados.forEach(p => {
            if (filtroBase && (p.baseTipo || 'TECNICA') !== filtroBase) {
                return;
            }
            const aulasProfessor = aulas.filter(a => a.professorId === p.id);
            const turnos = [...new Set(aulasProfessor.map(a => a.turno))];
            const disciplinasUsadas = [...new Set(
                aulasProfessor
                    .map(a => a.disciplina)
                    .filter(Boolean)
            )];
            const carga = calcularCargaHorariaProfessor(aulasProfessor);
            const temposVagos = calcularTemposVagosProfessor(aulasProfessor, p);
            const diasAtuantes = new Set(aulasProfessor.map(a => a.dia)).size;
            const diasDisponiveis = contarDiasDisponiveisProfessor(p);
            const turmasAtuantes = new Set(aulasProfessor.map(a => a.turmaId).filter(Boolean)).size;
            let numDisciplinas = disciplinasUsadas.length;
            if (!numDisciplinas && Array.isArray(p.disciplinas)) {
                numDisciplinas = new Set(p.disciplinas.filter(Boolean)).size;
            }

            const cargaPorDisciplinaMap = {};
            aulasProfessor.forEach(a => {
                const disc = a.disciplina || '';
                if (!disc) return;
                if (!cargaPorDisciplinaMap[disc]) cargaPorDisciplinaMap[disc] = 0;
                cargaPorDisciplinaMap[disc] += 1;
            });
            const cargaPorDisciplinaTexto = Object.keys(cargaPorDisciplinaMap).length
                ? Object.keys(cargaPorDisciplinaMap)
                    .sort((d1, d2) => d1.localeCompare(d2, 'pt-BR', { sensitivity: 'base' }))
                    .map(disc => {
                        const tempos = cargaPorDisciplinaMap[disc];
                        const horasMes = tempos * 5;
                        return `${disc}: ${tempos} tempos (${horasMes}h/mês)`;
                    })
                    .join(' | ')
                : '-';
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${p.nome}</strong></td>
                <td>${disciplinasUsadas.join(', ') || '-'}</td>
                <td style="text-align:left;">${cargaPorDisciplinaTexto}</td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;">${aulasProfessor.length}</span></td>
                <td>${turnos.map(t => textoTurno(t)).join(', ') || '-'}</td>
                <td><span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${carga.horasMes}</span></td>
                <td class="col-ch-janelas"><span style="background:#fff8e1;padding:2px 8px;border-radius:12px;">${temposVagos.vagosMes}</span></td>
                <td>${diasAtuantes}</td>
                <td>${diasDisponiveis}</td>
                <td>${turmasAtuantes}</td>
                <td>${numDisciplinas}</td>
            `;
            tbody.appendChild(tr);
        });

        const ocultarCH = document.getElementById('ocultarCHJanelasConsulta')?.checked;
        const tabela = document.getElementById('tabelaConsulta');
        if (tabela) {
            const display = ocultarCH ? 'none' : '';
            tabela.querySelectorAll('.col-ch-janelas').forEach(el => {
                el.style.display = display;
            });
        }
    } 
    else if (tipo === 'disponibilidadeProfessores') {
        const diasTabela = ['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA'];
        const turnosTabela = ['MANHA','TARDE','NOITE'];

        thead.innerHTML = `
            <tr>
                <th rowspan="2">Nº</th>
                <th rowspan="2">Professor(a)</th>
                ${diasTabela.map(d => `<th colspan="${turnosTabela.length}">${textoDia(d)}</th>`).join('')}
            </tr>
            <tr>
                ${diasTabela.map(() => `
                    <th>Manhã</th>
                    <th>Tarde</th>
                    <th>Noite</th>
                `).join('')}
            </tr>
        `;

        const filtroBaseEfetivo = filtroBase || '';
        const filtroTurnoEfetivo = filtroTurno || '';

        const ordenados = [...professores].sort((a, b) =>
            (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' })
        );

        let indice = 1;

        ordenados.forEach(p => {
            const baseTipo = p.baseTipo || 'TECNICA';
            if (filtroBaseEfetivo && baseTipo !== filtroBaseEfetivo) {
                return;
            }

            let temAlguma = false;
            let cellsHtml = '';

            diasTabela.forEach(dia => {
                turnosTabela.forEach(turno => {
                    if (filtroTurnoEfetivo && turno !== filtroTurnoEfetivo) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
                    if (!temposTurno.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    const temposDisponiveis = temposTurno.filter(t => {
                        if (!professorDisponivelNoHorario(p, turno, dia, t.id)) return false;
                        const conflito = aulas.some(a =>
                            a.professorId === p.id &&
                            a.turno === turno &&
                            a.dia === dia &&
                            a.tempoId === t.id
                        );
                        if (conflito) return false;
                        return true;
                    });

                    if (!temposDisponiveis.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    temAlguma = true;

                    const ordenadosTempos = temposDisponiveis
                        .slice()
                        .sort((t1, t2) => t1.id - t2.id);

                    const primeiro = ordenadosTempos[0];
                    const ultimo = ordenadosTempos[ordenadosTempos.length - 1];

                    let descTempos = '';
                    if (primeiro.id === ultimo.id) {
                        descTempos = (primeiro.etiqueta || '').replace(' Tempo', '').trim();
                    } else {
                        const inicioTexto = (primeiro.etiqueta || '').replace(' Tempo', '').trim();
                        const fimTexto = (ultimo.etiqueta || '').replace(' Tempo', '').trim();
                        descTempos = `${inicioTexto} ao ${fimTexto}`;
                    }

                    cellsHtml += `<td>${descTempos}</td>`;
                });
            });

            if (!temAlguma) return;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${indice}</td>
                <td><strong>${p.nome}</strong></td>
                ${cellsHtml}
            `;
            tbody.appendChild(tr);
            indice++;
        });
    }
    else if (tipo === 'disponibilidadeProfessoresInformada') {
        const diasTabela = ['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA'];
        const turnosTabela = ['MANHA','TARDE','NOITE'];

        thead.innerHTML = `
            <tr>
                <th rowspan="2">Nº</th>
                <th rowspan="2">Professor(a)</th>
                ${diasTabela.map(d => `<th colspan="${turnosTabela.length}">${textoDia(d)}</th>`).join('')}
            </tr>
            <tr>
                ${diasTabela.map(() => `
                    <th>Manhã</th>
                    <th>Tarde</th>
                    <th>Noite</th>
                `).join('')}
            </tr>
        `;

        const filtroBaseEfetivo = filtroBase || '';
        const filtroTurnoEfetivo = filtroTurno || '';

        const ordenados = [...professores].sort((a, b) =>
            (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' })
        );

        let indice = 1;

        ordenados.forEach(p => {
            const baseTipo = p.baseTipo || 'TECNICA';
            if (filtroBaseEfetivo && baseTipo !== filtroBaseEfetivo) {
                return;
            }

            const dispDia = p.disponibilidadePorDia || {};
            const temposMap = p.temposPorDiaTurno || null;

            let temAlguma = false;
            let cellsHtml = '';

            diasTabela.forEach(dia => {
                const listaTurnosDia = Array.isArray(dispDia[dia]) ? dispDia[dia] : [];

                turnosTabela.forEach(turno => {
                    if (filtroTurnoEfetivo && turno !== filtroTurnoEfetivo) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    if (!listaTurnosDia.includes(turno)) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    let ids = [];
                    if (temposMap && temposMap[dia] && Array.isArray(temposMap[dia][turno]) && temposMap[dia][turno].length) {
                        ids = temposMap[dia][turno];
                    } else {
                        ids = getTempos(turno).filter(t => !t.intervalo).map(t => t.id);
                    }

                    if (!ids.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
                    const usados = temposTurno.filter(tp => ids.includes(tp.id));

                    if (!usados.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    temAlguma = true;

                    const ordenadosTempos = usados
                        .slice()
                        .sort((t1, t2) => t1.id - t2.id);

                    const primeiro = ordenadosTempos[0];
                    const ultimo = ordenadosTempos[ordenadosTempos.length - 1];

                    let descTempos = '';
                    if (primeiro.id === ultimo.id) {
                        descTempos = (primeiro.etiqueta || '').replace(' Tempo', '').trim();
                    } else {
                        const inicioTexto = (primeiro.etiqueta || '').replace(' Tempo', '').trim();
                        const fimTexto = (ultimo.etiqueta || '').replace(' Tempo', '').trim();
                        descTempos = `${inicioTexto} ao ${fimTexto}`;
                    }

                    cellsHtml += `<td>${descTempos}</td>`;
                });
            });

            if (!temAlguma) return;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${indice}</td>
                <td><strong>${p.nome}</strong></td>
                ${cellsHtml}
            `;
            tbody.appendChild(tr);
            indice++;
        });
    }
    else if (tipo === 'tempoProfessoresSala') {
        const diasTabela = ['SEGUNDA','TERCA','QUARTA','QUINTA','SEXTA'];
        const turnosTabela = ['MANHA','TARDE','NOITE'];

        thead.innerHTML = `
            <tr>
                <th rowspan="2">Nº</th>
                <th rowspan="2">Professor(a)</th>
                ${diasTabela.map(d => `<th colspan="${turnosTabela.length}">${textoDia(d)}</th>`).join('')}
            </tr>
            <tr>
                ${diasTabela.map(() => `
                    <th>Manhã</th>
                    <th>Tarde</th>
                    <th>Noite</th>
                `).join('')}
            </tr>
        `;

        const filtroBaseEfetivo = filtroBase || '';
        const filtroTurnoEfetivo = filtroTurno || '';

        const ordenados = [...professores].sort((a, b) =>
            (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' })
        );

        let indice = 1;

        ordenados.forEach(p => {
            const baseTipo = p.baseTipo || 'TECNICA';
            if (filtroBaseEfetivo && baseTipo !== filtroBaseEfetivo) {
                return;
            }

            const aulasProfessor = aulas.filter(a => a.professorId === p.id);

            let temAlguma = false;
            let cellsHtml = '';

            diasTabela.forEach(dia => {
                turnosTabela.forEach(turno => {
                    if (filtroTurnoEfetivo && turno !== filtroTurnoEfetivo) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    const aulasDiaTurno = aulasProfessor.filter(a => a.dia === dia && a.turno === turno);
                    if (!aulasDiaTurno.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    const ids = Array.from(new Set(aulasDiaTurno.map(a => a.tempoId).filter(id => id !== null && id !== undefined))).sort((a, b) => a - b);
                    if (!ids.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    const temposTurno = getTempos(turno).filter(t => !t.intervalo);
                    const usados = temposTurno.filter(tp => ids.includes(tp.id));

                    if (!usados.length) {
                        cellsHtml += '<td></td>';
                        return;
                    }

                    temAlguma = true;

                    const ordenadosTempos = usados
                        .slice()
                        .sort((t1, t2) => t1.id - t2.id);

                    const primeiro = ordenadosTempos[0];
                    const ultimo = ordenadosTempos[ordenadosTempos.length - 1];

                    let descTempos = '';
                    if (primeiro.id === ultimo.id) {
                        descTempos = (primeiro.etiqueta || '').replace(' Tempo', '').trim();
                    } else {
                        const inicioTexto = (primeiro.etiqueta || '').replace(' Tempo', '').trim();
                        const fimTexto = (ultimo.etiqueta || '').replace(' Tempo', '').trim();
                        descTempos = `${inicioTexto} ao ${fimTexto}`;
                    }

                    cellsHtml += `<td>${descTempos}</td>`;
                });
            });

            if (!temAlguma) return;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${indice}</td>
                <td><strong>${p.nome}</strong></td>
                ${cellsHtml}
            `;
            tbody.appendChild(tr);
            indice++;
        });
    }
    else if (tipo === 'turmas') {
        thead.innerHTML = `
            <tr>
                <th>Curso / Área</th>
                <th>Turno</th>
                <th>Turma</th>
                <th>Tipo da turma</th>
                <th>Ano</th>
                <th>Fase</th>
                <th>Descrição</th>
                <th>Nº Professores</th>
                <th>Total de Aulas</th>
            </tr>
        `;
        
        // Aplica filtro de turma (todas ou turma específica)
        const baseTurmas = filtroTurmaId
            ? turmas.filter(t => t.id === filtroTurmaId)
            : turmas.slice();

        // Agrupa turmas por curso
        const turmasPorCurso = {};
        baseTurmas.forEach(t => {
            if (!turmasPorCurso[t.curso]) {
                turmasPorCurso[t.curso] = [];
            }
            turmasPorCurso[t.curso].push(t);
        });
        
        Object.keys(turmasPorCurso).sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' })).forEach(curso => {
            const turmasCurso = turmasPorCurso[curso].slice().sort((t1, t2) => {
                const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
                const ta = ordemTurno[t1.turno] || 99;
                const tb = ordemTurno[t2.turno] || 99;
                if (ta !== tb) return ta - tb;
                return (t1.nome || '').localeCompare(t2.nome || '', 'pt-BR', { sensitivity: 'base' });
            });

            // Agrupa por turno e tipo dentro do curso
            const turmasPorTurnoTipo = {};
            turmasCurso.forEach(t => {
                const key = `${t.turno || ''}__${t.tipoTurma || ''}`;
                if (!turmasPorTurnoTipo[key]) turmasPorTurnoTipo[key] = [];
                turmasPorTurnoTipo[key].push(t);
            });

            const chavesGrupo = Object.keys(turmasPorTurnoTipo).sort((ka, kb) => {
                const [turnoA, tipoA] = ka.split('__');
                const [turnoB, tipoB] = kb.split('__');
                const ordemTurno = { MANHA: 1, TARDE: 2, NOITE: 3 };
                const ta = ordemTurno[turnoA] || 99;
                const tb = ordemTurno[turnoB] || 99;
                if (ta !== tb) return ta - tb;
                return (textoTipoTurma(tipoA) || '').localeCompare(textoTipoTurma(tipoB) || '', 'pt-BR', { sensitivity: 'base' });
            });

            const totalLinhasCurso = chavesGrupo.reduce((acc, key) => acc + (turmasPorTurnoTipo[key] || []).length, 0);
            let cursoCellCriado = false;

            chavesGrupo.forEach(key => {
                const [turno, tipoTurma] = key.split('__');
                const turmasDoGrupo = turmasPorTurnoTipo[key];
                
                turmasDoGrupo.forEach((t, idx) => {
                    const aulasTurma = aulas.filter(a => a.turmaId === t.id);
                    const totalAulas = aulasTurma.length;
                    const profsTurma = new Set(aulasTurma.map(a => a.professorId).filter(Boolean)).size;
                    const tr = document.createElement('tr');
                    
                    // Coluna Curso / Área (uma célula mesclada para o curso inteiro)
                    if (!cursoCellCriado) {
                        const tdCurso = document.createElement('td');
                        tdCurso.rowSpan = totalLinhasCurso;
                        tdCurso.style.verticalAlign = 'middle';
                        tdCurso.style.fontWeight = 'bold';
                        tdCurso.style.textAlign = 'center';
                        tdCurso.innerHTML = `
                            <div>${curso}</div>
                            <div style="font-size:0.85em;font-weight:normal;">(${turmasCurso.length} turmas)</div>
                        `;
                        tr.appendChild(tdCurso);
                        cursoCellCriado = true;
                    }

                    // Coluna Turno (apenas na primeira do grupo de turno/tipo)
                    const tdTurno = document.createElement('td');
                    if (idx === 0) {
                        tdTurno.textContent = textoTurno(turno);
                        tdTurno.rowSpan = turmasDoGrupo.length;
                        tdTurno.style.verticalAlign = 'middle';
                        tdTurno.style.fontWeight = 'bold';
                        tdTurno.style.textAlign = 'left';
                        tr.appendChild(tdTurno);
                    } else if (idx === 0) { 
                        // nunca entra aqui, mas garante logica
                    }

                    // Turma
                    const tdTurma = document.createElement('td');
                    tdTurma.innerHTML = `<strong>${t.nome}</strong>`;
                    tdTurma.style.textAlign = 'left';
                    tr.appendChild(tdTurma);

                    const tdTipo = document.createElement('td');
                    if (idx === 0) {
                        tdTipo.textContent = textoTipoTurma(tipoTurma) || '-';
                        tdTipo.rowSpan = turmasDoGrupo.length;
                        tdTipo.style.verticalAlign = 'middle';
                        tdTipo.style.fontWeight = 'bold';
                        tr.appendChild(tdTipo);
                    }

                    const tdAno = document.createElement('td');
                    tdAno.textContent = textoAnoTurma(t.anoTurma) || '-';
                    tr.appendChild(tdAno);

                    const tdFase = document.createElement('td');
                    tdFase.textContent = textoFaseTurma(t.faseTurma) || '-';
                    tr.appendChild(tdFase);

                    // Descrição
                    const tdDesc = document.createElement('td');
                    tdDesc.textContent = t.descricao || '-';
                    tr.appendChild(tdDesc);

                    const tdProfs = document.createElement('td');
                    tdProfs.innerHTML = `<span style="background:#fff8e1;padding:2px 8px;border-radius:12px;">${profsTurma}</span>`;
                    tr.appendChild(tdProfs);

                    const tdTotal = document.createElement('td');
                    tdTotal.innerHTML = `<span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${totalAulas}</span>`;
                    tr.appendChild(tdTotal);

                    tbody.appendChild(tr);
                });
            });
        });
    } 
    else if (tipo === 'turnos') {
        thead.innerHTML = `
            <tr>
                <th>Turno</th>
                <th>Quantidade de Turmas</th>
                <th>Total de Aulas</th>
                <th>Nº Professores</th>
                <th>Cursos Presentes</th>
            </tr>
        `;
        
        const turnosParaExibir = filtroTurno ? [filtroTurno] : TURNOS;
        turnosParaExibir.forEach(turno => {
            const turmasTurno = turmas.filter(t => t.turno === turno);
            const idsTurmas = turmasTurno.map(t => t.id);
            const aulasTurno = aulas.filter(a => idsTurmas.includes(a.turmaId));
            const totalAulas = aulasTurno.length;
            const profsTurno = new Set(aulasTurno.map(a => a.professorId).filter(Boolean)).size;
            const cursosTurno = [...new Set(turmasTurno.map(t => t.curso))];
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${textoTurno(turno)}</strong></td>
                <td><span style="background:#fff8e1;padding:2px 8px;border-radius:12px;">${turmasTurno.length}</span></td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;">${totalAulas}</span></td>
                <td><span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${profsTurno}</span></td>
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
                <th>Nº Professores</th>
                <th>Total de Aulas</th>
            </tr>
        `;
        
        const cursosParaExibir = filtroCurso ? cursos.filter(c => c === filtroCurso) : cursos;
        cursosParaExibir.forEach(curso => {
            const turmasCurso = turmas.filter(t => t.curso === curso);
            const turmasManha = turmasCurso.filter(t => t.turno === 'MANHA').length;
            const turmasTarde = turmasCurso.filter(t => t.turno === 'TARDE').length;
            const turmasNoite = turmasCurso.filter(t => t.turno === 'NOITE').length;
            const idsTurmas = turmasCurso.map(t => t.id);
            const aulasCurso = aulas.filter(a => idsTurmas.includes(a.turmaId));
            const totalAulas = aulasCurso.length;
            const profsCurso = new Set(aulasCurso.map(a => a.professorId).filter(Boolean)).size;
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${curso}</strong></td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;">${turmasManha}</span></td>
                <td><span style="background:#fff8e1;padding:2px 8px;border-radius:12px;">${turmasTarde}</span></td>
                <td><span style="background:#e8f4fc;padding:2px 8px;border-radius:12px;background:#2c3e50;color:white;">${turmasNoite}</span></td>
                <td><span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${turmasCurso.length}</span></td>
                <td><span style="background:#e8f6e8;padding:2px 8px;border-radius:12px;">${profsCurso}</span></td>
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
        temposPorTurno: TEMPOS_POR_TURNO,
        exportadoEm: new Date().toISOString(),
        versao: '1.1'
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
            if (!cursos.includes(CURSO_TEMPO_INTEGRAL)) {
                cursos.push(CURSO_TEMPO_INTEGRAL);
            }
            if (dados.temposPorTurno) {
                TEMPOS_POR_TURNO = dados.temposPorTurno;
            }
            
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

    const selCursoTurma = document.getElementById('cursoTurma');
    if (selCursoTurma) {
        selCursoTurma.addEventListener('change', () => {
            const hiddenTurno = document.getElementById('turnoTurma');
            if (hiddenTurno) hiddenTurno.value = '';
            const chips = document.querySelectorAll('#chips-turno-turma .chip');
            chips.forEach(c => c.classList.remove('selecionado'));
        });
    }

    const formHorario = document.getElementById('formHorario');
    if (formHorario) formHorario.addEventListener('submit', salvarHorario);

    const formNovoCurso = document.getElementById('formNovoCurso');
    if (formNovoCurso) formNovoCurso.addEventListener('submit', salvarNovoCurso);

    const formApelido = document.getElementById('formApelidoDisciplina');
    if (formApelido) formApelido.addEventListener('submit', salvarApelidoDisciplina);

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
    preencherSelectDisciplinasApelido();
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
    atualizarIndicadorDesfazerTurno();

    const selTipoConsulta = document.getElementById('tipoConsulta');
    if (selTipoConsulta) {
        onChangeTipoConsulta();
    }

    const rapidoTurno = document.getElementById('rapidoTurno');
    if (rapidoTurno) {
        rapidoTurno.addEventListener('change', () => {
            atualizarSelectTurmasInsercaoRapida();
            const valor = rapidoTurno.value;
            if (valor) {
                const group = document.getElementById('chips-turno-visual');
                const btn = group ? group.querySelector(`.chip[data-value="${valor}"]`) : null;
                if (btn) {
                    mudarTurnoVisual(btn);
                }
            }
        });
    }

    atualizarSelectTurmasInsercaoRapida();

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
    const baseValor = p.baseTipo || 'TECNICA';
    const hiddenBase = document.getElementById('modalBaseProfessor');
    if (hiddenBase) hiddenBase.value = baseValor;
    const groupBase = document.getElementById('chips-modal-base-prof');
    if (groupBase) {
        const btnSelecionado = Array.from(groupBase.querySelectorAll('.chip'))
            .find(b => b.getAttribute('data-value') === baseValor);
        groupBase.querySelectorAll('.chip').forEach(c => c.classList.remove('selecionado'));
        if (btnSelecionado) {
            btnSelecionado.classList.add('selecionado');
        }
    }
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
    const baseValor = document.getElementById('modalBaseProfessor')?.value || 'TECNICA';
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
        baseTipo: baseValor,
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
    p.baseTipo = novoProf.baseTipo;
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
    const formModalTurma = document.getElementById('formModalTurma');
    if (formModalTurma) {
        formModalTurma.addEventListener('submit', salvarModalTurma);
    }
    const modalTurma = document.getElementById('modalTurma');
    if (modalTurma) {
        modalTurma.addEventListener('click', (e) => {
            if (e.target === modalTurma) fecharModalTurma();
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
            const dispAtual = coletarDisponibilidade(containerId);
            const temposAtuais = coletarTemposDisponibilidade(containerId);
            renderTemposDia(containerId, dia, dispAtual, temposAtuais);
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
        const temMapa = temposMap && temposMap[dia] && Array.isArray(temposMap[dia][t]) && temposMap[dia][t].length > 0;
        const lista = temMapa ? temposMap[dia][t] : getTempos(t).filter(x=>!x.intervalo).map(x=>x.id);
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
