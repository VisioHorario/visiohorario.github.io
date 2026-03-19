const http = require('http');

const PORT = 3001;

function sendJson(res, statusCode, data) {
    const body = JSON.stringify(data);
    res.writeHead(statusCode, {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
    });
    res.end(body);
}

function handleRequest(req, res) {
    if (req.method === 'OPTIONS') {
        res.writeHead(204, {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
        });
        return res.end();
    }

    if (req.method === 'POST' && req.url === '/ajuste-ia') {
        let body = '';
        req.on('data', chunk => {
            body += chunk;
        });
        req.on('end', () => {
            let payload;
            try {
                payload = JSON.parse(body || '{}');
            } catch (e) {
                return sendJson(res, 400, {
                    status: 'erro',
                    mensagem: 'JSON inválido no corpo da requisição.'
                });
            }

            if (!payload || !payload.turno || !Array.isArray(payload.aulas)) {
                return sendJson(res, 400, {
                    status: 'erro',
                    mensagem: 'Payload inválido. Esperado: { escolaId, turno, turmas, aulas, professores }.'
                });
            }

            try {
                const resultado = gerarHorarioIA(payload);
                return sendJson(res, 200, {
                    status: 'ok',
                    mensagem: 'Ajuste IA calculado com sucesso (heurística inicial).',
                    escolaId: payload.escolaId || null,
                    turno: payload.turno,
                    aulas: resultado.aulas
                });
            } catch (e) {
                console.error('Erro no ajuste IA:', e);
                return sendJson(res, 500, {
                    status: 'erro',
                    mensagem: 'Erro interno ao calcular o ajuste IA.'
                });
            }
        });
        return;
    }

    sendJson(res, 404, { status: 'erro', mensagem: 'Rota não encontrada.' });
}

function gerarHorarioIA(payload) {
    const turno = payload.turno;
    const aulas = payload.aulas || [];

    const aulasTurno = aulas.filter(a => a.turno === turno).map(a => ({ ...a }));

    return { aulas: aulasTurno };
}

const server = http.createServer(handleRequest);

server.listen(PORT, () => {
    console.log('Backend IA ouvindo em http://localhost:' + PORT);
});
