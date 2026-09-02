const {
  adminFetch,
  json,
  parseJsonResponse
} = require('../_lib/supabase-admin');

function getTenantCode(req) {
  const raw = req.query?.tenant
    || req.query?.escola
    || req.query?.code
    || '';
  return String(raw).trim().toUpperCase();
}

module.exports = async (req, res) => {
  if (req.method === 'OPTIONS') {
    return json(res, 200, { ok: true });
  }
  if (req.method !== 'GET') {
    return json(res, 405, { error: 'Método não permitido.' });
  }

  try {
    const tenantCode = getTenantCode(req);
    if (!tenantCode) {
      return json(res, 400, { error: 'Informe o código da escola no parâmetro tenant.' });
    }
    if (tenantCode === 'GLOBAL') {
      return json(res, 403, { error: 'O tenant GLOBAL não possui consulta pública por QR Code.' });
    }

    const tenantResponse = await adminFetch(
      `/rest/v1/tenants?code=eq.${encodeURIComponent(tenantCode)}&select=id,code,name&limit=1`
    );
    const tenantData = await parseJsonResponse(tenantResponse);
    const tenant = Array.isArray(tenantData) ? tenantData[0] : null;
    if (!tenantResponse.ok || !tenant || !tenant.id) {
      return json(res, 404, { error: 'Escola não encontrada para o QR Code informado.' });
    }

    const snapshotResponse = await adminFetch(
      `/rest/v1/visio_snapshots?tenant_id=eq.${encodeURIComponent(tenant.id)}&select=payload,updated_at&limit=1`
    );
    const snapshotData = await parseJsonResponse(snapshotResponse);
    const snapshot = Array.isArray(snapshotData) ? snapshotData[0] : null;
    if (!snapshotResponse.ok || !snapshot || !snapshot.payload) {
      return json(res, 404, { error: 'Não há dados publicados para esta escola no QR Code.' });
    }

    return json(res, 200, {
      tenant,
      updatedAt: snapshot.updated_at || null,
      payload: snapshot.payload
    });
  } catch (error) {
    return json(res, 500, { error: error.message || 'Erro ao carregar dados públicos do QR Code.' });
  }
};
