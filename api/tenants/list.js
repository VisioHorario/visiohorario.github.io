const {
  adminFetch,
  getRequesterContext,
  json,
  parseJsonResponse
} = require('../_lib/supabase-admin');

module.exports = async (req, res) => {
  if (req.method === 'OPTIONS') {
    return json(res, 200, { ok: true });
  }
  if (req.method !== 'GET') {
    return json(res, 405, { error: 'Método não permitido.' });
  }

  try {
    const context = await getRequesterContext(req);
    if (!context.isSuperAdmin) {
      return json(res, 403, { error: 'Apenas o superadministrador pode listar escolas.' });
    }

    const response = await adminFetch(
      '/rest/v1/tenants?code=neq.GLOBAL&select=id,code,name,created_at,updated_at&order=created_at.asc'
    );
    const tenants = await parseJsonResponse(response);
    if (!response.ok) {
      return json(res, 500, { error: 'Não foi possível listar as escolas.' });
    }

    return json(res, 200, { tenants: Array.isArray(tenants) ? tenants : [] });
  } catch (error) {
    return json(res, error.status || 500, { error: error.message || 'Erro ao listar escolas.' });
  }
};
