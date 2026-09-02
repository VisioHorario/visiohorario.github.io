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
    const requestedTenantId = String(req.query?.tenantId || '').trim();
    const targetTenantId = context.isSuperAdmin && requestedTenantId
      ? requestedTenantId
      : context.tenantId;
    const response = await adminFetch(
      `/rest/v1/app_users?tenant_id=eq.${encodeURIComponent(targetTenantId)}&select=id,tenant_id,auth_user_id,name,email,username,role,shift,is_active,created_at,updated_at&order=created_at.asc`
    );
    const users = await parseJsonResponse(response);
    if (!response.ok) {
      return json(res, 500, { error: 'Não foi possível listar os usuários da escola.' });
    }
    return json(res, 200, { users: Array.isArray(users) ? users : [] });
  } catch (error) {
    return json(res, error.status || 500, { error: error.message || 'Erro ao listar usuários.' });
  }
};
