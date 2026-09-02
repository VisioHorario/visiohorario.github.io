const {
  adminFetch,
  getRequesterContext,
  json,
  parseJsonResponse,
  readJsonBody
} = require('../_lib/supabase-admin');

function normalizeTenantCode(code) {
  return String(code || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}

module.exports = async (req, res) => {
  if (req.method === 'OPTIONS') {
    return json(res, 200, { ok: true });
  }
  if (req.method !== 'POST') {
    return json(res, 405, { error: 'Método não permitido.' });
  }

  try {
    const context = await getRequesterContext(req);
    if (!context.isSuperAdmin) {
      return json(res, 403, { error: 'Apenas o superadministrador pode cadastrar escolas.' });
    }

    const body = await readJsonBody(req);
    const tenantCode = normalizeTenantCode(body.code);
    const tenantName = String(body.name || '').trim();
    const adminName = String(body.adminName || '').trim();
    const adminEmail = String(body.adminEmail || '').trim().toLowerCase();
    const adminUsername = String(body.adminUsername || '').trim();
    const adminPassword = String(body.adminPassword || '');

    if (!tenantCode || !tenantName) {
      return json(res, 400, { error: 'Código e nome da escola são obrigatórios.' });
    }
    if (tenantCode === 'GLOBAL') {
      return json(res, 400, { error: 'O código GLOBAL é reservado ao superadministrador.' });
    }
    if (!adminName || !adminEmail) {
      return json(res, 400, { error: 'Informe nome e e-mail do administrador inicial.' });
    }
    if (adminPassword.length < 6) {
      return json(res, 400, { error: 'A senha do administrador inicial deve ter pelo menos 6 caracteres.' });
    }

    const tenantResponse = await adminFetch('/rest/v1/tenants', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Prefer: 'return=representation'
      },
      body: JSON.stringify({
        code: tenantCode,
        name: tenantName
      })
    });
    const tenantData = await parseJsonResponse(tenantResponse);
    const tenant = Array.isArray(tenantData) ? tenantData[0] : tenantData;
    if (!tenantResponse.ok || !tenant || !tenant.id) {
      return json(res, 400, {
        error: tenantData?.message || tenantData?.error || 'Não foi possível cadastrar a escola.'
      });
    }

    const userMetadata = {
      tenant_id: tenant.id,
      name: adminName,
      shift: null
    };

    const authResponse = await adminFetch('/auth/v1/admin/users', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        email: adminEmail,
        password: adminPassword,
        email_confirm: true,
        user_metadata: userMetadata
      })
    });
    const authData = await parseJsonResponse(authResponse);
    if (!authResponse.ok || !authData || !authData.id) {
      return json(res, 400, {
        error: authData?.msg || authData?.error || 'Não foi possível criar o usuário administrador no Auth.'
      });
    }

    const appUserResponse = await adminFetch('/rest/v1/app_users', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Prefer: 'return=representation'
      },
      body: JSON.stringify({
        tenant_id: tenant.id,
        auth_user_id: authData.id,
        name: adminName,
        email: adminEmail,
        username: adminUsername || null,
        role: 'ADMIN',
        shift: null,
        is_active: true
      })
    });
    const appUserData = await parseJsonResponse(appUserResponse);
    if (!appUserResponse.ok) {
      return json(res, 400, {
        error: appUserData?.message || appUserData?.error || 'A escola foi criada, mas houve erro ao vincular o administrador.'
      });
    }

    return json(res, 200, {
      tenant,
      adminUser: Array.isArray(appUserData) ? appUserData[0] : appUserData
    });
  } catch (error) {
    return json(res, error.status || 500, { error: error.message || 'Erro ao salvar escola.' });
  }
};
