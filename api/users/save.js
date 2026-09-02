const {
  adminFetch,
  getRequesterContext,
  json,
  parseJsonResponse,
  readJsonBody,
  sanitizeRole,
  sanitizeShift
} = require('../_lib/supabase-admin');

function montarPayloadAppUser(targetTenantId, authUserId, body, role, shift) {
  return {
    tenant_id: targetTenantId,
    auth_user_id: authUserId,
    name: body.name,
    email: body.email,
    username: body.username || null,
    role,
    shift,
    is_active: body.is_active !== false
  };
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
    const body = await readJsonBody(req);
    const requestedTenantId = String(body.tenantId || '').trim();
    const targetTenantId = context.isSuperAdmin && requestedTenantId
      ? requestedTenantId
      : context.tenantId;
    const name = String(body.name || '').trim();
    const email = String(body.email || '').trim().toLowerCase();
    const username = String(body.username || '').trim();
    const password = String(body.password || '');
    const role = sanitizeRole(body.role);
    const shift = role === 'COORD_TURNO' ? sanitizeShift(body.shift) : null;

    if (!name || !email) {
      return json(res, 400, { error: 'Nome e e-mail são obrigatórios.' });
    }
    if (!targetTenantId) {
      return json(res, 400, { error: 'Selecione uma escola válida para o usuário.' });
    }
    if (role === 'COORD_TURNO' && !shift) {
      return json(res, 400, { error: 'Selecione um turno válido para o coordenador.' });
    }

    let authUserId = body.authUserId || null;
    let created = false;
    const userMetadata = {
      tenant_id: targetTenantId,
      name,
      shift
    };

    if (!authUserId) {
      if (password.length < 6) {
        return json(res, 400, { error: 'A senha do novo usuário deve ter pelo menos 6 caracteres.' });
      }
      const createResponse = await adminFetch('/auth/v1/admin/users', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          email,
          password,
          email_confirm: true,
          user_metadata: userMetadata
        })
      });
      const createData = await parseJsonResponse(createResponse);
      if (!createResponse.ok || !createData || !createData.id) {
        return json(res, 400, { error: createData?.msg || createData?.error || 'Não foi possível criar o usuário no Auth.' });
      }
      authUserId = createData.id;
      created = true;
    } else {
      const updatePayload = {
        email,
        user_metadata: userMetadata
      };
      if (password) {
        if (password.length < 6) {
          return json(res, 400, { error: 'A nova senha deve ter pelo menos 6 caracteres.' });
        }
        updatePayload.password = password;
      }
      const updateResponse = await adminFetch(`/auth/v1/admin/users/${authUserId}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(updatePayload)
      });
      const updateData = await parseJsonResponse(updateResponse);
      if (!updateResponse.ok) {
        return json(res, 400, { error: updateData?.msg || updateData?.error || 'Não foi possível atualizar o usuário no Auth.' });
      }
    }

    const appUserPayload = montarPayloadAppUser(targetTenantId, authUserId, {
      name,
      email,
      username,
      is_active: body.is_active !== false
    }, role, shift);

    const appUserResponse = await adminFetch('/rest/v1/app_users?on_conflict=auth_user_id', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Prefer: 'resolution=merge-duplicates,return=representation'
      },
      body: JSON.stringify(appUserPayload)
    });
    const appUserData = await parseJsonResponse(appUserResponse);
    if (!appUserResponse.ok) {
      return json(res, 400, { error: appUserData?.message || appUserData?.error || 'Não foi possível salvar o usuário da escola.' });
    }

    return json(res, 200, {
      created,
      user: Array.isArray(appUserData) ? appUserData[0] : appUserData
    });
  } catch (error) {
    return json(res, error.status || 500, { error: error.message || 'Erro ao salvar usuário.' });
  }
};
