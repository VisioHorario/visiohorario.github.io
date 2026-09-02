const ROLES_VALIDOS = new Set(['ADMIN', 'COORD_TURNO']);
const TURNOS_VALIDOS = new Set(['MANHA', 'TARDE', 'NOITE']);

function getConfig() {
  const url = process.env.VISIO_SUPABASE_URL || '';
  const anonKey = process.env.VISIO_SUPABASE_ANON_KEY || '';
  const serviceRoleKey = process.env.VISIO_SUPABASE_SERVICE_ROLE_KEY || '';
  if (!url || !serviceRoleKey) {
    throw new Error('Variáveis do Supabase incompletas no backend.');
  }
  return { url, anonKey, serviceRoleKey };
}

function json(res, status, payload) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.status(status).json(payload);
}

async function readJsonBody(req) {
  if (!req.body) return {};
  if (typeof req.body === 'object') return req.body;
  try {
    return JSON.parse(req.body);
  } catch (error) {
    return {};
  }
}

function getBearerToken(req) {
  const header = req.headers.authorization || req.headers.Authorization || '';
  if (!header.startsWith('Bearer ')) return '';
  return header.slice('Bearer '.length).trim();
}

async function parseJsonResponse(response) {
  const text = await response.text();
  try {
    return text ? JSON.parse(text) : null;
  } catch (error) {
    return text || null;
  }
}

async function adminFetch(path, options = {}, config = getConfig()) {
  const response = await fetch(`${config.url}${path}`, {
    ...options,
    headers: {
      apikey: config.serviceRoleKey,
      Authorization: `Bearer ${config.serviceRoleKey}`,
      ...(options.headers || {})
    }
  });
  return response;
}

async function getRequesterContext(req) {
  const config = getConfig();
  const token = getBearerToken(req);
  if (!token) {
    const error = new Error('Sessão ausente.');
    error.status = 401;
    throw error;
  }

  const userResponse = await fetch(`${config.url}/auth/v1/user`, {
    headers: {
      apikey: config.anonKey || config.serviceRoleKey,
      Authorization: `Bearer ${token}`
    }
  });
  const authUser = await parseJsonResponse(userResponse);
  if (!userResponse.ok || !authUser || !authUser.id) {
    const error = new Error('Sessão inválida ou expirada.');
    error.status = 401;
    throw error;
  }

  const tenantId = authUser.user_metadata?.tenant_id || null;
  if (!tenantId) {
    const error = new Error('Usuário sem tenant vinculado.');
    error.status = 403;
    throw error;
  }

  const appUserResponse = await adminFetch(
    `/rest/v1/app_users?auth_user_id=eq.${encodeURIComponent(authUser.id)}&tenant_id=eq.${encodeURIComponent(tenantId)}&is_active=eq.true&select=id,tenant_id,auth_user_id,role,shift,name,email,username,is_active`,
    {
      headers: {
        Accept: 'application/json'
      }
    },
    config
  );
  const appUsers = await parseJsonResponse(appUserResponse);
  const appUser = Array.isArray(appUsers) ? appUsers[0] : null;
  if (!appUser || !['ADMIN', 'SUPER_ADMIN'].includes(appUser.role)) {
    const error = new Error('Apenas administradores podem acessar esta área.');
    error.status = 403;
    throw error;
  }

  return {
    config,
    authUser,
    appUser,
    tenantId,
    isSuperAdmin: appUser.role === 'SUPER_ADMIN'
  };
}

function sanitizeRole(role) {
  return ROLES_VALIDOS.has(role) ? role : 'ADMIN';
}

function sanitizeShift(shift) {
  return TURNOS_VALIDOS.has(shift) ? shift : null;
}

module.exports = {
  adminFetch,
  getConfig,
  getRequesterContext,
  json,
  parseJsonResponse,
  readJsonBody,
  sanitizeRole,
  sanitizeShift
};
