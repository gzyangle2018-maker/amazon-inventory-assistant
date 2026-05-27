// Cloudflare Worker Backend for Amazon Inventory Assistant
// Handles auth, data persistence, and admin APIs

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
  });
}

function errorResponse(message, status = 400) {
  return jsonResponse({ error: message }, status);
}

// Simple SHA-256 hash
async function hashPassword(pw) {
  const encoder = new TextEncoder();
  const data = encoder.encode(pw);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// JWT-like token (signed with worker secret)
async function signToken(payload, secret) {
  const header = btoa(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const body = btoa(JSON.stringify(payload));
  const data = new TextEncoder().encode(`${header}.${body}`);
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const signature = await crypto.subtle.sign('HMAC', key, data);
  const sig = btoa(String.fromCharCode(...new Uint8Array(signature)));
  return `${header}.${body}.${sig}`;
}

async function verifyToken(token, secret) {
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [header, body, sig] = parts;
  const data = new TextEncoder().encode(`${header}.${body}`);
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']);
  const signature = Uint8Array.from(atob(sig), c => c.charCodeAt(0));
  const valid = await crypto.subtle.verify('HMAC', key, signature, data);
  if (!valid) return null;
  return JSON.parse(atob(body));
}

// Auth middleware
async function getUser(request, env) {
  const auth = request.headers.get('Authorization');
  if (!auth || !auth.startsWith('Bearer ')) return null;
  const token = auth.slice(7);
  return await verifyToken(token, env.JWT_SECRET || 'default-secret-change-me');
}

// ========== API Routes ==========

async function handleLogin(request, env) {
  const { username, password } = await request.json();
  if (!username || !password) return errorResponse('用户名和密码不能为空');

  const user = await env.DB.prepare('SELECT * FROM users WHERE username = ?').bind(username).first();
  if (!user || user.is_active !== 1) {
    await env.DB.prepare('INSERT INTO login_logs (username, login_time, success) VALUES (?, datetime("now"), 0)').bind(username).run();
    return errorResponse('用户名或密码错误，或账号已禁用', 401);
  }

  const hash = await hashPassword(password);
  if (hash !== user.password_hash) {
    await env.DB.prepare('INSERT INTO login_logs (username, login_time, success) VALUES (?, datetime("now"), 0)').bind(username).run();
    return errorResponse('用户名或密码错误', 401);
  }

  const token = await signToken({ username: user.username, role: user.role }, env.JWT_SECRET || 'default-secret-change-me');
  await env.DB.prepare('INSERT INTO login_logs (username, login_time, success) VALUES (?, datetime("now"), 1)').bind(username).run();

  return jsonResponse({ token, username: user.username, role: user.role });
}

async function handleRegister(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  const { username, password, role = 'user', department = '', created_by = '' } = await request.json();
  if (!username || !password) return errorResponse('用户名和密码不能为空');

  const hash = await hashPassword(password);
  try {
    await env.DB.prepare(
      'INSERT INTO users (username, password_hash, role, created_at, is_active, department, created_by) VALUES (?, ?, ?, datetime("now"), 1, ?, ?)'
    ).bind(username, hash, role, department, created_by).run();
    return jsonResponse({ success: true, message: '用户创建成功' });
  } catch (e) {
    return errorResponse('用户名已存在');
  }
}

async function handleUsers(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  const { results } = await env.DB.prepare('SELECT id, username, role, created_at, is_active, machine_limit, department, created_by FROM users ORDER BY id').all();
  return jsonResponse(results || []);
}

async function handleToggleUser(request, env, username) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  const { active } = await request.json();
  await env.DB.prepare('UPDATE users SET is_active = ? WHERE username = ? AND role != "admin"').bind(active ? 1 : 0, username).run();
  return jsonResponse({ success: true });
}

async function handleDeleteUser(request, env, username) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  await env.DB.prepare('DELETE FROM users WHERE username = ? AND role != "admin"').bind(username).run();
  return jsonResponse({ success: true });
}

async function handleHistory(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  let query;
  if (currentUser.role === 'admin') {
    query = env.DB.prepare('SELECT * FROM upload_history ORDER BY id DESC LIMIT 200');
  } else {
    query = env.DB.prepare('SELECT * FROM upload_history WHERE username = ? ORDER BY id DESC LIMIT 200').bind(currentUser.username);
  }
  const { results } = await query.all();
  return jsonResponse(results || []);
}

async function handleSaveHistory(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  const { filename, row_count } = await request.json();
  await env.DB.prepare(
    'INSERT INTO upload_history (username, filename, uploaded_at, row_count) VALUES (?, ?, datetime("now"), ?)'
  ).bind(currentUser.username, filename, row_count).run();
  return jsonResponse({ success: true });
}

async function handleSeckill(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  if (request.method === 'GET') {
    let query;
    if (currentUser.role === 'admin') {
      query = env.DB.prepare('SELECT * FROM seckill_reports ORDER BY id DESC LIMIT 200');
    } else {
      query = env.DB.prepare('SELECT * FROM seckill_reports WHERE username = ? ORDER BY id DESC LIMIT 200').bind(currentUser.username);
    }
    const { results } = await query.all();
    return jsonResponse(results || []);
  }

  if (request.method === 'POST') {
    const { items, ziniao_info } = await request.json();
    await env.DB.prepare(
      'INSERT INTO seckill_reports (username, created_at, items, ziniao_info) VALUES (?, datetime("now"), ?, ?)'
    ).bind(currentUser.username, JSON.stringify(items), JSON.stringify(ziniao_info)).run();
    return jsonResponse({ success: true });
  }
}

async function handleVersions(request, env) {
  if (request.method === 'GET') {
    const { results } = await env.DB.prepare('SELECT * FROM versions ORDER BY created_at DESC').all();
    return jsonResponse(results || []);
  }

  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  if (request.method === 'POST') {
    const { version, description = '' } = await request.json();
    try {
      await env.DB.prepare('INSERT INTO versions (version, description, is_active, created_at) VALUES (?, ?, 1, datetime("now"))').bind(version, description).run();
      return jsonResponse({ success: true });
    } catch (e) {
      return errorResponse('版本号已存在');
    }
  }

  if (request.method === 'PUT') {
    const { id, active } = await request.json();
    await env.DB.prepare('UPDATE versions SET is_active = ? WHERE id = ?').bind(active ? 1 : 0, id).run();
    return jsonResponse({ success: true });
  }

  if (request.method === 'DELETE') {
    const { id } = await request.json();
    await env.DB.prepare('DELETE FROM versions WHERE id = ?').bind(id).run();
    return jsonResponse({ success: true });
  }
}

async function handleCheckVersion(request, env) {
  const { version } = await request.json();
  const count = await env.DB.prepare('SELECT COUNT(*) as count FROM versions').first();

  if (!count || count.count === 0) {
    return jsonResponse({ allowed: true });
  }

  const ver = await env.DB.prepare('SELECT is_active FROM versions WHERE version = ?').bind(version).first();
  if (ver) {
    return jsonResponse({ allowed: ver.is_active === 1 });
  }

  const activeCount = await env.DB.prepare('SELECT COUNT(*) as count FROM versions WHERE is_active = 1').first();
  return jsonResponse({ allowed: activeCount && activeCount.count === 0 });
}

async function handleLogs(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  const { results } = await env.DB.prepare('SELECT * FROM login_logs ORDER BY id DESC LIMIT 200').all();
  return jsonResponse(results || []);
}

async function handleLLMConfig(request, env) {
  if (request.method === 'GET') {
    const config = await env.DB.prepare('SELECT api_key, base_url, model_name FROM llm_config WHERE id = 1').first();
    return jsonResponse(config || { api_key: '', base_url: '', model_name: '' });
  }

  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  const { api_key, base_url, model_name } = await request.json();
  const exists = await env.DB.prepare('SELECT COUNT(*) as count FROM llm_config').first();
  if (exists && exists.count > 0) {
    await env.DB.prepare('UPDATE llm_config SET api_key = ?, base_url = ?, model_name = ?, updated_at = datetime("now") WHERE id = 1').bind(api_key, base_url, model_name).run();
  } else {
    await env.DB.prepare('INSERT INTO llm_config (api_key, base_url, model_name, updated_at) VALUES (?, ?, ?, datetime("now"))').bind(api_key, base_url, model_name).run();
  }
  return jsonResponse({ success: true });
}

// ========== Main Router ==========

export default {
  async fetch(request, env, ctx) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    const url = new URL(request.url);
    const path = url.pathname;

    try {
      // API Routes
      if (path === '/api/login' && request.method === 'POST') return await handleLogin(request, env);
      if (path === '/api/register' && request.method === 'POST') return await handleRegister(request, env);
      if (path === '/api/users') {
        if (request.method === 'GET') return await handleUsers(request, env);
      }
      if (path.startsWith('/api/users/')) {
        const username = decodeURIComponent(path.slice(11));
        if (request.method === 'DELETE') return await handleDeleteUser(request, env, username);
        if (request.method === 'PUT') return await handleToggleUser(request, env, username);
      }
      if (path === '/api/history') {
        if (request.method === 'GET') return await handleHistory(request, env);
        if (request.method === 'POST') return await handleSaveHistory(request, env);
      }
      if (path === '/api/seckill') return await handleSeckill(request, env);
      if (path === '/api/versions') return await handleVersions(request, env);
      if (path === '/api/check-version' && request.method === 'POST') return await handleCheckVersion(request, env);
      if (path === '/api/logs' && request.method === 'GET') return await handleLogs(request, env);
      if (path === '/api/llm-config') return await handleLLMConfig(request, env);

      // Static files from R2 or KV would go here
      // For Pages deployment, this is handled automatically
      return errorResponse('Not found', 404);
    } catch (e) {
      return errorResponse(e.message || 'Internal error', 500);
    }
  },
};
