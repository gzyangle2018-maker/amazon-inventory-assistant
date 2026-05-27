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

  // Load user permissions
  let permissions = null;
  try { permissions = user.permissions_json ? JSON.parse(user.permissions_json) : null; } catch(e) {}
  return jsonResponse({ token, username: user.username, role: user.role, permissions });
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

  const { results } = await env.DB.prepare('SELECT id, username, role, created_at, is_active, machine_limit, department, created_by, permissions_json FROM users ORDER BY id').all();
  // Parse permissions for each user
  for (const u of (results || [])) {
    try { u.permissions = u.permissions_json ? JSON.parse(u.permissions_json) : null; } catch(e) { u.permissions = null; }
    delete u.permissions_json;
  }
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
  const result = await env.DB.prepare(
    'INSERT INTO upload_history (username, filename, uploaded_at, row_count) VALUES (?, ?, datetime("now"), ?)'
  ).bind(currentUser.username, filename, row_count).run();
  // Return the inserted ID so frontend can save file data
  const lastId = result.meta?.last_row_id || 0;
  return jsonResponse({ success: true, id: lastId });
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

// ========== Column Mappings ==========

async function handleGetColumnMappings(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  // Try to get user-specific mappings first, fall back to global
  let result = {};
  try {
    const { results } = await env.DB.prepare(
      'SELECT map_type, columns_json FROM column_mappings WHERE username = ? OR username = \'global\' ORDER BY username DESC'
    ).bind(currentUser.username).all();
    for (const row of (results || [])) {
      if (!result[row.map_type]) {
        try { result[row.map_type] = JSON.parse(row.columns_json || '[]'); }
        catch(e) { result[row.map_type] = []; }
      }
    }
  } catch(e) {
    // Table might not exist yet
  }
  return jsonResponse(result);
}

async function handleSaveColumnMappings(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  const mappings = await request.json();
  const now = new Date().toISOString();

  for (const [mapType, colsList] of Object.entries(mappings)) {
    const colsJson = JSON.stringify(colsList || []);
    try {
      const existing = await env.DB.prepare(
        'SELECT id FROM column_mappings WHERE username = ? AND map_type = ?'
      ).bind(currentUser.username, mapType).first();

      if (existing) {
        await env.DB.prepare(
          'UPDATE column_mappings SET columns_json = ?, updated_by = ?, updated_at = ? WHERE username = ? AND map_type = ?'
        ).bind(colsJson, currentUser.username, now, currentUser.username, mapType).run();
      } else {
        await env.DB.prepare(
          'INSERT INTO column_mappings (username, map_type, columns_json, updated_by, updated_at) VALUES (?, ?, ?, ?, ?)'
        ).bind(currentUser.username, mapType, colsJson, currentUser.username, now).run();
      }
    } catch(e) {
      // Table might not exist, try creating it
      try {
        await env.DB.prepare(
          'CREATE TABLE IF NOT EXISTS column_mappings (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL DEFAULT \'global\', map_type TEXT NOT NULL, columns_json TEXT, updated_by TEXT DEFAULT \'\', updated_at TEXT, UNIQUE(username, map_type))'
        ).run();
        await env.DB.prepare(
          'INSERT INTO column_mappings (username, map_type, columns_json, updated_by, updated_at) VALUES (?, ?, ?, ?, ?)'
        ).bind(currentUser.username, mapType, colsJson, currentUser.username, now).run();
      } catch(e2) { /* ignore */ }
    }
  }
  return jsonResponse({ success: true });
}

// ========== Upload File Data (for admin online viewing) ==========

async function handleSaveUploadData(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  const { history_id, headers, data, green_rows } = await request.json();
  if (!history_id || !headers || !data) return errorResponse('参数不完整');

  const headersJson = JSON.stringify(headers);
  const dataJson = JSON.stringify(data);
  const greenRowsJson = JSON.stringify(green_rows || []);
  const now = new Date().toISOString();

  try {
    // Delete old data if exists, then insert
    await env.DB.prepare('DELETE FROM upload_file_data WHERE history_id = ?').bind(history_id).run();
    await env.DB.prepare(
      'INSERT INTO upload_file_data (history_id, headers_json, data_json, green_rows_json, created_at) VALUES (?, ?, ?, ?, ?)'
    ).bind(history_id, headersJson, dataJson, greenRowsJson, now).run();
  } catch(e) {
    // Table might not exist, try creating it
    try {
      await env.DB.prepare(
        'CREATE TABLE IF NOT EXISTS upload_file_data (id INTEGER PRIMARY KEY AUTOINCREMENT, history_id INTEGER NOT NULL UNIQUE, headers_json TEXT NOT NULL, data_json TEXT NOT NULL, green_rows_json TEXT DEFAULT \'[]\', created_at TEXT)'
      ).run();
      await env.DB.prepare(
        'INSERT INTO upload_file_data (history_id, headers_json, data_json, green_rows_json, created_at) VALUES (?, ?, ?, ?, ?)'
      ).bind(history_id, headersJson, dataJson, greenRowsJson, now).run();
    } catch(e2) {
      return errorResponse('保存失败: ' + e2.message);
    }
  }
  return jsonResponse({ success: true });
}

async function handleGetUploadData(request, env, historyId) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);

  try {
    const row = await env.DB.prepare(
      'SELECT headers_json, data_json, green_rows_json FROM upload_file_data WHERE history_id = ?'
    ).bind(historyId).first();

    if (!row) return errorResponse('未找到该文件数据', 404);

    return jsonResponse({
      headers: JSON.parse(row.headers_json || '[]'),
      data: JSON.parse(row.data_json || '[]'),
      green_rows: JSON.parse(row.green_rows_json || '[]'),
    });
  } catch(e) {
    return errorResponse('读取失败: ' + e.message);
  }
}

// ========== Restock Config (Weighted Algorithm) ==========

async function handleGetRestockConfig(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);
  try {
    const row = await env.DB.prepare('SELECT * FROM restock_config WHERE id = 1').first();
    if (!row) return jsonResponse({
      stock_months: 4, weight_3d: 20, weight_7d: 30, weight_15d: 30, weight_30d: 20,
      deduct_unshipped: 0, deduct_week_outbound: 0
    });
    return jsonResponse({
      stock_months: row.stock_months || 4,
      weight_3d: row.weight_3d || 20, weight_7d: row.weight_7d || 30,
      weight_15d: row.weight_15d || 30, weight_30d: row.weight_30d || 20,
      deduct_unshipped: row.deduct_unshipped || 0, deduct_week_outbound: row.deduct_week_outbound || 0
    });
  } catch(e) {
    return jsonResponse({ stock_months: 4, weight_3d: 20, weight_7d: 30, weight_15d: 30, weight_30d: 20, deduct_unshipped: 0, deduct_week_outbound: 0 });
  }
}

async function handleSaveRestockConfig(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);
  const cfg = await request.json();
  const total = (cfg.weight_3d||0) + (cfg.weight_7d||0) + (cfg.weight_15d||0) + (cfg.weight_30d||0);
  if (total !== 100) return errorResponse(`权重总和必须为100%，当前为${total}%`);
  try {
    const existing = await env.DB.prepare('SELECT id FROM restock_config WHERE id = 1').first();
    if (existing) {
      await env.DB.prepare(
        'UPDATE restock_config SET stock_months=?, weight_3d=?, weight_7d=?, weight_15d=?, weight_30d=?, deduct_unshipped=?, deduct_week_outbound=?, updated_at=datetime("now") WHERE id=1'
      ).bind(cfg.stock_months||4, cfg.weight_3d||20, cfg.weight_7d||30, cfg.weight_15d||30, cfg.weight_30d||20,
        cfg.deduct_unshipped?1:0, cfg.deduct_week_outbound?1:0).run();
    } else {
      await env.DB.prepare(
        'INSERT INTO restock_config (id, stock_months, weight_3d, weight_7d, weight_15d, weight_30d, deduct_unshipped, deduct_week_outbound, updated_at) VALUES (1,?,?,?,?,?,?,?,datetime("now"))'
      ).bind(cfg.stock_months||4, cfg.weight_3d||20, cfg.weight_7d||30, cfg.weight_15d||30, cfg.weight_30d||20,
        cfg.deduct_unshipped?1:0, cfg.deduct_week_outbound?1:0).run();
    }
    return jsonResponse({ success: true });
  } catch(e) {
    return errorResponse('保存失败: ' + e.message);
  }
}

// ========== Stocking Rules ==========

async function handleStockingRules(request, env) {
  if (request.method === 'GET') {
    const currentUser = await getUser(request, env);
    if (!currentUser) return errorResponse('未登录', 401);
    try {
      const { results } = await env.DB.prepare('SELECT * FROM stocking_rules ORDER BY priority DESC, id ASC').all();
      return jsonResponse(results || []);
    } catch(e) { return jsonResponse([]); }
  }
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  if (request.method === 'POST') {
    const body = await request.json();
    // Support bulk save: { rules: [...] } or single add: { min_monthly_sales, ... }
    if (body.rules && Array.isArray(body.rules)) {
      // Bulk save: delete all existing and re-insert
      try {
        await env.DB.prepare('DELETE FROM stocking_rules').run();
        for (const rule of body.rules) {
          await env.DB.prepare(
            'INSERT INTO stocking_rules (name, min_monthly_sales, max_monthly_sales, stock_multiplier, priority, is_active, created_at) VALUES (?,?,?,?,?,?,datetime("now"))'
          ).bind(rule.name || '', rule.min_monthly_sales || 0, rule.max_monthly_sales || 99999,
            rule.multiplier || rule.stock_multiplier || 2.0, rule.priority || 0, 1).run();
        }
        return jsonResponse({ success: true, count: body.rules.length });
      } catch(e) { return errorResponse('批量保存失败: ' + e.message); }
    }
    // Single add
    try {
      await env.DB.prepare(
        'INSERT INTO stocking_rules (name, min_monthly_sales, max_monthly_sales, stock_multiplier, priority, is_active, created_at) VALUES (?,?,?,?,?,?,datetime("now"))'
      ).bind(body.name||'', body.min_monthly_sales||0, body.max_monthly_sales||0,
        body.multiplier || body.stock_multiplier || 2.0, body.priority||0, body.is_active!==undefined?body.is_active:1).run();
      return jsonResponse({ success: true });
    } catch(e) { return errorResponse('添加失败: ' + e.message); }
  }
  if (request.method === 'PUT') {
    const rule = await request.json();
    await env.DB.prepare(
      'UPDATE stocking_rules SET name=?, min_monthly_sales=?, max_monthly_sales=?, stock_multiplier=?, priority=?, is_active=? WHERE id=?'
    ).bind(rule.name||'', rule.min_monthly_sales||0, rule.max_monthly_sales||0,
      rule.multiplier || rule.stock_multiplier || 2.0, rule.priority||0, rule.is_active!==undefined?rule.is_active:1, rule.id).run();
    return jsonResponse({ success: true });
  }
  if (request.method === 'DELETE') {
    const { id } = await request.json();
    await env.DB.prepare('DELETE FROM stocking_rules WHERE id = ?').bind(id).run();
    return jsonResponse({ success: true });
  }
}

// ========== User Permissions ==========

async function handleUserPermissions(request, env, username) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  if (request.method === 'GET') {
    const user = await env.DB.prepare('SELECT permissions_json FROM users WHERE username = ?').bind(username).first();
    if (!user) return errorResponse('用户不存在', 404);
    let perms = null;
    try { perms = user.permissions_json ? JSON.parse(user.permissions_json) : null; } catch(e) {}
    return jsonResponse({ username, permissions: perms });
  }

  if (request.method === 'PUT') {
    const { permissions } = await request.json();
    const permsJson = permissions ? JSON.stringify(permissions) : null;
    await env.DB.prepare('UPDATE users SET permissions_json = ? WHERE username = ?').bind(permsJson, username).run();
    return jsonResponse({ success: true });
  }
}

// ========== Additional API Routes ==========

async function handleMe(request, env) {
  const user = await getUser(request, env);
  if (!user) return errorResponse('未登录', 401);
  const dbUser = await env.DB.prepare('SELECT username, role, is_active FROM users WHERE username = ?').bind(user.username).first();
  if (!dbUser || dbUser.is_active !== 1) return errorResponse('用户不存在或已禁用', 401);
  return jsonResponse({ username: dbUser.username, role: dbUser.role });
}

async function handleResetPassword(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);

  const { username, new_password } = await request.json();
  if (!username || !new_password) return errorResponse('用户名和新密码不能为空');
  if (new_password.length < 4) return errorResponse('密码至少4个字符');

  const hash = await hashPassword(new_password);
  await env.DB.prepare('UPDATE users SET password_hash = ? WHERE username = ? AND role != "admin"').bind(hash, username).run();
  return jsonResponse({ success: true, message: '密码已重置' });
}

async function handleDeleteSeckill(request, env, id) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);
  await env.DB.prepare('DELETE FROM seckill_reports WHERE id = ?').bind(id).run();
  return jsonResponse({ success: true });
}

// ========== Admin Notes ==========

async function ensureNotesTable(env) {
  try {
    await env.DB.prepare(
      'CREATE TABLE IF NOT EXISTS admin_notes (id INTEGER PRIMARY KEY AUTOINCREMENT, history_id INTEGER NOT NULL, author TEXT NOT NULL, content TEXT NOT NULL, note_type TEXT DEFAULT "info", target_user TEXT DEFAULT "", created_at TEXT)'
    ).run();
  } catch(e) { /* table exists */ }
}

async function handleGetNotes(request, env, historyId) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);
  await ensureNotesTable(env);
  try {
    const { results } = await env.DB.prepare(
      'SELECT * FROM admin_notes WHERE history_id = ? ORDER BY id DESC'
    ).bind(historyId).all();
    return jsonResponse(results || []);
  } catch(e) { return jsonResponse([]); }
}

async function handleGetNoteCounts(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);
  await ensureNotesTable(env);
  try {
    const { results } = await env.DB.prepare(
      'SELECT history_id, COUNT(*) as count FROM admin_notes GROUP BY history_id'
    ).all();
    const map = {};
    for (const r of (results || [])) map[r.history_id] = r.count;
    return jsonResponse(map);
  } catch(e) { return jsonResponse({}); }
}

async function handleAddNote(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser || currentUser.role !== 'admin') return errorResponse('无权限', 403);
  await ensureNotesTable(env);

  const { history_id, content, note_type = 'info', target_user = '' } = await request.json();
  if (!history_id || !content) return errorResponse('参数不完整');

  try {
    await env.DB.prepare(
      'INSERT INTO admin_notes (history_id, author, content, note_type, target_user, created_at) VALUES (?,?,?,?,?,datetime("now"))'
    ).bind(history_id, currentUser.username, content, note_type, target_user).run();
    return jsonResponse({ success: true });
  } catch(e) { return errorResponse('添加备注失败: ' + e.message); }
}

// ========== Chat Messages ==========

async function ensureMessagesTable(env) {
  try {
    await env.DB.prepare(
      'CREATE TABLE IF NOT EXISTS chat_messages (id INTEGER PRIMARY KEY AUTOINCREMENT, sender TEXT NOT NULL, recipient TEXT DEFAULT "", content TEXT NOT NULL, image_data TEXT DEFAULT "", created_at TEXT)'
    ).run();
    // Add image_data column if it doesn't exist (migration)
    try { await env.DB.prepare('ALTER TABLE chat_messages ADD COLUMN image_data TEXT DEFAULT ""').run(); } catch(e) { /* column exists */ }
  } catch(e) { /* table exists */ }
}

async function handleGetMessages(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);
  await ensureMessagesTable(env);

  const url = new URL(request.url);
  const withUser = url.searchParams.get('with');

  try {
    let results;
    if (withUser) {
      // Get messages between current user and the specified user
      const { results: r } = await env.DB.prepare(
        'SELECT * FROM chat_messages WHERE (sender = ? AND recipient = ?) OR (sender = ? AND recipient = ?) ORDER BY id DESC LIMIT 100'
      ).bind(currentUser.username, withUser, withUser, currentUser.username).all();
      results = r;
    } else {
      // Get all "public" messages (no specific recipient) + messages involving current user
      const { results: r } = await env.DB.prepare(
        'SELECT * FROM chat_messages WHERE recipient = "" OR sender = ? OR recipient = ? ORDER BY id DESC LIMIT 100'
      ).bind(currentUser.username, currentUser.username).all();
      results = r;
    }
    return jsonResponse((results || []).reverse());
  } catch(e) { return jsonResponse([]); }
}

async function handleSendMessage(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);
  await ensureMessagesTable(env);

  const { content, recipient = '', image_data = '' } = await request.json();
  if (!content && !image_data) return errorResponse('消息不能为空');

  try {
    await env.DB.prepare(
      'INSERT INTO chat_messages (sender, recipient, content, image_data, created_at) VALUES (?,?,?,?,datetime("now"))'
    ).bind(currentUser.username, recipient, content || '[图片]', image_data).run();
    return jsonResponse({ success: true });
  } catch(e) { return errorResponse('发送失败: ' + e.message); }
}

// Lightweight user list for chat (accessible to all authenticated users)
async function handleChatUserList(request, env) {
  const currentUser = await getUser(request, env);
  if (!currentUser) return errorResponse('未登录', 401);
  try {
    const { results } = await env.DB.prepare(
      'SELECT username, role, is_active FROM users WHERE is_active = 1 ORDER BY id'
    ).all();
    return jsonResponse(results || []);
  } catch(e) { return jsonResponse([]); }
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
      if (path === '/api/me' && request.method === 'GET') return await handleMe(request, env);
      if (path === '/api/register' && request.method === 'POST') return await handleRegister(request, env);
      if (path === '/api/reset-password' && request.method === 'POST') return await handleResetPassword(request, env);
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
      if (path.startsWith('/api/seckill/')) {
        const id = parseInt(path.split('/').pop());
        if (request.method === 'DELETE') return await handleDeleteSeckill(request, env, id);
      }
      if (path === '/api/versions') return await handleVersions(request, env);
      if (path === '/api/check-version' && request.method === 'POST') return await handleCheckVersion(request, env);
      if (path === '/api/logs' && request.method === 'GET') return await handleLogs(request, env);
      if (path === '/api/llm-config') return await handleLLMConfig(request, env);

      // Column mappings
      if (path === '/api/column-mappings') {
        if (request.method === 'GET') return await handleGetColumnMappings(request, env);
        if (request.method === 'POST') return await handleSaveColumnMappings(request, env);
      }

      // Upload file data (for admin online viewing)
      if (path === '/api/upload-data' && request.method === 'POST') return await handleSaveUploadData(request, env);
      if (path.startsWith('/api/upload-data/')) {
        const hid = parseInt(path.split('/').pop());
        if (request.method === 'GET' && !isNaN(hid)) return await handleGetUploadData(request, env, hid);
      }

      // Restock config (weighted algorithm)
      if (path === '/api/restock-config') {
        if (request.method === 'GET') return await handleGetRestockConfig(request, env);
        if (request.method === 'POST') return await handleSaveRestockConfig(request, env);
      }

      // Stocking rules
      if (path === '/api/stocking-rules') return await handleStockingRules(request, env);

      // User permissions
      if (path.startsWith('/api/user-permissions/')) {
        const uname = decodeURIComponent(path.slice('/api/user-permissions/'.length));
        return await handleUserPermissions(request, env, uname);
      }

      // Admin notes
      if (path === '/api/notes' && request.method === 'POST') return await handleAddNote(request, env);
      if (path === '/api/notes/counts' && request.method === 'GET') return await handleGetNoteCounts(request, env);
      if (path.startsWith('/api/notes/') && request.method === 'GET') {
        const hid = parseInt(path.split('/').pop());
        if (!isNaN(hid)) return await handleGetNotes(request, env, hid);
      }

      // Chat messages
      if (path === '/api/messages') {
        if (request.method === 'GET') return await handleGetMessages(request, env);
        if (request.method === 'POST') return await handleSendMessage(request, env);
      }
      if (path === '/api/chat-users' && request.method === 'GET') return await handleChatUserList(request, env);

      return errorResponse('Not found', 404);
    } catch (e) {
      return errorResponse(e.message || 'Internal error', 500);
    }
  },
};
