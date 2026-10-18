const SPREADSHEET_ID = '1Qbpyi39qV4bqyCu9bh4gr0Igy8IbXbKmW-QIJruGv6s';

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    try {
      var action = e.parameter.action;
      var result;
      if (action === 'getDashboardData') {
        result = getDashboardData(e.parameter);
      } else if (action === 'getProducts') {
        result = getSheetRecords('Products');
      } else if (action === 'ping') {
        result = { status: 'ok', timestamp: new Date().toISOString() };
      } else {
        result = { status: 'ready', message: 'ThreadAffiliate Engine API Active' };
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, message: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ThreadAffiliate Engine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var contents = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var body = JSON.parse(contents);
    var action = body.action;
    var payload = body.payload || {};
    var result;

    if (action === 'getDashboardData') {
      result = getDashboardData(payload);
    } else if (action === 'getProducts') {
      result = getSheetRecords('Products');
    } else if (action === 'syncThreadsAccount') {
      result = syncThreadsAccount(payload);
    } else if (action === 'pingThreadsAccount') {
      result = pingThreadsAccount(payload);
    } else if (action === 'saveProduct') {
      result = saveProduct(payload);
    } else if (action === 'saveMediaBatch') {
      result = saveMediaBatch(payload);
    } else if (action === 'generateThreadOptions') {
      result = generateThreadOptions(payload);
    } else if (action === 'createPost') {
      result = createPost(payload);
    } else if (action === 'publishPostNow') {
      result = publishPostNow(payload);
    } else if (action === 'checkViewTriggersAndReply') {
      result = syncAllPublishedPosts();
    } else if (action === 'syncPostInsights') {
      result = syncPostInsights(payload);
    } else if (action === 'syncAllPublishedPosts') {
      result = syncAllPublishedPosts();
    } else if (action === 'quickReplyComment') {
      result = quickReplyComment(payload);
    } else if (action === 'syncInboxComments') {
      result = syncInboxComments(payload);
    } else if (action === 'saveSettings') {
      result = saveSettings(payload);
    } else if (action === 'checkAndPublishQueue') {
      result = checkAndPublishQueue();
    } else if (action === 'checkWaitingViewsQueue') {
      result = checkWaitingViewsQueue();
    } else if (action === 'checkAndExecuteQueueAndViews') {
      result = checkAndExecuteQueueAndViews();
    } else if (action === 'setupTimeDrivenTrigger') {
      result = setupTimeDrivenTrigger(payload);
    } else if (action === 'saveHookTemplate') {
      result = saveHookTemplate(payload);
    } else if (action === 'deleteHookTemplate') {
      result = deleteHookTemplate(payload);
    } else if (action === 'saveViralReference') {
      result = saveViralReference(payload);
    } else if (action === 'deleteViralReference') {
      result = deleteViralReference(payload);
    } else if (action === 'generateSumopodRemake') {
      result = generateSumopodRemake(payload);
    } else if (action === 'scheduleAutopilotBatch') {
      result = scheduleAutopilotBatch(payload);
    } else if (action === 'saveAutopilotRule') {
      result = saveAutopilotRule(payload);
    } else if (action === 'toggleAutopilotRule') {
      result = toggleAutopilotRule(payload);
    } else if (action === 'deleteAutopilotRule') {
      result = deleteAutopilotRule(payload);
    } else if (action === 'runAutonomousAutopilotWorker') {
      result = runAutonomousAutopilotWorker(payload ? payload.rule_id : null);
    } else if (action === 'deleteProduct') {
      result = deleteProduct(payload);
    } else if (action === 'deleteMedia') {
      result = deleteMedia(payload);
    } else if (action === 'deleteMediaSession') {
      result = deleteMediaSession(payload);
    } else if (action === 'updateQueuePost') {
      result = updateQueuePost(payload);
    } else if (action === 'deleteQueuePost') {
      result = deleteQueuePost(payload);
    } else if (action === 'deleteThreadsAccount') {
      result = deleteThreadsAccount(payload);
    } else if (action === 'updateAccountCategory') {
      result = updateAccountCategory(payload);
    } else if (action === 'getCloudinaryUsage') {
      result = getCloudinaryUsage(payload);
    } else if (action === 'getAccountScheduleRadar') {
      result = getAccountScheduleRadar(payload);
    } else if (action === 'reorderMediaSession') {
      result = reorderMediaSession(payload);
    } else if (action === 'setupDb') {
      result = setupDatabase();
    } else {
      throw new Error('Aksi tidak dikenali: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Inisialisasi & sinkronisasi skema header sheet
function setupDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var schema = {
    'Accounts': ['account_id', 'username', 'profile_pic_url', 'access_token', 'token_expires_at', 'status', 'updated_at', 'category'],
    'Products': ['product_id', 'product_name', 'category', 'price', 'product_image_url', 'shopee_raw_url', 'shopee_affiliate_url', 'selling_points', 'allowed_accounts', 'created_at'],
    'Media': ['media_id', 'product_id', 'session_id', 'media_url', 'media_type', 'file_hash', 'uploaded_at', 'cooldown_until'],
    'Posts': ['post_id', 'account_id', 'thread_content_json', 'media_ids_json', 'product_id', 'affiliate_sub_id', 'views_threshold', 'current_views', 'status', 'threads_post_id', 'auto_reply_content', 'auto_reply_sent', 'created_at', 'likes_count', 'replies_count', 'permalink', 'drip_status', 'source_type', 'rule_id'],
    'Comments': ['comment_id', 'post_id', 'account_id', 'parent_comment_id', 'sender_username', 'content', 'intent_flag', 'status', 'created_at'],
    'Thread_Interactions': ['interaction_id', 'post_id', 'threads_post_id', 'account_id', 'comment_id', 'parent_comment_id', 'sender_username', 'author_type', 'content', 'intent_flag', 'reply_status', 'created_at'],
    'Posting_Queue': ['queue_id', 'post_id', 'account_id', 'scheduled_time', 'thread_content_json', 'media_ids_json', 'status', 'error_log', 'executed_at', 'reply_to_id', 'drip_step', 'source_type', 'rule_id'],
    'Hook_Bank': ['hook_id', 'category', 'hook_text', 'formula', 'created_at'],
    'Viral_References': ['ref_id', 'product_id', 'title', 'hook_thread_1', 'spill_thread_2', 'original_url', 'media_urls_json', 'created_at', 'threads_json'],
    'Autopilot_Rules': ['rule_id', 'name', 'account_ids_json', 'product_id', 'preferred_session_id', 'frequency_per_day', 'views_threshold', 'jitter_minutes', 'status', 'last_run_at', 'created_at', 'max_posts_per_day', 'min_interval_hours', 'last_session_id', 'time_windows_json'],
    'Settings': ['key', 'value']
  };

  for (var tab in schema) {
    var sheet = ss.getSheetByName(tab);
    if (!sheet) {
      sheet = ss.insertSheet(tab);
    }
    var headers = schema[tab];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (tab === 'Products') {
      var aCol = headers.indexOf('allowed_accounts') + 1;
      if (aCol > 0) {
        var lastR = Math.max(sheet.getLastRow(), 2);
        sheet.getRange(1, aCol, lastR, 1).setNumberFormat('@');
      }
    }
  }

  // Isi default Hook Bank jika masih kosong
  var hookSheet = ss.getSheetByName('Hook_Bank');
  if (hookSheet && hookSheet.getLastRow() <= 1) {
    var defaultHooks = [
      ['HK_001', 'Curhat/Story', 'Nyesel banget baru tau barang ini sekarang, padahal udah capek-capek nyari...', '[Penyesalan] + [Solusi Masalah]', new Date().toISOString()],
      ['HK_002', 'Racun Promo', 'Jangan bilang-bilang olshop sebelah, tapi barang seworth-it ini lagi flash sale 70%...', '[Kerahasiaan/FOMO] + [Diskon Besar]', new Date().toISOString()],
      ['HK_003', 'Kontroversial', 'Stop buang duit beli merk mahal kalau produk lokal 50 ribuan ini kualitasnya 11-12.', '[Tantang Persepsi] + [Komparasi Nilai]', new Date().toISOString()],
      ['HK_004', 'Sebelum-Sesudah', 'Tadinya skeptis liat review sliweran di FYP, tapi pas paketnya dateng dan dicoba langsung bengong...', '[Skeptis] + [Pembuktian Nyata]', new Date().toISOString()],
      ['HK_005', 'Spill Toko', 'Buat yang kemarin DM nanyain belinya di mana, mending cek thread ini sebelum stoknya abis lagi.', '[Sosial Proof] + [Urgensi Stok]', new Date().toISOString()]
    ];
    hookSheet.getRange(2, 1, defaultHooks.length, defaultHooks[0].length).setValues(defaultHooks);
  }

  // Isi default Settings branding jika belum ada
  var settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet && settingsSheet.getLastRow() <= 1) {
    var defaultSettings = [
      ['app_name', 'ThreadAffiliate'],
      ['app_tagline', 'Threads Automation'],
      ['browser_title', 'ThreadAffiliate Engine'],
      ['favicon_url', '']
    ];
    settingsSheet.getRange(2, 1, defaultSettings.length, defaultSettings[0].length).setValues(defaultSettings);
  }

  return { status: 'Database siap digunakan' };
}

// Sanitizer token Meta: bersihkan kutip, baris baru, spasi, prefiks Bearer, dan unwrap JSON/URL
function cleanToken(raw) {
  if (!raw) return '';
  var s = String(raw).trim();

  // Ekstrak jika user menempelkan respons JSON lengkap {"access_token":"..."}
  if (s.indexOf('{') !== -1 && s.indexOf('access_token') !== -1) {
    try {
      var parsed = JSON.parse(s);
      if (parsed.access_token) {
        s = String(parsed.access_token);
      }
    } catch (e) {
      var match = s.match(/"access_token"\s*:\s*"([^"]+)"/);
      if (match && match[1]) {
        s = match[1];
      }
    }
  }

  // Ekstrak jika user menempelkan query string atau URL dengan ?access_token=
  if (s.indexOf('access_token=') !== -1) {
    var parts = s.split('access_token=');
    if (parts.length > 1) {
      s = parts[1].split('&')[0];
    }
  }

  // Decode URI jika token ter-encode sebelumnya
  try {
    if (s.indexOf('%') !== -1) {
      s = decodeURIComponent(s);
    }
  } catch (e) {}

  s = s.replace(/^["'“”‘’`]+|["'“”‘’`]+$/g, '');
  s = s.replace(/^Bearer\s+/i, '');
  s = s.replace(/[\r\n\t\s\u200B-\u200D\uFEFF]/g, '');
  return s.trim();
}

// Sanitizer account_id: cegah distorsi scientific notation angka 17 digit Threads
function cleanAccountId(id) {
  if (id === null || id === undefined) return '';
  var s = String(id).replace(/^'/, '').trim();
  if (s.indexOf('e+') !== -1 || s.indexOf('E+') !== -1) {
    try {
      s = Number(s).toFixed(0);
    } catch(e) {}
  }
  return s;
}

// Serializer anti-error: konversi Date rekursif menjadi ISO String
function serializeForClient(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? '' : val.toISOString();
  }
  if (Array.isArray(val)) {
    return val.map(function(item) { return serializeForClient(item); });
  }
  if (typeof val === 'object') {
    var out = {};
    for (var k in val) {
      if (val.hasOwnProperty(k)) {
        out[k] = serializeForClient(val[k]);
      }
    }
    return out;
  }
  return val;
}

function getSheetRecords(tabName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var item = {};
    for (var j = 0; j < headers.length; j++) {
      item[headers[j]] = serializeForClient(data[i][j]);
    }
    if (tabName === 'Products' && item.allowed_accounts) {
      var rawParts = String(item.allowed_accounts).split(',');
      var cleanParts = [];
      for (var p = 0; p < rawParts.length; p++) {
        var part = rawParts[p].replace(/^['@\s]+|[\s']+$/g, '').trim().toLowerCase();
        if (part && !/e\+/i.test(part) && !/^\d+$/.test(part) && /^[a-z0-9._]+$/.test(part)) {
          if (cleanParts.indexOf(part) === -1) cleanParts.push(part);
        }
      }
      item.allowed_accounts = cleanParts.join(', ');
    }
    rows.push(item);
  }
  return rows;
}

function getDashboardData(payload) {
  var accounts = getSheetRecords('Accounts');
  var products = getSheetRecords('Products');
  var media = getSheetRecords('Media');
  var posts = getSheetRecords('Posts');
  var rawComments = getSheetRecords('Comments');
  var settingsRows = getSheetRecords('Settings');

  var settings = {};
  for (var s = 0; s < settingsRows.length; s++) {
    settings[settingsRows[s].key] = settingsRows[s].value;
  }

  var queue = getSheetRecords('Posting_Queue');
  var hooks = getSheetRecords('Hook_Bank');
  var viralRefs = getSheetRecords('Viral_References');
  var autopilotRules = getSheetRecords('Autopilot_Rules');
  var interactions = getSheetRecords('Thread_Interactions');

  // Urutkan postingan & antrean dari yang paling baru di atas (Descending)
  posts.sort(function(a, b) {
    var tA = a.created_at ? new Date(a.created_at).getTime() : 0;
    var tB = b.created_at ? new Date(b.created_at).getTime() : 0;
    return tB - tA;
  });

  queue.sort(function(a, b) {
    var tA = new Date(a.scheduled_time || a.executed_at || 0).getTime();
    var tB = new Date(b.scheduled_time || b.executed_at || 0).getTime();
    return tB - tA;
  });

  // Peta akun milik sendiri untuk memfilter self-reply / Utas 2 dari komentar audiens
  var allOwnedUsers = {};
  var allOwnedIds = {};
  for (var au = 0; au < accounts.length; au++) {
    var uName = String(accounts[au].username || '').toLowerCase().trim();
    if (uName) allOwnedUsers[uName] = true;
    var aId = cleanAccountId(accounts[au].account_id);
    if (aId) allOwnedIds[aId] = true;
  }

  // Gabungkan interactions dan comments (Hanya untuk balasan audiens luar)
  var audienceInteractions = [];
  var interactionIdMap = {};

  for (var t = 0; t < interactions.length; t++) {
    var inter = interactions[t];
    var senderUser = String(inter.sender_username || '').toLowerCase().trim();
    var isOwn = (inter.author_type === 'OWNER') || (inter.reply_status === 'OWNER_POST') || (senderUser && allOwnedUsers[senderUser]) || (senderUser === 'adminagent');
    if (!isOwn) {
      var key = String(inter.comment_id || inter.interaction_id).trim();
      if (key) {
        interactionIdMap[key] = true;
        audienceInteractions.push(inter);
      }
    }
  }

  for (var r = 0; r < rawComments.length; r++) {
    var cKey = String(rawComments[r].comment_id).trim();
    var rSender = String(rawComments[r].sender_username || '').toLowerCase().trim();
    var isOwnComment = (rawComments[r].author_type === 'OWNER') || (rawComments[r].reply_status === 'OWNER_POST') || (rSender && allOwnedUsers[rSender]) || (rSender === 'adminagent');

    if (!isOwnComment && cKey && !interactionIdMap[cKey]) {
      var audObj = {
        interaction_id: 'INT_' + cKey,
        post_id: rawComments[r].post_id,
        threads_post_id: '',
        account_id: rawComments[r].account_id,
        comment_id: cKey,
        parent_comment_id: rawComments[r].parent_comment_id || '',
        sender_username: rawComments[r].sender_username || 'user',
        author_type: 'AUDIENCE',
        content: rawComments[r].content || '',
        intent_flag: rawComments[r].intent_flag,
        reply_status: rawComments[r].reply_status || 'UNREPLIED',
        created_at: rawComments[r].created_at
      };
      audienceInteractions.push(audObj);
      interactionIdMap[cKey] = true;
    }
  }

  // Bangun struktur pohon komentar berjenjang (nested tree) rekursif
  var commentMap = {};
  for (var c = 0; c < audienceInteractions.length; c++) {
    var item = audienceInteractions[c];
    item.replies = [];
    commentMap[String(item.comment_id || item.interaction_id)] = item;
  }

  var rootComments = [];
  for (var k = 0; k < audienceInteractions.length; k++) {
    var cmtItem = audienceInteractions[k];
    var parentId = cmtItem.parent_comment_id ? String(cmtItem.parent_comment_id).trim() : '';
    if (parentId && commentMap[parentId]) {
      commentMap[parentId].replies.push(cmtItem);
    } else {
      rootComments.push(cmtItem);
    }
  }

  // Hitung ulang metrik replies_count per post khusus untuk balasan audiens luar
  var audienceRepliesCountMap = {};
  for (var ai = 0; ai < interactions.length; ai++) {
    if (interactions[ai].author_type !== 'OWNER' && interactions[ai].reply_status !== 'OWNER_POST') {
      var pIdKey = String(interactions[ai].post_id || '').trim();
      var thIdKey = String(interactions[ai].threads_post_id || '').trim();
      if (pIdKey) audienceRepliesCountMap[pIdKey] = (audienceRepliesCountMap[pIdKey] || 0) + 1;
      if (thIdKey) audienceRepliesCountMap[thIdKey] = (audienceRepliesCountMap[thIdKey] || 0) + 1;
    }
  }
  for (var pi = 0; pi < posts.length; pi++) {
    var pKey = String(posts[pi].post_id || '').trim();
    var pThKey = String(posts[pi].threads_post_id || '').trim();
    if (audienceRepliesCountMap[pKey] !== undefined) {
      posts[pi].replies_count = audienceRepliesCountMap[pKey];
    } else if (audienceRepliesCountMap[pThKey] !== undefined) {
      posts[pi].replies_count = audienceRepliesCountMap[pThKey];
    }
  }

  return {
    accounts: accounts || [],
    products: products || [],
    media: media || [],
    posts: posts || [],
    comments: rootComments || [],
    interactions: interactions || [],
    hooks: hooks || [],
    viral_references: viralRefs || [],
    autopilot_rules: autopilotRules || [],
    queue: queue || [],
    settings: settings || {}
  };
}

function syncThreadsAccount(payload) {
  var token = cleanToken(payload.access_token);
  if (!token) throw new Error('Access Token Meta wajib diisi');

  // Request otentikasi tunggal sesuai spek Meta Threads Graph API
  var url = 'https://graph.threads.net/v1.0/me?fields=id,username,name,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });
  var status = resp.getResponseCode();
  var resJson = JSON.parse(resp.getContentText() || '{}');

  if (status !== 200 || !resJson.id) {
    var fallbackResp = UrlFetchApp.fetch('https://graph.threads.net/v1.0/me?fields=id,username,name,threads_profile_picture_url', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var fbStatus = fallbackResp.getResponseCode();
    var fbJson = JSON.parse(fallbackResp.getContentText() || '{}');
    if (fbStatus === 200 && fbJson.id) {
      resJson = fbJson;
      status = fbStatus;
    }
  }

  if (status !== 200 || !resJson.id) {
    throw new Error('Gagal verifikasi token Threads: ' + (resJson.error ? resJson.error.message : 'Invalid Token'));
  }

  var rawId = String(resJson.id).trim();
  var newUsername = String(resJson.username || '').toLowerCase().trim();
  var profilePic = resJson.threads_profile_picture_url ? String(resJson.threads_profile_picture_url).trim() : '';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var userCol = headers.indexOf('username');
  if (userCol === -1) userCol = 1;
  var catCol = headers.indexOf('category');
  if (catCol === -1) {
    catCol = headers.length;
    sheet.getRange(1, catCol + 1).setValue('category');
  }

  var foundRow = -1;
  var existingCategory = 'General';

  for (var i = 1; i < data.length; i++) {
    var rowId = cleanAccountId(data[i][0]);
    var rowUser = String(data[i][userCol] || '').toLowerCase().trim();
    if (rowId === cleanAccountId(rawId) || (newUsername && rowUser === newUsername)) {
      foundRow = i + 1;
      if (catCol < data[i].length && data[i][catCol]) {
        existingCategory = String(data[i][catCol]).trim();
      }
      break;
    }
  }

  var nowIso = new Date().toISOString();
  var expiryIso = new Date(Date.now() + 60 * 24 * 3600 * 1000).toISOString(); // 60 hari
  var rowValues = [
    "'" + rawId,
    resJson.username || '',
    profilePic,
    token,
    expiryIso,
    'ACTIVE',
    nowIso,
    existingCategory
  ];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return {
    account_id: String(resJson.id),
    username: resJson.username || '',
    profile_pic_url: profilePic,
    category: existingCategory
  };
}

function pingThreadsAccount(payload) {
  var rawAccId = cleanAccountId(payload.account_id);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  if (!sheet) throw new Error('Sheet Accounts tidak ditemukan');

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error('Belum ada akun di database');

  var headers = data[0];
  var tokenCol = headers.indexOf('access_token');
  if (tokenCol === -1) tokenCol = 3;
  var idCol = headers.indexOf('account_id');
  if (idCol === -1) idCol = 0;
  var statusCol = headers.indexOf('status');
  if (statusCol === -1) statusCol = 5;
  var updatedCol = headers.indexOf('updated_at');
  if (updatedCol === -1) updatedCol = 6;

  var token = '';
  var targetRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (cleanAccountId(data[i][idCol]) === rawAccId) {
      token = cleanToken(data[i][tokenCol]);
      targetRow = i + 1;
      break;
    }
  }

  if (!token && data.length >= 2) {
    token = cleanToken(data[1][tokenCol]);
    targetRow = 2;
  }

  if (!token) throw new Error('Access Token tidak ditemukan untuk akun ini');

  // Panggil endpoint otentikasi Threads
  var url = 'https://graph.threads.net/v1.0/me?fields=id,username,name,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });
  var status = resp.getResponseCode();
  var resJson = JSON.parse(resp.getContentText() || '{}');

  if (status !== 200 || !resJson.id) {
    var fallbackUrl = 'https://graph.threads.net/me?fields=id,username,name,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
    var fallbackResp = UrlFetchApp.fetch(fallbackUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
    if (fallbackResp.getResponseCode() === 200) {
      var fbJson = JSON.parse(fallbackResp.getContentText() || '{}');
      if (fbJson.id) {
        resJson = fbJson;
        status = 200;
      }
    }
  }

  if (status !== 200 || !resJson.id) {
    if (targetRow > 0) {
      sheet.getRange(targetRow, statusCol + 1).setValue('EXPIRED');
    }
    throw new Error('Meta API Ping Gagal: ' + (resJson.error ? resJson.error.message : 'Invalid Token'));
  }

  var picCol = headers.indexOf('profile_pic_url');
  var picUrl = resJson.threads_profile_picture_url ? String(resJson.threads_profile_picture_url).trim() : '';

  if (targetRow > 0) {
    sheet.getRange(targetRow, tokenCol + 1).setValue(token);
    sheet.getRange(targetRow, statusCol + 1).setValue('ACTIVE');
    sheet.getRange(targetRow, updatedCol + 1).setValue(new Date().toISOString());
    if (picCol !== -1 && picUrl) {
      sheet.getRange(targetRow, picCol + 1).setValue(picUrl);
    }
  }

  return {
    account_id: String(resJson.id),
    username: resJson.username || '',
    profile_pic_url: picUrl,
    status: 'ACTIVE',
    message: 'Koneksi aktif! Token Meta @' + (resJson.username || '') + ' valid.'
  };
}

function saveProduct(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Products');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('Products');
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var defaultHeaders = ['product_id', 'product_name', 'category', 'price', 'product_image_url', 'shopee_raw_url', 'shopee_affiliate_url', 'selling_points', 'allowed_accounts', 'created_at'];

  // Pastikan skema header baru diperbarui jika belum lengkap
  if (headers.indexOf('allowed_accounts') === -1 || headers.indexOf('selling_points') === -1 || headers.indexOf('product_image_url') === -1 || headers.indexOf('shopee_affiliate_url') === -1) {
    headers = defaultHeaders;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var rawUrl = String(payload.shopee_raw_url || '').trim();
  if (rawUrl && rawUrl.indexOf('affiliate.shopee.co.id/offer/product_offer') === -1) {
    throw new Error('URL Halaman Produk wajib berformat resmi Shopee Affiliate Offer (https://affiliate.shopee.co.id/offer/product_offer/...)');
  }

  var idCol = headers.indexOf('product_id');
  if (idCol === -1) idCol = 0;
  var rawCol = headers.indexOf('shopee_raw_url');
  if (rawCol === -1) rawCol = 5;
  var affCol = headers.indexOf('shopee_affiliate_url');
  if (affCol === -1) affCol = 6;
  var nameCol = headers.indexOf('product_name');
  if (nameCol === -1) nameCol = 1;

  var existingRow = -1;
  var productId = payload.product_id ? String(payload.product_id).trim() : '';

  if (productId) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === productId) {
        existingRow = i + 1;
        break;
      }
    }
  }

  // Validasi anti-duplikasi URL Shopee (Raw & Affiliate)
  var normRaw = rawUrl.toLowerCase().replace(/^https?:\/\//i, '').replace(/\/+$/, '');
  var normAff = String(payload.shopee_affiliate_url || '').trim().toLowerCase().replace(/^https?:\/\//i, '').replace(/\/+$/, '');

  for (var r = 1; r < data.length; r++) {
    if (existingRow > 0 && (r + 1) === existingRow) continue;
    var rowRaw = String(data[r][rawCol] || '').trim().toLowerCase().replace(/^https?:\/\//i, '').replace(/\/+$/, '');
    var rowAff = String(data[r][affCol] || '').trim().toLowerCase().replace(/^https?:\/\//i, '').replace(/\/+$/, '');
    var existingName = String(data[r][nameCol] || data[r][idCol] || 'Lain');

    if (normRaw && rowRaw && normRaw === rowRaw) {
      throw new Error('URL Halaman Produk Asli sudah pernah terdaftar pada produk "' + existingName + '".');
    }
    if (normAff && rowAff && normAff === rowAff) {
      throw new Error('URL Shopee Affiliate (Shortlink) sudah pernah terdaftar pada produk "' + existingName + '".');
    }
    if (normAff && rowRaw && normAff === rowRaw) {
      throw new Error('Tautan Shopee Affiliate bentrok dengan URL Produk Asli milik "' + existingName + '".');
    }
    if (normRaw && rowAff && normRaw === rowAff) {
      throw new Error('URL Halaman Produk Asli bentrok dengan Tautan Affiliate milik "' + existingName + '".');
    }
  }

  if (!productId) {
    productId = 'PROD_' + Utilities.getUuid().substring(0, 8);
  }

  // Sanitasi ketat allowed_accounts: resolusi ID ke username sah, buang sampah angka & format ilmiah
  var accRecords = getSheetRecords('Accounts');
  var idToUserMap = {};
  for (var a = 0; a < accRecords.length; a++) {
    var aId = cleanAccountId(accRecords[a].account_id);
    var aUser = String(accRecords[a].username || '').toLowerCase().trim();
    if (aId && aUser) idToUserMap[aId] = aUser;
    if (aUser) idToUserMap[aUser] = aUser;
  }

  var rawAllowed = payload.allowed_accounts;
  var rawList = [];
  if (Array.isArray(rawAllowed)) {
    rawList = rawAllowed;
  } else if (rawAllowed) {
    rawList = String(rawAllowed).split(',');
  }

  var cleanUsernames = [];
  var seenUsers = {};
  for (var k = 0; k < rawList.length; k++) {
    var rawItem = String(rawList[k] || '')
      .replace(/^['"@\s]+|['"@\s]+$/g, '')
      .replace(/[\r\n\t\s]/g, '')
      .trim()
      .toLowerCase();
    if (!rawItem) continue;
    if (/e[+-]?\d+/i.test(rawItem) || /^\d+$/.test(rawItem)) continue;

    var resolvedUser = idToUserMap[cleanAccountId(rawItem)] || idToUserMap[rawItem] || rawItem;
    if (/e[+-]?\d+/i.test(resolvedUser) || /^\d+$/.test(resolvedUser)) continue;

    if (/^[a-z0-9._]+$/.test(resolvedUser) && !seenUsers[resolvedUser]) {
      seenUsers[resolvedUser] = true;
      cleanUsernames.push(resolvedUser);
    }
  }

  var allowedStr = cleanUsernames.join(', ');
  var allowedCol = headers.indexOf('allowed_accounts');
  if (allowedCol === -1) {
    allowedCol = headers.length;
    headers.push('allowed_accounts');
    sheet.getRange(1, allowedCol + 1).setValue('allowed_accounts');
  }

  var rowObj = {
    product_id: productId,
    product_name: payload.product_name || 'Produk Baru',
    category: payload.category || 'General',
    price: Number(payload.price || 0),
    product_image_url: payload.product_image_url || '',
    shopee_raw_url: payload.shopee_raw_url || '',
    shopee_affiliate_url: payload.shopee_affiliate_url || '',
    selling_points: payload.selling_points || '',
    allowed_accounts: allowedStr,
    created_at: payload.created_at || new Date().toISOString()
  };

  var rowValues = [];
  for (var h = 0; h < headers.length; h++) {
    var k = headers[h];
    rowValues.push(rowObj[k] !== undefined ? rowObj[k] : '');
  }

  if (existingRow > 0) {
    // Format Plain Text ('@') sebelum dan sesudah penulisan untuk mencegah konversi eksponensial
    sheet.getRange(existingRow, allowedCol + 1).setNumberFormat('@');
    sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(existingRow, allowedCol + 1).setNumberFormat('@').setValue(allowedStr);
  } else {
    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, allowedCol + 1).setNumberFormat('@');
    sheet.appendRow(rowValues);
    sheet.getRange(nextRow, allowedCol + 1).setNumberFormat('@').setValue(allowedStr);
  }

  return { product_id: productId };
}

function saveMediaBatch(payload) {
  var items = payload.items || [];
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('Daftar media batch kosong');
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Media');
  var now = new Date();
  var cooldown = new Date(now.getTime() + 14 * 24 * 3600 * 1000); // Cooldown 14 hari
  var rowsToAppend = [];
  var savedIds = [];

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var mediaId = 'MED_' + Utilities.getUuid().substring(0, 8);
    savedIds.push(mediaId);
    rowsToAppend.push([
      mediaId,
      it.product_id || '',
      it.session_id || ('REV_' + Date.now()),
      it.media_url,
      it.media_type || 'IMAGE',
      it.file_hash || '',
      now.toISOString(),
      cooldown.toISOString()
    ]);
  }

  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  return { saved_count: rowsToAppend.length, media_ids: savedIds };
}

// Helper pembentukan URL endpoint Sumopod
function buildSumopodEndpoint(customEndpoint) {
  var url = (customEndpoint || '').trim();
  if (!url || url.indexOf('api.sumopod.com') !== -1) {
    url = 'https://ai.sumopod.com/v1/chat/completions';
  }
  url = url.replace(/\/+$/, '');
  if (url.indexOf('/chat/completions') === -1) {
    if (url.indexOf('/v1') === -1) {
      url += '/v1';
    }
    url += '/chat/completions';
  }
  return url;
}

// Helper pemetaan detail error respons SumoPod AI transparan
function extractSumopodError(code, resJson, endpoint) {
  var detail = '';
  if (resJson && resJson.error) {
    detail = typeof resJson.error === 'object'
      ? (resJson.error.message || JSON.stringify(resJson.error))
      : String(resJson.error);
  } else if (resJson && resJson.message) {
    detail = resJson.message;
  } else {
    detail = 'Respons kosong dari provider';
  }

  if (code === 401) {
    return 'Autentikasi SumoPod Gagal (HTTP 401): API Key tidak valid atau belum diisi. Periksa menu Settings. Detail: ' + detail;
  }
  if (code === 404) {
    return 'Endpoint SumoPod Tidak Ditemukan (HTTP 404): URL "' + endpoint + '" tidak dikenali. Detail: ' + detail;
  }
  if (code === 429) {
    return 'Batas Kuota / Rate Limit SumoPod Terlampaui (HTTP 429): ' + detail;
  }
  return 'SumoPod API Error (HTTP ' + code + '): ' + detail;
}

// Generator 3 Opsi Caption Dinamis per Utas Menggunakan Sumopod AI (Fallback Gemini)
function generateThreadOptions(payload) {
  var settings = getDashboardData().settings || {};
  var apiKey = settings.sumopod_api_key;
  var isHook = (Number(payload.thread_index || 0) === 0);
  var prodName = payload.product_names || payload.product_name || 'Produk Shopee';
  var sellingPoints = payload.selling_points || '';
  var currentText = payload.current_text || '';

  // Fallback ke Gemini jika Sumopod API Key belum diisi
  if (!apiKey && settings.gemini_key) {
    return generateGeminiThreadOptions(payload, settings.gemini_key);
  }
  if (!apiKey) {
    throw new Error('Sumopod API Key belum dikonfigurasi di Pengaturan Sistem (Settings).');
  }

  var rawEndpoint = (settings.sumopod_endpoint || '').trim();
  var endpoint = buildSumopodEndpoint(rawEndpoint);
  var model = (settings.sumopod_model && settings.sumopod_model.trim()) ? settings.sumopod_model.trim() : 'claude-haiku-4-5';

  var systemPrompt = 'Anda adalah copywriter Meta Threads profesional untuk Shopee Affiliate Indonesia dengan prinsip STRICT STYLE MIRRORING. ' +
    'DILARANG keras membuat bullet list spesifikasi centang atau bahasa iklan kaku. ' +
    'Gunakan gaya percakapan santai, curhat emosional, pemenggalan jeda baris natural, penekanan kata dengan HURUF BESAR seperlunya, dan emoji yang selaras. ' +
    (isHook
      ? 'Ini adalah Utas #1 (Root Hook). Caption WAJIB BEBAS dari segala link eksternal agar distribusi jangkauan organik Threads maksimal. '
      : 'Ini adalah Utas Lanjutan/Spill. WAJIB merajut tag [LINK] atau [LINK_AFFILIATE] menyatu secara kontekstual di dalam kalimat (inline embedding), BUKAN ditaruh di baris terbawah seperti footer iklan. Contoh: "Bisa custom nama disini [LINK] biar ga ketuker" atau "aku ambil yg ukuran sedang di [LINK] pas bgt di wastafel". ') +
    'Keluarkan output HANYA JSON valid berformat: {"options": ["Opsi 1...", "Opsi 2...", "Opsi 3..."]}';

  var userPrompt = 'Target Produk: ' + prodName + '\n' +
    (sellingPoints ? ('Keunggulan Produk: ' + sellingPoints + '\n') : '') +
    (currentText ? ('Teks Draf / Referensi Pengguna:\n"' + currentText + '"\nTiru ritme, pemenggalan baris, dan gaya penekanan kata dari teks di atas!\n') : '') +
    'Buatkan 3 opsi variasi caption yang natural, santai, dan memikat audiens.';

  var body = {
    model: model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.75
  };

  var response;
  try {
    response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + cleanToken(apiKey),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
  } catch (netErr) {
    Logger.log('[Sumopod Network Error generateThreadOptions] ' + netErr.message);
    throw new Error('Gagal terhubung ke SumoPod: ' + netErr.message);
  }

  var code = response.getResponseCode();
  var rawText = (response.getContentText() || '').trim();
  Logger.log('[Sumopod generateThreadOptions response HTTP ' + code + ']: ' + rawText);

  if (rawText.indexOf('<') === 0) {
    throw new Error('SumoPod API Error (HTTP ' + code + '): Server AI mengembalikan format HTML. Periksa URL endpoint atau API Key.');
  }

  var resJson = null;
  try {
    resJson = JSON.parse(rawText);
  } catch(parseErr) {
    Logger.log('[Sumopod JSON Parse Error]: ' + rawText);
    throw new Error('SumoPod API Error (HTTP ' + code + '): Format respons bukan JSON: ' + rawText.substring(0, 150));
  }

  if (code !== 200 || !resJson || !resJson.choices || resJson.choices.length === 0) {
    var errMsg = extractSumopodError(code, resJson, endpoint);
    Logger.log('[Sumopod API Failure generateThreadOptions]: ' + errMsg);
    throw new Error(errMsg);
  }

  var replyContent = resJson.choices[0].message.content || '';
  var cleanJsonStr = replyContent.replace(/```json/gi, '').replace(/```/g, '').trim();

  try {
    var parsed = JSON.parse(cleanJsonStr);
    if (Array.isArray(parsed.options) && parsed.options.length > 0) {
      return { options: parsed.options.slice(0, 3) };
    }
  } catch(parseErr) {}

  var lines = replyContent.split('\n')
    .map(function(s) { return s.replace(/^[0-9\.\-\* ]+/, '').trim(); })
    .filter(function(s) { return s.length > 10; });

  return { options: lines.slice(0, 3) };
}

function generateGeminiThreadOptions(payload, apiKey) {
  var isHook = (Number(payload.thread_index || 0) === 0);
  var prodName = payload.product_names || payload.product_name || 'Produk Shopee';
  var sellingPoints = payload.selling_points || '';
  var promptText = 'Buatkan 3 opsi kalimat caption Meta Threads Indonesia untuk mempromosikan produk: "' + prodName + '"' +
    (sellingPoints ? (' dengan keunggulan: "' + sellingPoints + '"') : '') +
    (isHook ? '. Utas #1 (Root Hook pembuka viral, TANPA link eksternal).' : '. Utas Lanjutan dengan tag [LINK_AFFILIATE].') +
    ' Format balasan HANYA 3 baris teks terpisah per nomor tanpa intro atau outro.';

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] }),
    muteHttpExceptions: true
  });
  var resJson = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !resJson.candidates) {
    throw new Error('AI Error: Gagal menghasilkan opsi caption');
  }
  var rawOutput = resJson.candidates[0].content.parts[0].text || '';
  var lines = rawOutput.split('\n')
    .map(function(s) { return s.replace(/^[0-9\.\-\* ]+/, '').trim(); })
    .filter(function(s) { return s.length > 10; });
  return { options: lines.slice(0, 3) };
}

// Helper ekstraksi array media_id dari rangkaian nodes
function extractMediaIds(nodes) {
  var ids = [];
  if (!Array.isArray(nodes)) return ids;
  for (var i = 0; i < nodes.length; i++) {
    var n = nodes[i];
    if (!n) continue;
    if (Array.isArray(n.media_ids)) {
      for (var j = 0; j < n.media_ids.length; j++) {
        if (n.media_ids[j] && ids.indexOf(n.media_ids[j]) === -1) ids.push(n.media_ids[j]);
      }
    } else if (n.media_id && ids.indexOf(n.media_id) === -1) {
      ids.push(n.media_id);
    }
  }
  return ids;
}

function createPost(payload) {
  // STRICT ACCOUNT BINDING: Pastikan produk yang dikaitkan diizinkan secara eksplisit untuk akun ini
  if (payload.product_id && payload.account_id) {
    var prodsListCheck = getSheetRecords('Products');
    var targetP = null;
    for (var pi = 0; pi < prodsListCheck.length; pi++) {
      if (prodsListCheck[pi].product_id === payload.product_id) {
        targetP = prodsListCheck[pi];
        break;
      }
    }
    if (targetP) {
      var rawAllowed = String(targetP.allowed_accounts || '').trim();
      if (!rawAllowed) {
        throw new Error('Produk "' + (targetP.product_name || targetP.product_id) + '" berstatus Belum Terhubung dan tidak dapat diposting.');
      }
      var cleanTargetAcc = cleanAccountId(payload.account_id);
      var allowedItems = rawAllowed.split(',').map(function(s) { return s.trim().toLowerCase(); });
      var accRecords = getSheetRecords('Accounts');
      var userMatch = '';
      for (var ai = 0; ai < accRecords.length; ai++) {
        if (cleanAccountId(accRecords[ai].account_id) === cleanTargetAcc) {
          userMatch = String(accRecords[ai].username || '').toLowerCase().trim();
          break;
        }
      }
      var isPermitted = false;
      for (var aIdx = 0; aIdx < allowedItems.length; aIdx++) {
        var it = allowedItems[aIdx];
        if (cleanAccountId(it) === cleanTargetAcc || (userMatch && it === userMatch)) {
          isPermitted = true;
          break;
        }
      }
      if (!isPermitted) {
        throw new Error('Akun ini tidak memiliki izin untuk memposting produk "' + (targetP.product_name || targetP.product_id) + '".');
      }
    }
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Posts');
  var postId = 'POST_' + Utilities.getUuid().substring(0, 8);
  var subId = (payload.account_id || 'acc') + '_' + postId;
  var status = payload.status || 'SCHEDULED';
  var publishMode = payload.publish_mode || 'DRIP_ON_VIEWS';
  var publishedThreadId = '';

  var nodes = [];
  try {
    nodes = typeof payload.thread_content_json === 'string' ? JSON.parse(payload.thread_content_json) : payload.thread_content_json;
  } catch(e) {
    nodes = [{ content: String(payload.thread_content_json || ''), media_id: '' }];
  }
  if (!Array.isArray(nodes)) nodes = [nodes];

  var isMultiNode = nodes.length > 1;
  var isDripMode = (isMultiNode && publishMode === 'DRIP_ON_VIEWS');
  var dripStatus = isDripMode ? 'WAITING_VIEWS' : 'COMPLETED';
  var isReplySent = !isDripMode;

  // Jika user memilih posting langsung sekarang
  if (status === 'PUBLISH_NOW') {
    var accounts = getSheetRecords('Accounts');
    var token = '';
    var targetAccId = cleanAccountId(payload.account_id);
    for (var a = 0; a < accounts.length; a++) {
      if (cleanAccountId(accounts[a].account_id) === targetAccId) {
        token = cleanToken(accounts[a].access_token);
        break;
      }
    }
    if (!token && accounts.length === 1) {
      token = cleanToken(accounts[0].access_token);
      targetAccId = cleanAccountId(accounts[0].account_id);
    }
    if (!token) throw new Error('Access Token Meta tidak ditemukan untuk akun ini. Silakan sinkronkan ulang token di menu Akun Meta API.');

    var mediaRecords = getSheetRecords('Media');
    var mediaMap = {};
    for (var m = 0; m < mediaRecords.length; m++) {
      mediaMap[String(mediaRecords[m].media_id).trim()] = mediaRecords[m];
    }

    var prods = getSheetRecords('Products');

    if (isDripMode) {
      // Auto-Drip Mode: Terbitkan HANYA Utas 1 (Root Hook) sekarang tanpa link eksternal
      publishedThreadId = executeThreadChain(payload.account_id, token, [nodes[0]], mediaMap, prods, subId);
      for (var d = 1; d < nodes.length; d++) {
        if (typeof nodes[d] === 'object' && nodes[d] !== null) {
          nodes[d].status = 'WAITING_VIEWS';
        }
      }
      dripStatus = 'WAITING_VIEWS';
      isReplySent = false;
    } else {
      // Publish All Mode atau Utas Tunggal: Terbitkan seluruh rantai sekaligus
      publishedThreadId = executeThreadChain(payload.account_id, token, nodes, mediaMap, prods, subId);
      dripStatus = 'COMPLETED';
      isReplySent = true;
    }
    status = 'PUBLISHED';
    payload.thread_content_json = JSON.stringify(nodes);
  } else if (status === 'SCHEDULED' || status === 'Pending_Scheduled' || status === 'DRAFT') {
    if (isDripMode) {
      if (typeof nodes[0] === 'object' && nodes[0] !== null) nodes[0].status = status;
      for (var s = 1; s < nodes.length; s++) {
        if (typeof nodes[s] === 'object' && nodes[s] !== null) nodes[s].status = 'WAITING_VIEWS';
      }
    } else {
      for (var s2 = 0; s2 < nodes.length; s2++) {
        if (typeof nodes[s2] === 'object' && nodes[s2] !== null) nodes[s2].status = status;
      }
    }
    payload.thread_content_json = JSON.stringify(nodes);
  }

  var rootPermalink = (nodes[0] && nodes[0].permalink) ? nodes[0].permalink : '';

  var sourceType = payload.source_type || 'manual';
  var ruleId = payload.rule_id || '';

  var row = [
    postId,
    payload.account_id,
    payload.thread_content_json,
    payload.media_ids_json || '[]',
    payload.product_id || '',
    subId,
    Number(payload.views_threshold || 500),
    0,
    status,
    publishedThreadId,
    payload.auto_reply_content || '',
    isReplySent,
    new Date().toISOString(),
    0,
    0,
    rootPermalink,
    dripStatus,
    sourceType,
    ruleId
  ];
  sheet.appendRow(row);

  var qSheet = ss.getSheetByName('Posting_Queue');
  if (qSheet) {
    if (status === 'PUBLISHED' && isDripMode && publishedThreadId) {
      // Simpan Utas #2 ke dalam antrean dengan status khusus Waiting_Views_Threshold dan root_post_id sebagai reply_to_id
      var remainingNodes = nodes.slice(1);
      var qId = 'Q_' + Utilities.getUuid().substring(0, 8);
      qSheet.appendRow([
        qId,
        postId,
        payload.account_id,
        new Date().toISOString(),
        JSON.stringify(remainingNodes),
        JSON.stringify(extractMediaIds(remainingNodes)),
        'Waiting_Views_Threshold',
        '',
        '',
        publishedThreadId,
        2,
        sourceType,
        ruleId
      ]);
    } else if (status === 'SCHEDULED' || status === 'Pending_Scheduled') {
      var qIdSched = 'Q_' + Utilities.getUuid().substring(0, 8);
      var schedIso = payload.scheduled_time ? new Date(payload.scheduled_time).toISOString() : new Date().toISOString();
      qSheet.appendRow([
        qIdSched,
        postId,
        payload.account_id,
        schedIso,
        payload.thread_content_json,
        payload.media_ids_json || '[]',
        status === 'Pending_Scheduled' ? 'Pending_Scheduled' : 'Scheduled',
        '',
        '',
        '',
        isDripMode ? 1 : 0,
        sourceType,
        ruleId
      ]);
    }
  }

  return { post_id: postId, sub_id: subId, status: status, threads_post_id: publishedThreadId };
}

function publishPostNow(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var postId = payload.post_id;
  if (!postId) throw new Error('Post ID wajib disertakan');

  var postSheet = ss.getSheetByName('Posts');
  var postData = postSheet.getDataRange().getValues();
  var targetRow = -1;
  var targetPost = null;

  for (var i = 1; i < postData.length; i++) {
    if (String(postData[i][0]) === String(postId)) {
      targetRow = i + 1;
      targetPost = {
        post_id: postData[i][0],
        account_id: postData[i][1],
        thread_content_json: postData[i][2],
        media_ids_json: postData[i][3],
        drip_status: postData[i][16]
      };
      break;
    }
  }

  if (!targetPost) throw new Error('Postingan dengan ID tersebut tidak ditemukan');

  var accounts = getSheetRecords('Accounts');
  var token = '';
  var targetAccId = cleanAccountId(targetPost.account_id);
  for (var a = 0; a < accounts.length; a++) {
    if (cleanAccountId(accounts[a].account_id) === targetAccId) {
      token = cleanToken(accounts[a].access_token);
      break;
    }
  }
  if (!token && accounts.length === 1) {
    token = cleanToken(accounts[0].access_token);
    targetAccId = cleanAccountId(accounts[0].account_id);
  }
  if (!token) throw new Error('Access Token Meta tidak ditemukan untuk akun ini.');

  var mediaRecords = getSheetRecords('Media');
  var mediaMap = {};
  for (var m = 0; m < mediaRecords.length; m++) {
    mediaMap[String(mediaRecords[m].media_id).trim()] = mediaRecords[m];
  }

  var nodes = [];
  try {
    nodes = typeof targetPost.thread_content_json === 'string' ? JSON.parse(targetPost.thread_content_json) : targetPost.thread_content_json;
  } catch(e) {
    nodes = [{ content: String(targetPost.thread_content_json || ''), media_id: '' }];
  }
  if (!Array.isArray(nodes)) nodes = [nodes];

  var prods = getSheetRecords('Products');
  var subId = targetPost.account_id + '_' + targetPost.post_id;
  var isDripMode = (targetPost.drip_status === 'WAITING_VIEWS' && nodes.length > 1);

  // Jika mode Auto-Drip on Views, terbitkan HANYA Utas #1 sekarang
  var nodesToPublish = isDripMode ? [nodes[0]] : nodes;
  var publishedThreadId = executeThreadChain(targetPost.account_id, token, nodesToPublish, mediaMap, prods, subId);

  if (isDripMode) {
    for (var dn = 1; dn < nodes.length; dn++) {
      if (typeof nodes[dn] === 'object' && nodes[dn] !== null) {
        nodes[dn].status = 'WAITING_VIEWS';
      }
    }
  }

  var rootPermalink = (nodes[0] && nodes[0].permalink) ? nodes[0].permalink : '';
  postSheet.getRange(targetRow, 3).setValue(JSON.stringify(nodes));
  postSheet.getRange(targetRow, 9).setValue('PUBLISHED');
  postSheet.getRange(targetRow, 10).setValue(publishedThreadId);
  if (rootPermalink) postSheet.getRange(targetRow, 16).setValue(rootPermalink);

  var qSheet = ss.getSheetByName('Posting_Queue');
  if (qSheet) {
    var qData = qSheet.getDataRange().getValues();
    for (var q = 1; q < qData.length; q++) {
      var qStatusNorm = String(qData[q][6] || '').toLowerCase().trim();
      var isQWaiting = (qStatusNorm === 'waiting_views_threshold' || qStatusNorm === 'menunggu target' || qStatusNorm === 'waiting_views');
      if (String(qData[q][1]) === String(postId) && !isQWaiting) {
        qSheet.getRange(q + 1, 7).setValue('Posted');
        qSheet.getRange(q + 1, 8).setValue('');
        qSheet.getRange(q + 1, 9).setValue(new Date().toISOString());
      }
    }

    // Daftarkan Utas #2 ke dalam antrean berstatus Waiting_Views_Threshold
    if (isDripMode && publishedThreadId) {
      var remainingNodesNow = nodes.slice(1);
      var qIdDripNow = 'Q_' + Utilities.getUuid().substring(0, 8);
      qSheet.appendRow([
        qIdDripNow,
        postId,
        targetPost.account_id,
        new Date().toISOString(),
        JSON.stringify(remainingNodesNow),
        JSON.stringify(extractMediaIds(remainingNodesNow)),
        'Waiting_Views_Threshold',
        '',
        '',
        publishedThreadId,
        2
      ]);
    }
  }

  return { post_id: postId, threads_post_id: publishedThreadId, status: 'PUBLISHED' };
}

// Sinkronisasi metrik live (Views, Likes, Balasan) dan komentar dari Meta Threads API
function syncPostInsights(payload) {
  var postId = payload.post_id;
  if (!postId) throw new Error('Post ID wajib disertakan');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var postSheet = ss.getSheetByName('Posts');
  var postData = postSheet.getDataRange().getValues();
  var headers = postData[0];

  var idCol = headers.indexOf('post_id');
  var accCol = headers.indexOf('account_id');
  var viewsCol = headers.indexOf('current_views');
  var thresholdCol = headers.indexOf('views_threshold');
  var thIdCol = headers.indexOf('threads_post_id');
  var replySentCol = headers.indexOf('auto_reply_sent');
  var autoReplyContentCol = headers.indexOf('auto_reply_content');
  var likesCol = headers.indexOf('likes_count');
  var repliesCol = headers.indexOf('replies_count');
  var linkCol = headers.indexOf('permalink');

  // Pastikan kolom baru terpasang di sheet jika belum ada
  var lastCol = headers.length;
  if (likesCol === -1) {
    likesCol = lastCol++;
    postSheet.getRange(1, likesCol + 1).setValue('likes_count');
  }
  if (repliesCol === -1) {
    repliesCol = lastCol++;
    postSheet.getRange(1, repliesCol + 1).setValue('replies_count');
  }
  if (linkCol === -1) {
    linkCol = lastCol++;
    postSheet.getRange(1, linkCol + 1).setValue('permalink');
  }

  var targetRow = -1;
  var rowData = null;
  for (var i = 1; i < postData.length; i++) {
    if (String(postData[i][idCol]) === String(postId)) {
      targetRow = i + 1;
      rowData = postData[i];
      break;
    }
  }

  if (!rowData) throw new Error('Postingan tidak ditemukan di database');
  var threadsPostId = String(rowData[thIdCol] || '').trim();
  if (!threadsPostId) throw new Error('Postingan ini belum memiliki Threads Media ID (belum terbit di Threads)');

  var accounts = getSheetRecords('Accounts');
  var targetAccId = cleanAccountId(rowData[accCol]);
  var token = '';
  for (var a = 0; a < accounts.length; a++) {
    if (cleanAccountId(accounts[a].account_id) === targetAccId) {
      token = cleanToken(accounts[a].access_token);
      break;
    }
  }
  if (!token && accounts.length > 0) {
    token = cleanToken(accounts[0].access_token);
  }
  if (!token) throw new Error('Token Meta tidak ditemukan untuk akun ini');

  var views = Number(rowData[viewsCol] || 0);
  var likes = Number(rowData[likesCol] || 0);
  var replies = Number(rowData[repliesCol] || 0);
  var permalink = String(rowData[linkCol] || '');

  // 1. Ambil permalink postingan Threads
  var mediaUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '?fields=id,permalink&access_token=' + encodeURIComponent(token);
  try {
    var mResp = UrlFetchApp.fetch(mediaUrl, { method: 'get', muteHttpExceptions: true });
    if (mResp.getResponseCode() === 200) {
      var mJson = JSON.parse(mResp.getContentText() || '{}');
      if (mJson.permalink) permalink = mJson.permalink;
    }
  } catch (e) {}

  // 2. Ambil metrik Views, Likes, dan Balasan via Threads Insights API
  var insightsUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/insights?metric=views,likes,replies&access_token=' + encodeURIComponent(token);
  try {
    var insResp = UrlFetchApp.fetch(insightsUrl, { method: 'get', muteHttpExceptions: true });
    if (insResp.getResponseCode() === 200) {
      var insJson = JSON.parse(insResp.getContentText() || '{}');
      if (insJson.data && Array.isArray(insJson.data)) {
        for (var d = 0; d < insJson.data.length; d++) {
          var item = insJson.data[d];
          var val = (item.values && item.values.length > 0) ? Number(item.values[0].value || 0) : 0;
          if (item.name === 'views') views = val;
          if (item.name === 'likes') likes = val;
          if (item.name === 'replies') replies = val;
        }
      }
    }
  } catch (e) {}

  // 3. Sinkronkan balasan/komentar dari Threads ke tab Comments & Thread_Interactions
  var commentsSynced = 0;
  var repliesUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(token);
  try {
    var repResp = UrlFetchApp.fetch(repliesUrl, { method: 'get', muteHttpExceptions: true });
    if (repResp.getResponseCode() === 200) {
      var repJson = JSON.parse(repResp.getContentText() || '{}');
      if (repJson.data && Array.isArray(repJson.data)) {
        var cmtSheet = ss.getSheetByName('Comments');
        var interSheet = ss.getSheetByName('Thread_Interactions');
        var existingCmts = getSheetRecords('Comments');
        var cmtIdSet = {};
        for (var c = 0; c < existingCmts.length; c++) {
          cmtIdSet[String(existingCmts[c].comment_id)] = true;
        }
        var existingInter = getSheetRecords('Thread_Interactions');
        for (var ei = 0; ei < existingInter.length; ei++) {
          cmtIdSet[String(existingInter[ei].comment_id || existingInter[ei].interaction_id)] = true;
        }

        var allOwnedUsers = {};
        var allOwnedIds = {};
        for (var au = 0; au < accounts.length; au++) {
          var uName = String(accounts[au].username || '').toLowerCase().trim();
          if (uName) allOwnedUsers[uName] = true;
          var aId = cleanAccountId(accounts[au].account_id);
          if (aId) allOwnedIds[aId] = true;
        }

        for (var r = 0; r < repJson.data.length; r++) {
          var repItem = repJson.data[r];
          var repIdStr = String(repItem.id || '').trim();
          if (repIdStr && !cmtIdSet[repIdStr]) {
            var senderUser = String(repItem.username || '').toLowerCase().trim();
            var senderUserId = cleanAccountId(repItem.user_id || (repItem.from && repItem.from.id) || '');
            var isOwn = (senderUser && allOwnedUsers[senderUser]) || (senderUserId && allOwnedIds[senderUserId]) || (senderUser === 'adminagent');

            // Jika balasan akun sendiri (self-reply / Utas 2), lewati dan JANGAN masukkan ke database percakapan audiens
            if (isOwn) {
              cmtIdSet[repIdStr] = true;
              continue;
            }

            var intentFlag = /spill|link|beli|toko|harga|order|checkout|beli di mana|buy|mau/i.test(repItem.text || '');
            var createdAt = repItem.timestamp || new Date().toISOString();

            if (interSheet) {
              interSheet.appendRow([
                'INT_' + repIdStr.substring(0, 10),
                postId,
                threadsPostId,
                rowData[accCol],
                repIdStr,
                '',
                repItem.username || 'user',
                'AUDIENCE',
                repItem.text || '',
                intentFlag,
                'UNREPLIED',
                createdAt
              ]);
            }
            if (cmtSheet) {
              cmtSheet.appendRow([
                repIdStr,
                postId,
                rowData[accCol],
                '',
                repItem.username || 'user',
                repItem.text || '',
                intentFlag,
                'PUBLISHED',
                createdAt
              ]);
              commentsSynced++;
            }
            cmtIdSet[repIdStr] = true;
          }
        }
      }
    }
  } catch (e) {}

  // 4. Perbarui data ke baris Google Sheets
  postSheet.getRange(targetRow, viewsCol + 1).setValue(views);
  postSheet.getRange(targetRow, likesCol + 1).setValue(likes);
  postSheet.getRange(targetRow, repliesCol + 1).setValue(replies);
  if (permalink) postSheet.getRange(targetRow, linkCol + 1).setValue(permalink);

  // 5. STANDARISASI NUMERIK & STATUS: Jika current_views >= target_views dan Utas #2 masih menunggu, eksekusi instan
  var currentViewsNum = Number(views || 0);
  var targetViewsNum = Number(rowData[thresholdCol] || 0);
  if (targetViewsNum <= 0) targetViewsNum = 500;

  var dripColIdx = headers.indexOf('drip_status');
  var dripStatusRaw = String(dripColIdx !== -1 ? rowData[dripColIdx] : '').toLowerCase().trim();
  var isReplyAlreadySent = (rowData[replySentCol] === true || String(rowData[replySentCol]).toLowerCase() === 'true' || dripStatusRaw === 'completed' || dripStatusRaw === 'terkirim');
  var isWaitingDrip = (!isReplyAlreadySent) || (dripStatusRaw === 'waiting_views' || dripStatusRaw === 'waiting_views_threshold' || dripStatusRaw === 'menunggu target');

  var replyTriggered = false;
  var dripErrorMessage = '';

  if (currentViewsNum >= targetViewsNum && isWaitingDrip) {
    try {
      if (!threadsPostId) {
        throw new Error('ID Utas #1 (root_thread_id) belum ada di Meta Threads.');
      }

      var affiliateLink = 's.shopee.co.id/affiliate';
      var prodId = rowData[headers.indexOf('product_id')];
      var prodsList = getSheetRecords('Products');
      var prodLookup = {};
      for (var pl = 0; pl < prodsList.length; pl++) {
        prodLookup[prodsList[pl].product_id] = prodsList[pl];
        if (prodsList[pl].product_id === prodId && prodsList[pl].shopee_affiliate_url) {
          affiliateLink = prodsList[pl].shopee_affiliate_url;
        }
      }
      var cleanShortAff = stripProtocol(affiliateLink).replace(/^https?:\/\//i, '').trim();

      var nodes = [];
      try {
        var rawNodes = rowData[headers.indexOf('thread_content_json')];
        nodes = typeof rawNodes === 'string' ? JSON.parse(rawNodes) : rawNodes;
      } catch(e) {}
      if (!Array.isArray(nodes)) nodes = [nodes];

      if (nodes.length > 1) {
        var mediaRecords = getSheetRecords('Media');
        var mediaMap = {};
        for (var m = 0; m < mediaRecords.length; m++) {
          mediaMap[String(mediaRecords[m].media_id).trim()] = mediaRecords[m];
        }

        var chainParentId = threadsPostId;
        for (var step = 1; step < nodes.length; step++) {
          var subNode = nodes[step];
          var subText = typeof subNode === 'string' ? subNode : (subNode.content || '');
          subText = resolveAffiliateLinks(subText, subNode, prodLookup, rowData[headers.indexOf('affiliate_sub_id')], prodId);

          if (subText.indexOf('[LINK') !== -1) {
            subText = subText.replace(/\[(?:LINK_AFFILIATE|LINK)\]/g, cleanShortAff);
          } else if (subText.indexOf(cleanShortAff) === -1 && cleanShortAff) {
            subText = subText + ' cek di ' + cleanShortAff;
          }
          subText = subText.replace(/https?:\/\/(shope\.ee\/|s\.shopee\.co\.id\/|shopee\.co\.id\/)/gi, '$1');

          var mIds = [];
          if (subNode && Array.isArray(subNode.media_ids)) {
            mIds = subNode.media_ids;
          } else if (subNode && subNode.media_id) {
            mIds = [String(subNode.media_id).trim()];
          }

          var mediaItems = [];
          for (var mi = 0; mi < mIds.length; mi++) {
            if (mediaMap[mIds[mi]]) mediaItems.push(mediaMap[mIds[mi]]);
          }

          if (mediaItems.length > 1) {
            chainParentId = publishThreadsCarousel(rowData[accCol], token, subText, mediaItems, chainParentId);
          } else {
            var sMedia = mediaItems.length === 1 ? mediaItems[0] : null;
            chainParentId = publishThreadsNode(rowData[accCol], token, subText, sMedia, chainParentId);
          }

          if (typeof subNode === 'object' && subNode !== null) {
            subNode.threads_post_id = chainParentId;
            subNode.permalink = getThreadPermalink(chainParentId, token);
            subNode.status = 'PUBLISHED';
          }

          Utilities.sleep(1500);
        }

        if (nodes[0] && !nodes[0].permalink) {
          nodes[0].threads_post_id = threadsPostId;
          nodes[0].permalink = permalink || getThreadPermalink(threadsPostId, token);
          nodes[0].status = 'PUBLISHED';
        }

        postSheet.getRange(targetRow, headers.indexOf('thread_content_json') + 1).setValue(JSON.stringify(nodes));
        postSheet.getRange(targetRow, replySentCol + 1).setValue(true);
        if (dripColIdx !== -1) postSheet.getRange(targetRow, dripColIdx + 1).setValue('COMPLETED');
        replyTriggered = true;

        // Tandai baris antrean Utas #2 di Posting_Queue menjadi 'Posted'
        var qSheet = ss.getSheetByName('Posting_Queue');
        if (qSheet) {
          var qData = qSheet.getDataRange().getValues();
          for (var q = 1; q < qData.length; q++) {
            var qPostId = String(qData[q][1] || '').trim();
            var qStatusNorm = String(qData[q][6] || '').toLowerCase().trim();
            var isWaitingTarget = (qStatusNorm === 'waiting_views_threshold' || qStatusNorm === 'menunggu target' || qStatusNorm === 'waiting_views');
            if (qPostId === String(postId) && isWaitingTarget) {
              qSheet.getRange(q + 1, 7).setValue('Posted');
              qSheet.getRange(q + 1, 8).setValue('');
              qSheet.getRange(q + 1, 9).setValue(new Date().toISOString());
            }
          }
        }
      } else if (rowData[autoReplyContentCol]) {
        var replyText = String(rowData[autoReplyContentCol]).replace(/\[(?:LINK_AFFILIATE|LINK)\]/g, cleanShortAff);
        replyText = replyText.replace(/https?:\/\/(shope\.ee\/|s\.shopee\.co\.id\/|shopee\.co\.id\/)/gi, '$1');
        var singleChildId = publishThreadsNode(rowData[accCol], token, replyText, null, threadsPostId);
        postSheet.getRange(targetRow, replySentCol + 1).setValue(true);
        if (dripColIdx !== -1) postSheet.getRange(targetRow, dripColIdx + 1).setValue('COMPLETED');
        replyTriggered = true;

        var qSheet2 = ss.getSheetByName('Posting_Queue');
        if (qSheet2) {
          var qData2 = qSheet2.getDataRange().getValues();
          for (var q2 = 1; q2 < qData2.length; q2++) {
            var qPostId2 = String(qData2[q2][1] || '').trim();
            var qStatusNorm2 = String(qData2[q2][6] || '').toLowerCase().trim();
            var isWaitingTarget2 = (qStatusNorm2 === 'waiting_views_threshold' || qStatusNorm2 === 'menunggu target' || qStatusNorm2 === 'waiting_views');
            if (qPostId2 === String(postId) && isWaitingTarget2) {
              qSheet2.getRange(q2 + 1, 7).setValue('Posted');
              qSheet2.getRange(q2 + 1, 8).setValue('');
              qSheet2.getRange(q2 + 1, 9).setValue(new Date().toISOString());
            }
          }
        }
      }
    } catch (dripErr) {
      dripErrorMessage = (dripErr && dripErr.message) ? dripErr.message : String(dripErr);
      Logger.log('[Auto-Drip Views Error post_id: ' + postId + '] ' + dripErrorMessage);
      var qSheetErr = ss.getSheetByName('Posting_Queue');
      if (qSheetErr) {
        var qDataErr = qSheetErr.getDataRange().getValues();
        for (var qe = 1; qe < qDataErr.length; qe++) {
          var qpId = String(qDataErr[qe][1] || '').trim();
          var qStat = String(qDataErr[qe][6] || '').toLowerCase().trim();
          if (qpId === String(postId) && (qStat === 'waiting_views_threshold' || qStat === 'menunggu target' || qStat === 'waiting_views')) {
            qSheetErr.getRange(qe + 1, 8).setValue('Gagal Drip: ' + dripErrorMessage);
          }
        }
      }
    }
  }

  return {
    post_id: postId,
    views: views,
    likes: likes,
    replies: replies,
    permalink: permalink,
    comments_synced: commentsSynced,
    reply_triggered: replyTriggered
  };
}

// Sinkronkan postingan aktif secara cerdas untuk menghemat kuota UrlFetch & Meta API
function syncAllPublishedPosts() {
  var posts = getSheetRecords('Posts');
  var now = Date.now();
  var twoDaysMs = 48 * 60 * 60 * 1000;

  // Filter kandidat postingan:
  // 1. Yang belum kirim auto-reply (prioritas monetisasi utama)
  // 2. Atau yang berumur di bawah 48 jam (masih aktif mendapat views)
  var candidates = [];
  for (var i = 0; i < posts.length; i++) {
    var p = posts[i];
    if (p.status === 'PUBLISHED' && p.threads_post_id) {
      var isReplySent = (p.auto_reply_sent === true || String(p.auto_reply_sent).toLowerCase() === 'true');
      var postTime = p.created_at ? new Date(p.created_at).getTime() : 0;
      var isFresh = (now - postTime) < twoDaysMs;

      if (!isReplySent || isFresh) {
        candidates.push({ post: p, priority: !isReplySent ? 1 : 2 });
      }
    }
  }

  // Urutkan: prioritas auto-reply belum terkirim paling atas
  candidates.sort(function(a, b) { return a.priority - b.priority; });

  // Batasi maks 4 postingan per eksekusi agar hemat kuota & durasi < 8 detik
  var maxBatch = Math.min(candidates.length, 4);
  var updated = 0;
  for (var k = 0; k < maxBatch; k++) {
    try {
      syncPostInsights({ post_id: candidates[k].post.post_id });
      updated++;
    } catch (e) {}
  }

  return {
    updated: updated,
    message: 'Sinkronisasi cerdas: ' + updated + ' postingan aktif diproses.'
  };
}

// Sinkronkan interaksi audiens dari endpoint Meta Threads API (/conversation & /replies)
function syncInboxComments(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var interSheet = ss.getSheetByName('Thread_Interactions');
  if (!interSheet) {
    setupDatabase();
    interSheet = ss.getSheetByName('Thread_Interactions');
  }

  var cmtSheet = ss.getSheetByName('Comments');
  var posts = getSheetRecords('Posts');
  var accounts = getSheetRecords('Accounts');

  // Peta username dan ID akun sendiri untuk klasifikasi author_type
  var allOwnedUsers = {};
  var allOwnedIds = {};
  var tokenMap = {};
  for (var a = 0; a < accounts.length; a++) {
    var rawAcc = accounts[a].account_id;
    var accId = cleanAccountId(rawAcc);
    var uName = String(accounts[a].username || '').toLowerCase().trim();
    var tok = cleanToken(accounts[a].access_token);
    if (uName) allOwnedUsers[uName] = true;
    if (accId) allOwnedIds[accId] = true;
    tokenMap[accId] = tok;
    tokenMap[String(rawAcc).trim()] = tok;
  }

  // Ambil postingan terbit yang memiliki threads_post_id
  var publishedPosts = [];
  for (var p = 0; p < posts.length; p++) {
    if (posts[p].status === 'PUBLISHED' && posts[p].threads_post_id) {
      if (!payload || !payload.post_id || String(posts[p].post_id) === String(payload.post_id)) {
        publishedPosts.push(posts[p]);
      }
    }
  }

  if (publishedPosts.length === 0) {
    return { synced_count: 0, message: 'Tidak ada postingan terbit aktif untuk disinkronkan' };
  }

  var existingInteractions = getSheetRecords('Thread_Interactions');
  var existingIds = {};
  for (var e = 0; e < existingInteractions.length; e++) {
    var key = String(existingInteractions[e].comment_id || existingInteractions[e].interaction_id).trim();
    if (key) existingIds[key] = true;
  }

  var syncedCount = 0;
  var maxPostsToScan = Math.min(publishedPosts.length, 10);

  for (var i = 0; i < maxPostsToScan; i++) {
    var post = publishedPosts[i];
    var accId = cleanAccountId(post.account_id);
    var token = tokenMap[accId] || tokenMap[String(post.account_id).trim()];
    if (!token && accounts.length > 0) token = cleanToken(accounts[0].access_token);
    if (!token) continue;

    var threadsPostId = String(post.threads_post_id).trim();
    var rawList = [];

    // 1. Coba endpoint /conversation dengan fields valid Threads API
    var convUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/conversation?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(token);
    try {
      var resp = UrlFetchApp.fetch(convUrl, { method: 'get', muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        var json = JSON.parse(resp.getContentText() || '{}');
        if (json.data && Array.isArray(json.data) && json.data.length > 0) {
          rawList = json.data;
        }
      }
    } catch (errConv) {}

    // 2. Fallback / gabungkan dengan endpoint /replies
    try {
      var repUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(token);
      var resp2 = UrlFetchApp.fetch(repUrl, { method: 'get', muteHttpExceptions: true });
      if (resp2.getResponseCode() === 200) {
        var json2 = JSON.parse(resp2.getContentText() || '{}');
        if (json2.data && Array.isArray(json2.data)) {
          var seenInBatch = {};
          for (var b = 0; b < rawList.length; b++) {
            seenInBatch[String(rawList[b].id)] = true;
          }
          for (var r = 0; r < json2.data.length; r++) {
            if (!seenInBatch[String(json2.data[r].id)]) {
              rawList.push(json2.data[r]);
              seenInBatch[String(json2.data[r].id)] = true;
            }
          }
        }
      }
    } catch (errRep) {}

    // 3. Simpan interaksi ke database Google Sheets & tarik sub-replies berjenjang
    for (var k = 0; k < rawList.length; k++) {
      var item = rawList[k];
      var commentId = String(item.id || '').trim();
      if (!commentId || existingIds[commentId]) continue;

      var sender = String(item.username || 'user').trim();
      var senderUser = sender.toLowerCase();
      var senderUserId = cleanAccountId(item.user_id || (item.from && item.from.id) || '');
      var isOwnAccount = (senderUser && allOwnedUsers[senderUser]) || (senderUserId && allOwnedIds[senderUserId]) || (senderUser === 'adminagent');

      // Lewati balasan akun sendiri agar Inbox hanya berisi percakapan audiens luar
      if (isOwnAccount) {
        existingIds[commentId] = true;
        continue;
      }

      var text = item.text || '';
      var intentFlag = /spill|link|beli|toko|harga|order|checkout|beli di mana|buy|mau/i.test(text);
      var interId = 'INT_' + commentId.substring(0, 10);
      var createdAt = item.timestamp || new Date().toISOString();

      interSheet.appendRow([
        interId,
        post.post_id,
        threadsPostId,
        post.account_id,
        commentId,
        '',
        sender,
        'AUDIENCE',
        text,
        intentFlag,
        'UNREPLIED',
        createdAt
      ]);

      if (cmtSheet) {
        cmtSheet.appendRow([
          commentId,
          post.post_id,
          post.account_id,
          '',
          sender,
          text,
          intentFlag,
          'PUBLISHED',
          createdAt
        ]);
        syncedCount++;
      }

      existingIds[commentId] = true;

      // Tarik sub-replies audiens level 2 jika ada
      try {
        var subRepUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(commentId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(token);
        var subResp = UrlFetchApp.fetch(subRepUrl, { method: 'get', muteHttpExceptions: true });
        if (subResp.getResponseCode() === 200) {
          var subJson = JSON.parse(subResp.getContentText() || '{}');
          if (subJson.data && Array.isArray(subJson.data)) {
            for (var s = 0; s < subJson.data.length; s++) {
              var sItem = subJson.data[s];
              var sId = String(sItem.id || '').trim();
              if (!sId || existingIds[sId]) continue;

              var sSender = String(sItem.username || 'user').trim();
              var sSenderUser = sSender.toLowerCase();
              var sSenderUserId = cleanAccountId(sItem.user_id || (sItem.from && sItem.from.id) || '');
              var sIsOwn = (sSenderUser && allOwnedUsers[sSenderUser]) || (sSenderUserId && allOwnedIds[sSenderUserId]) || (sSenderUser === 'adminagent');

              if (sIsOwn) {
                existingIds[sId] = true;
                continue;
              }

              var sText = sItem.text || '';
              var sIntent = /spill|link|beli|toko|harga|order|checkout|beli di mana|buy|mau/i.test(sText);

              interSheet.appendRow([
                'INT_' + sId.substring(0, 10),
                post.post_id,
                threadsPostId,
                post.account_id,
                sId,
                commentId,
                sSender,
                'AUDIENCE',
                sText,
                sIntent,
                'UNREPLIED',
                sItem.timestamp || new Date().toISOString()
              ]);
              if (cmtSheet) {
                cmtSheet.appendRow([
                  sId,
                  post.post_id,
                  post.account_id,
                  commentId,
                  sSender,
                  sText,
                  sIntent,
                  'PUBLISHED',
                  sItem.timestamp || new Date().toISOString()
                ]);
                syncedCount++;
              }
              existingIds[sId] = true;
            }
          }
        }
      } catch (errSub) {}
    }
  }

  return {
    synced_count: syncedCount,
    message: 'Sinkronisasi selesai: ' + syncedCount + ' interaksi baru dari Meta Threads API.'
  };
}

function quickReplyComment(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cmtSheet = ss.getSheetByName('Comments');
  var interSheet = ss.getSheetByName('Thread_Interactions');

  var posts = getSheetRecords('Posts');
  var prods = getSheetRecords('Products');
  var accounts = getSheetRecords('Accounts');

  var targetPost = null;
  for (var p = 0; p < posts.length; p++) {
    if (posts[p].post_id === payload.post_id) {
      targetPost = posts[p];
      break;
    }
  }

  var affiliateLink = 'shope.ee/affiliate';
  if (targetPost) {
    for (var r = 0; r < prods.length; r++) {
      if (prods[r].product_id === targetPost.product_id) {
        affiliateLink = prods[r].shopee_affiliate_url || affiliateLink;
        break;
      }
    }
    var sep = affiliateLink.indexOf('?') === -1 ? '?' : '&';
    affiliateLink += sep + 'sub_id=' + (targetPost.affiliate_sub_id || 'aff_sub');
  }

  var finalContent = (payload.reply_content || '').replace(/\[(?:LINK_AFFILIATE|LINK)\]/g, stripProtocol(affiliateLink));
  finalContent = finalContent.replace(/https?:\/\/(shope\.ee\/|s\.shopee\.co\.id\/|shopee\.co\.id\/)/gi, '$1');
  var parentCommentId = payload.comment_id;
  var targetAccId = cleanAccountId(payload.account_id);

  var token = '';
  var username = 'AdminAgent';
  for (var a = 0; a < accounts.length; a++) {
    if (cleanAccountId(accounts[a].account_id) === targetAccId) {
      token = cleanToken(accounts[a].access_token);
      username = accounts[a].username || username;
      break;
    }
  }
  if (!token && accounts.length > 0) {
    token = cleanToken(accounts[0].access_token);
    username = accounts[0].username || username;
  }

  var liveReplyId = 'CMT_' + Utilities.getUuid().substring(0, 8);
  if (token && parentCommentId) {
    try {
      liveReplyId = publishThreadsNode(targetAccId, token, finalContent, null, parentCommentId);
    } catch (e) {
      // jika gagal live tetap catat id lokal
    }
  }

  var nowIso = new Date().toISOString();

  if (cmtSheet) {
    cmtSheet.appendRow([
      liveReplyId,
      payload.post_id,
      payload.account_id,
      parentCommentId,
      username,
      finalContent,
      false,
      'PUBLISHED',
      nowIso
    ]);
  }

  // Update status balasan pada sheet Thread_Interactions
  if (interSheet) {
    var data = interSheet.getDataRange().getValues();
    var headers = data[0];
    var cmtCol = headers.indexOf('comment_id');
    var statusCol = headers.indexOf('reply_status');

    if (cmtCol !== -1 && statusCol !== -1) {
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][cmtCol]) === String(parentCommentId)) {
          interSheet.getRange(r + 1, statusCol + 1).setValue('REPLIED');
          break;
        }
      }
    }

    // Catat baris balasan OWNER
    interSheet.appendRow([
      'INT_' + Utilities.getUuid().substring(0, 8),
      payload.post_id,
      '',
      payload.account_id,
      liveReplyId,
      parentCommentId,
      username,
      'OWNER',
      finalContent,
      false,
      'SENT',
      nowIso
    ]);
  }

  return { comment_id: liveReplyId, content: finalContent, reply_status: 'REPLIED' };
}

// Modul Autopilot & Referensi Viral (Dukungan Rangkaian Utas Dinamis)
function saveViralReference(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Viral_References');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('Viral_References');
  }

  var refId = payload.ref_id ? String(payload.ref_id).trim() : ('VREF_' + Utilities.getUuid().substring(0, 8));
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === refId) {
      foundRow = i + 1;
      break;
    }
  }

  // Ekstrak rangkaian utas dinamis (Utas 1, 2, 3, dst.) dengan dukungan konten teks & media
  var rawThreads = [];
  var allMediaUrls = [];
  if (Array.isArray(payload.threads) && payload.threads.length > 0) {
    for (var ti = 0; ti < payload.threads.length; ti++) {
      var item = payload.threads[ti];
      if (typeof item === 'string') {
        var str = item.trim();
        if (str) rawThreads.push({ content: str, media_urls: [] });
      } else if (item && typeof item === 'object') {
        var cText = String(item.content || '').trim();
        var mList = Array.isArray(item.media_urls) ? item.media_urls.filter(Boolean) : [];
        for (var mi = 0; mi < mList.length; mi++) {
          if (allMediaUrls.indexOf(mList[mi]) === -1) allMediaUrls.push(mList[mi]);
        }
        if (cText || mList.length > 0) {
          rawThreads.push({ content: cText, media_urls: mList });
        }
      }
    }
  } else {
    if (payload.hook_thread_1) rawThreads.push({ content: String(payload.hook_thread_1).trim(), media_urls: [] });
    if (payload.spill_thread_2) rawThreads.push({ content: String(payload.spill_thread_2).trim(), media_urls: [] });
  }
  if (rawThreads.length === 0) rawThreads = [{ content: '', media_urls: [] }];

  var hook1 = rawThreads[0] ? (rawThreads[0].content || '') : '';
  var spill2 = (rawThreads.length > 1 && rawThreads[1]) ? (rawThreads[1].content || '') : '';
  var threadsJson = JSON.stringify(rawThreads);

  var title = (payload.title && String(payload.title).trim()) ? String(payload.title).trim() : '';
  if (!title) {
    var cleanHook = hook1.replace(/[\r\n\t]+/g, ' ').trim();
    title = cleanHook ? (cleanHook.length > 30 ? (cleanHook.substring(0, 30) + '...') : cleanHook) : 'Referensi Viral';
  }

  var rowVals = [
    refId,
    payload.product_id || '',
    title,
    hook1,
    spill2,
    payload.original_url || '',
    JSON.stringify(allMediaUrls),
    payload.created_at || new Date().toISOString(),
    threadsJson
  ];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowVals.length).setValues([rowVals]);
  } else {
    sheet.appendRow(rowVals);
  }

  return { ref_id: refId };
}

function deleteViralReference(payload) {
  var refId = payload.ref_id ? String(payload.ref_id).trim() : '';
  if (!refId) throw new Error('Reference ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Viral_References');
  if (!sheet) throw new Error('Sheet Viral_References belum ada');

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === refId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error('Referensi viral tidak ditemukan');
}

// Generate Variasi Remake dengan Sumopod API
function generateSumopodRemake(payload) {
  var settings = getDashboardData().settings || {};
  var apiKey = settings.sumopod_api_key || (payload ? payload.api_key : '');
  if (!apiKey) throw new Error('Sumopod API Key belum dikonfigurasi di Pengaturan Sistem (Settings).');

  var rawEndpoint = (settings.sumopod_endpoint || '').trim();
  var endpoint = buildSumopodEndpoint(rawEndpoint);
  var model = (settings.sumopod_model && settings.sumopod_model.trim()) ? settings.sumopod_model.trim() : 'claude-haiku-4-5';

  var systemPrompt = 'Anda adalah spesialis Meta Threads copywriter Shopee Affiliate dengan metode STRICT STYLE MIRRORING (Meniru Gaya Referensi 1:1).\n' +
    'ATURAN KETAT:\n' +
    '1. DILARANG membuat template kaku atau bullet list spesifikasi brosur.\n' +
    '2. DILARANG memunculkan URL/link mentah apa pun. Gunakan HANYA tag placeholder [LINK] bila dibutuhkan.\n' +
    '3. Utas 1 (Root Hook): WAJIB 100% BEBAS DARI SEGALA LINK atau tag link agar distribusi jangkauan organik Threads maksimal.\n' +
    '4. Utas 2 (Lanjutan / Spill): WAJIB SANGAT RINGKAS (maksimal 1-2 kalimat santai mengalir, DILARANG panjang lebar atau berbau brosur/iklan). Sisipkan HANYA tag placeholder [LINK] menyatu secara natural di tengah kalimat (inline context), BUKAN di baris terbawah seperti footer iklan.\n' +
    '   Contoh pola yang diinginkan:\n' +
    '   - "Bisa custom nama disini [LINK] biar ga ketuker"\n' +
    '   - "aku ambil yg ukuran sedang di [LINK] pas bgt di wastafel"\n' +
    '   - "kalo mau checkout lgsg di toko ini ya [LINK] mumpung gratis ongkir"\n' +
    '5. Format balasan WAJIB berupa JSON valid murni tanpa teks pengantar:\n' +
    '{"hook_1": "...", "spill_2": "...", "threads": ["Utas 1...", "Utas 2..."]}';

  var winnerContext = '';
  if (payload && payload.winner_hook) {
    winnerContext = '\n[SELF-LEARNING WINNER CONTEXT - Postingan Juara Sebelumnya Menembus ' +
      (payload.winner_views || 0) + ' Views]:\n"' + payload.winner_hook + '"\n' +
      'Tiru struktur emosi dan penekanan kata dari postingan juara di atas!\n';
  }

  var threadsSourceText = '';
  if (payload && Array.isArray(payload.threads) && payload.threads.length > 0) {
    threadsSourceText = payload.threads.map(function(t, idx) {
      return 'Referensi Asli Utas #' + (idx + 1) + ':\n"' + (typeof t === 'string' ? t : (t.content || '')) + '"';
    }).join('\n\n');
  } else {
    threadsSourceText = 'Referensi Asli Utas #1 (Hook):\n"' + (payload ? (payload.hook_thread_1 || '') : '') + '"\n\n' +
      'Referensi Asli Utas #2 (Spill):\n"' + (payload ? (payload.spill_thread_2 || '') : '') + '"';
  }

  // Isolasi ketat: bersihkan URL penawaran/dashboard agar tidak bocor ke prompt AI
  var cleanProdName = String(payload ? (payload.product_name || '') : '').replace(/https?:\/\/[^\s]+/gi, '').trim();
  var cleanSellingPoints = String(payload ? (payload.selling_points || '-') : '-').replace(/https?:\/\/[^\s]+/gi, '').trim();

  var userPrompt = '[PRODUK SHOPEE TARGET]:\n' +
    'Nama Produk: ' + cleanProdName + '\n' +
    'Keunggulan / Selling Points: ' + cleanSellingPoints + '\n\n' +
    '[TEKS REFERENSI VIRAL YANG WAJIB DITIRU 1:1 GAYANYA]:\n' +
    threadsSourceText + '\n' +
    winnerContext + '\n' +
    'Instruksi: Terapkan STRICT STYLE MIRRORING sekarang. Tukar konteks masalah dan nama barang ke produk Shopee target. Utas 2 WAJIB SANGAT RINGKAS (1-2 kalimat santai) dan merajut tag [LINK] menyatu di dalam kalimat, bukan di baris paling bawah.';

  var body = {
    model: model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.7
  };

  var response;
  try {
    response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + cleanToken(apiKey),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
  } catch (netErr) {
    Logger.log('[Sumopod Network Error generateSumopodRemake] ' + netErr.message);
    throw new Error('Gagal terhubung ke server SumoPod: ' + netErr.message);
  }

  var code = response.getResponseCode();
  var rawText = (response.getContentText() || '').trim();
  Logger.log('[Sumopod generateSumopodRemake response HTTP ' + code + ']: ' + rawText);

  if (rawText.indexOf('<') === 0) {
    throw new Error('SumoPod API Error (HTTP ' + code + '): Server AI mengembalikan format HTML. Periksa URL endpoint atau status kuota.');
  }

  var resJson = {};
  try {
    resJson = JSON.parse(rawText);
  } catch(e) {
    Logger.log('[Sumopod JSON Parse Error generateSumopodRemake]: ' + rawText);
    throw new Error('SumoPod API Error (HTTP ' + code + '): Format respons bukan JSON: ' + rawText.substring(0, 150));
  }

  if (code !== 200 || !resJson || !resJson.choices || resJson.choices.length === 0) {
    var errMsg = extractSumopodError(code, resJson, endpoint);
    Logger.log('[Sumopod API Failure generateSumopodRemake]: ' + errMsg);
    throw new Error(errMsg);
  }

  var replyContent = resJson.choices[0].message.content || '';
  var cleanJsonStr = replyContent.replace(/```json/gi, '').replace(/```/g, '').trim();

  try {
    var parsed = JSON.parse(cleanJsonStr);
    var outThreads = Array.isArray(parsed.threads) ? parsed.threads : [];
    if (outThreads.length === 0) {
      if (parsed.hook_1) outThreads.push(parsed.hook_1);
      if (parsed.spill_2) outThreads.push(parsed.spill_2);
    }
    return {
      hook_1: parsed.hook_1 || (outThreads[0] || ''),
      spill_2: parsed.spill_2 || (outThreads[1] || ''),
      threads: outThreads
    };
  } catch(parseErr) {
    return {
      hook_1: replyContent,
      spill_2: payload.spill_thread_2 || '',
      threads: [replyContent, payload.spill_thread_2 || ''].filter(Boolean)
    };
  }
}

// Jadwalkan Batch Autopilot dengan Strict Session Binding & Jitter Delay
function scheduleAutopilotBatch(payload) {
  var prodId = payload.product_id ? String(payload.product_id).trim() : '';
  var targetSessionId = payload.session_id ? String(payload.session_id).trim() : '';
  var accounts = payload.account_ids || [];
  if (!prodId) throw new Error('Pilih produk target untuk Autopilot');
  if (!accounts || accounts.length === 0) throw new Error('Centang minimal 1 akun target');

  var allMedia = getSheetRecords('Media');
  // Ambil paket media utuh HANYA dari sesi ulasan terpilih (No-Random Media)
  var sessionMedia = [];
  for (var m = 0; m < allMedia.length; m++) {
    var row = allMedia[m];
    if (row.product_id === prodId) {
      if (targetSessionId && row.session_id === targetSessionId) {
        sessionMedia.push(row.media_id);
      } else if (!targetSessionId) {
        targetSessionId = row.session_id; // kunci ke sesi pertama yang ditemukan
        sessionMedia.push(row.media_id);
      }
    }
  }

  if (sessionMedia.length === 0) {
    throw new Error('Produk ini belum memiliki file media dari sesi ulasan yang valid');
  }

  var prods = getSheetRecords('Products');
  var prodObj = null;
  for (var p = 0; p < prods.length; p++) {
    if (prods[p].product_id === prodId) {
      prodObj = prods[p];
      break;
    }
  }

  if (prodObj) {
    var rawAllowed = String(prodObj.allowed_accounts || '').trim();
    if (!rawAllowed) {
      throw new Error('Produk "' + (prodObj.product_name || prodId) + '" berstatus Belum Terhubung dan tidak diizinkan untuk diposting.');
    }
  }

  var baseTime = payload.start_time ? new Date(payload.start_time).getTime() : (Date.now() + 10 * 60 * 1000);
  if (isNaN(baseTime)) baseTime = Date.now() + 10 * 60 * 1000;
  var jitterBaseMin = Number(payload.jitter_minutes || 15);
  var threshold = Number(payload.views_threshold || 500);

  var createdPosts = [];

  for (var a = 0; a < accounts.length; a++) {
    var accId = accounts[a];
    // Jitter delay acak antar akun (contoh: offset 12-18 menit per giliran)
    var jitterOffsetMs = (a * jitterBaseMin * 60000) + Math.floor(Math.random() * (jitterBaseMin * 0.4 * 60000));
    var scheduledDate = new Date(baseTime + jitterOffsetMs);

    var cleanSpill = payload.spill_thread_2 || '';
    if (cleanSpill.indexOf('[LINK_AFFILIATE') === -1) {
      var shortAff = prodObj ? (prodObj.shopee_affiliate_url || 'shope.ee/affiliate') : 'shope.ee/affiliate';
      cleanSpill = cleanSpill + '\n\n' + stripProtocol(shortAff);
    }

    var threadNodes = [
      {
        content: payload.hook_thread_1 || 'Rekomendasi terbaik wajib punya!',
        product_ids: [prodId],
        product_id: prodId,
        media_ids: sessionMedia,
        media_id: sessionMedia[0]
      },
      {
        content: cleanSpill,
        product_ids: [prodId],
        product_id: prodId,
        media_ids: [],
        media_id: ''
      }
    ];

    var postResult = createPost({
      account_id: accId,
      product_id: prodId,
      publish_mode: 'DRIP_ON_VIEWS',
      thread_content_json: JSON.stringify(threadNodes),
      media_ids_json: JSON.stringify(sessionMedia),
      views_threshold: threshold,
      scheduled_time: scheduledDate.toISOString(),
      status: 'SCHEDULED'
    });

    createdPosts.push({
      post_id: postResult.post_id,
      account_id: accId,
      scheduled_time: scheduledDate.toISOString()
    });
  }

  return {
    scheduled_count: createdPosts.length,
    session_id: targetSessionId,
    media_count: sessionMedia.length,
    posts: createdPosts
  };
}

// Mesin Autopilot Otonom: Simpan, Toggle & Hapus Aturan Autopilot
function saveAutopilotRule(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Autopilot_Rules');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('Autopilot_Rules');
  }

  var ruleId = payload.rule_id ? String(payload.rule_id).trim() : ('RULE_' + Utilities.getUuid().substring(0, 8));
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === ruleId) {
      foundRow = i + 1;
      break;
    }
  }

  var accJson = '[]';
  if (Array.isArray(payload.account_ids)) {
    accJson = JSON.stringify(payload.account_ids);
  } else if (payload.account_ids_json) {
    accJson = String(payload.account_ids_json);
  }

  var existingLastSession = '';
  if (foundRow > 0 && data[foundRow - 1].length >= 14) {
    existingLastSession = data[foundRow - 1][13] || '';
  }

  var windowsJson = '["06:00-08:30","11:30-13:30","16:30-18:30","19:30-21:30"]';
  if (Array.isArray(payload.time_windows)) {
    windowsJson = JSON.stringify(payload.time_windows);
  } else if (payload.time_windows_json) {
    windowsJson = String(payload.time_windows_json);
  }

  var targetPosts = Number(payload.target_posts || payload.max_posts_per_day || payload.frequency_per_day || 3);
  if (targetPosts <= 0) targetPosts = 3;

  var rowVals = [
    ruleId,
    payload.name || ('Aturan Autopilot ' + (payload.product_id || '')),
    accJson,
    payload.product_id || '',
    payload.preferred_session_id || '',
    targetPosts,
    Number(payload.views_threshold || 500),
    Number(payload.jitter_minutes || 15),
    payload.status || 'ACTIVE',
    payload.last_run_at || '',
    payload.created_at || new Date().toISOString(),
    targetPosts,
    0,
    payload.last_session_id || existingLastSession,
    windowsJson
  ];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowVals.length).setValues([rowVals]);
  } else {
    sheet.appendRow(rowVals);
  }

  return { rule_id: ruleId, status: 'saved' };
}

function toggleAutopilotRule(payload) {
  var ruleId = payload.rule_id ? String(payload.rule_id).trim() : '';
  if (!ruleId) throw new Error('Rule ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Autopilot_Rules');
  if (!sheet) throw new Error('Sheet Autopilot_Rules belum ada');

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === ruleId) {
      var currentStatus = String(data[i][8] || 'ACTIVE').toUpperCase();
      var newStatus = currentStatus === 'ACTIVE' ? 'PAUSED' : 'ACTIVE';
      sheet.getRange(i + 1, 9).setValue(newStatus);
      return { rule_id: ruleId, status: newStatus };
    }
  }
  throw new Error('Aturan tidak ditemukan');
}

function deleteAutopilotRule(payload) {
  var ruleId = payload.rule_id ? String(payload.rule_id).trim() : '';
  if (!ruleId) throw new Error('Rule ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Autopilot_Rules');
  if (!sheet) throw new Error('Sheet Autopilot_Rules belum ada');

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === ruleId) {
      sheet.deleteRow(i + 1);
      return { success: true, rule_id: ruleId };
    }
  }
  throw new Error('Aturan tidak ditemukan');
}

// Helper pembagian titik jadwal terdistribusi dinamis harian (Batch Daily Scheduling)
function generateDistributedScheduleSlots(windowsList, targetPosts, jitterMin, accIdx, now) {
  var windows = Array.isArray(windowsList) && windowsList.length > 0
    ? windowsList
    : ['06:00-08:30', '11:30-13:30', '16:30-18:30', '19:30-21:30'];

  var parsedWindows = [];
  var totalActiveMinutes = 0;

  for (var w = 0; w < windows.length; w++) {
    var parts = String(windows[w]).split('-');
    var startParts = (parts[0] || '06:00').trim().split(':');
    var endParts = (parts[1] || '08:30').trim().split(':');
    var startMin = (parseInt(startParts[0], 10) || 0) * 60 + (parseInt(startParts[1], 10) || 0);
    var endMin = (parseInt(endParts[0], 10) || 0) * 60 + (parseInt(endParts[1], 10) || 0);
    if (endMin <= startMin) endMin = startMin + 60;
    var dur = endMin - startMin;
    parsedWindows.push({
      startMin: startMin,
      endMin: endMin,
      duration: dur,
      cumStart: totalActiveMinutes,
      cumEnd: totalActiveMinutes + dur
    });
    totalActiveMinutes += dur;
  }

  if (totalActiveMinutes <= 0) {
    totalActiveMinutes = 480;
    parsedWindows = [{ startMin: 480, endMin: 960, duration: 480, cumStart: 0, cumEnd: 480 }];
  }

  var count = Math.max(1, parseInt(targetPosts, 10) || 3);
  var interval = totalActiveMinutes / count;
  var slots = [];
  var baseNow = (now instanceof Date) ? now : new Date(now || Date.now());

  for (var i = 0; i < count; i++) {
    var targetOffset = (i + 0.5) * interval;

    var chosenWin = parsedWindows[0];
    for (var j = 0; j < parsedWindows.length; j++) {
      if (targetOffset >= parsedWindows[j].cumStart && targetOffset < parsedWindows[j].cumEnd) {
        chosenWin = parsedWindows[j];
        break;
      }
    }

    var minIntoWin = targetOffset - chosenWin.cumStart;
    var minuteOfDay = chosenWin.startMin + minIntoWin;

    // Tambah jitter delay acak (+/- 2 menit)
    var randJitter = (Math.random() * 4 - 2);
    var accOffsetMin = (accIdx * 1.5);
    var finalMinute = minuteOfDay + randJitter + accOffsetMin;

    var slotDate = new Date(baseNow.getTime());
    slotDate.setHours(Math.floor(finalMinute / 60));
    slotDate.setMinutes(Math.floor(finalMinute % 60));
    slotDate.setSeconds(0);
    slotDate.setMilliseconds(0);

    // Jika waktu pada hari ini sudah lewat dari sekarang, jadwalkan untuk besok di jam yang sama
    if (slotDate.getTime() <= (baseNow.getTime() + 2 * 60 * 1000)) {
      slotDate.setDate(slotDate.getDate() + 1);
    }

    slots.push(slotDate);
  }

  slots.sort(function(a, b) { return a.getTime() - b.getTime(); });
  return slots;
}

// Background Batch Worker: Mesin Autopilot Batch Daily Scheduling
function runAutonomousAutopilotWorker(ruleIdFilter) {
  var rules = getSheetRecords('Autopilot_Rules');
  if (!rules || rules.length === 0) {
    return { generated_count: 0, message: 'Belum ada aturan autopilot tersimpan.' };
  }

  var activeRules = [];
  for (var r = 0; r < rules.length; r++) {
    var rObj = rules[r];
    if (ruleIdFilter && String(rObj.rule_id).trim() === String(ruleIdFilter).trim()) {
      activeRules.push(rObj);
      break;
    }
    if (String(rObj.status || '').toUpperCase() === 'ACTIVE') {
      activeRules.push(rObj);
    }
  }

  if (activeRules.length === 0) {
    return { generated_count: 0, message: 'Tidak ada aturan autopilot yang aktif saat ini.' };
  }

  var allProducts = getSheetRecords('Products');
  var allMedia = getSheetRecords('Media');
  var allViralRefs = getSheetRecords('Viral_References');
  var allAccounts = getSheetRecords('Accounts');
  var allPosts = getSheetRecords('Posts');
  var allQueue = getSheetRecords('Posting_Queue');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ruleSheet = ss.getSheetByName('Autopilot_Rules');
  var ruleData = ruleSheet ? ruleSheet.getDataRange().getValues() : [];
  var ruleHeaders = ruleData.length > 0 ? ruleData[0] : [];
  var lastRunCol = ruleHeaders.indexOf('last_run_at');
  var ruleIdCol = ruleHeaders.indexOf('rule_id');
  var lastSessionCol = ruleHeaders.indexOf('last_session_id');

  var totalGenerated = 0;
  var now = new Date();
  var nowMs = now.getTime();

  for (var i = 0; i < activeRules.length; i++) {
    var rule = activeRules[i];
    var isManualTrigger = !!ruleIdFilter;

    // Baca target_posts harian dinamis dari input user (misal 3, 5, 8, 10, dsb.)
    var targetPosts = Number(rule.target_posts || rule.max_posts_per_day || rule.frequency_per_day || 3);
    if (targetPosts <= 0) targetPosts = 3;

    var lastRunTime = rule.last_run_at ? new Date(rule.last_run_at).getTime() : 0;
    // Throttle cron background: jika bukan manual trigger, batasi eksekusi batch per 12 jam
    if (!isManualTrigger && lastRunTime && (nowMs - lastRunTime) < (12 * 3600 * 1000)) {
      continue;
    }

    var rawAccIds = [];
    try {
      rawAccIds = typeof rule.account_ids_json === 'string' ? JSON.parse(rule.account_ids_json) : (rule.account_ids_json || []);
    } catch(eAcc) {
      rawAccIds = String(rule.account_ids_json || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }
    if (!Array.isArray(rawAccIds) || rawAccIds.length === 0) {
      continue;
    }

    var jitterMin = Number(rule.jitter_minutes || 15);
    var thresh = Number(rule.views_threshold || 500);
    var ruleLastSession = rule.last_session_id || '';

    var mediaToSessionMap = {};
    for (var mi = 0; mi < allMedia.length; mi++) {
      var medRow = allMedia[mi];
      if (medRow.media_id && medRow.session_id) {
        mediaToSessionMap[String(medRow.media_id).trim()] = String(medRow.session_id).trim();
      }
    }

    var ruleWindows = [];
    try {
      ruleWindows = typeof rule.time_windows_json === 'string' ? JSON.parse(rule.time_windows_json) : (rule.time_windows_json || []);
    } catch(eWin) {}
    if (!Array.isArray(ruleWindows) || ruleWindows.length === 0) {
      ruleWindows = ['06:00-08:30', '11:30-13:30', '16:30-18:30', '19:30-21:30'];
    }

    for (var accIdx = 0; accIdx < rawAccIds.length; accIdx++) {
      var targetAccId = cleanAccountId(rawAccIds[accIdx]);
      var targetUsername = '';
      for (var aci = 0; aci < allAccounts.length; aci++) {
        if (cleanAccountId(allAccounts[aci].account_id) === targetAccId) {
          targetUsername = String(allAccounts[aci].username || '').toLowerCase().trim();
          break;
        }
      }

      // Filter produk terhubung sah ke akun ini
      var eligibleProductsForAcc = [];
      for (var pr = 0; pr < allProducts.length; pr++) {
        var pCandidate = allProducts[pr];
        var allowedStr = String(pCandidate.allowed_accounts || '').trim();
        if (!allowedStr) continue;

        var allowedList = allowedStr.split(',').map(function(s) { return s.trim().toLowerCase(); });
        var isAllowed = false;
        for (var al = 0; al < allowedList.length; al++) {
          var aItem = allowedList[al];
          if (cleanAccountId(aItem) === targetAccId || (targetUsername && aItem === targetUsername)) {
            isAllowed = true;
            break;
          }
        }
        if (!isAllowed) continue;

        var hasMedia = false;
        for (var md = 0; md < allMedia.length; md++) {
          if (allMedia[md].product_id === pCandidate.product_id) {
            hasMedia = true;
            break;
          }
        }
        if (hasMedia) eligibleProductsForAcc.push(pCandidate);
      }

      if (eligibleProductsForAcc.length === 0) continue;

      // Hitung titik waktu jadwal terdistribusi proporsional untuk seluruh kuota target_posts
      var distributedSlots = generateDistributedScheduleSlots(ruleWindows, targetPosts, jitterMin, accIdx, now);

      for (var slotIdx = 0; slotIdx < distributedSlots.length; slotIdx++) {
        var scheduledDate = distributedSlots[slotIdx];
        var chosenProd = eligibleProductsForAcc[slotIdx % eligibleProductsForAcc.length];

        // Kelompokkan sesi media untuk produk terpilih
        var prodSessions = {};
        var sessionOrder = [];
        for (var m = 0; m < allMedia.length; m++) {
          var med = allMedia[m];
          if (med.product_id === chosenProd.product_id) {
            var sId = med.session_id || 'SESI_DEFAULT';
            if (!prodSessions[sId]) {
              prodSessions[sId] = [];
              sessionOrder.push(sId);
            }
            prodSessions[sId].push(med.media_id);
          }
        }

        if (sessionOrder.length === 0) continue;

        var chosenSessionId = sessionOrder[slotIdx % sessionOrder.length];
        var chosenMediaIds = prodSessions[chosenSessionId];
        ruleLastSession = chosenSessionId;

        // Ambil referensi viral & context
        var matchingRefs = [];
        for (var v = 0; v < allViralRefs.length; v++) {
          if (allViralRefs[v].product_id === chosenProd.product_id) {
            matchingRefs.push(allViralRefs[v]);
          }
        }
        var chosenRef = matchingRefs.length > 0
          ? matchingRefs[slotIdx % matchingRefs.length]
          : (allViralRefs.length > 0 ? allViralRefs[slotIdx % allViralRefs.length] : null);

        var origHook = chosenRef ? (chosenRef.hook_thread_1 || '') : ('Rekomendasi ' + chosenProd.product_name + ' paling worth it!');
        var origSpill = chosenRef ? (chosenRef.spill_thread_2 || '') : 'Spill belinya di toko resmi ini: [LINK]';

        var remakeResult = { hook_1: origHook, spill_2: origSpill };
        try {
          remakeResult = generateSumopodRemake({
            product_name: chosenProd.product_name,
            selling_points: chosenProd.selling_points || '',
            hook_thread_1: origHook,
            spill_thread_2: origSpill
          });
        } catch(aiErr) {
          Logger.log('[Autopilot AI Remake Fallback]: ' + aiErr.message);
        }

        var hook1 = remakeResult.hook_1 || origHook;
        var spill2 = remakeResult.spill_2 || origSpill;

        // Isolasi ketat: ambil HANYA shopee_affiliate_url murni (shortlink), JANGAN PERNAH gunakan raw offer url
        var rawAff = chosenProd.shopee_affiliate_url ? String(chosenProd.shopee_affiliate_url).trim() : '';
        if (!rawAff) rawAff = 's.shopee.co.id/affiliate';
        var cleanShortAff = stripProtocol(rawAff).replace(/^https?:\/\//i, '').trim();

        var cleanSpill = spill2;
        if (cleanSpill.indexOf('[LINK') !== -1) {
          cleanSpill = cleanSpill.replace(/\[(?:LINK_AFFILIATE|LINK)\]/g, cleanShortAff);
        } else if (cleanSpill.indexOf(cleanShortAff) === -1) {
          cleanSpill = cleanSpill + ' cek di ' + cleanShortAff;
        }
        cleanSpill = cleanSpill.replace(/https?:\/\/(shope\.ee\/|s\.shopee\.co\.id\/|shopee\.co\.id\/)/gi, '$1');

        var threadNodes = [
          {
            content: hook1,
            product_ids: [chosenProd.product_id],
            product_id: chosenProd.product_id,
            media_ids: chosenMediaIds,
            media_id: chosenMediaIds[0],
            status: 'Pending_Scheduled'
          },
          {
            content: cleanSpill,
            product_ids: [chosenProd.product_id],
            product_id: chosenProd.product_id,
            media_ids: [],
            media_id: '',
            status: 'WAITING_VIEWS'
          }
        ];

        // Masukkan ke database & Posting_Queue berstatus Pending_Scheduled (TIDAK terbit saat itu juga)
        createPost({
          account_id: targetAccId,
          product_id: chosenProd.product_id,
          publish_mode: 'DRIP_ON_VIEWS',
          thread_content_json: JSON.stringify(threadNodes),
          media_ids_json: JSON.stringify(chosenMediaIds),
          views_threshold: thresh,
          scheduled_time: scheduledDate.toISOString(),
          status: 'Pending_Scheduled',
          source_type: 'autopilot',
          rule_id: rule.rule_id || ''
        });

        totalGenerated++;
      }
    }

    if (ruleSheet && ruleIdCol !== -1) {
      for (var rRow = 1; rRow < ruleData.length; rRow++) {
        if (String(ruleData[rRow][ruleIdCol]) === String(rule.rule_id)) {
          if (lastRunCol !== -1) {
            ruleSheet.getRange(rRow + 1, lastRunCol + 1).setValue(new Date().toISOString());
          }
          if (lastSessionCol !== -1 && ruleLastSession) {
            ruleSheet.getRange(rRow + 1, lastSessionCol + 1).setValue(ruleLastSession);
          }
          break;
        }
      }
    }
  }

  return {
    generated_count: totalGenerated,
    message: totalGenerated > 0
      ? ('Autopilot Batch Scheduling berhasil: ' + totalGenerated + ' postingan terjadwal dimasukkan ke antrean (Pending_Scheduled) sesuai jendela jam aktif.')
      : 'Semua akun autopilot berada dalam batas aman atau belum memiliki produk terhubung.'
  };
}

// Kelola Master Bank Hook
function saveHookTemplate(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Hook_Bank');
  var hookId = payload.hook_id || ('HK_' + Utilities.getUuid().substring(0, 6));
  sheet.appendRow([
    hookId,
    payload.category || 'General',
    payload.hook_text || '',
    payload.formula || '',
    new Date().toISOString()
  ]);
  return { hook_id: hookId };
}

function deleteHookTemplate(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Hook_Bank');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.hook_id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error('Hook template tidak ditemukan');
}

// Background Worker: Eksekusi antrean tayang DAN auto-reply saat views tembus
function checkAndExecuteQueueAndViews() {
  var queueProcessed = 0;
  var postsSynced = 0;
  var autopilotGenerated = 0;
  var dripProcessed = 0;

  try {
    var autoResult = runAutonomousAutopilotWorker();
    autopilotGenerated = autoResult && autoResult.generated_count ? autoResult.generated_count : 0;
  } catch (aErr) {
    // Isolasi error cron autopilot
  }

  try {
    var queueResult = checkAndPublishQueue();
    queueProcessed = queueResult && queueResult.processed ? queueResult.processed : 0;
  } catch (qErr) {
    // Isolasi error antrean agar views tetap terproses
  }

  try {
    var viewsResult = syncAllPublishedPosts();
    postsSynced = viewsResult && viewsResult.updated ? viewsResult.updated : 0;
  } catch (vErr) {
    // Isolasi error views
  }

  try {
    var dripResult = checkWaitingViewsQueue();
    dripProcessed = dripResult && dripResult.processed ? dripResult.processed : 0;
  } catch (dErr) {
    // Isolasi error antrean views
  }

  return {
    autopilot_generated: autopilotGenerated,
    queue_processed: queueProcessed,
    posts_synced: postsSynced,
    drip_processed: dripProcessed,
    timestamp: new Date().toISOString()
  };
}

// Memeriksa dan mengeksekusi Utas #2 dalam antrean jika target views root_post_id sudah tercapai
function checkWaitingViewsQueue() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var qSheet = ss.getSheetByName('Posting_Queue');
  if (!qSheet) return { processed: 0, checked: 0 };

  var qData = qSheet.getDataRange().getValues();
  if (qData.length <= 1) return { processed: 0, checked: 0 };

  var processed = 0;
  var checked = 0;

  for (var i = 1; i < qData.length; i++) {
    var rawStatus = String(qData[i][6] || '').toLowerCase().trim();
    var isWaiting = (rawStatus === 'waiting_views_threshold' || rawStatus === 'menunggu target' || rawStatus === 'waiting_views');
    if (!isWaiting) continue;

    var postId = String(qData[i][1] || '').trim();
    if (!postId) continue;
    checked++;

    try {
      var syncRes = syncPostInsights({ post_id: postId });
      if (syncRes && syncRes.reply_triggered) {
        processed++;
      }
    } catch (e) {
      var errDetail = e && e.message ? e.message : String(e);
      Logger.log('[checkWaitingViewsQueue Error] Post ' + postId + ': ' + errDetail);
      try {
        qSheet.getRange(i + 1, 8).setValue('Gagal Cek Views: ' + errDetail);
      } catch (logErr) {}
    }
  }

  return { processed: processed, checked: checked };
}

function saveSettings(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Settings');
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[data[i][0]] = i + 1;
  }

  for (var key in payload) {
    if (map[key]) {
      sheet.getRange(map[key], 2).setValue(payload[key]);
    } else {
      sheet.appendRow([key, payload[key]]);
    }
  }
  return { status: 'Pengaturan tersimpan' };
}

// Pasang Google Apps Script time-driven trigger (Default: 5 menit)
function setupTimeDrivenTrigger(payload) {
  var interval = (payload && payload.interval) ? Number(payload.interval) : 5;
  if ([1, 5, 10, 15].indexOf(interval) === -1) interval = 5;

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'checkAndPublishQueue' || fn === 'checkAndExecuteQueueAndViews') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('checkAndExecuteQueueAndViews')
    .timeBased()
    .everyMinutes(interval)
    .create();

  try {
    saveSettings({ trigger_interval: String(interval) });
  } catch(e) {}

  return {
    interval: interval,
    status: 'Trigger background checkAndExecuteQueueAndViews berhasil didaftarkan aktif setiap ' + interval + ' menit!'
  };
}

// Mengambil permalink resmi dari Meta Threads Graph API
function getThreadPermalink(threadsMediaId, token) {
  if (!threadsMediaId || !token) return '';
  try {
    var url = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsMediaId) + '?fields=id,permalink&access_token=' + encodeURIComponent(cleanToken(token));
    var resp = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      var json = JSON.parse(resp.getContentText() || '{}');
      return json.permalink || '';
    }
  } catch (e) {}
  return '';
}

// Validasi URL video Cloudinary wajib berakhiran .mp4
function validateAndFormatVideoUrl(rawUrl) {
  if (!rawUrl) throw new Error('URL video wajib diisi.');
  var url = String(rawUrl).trim();
  var parts = url.split('?');
  var basePath = parts[0];
  var query = parts.length > 1 ? ('?' + parts[1]) : '';

  // Jika URL Cloudinary tanpa ekstensi .mp4, transformasikan otomatis
  if (url.indexOf('cloudinary.com') !== -1 && !basePath.toLowerCase().endsWith('.mp4')) {
    if (/\.(mov|webm|avi|mkv|m4v)$/i.test(basePath)) {
      basePath = basePath.replace(/\.(mov|webm|avi|mkv|m4v)$/i, '.mp4');
    } else {
      basePath = basePath + '.mp4';
    }
    url = basePath + query;
  }

  if (!url.split('?')[0].toLowerCase().endsWith('.mp4')) {
    throw new Error('Format video tidak valid: Meta Threads mewajibkan file video berakhiran .mp4');
  }
  return url;
}

// Helper polling kesiapan kontainer media sebelum dipublikasi (2 tahap publikasi video)
function waitForContainerReady(containerId, token, isVideo) {
  var cleanTok = cleanToken(token);
  var statusUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(containerId) + '?fields=status,error_message&access_token=' + encodeURIComponent(cleanTok);
  var maxAttempts = isVideo ? 15 : 6;
  var delayMs = isVideo ? 2500 : 1500;

  for (var i = 0; i < maxAttempts; i++) {
    Utilities.sleep(delayMs);
    var res;
    try {
      res = UrlFetchApp.fetch(statusUrl, {
        method: 'get',
        muteHttpExceptions: true
      });
    } catch (netErr) {
      continue;
    }

    var json = {};
    try {
      json = JSON.parse(res.getContentText() || '{}');
    } catch (parseErr) {
      continue;
    }

    var status = json.status || '';
    if (status === 'FINISHED') {
      return true;
    }
    if (status === 'ERROR') {
      var detail = json.error_message || '';
      var advice = isVideo ? ' Pastikan durasi video maks 5 menit, resolusi maks 1920x1080, ukuran di bawah 1GB, dan codec video H.264/AAC.' : '';
      throw new Error('Meta menolak media: ' + (detail || 'Spesifikasi tidak sesuai.') + advice);
    }
    if (status === 'EXPIRED') {
      throw new Error('Kontainer media Meta kedaluwarsa sebelum dipublikasi.');
    }
  }

  if (isVideo) {
    throw new Error('Pemrosesan video di server Meta timeout (status masih IN_PROGRESS). Kurangi resolusi atau durasi video.');
  }
  return true;
}

// Helper penerbitan karusel multi-media ke Meta Threads Graph API
function publishThreadsCarousel(accountId, token, text, mediaArray, replyToId) {
  var cleanTok = cleanToken(token);
  var cleanAcc = cleanAccountId(accountId);

  if (!cleanTok) throw new Error('Access Token Meta kosong.');
  if (!mediaArray || mediaArray.length === 0) throw new Error('Daftar media karusel kosong.');
  if (mediaArray.length === 1) {
    return publishThreadsNode(accountId, token, text, mediaArray[0], replyToId);
  }

  // Batasi maksimal 10 item media per karusel sesuai batasan Meta Threads API
  var items = mediaArray.slice(0, 10);
  var childIds = [];

  for (var i = 0; i < items.length; i++) {
    var m = items[i];
    var isVideo = (m.media_type === 'VIDEO');
    var itemPayload = {
      access_token: cleanTok,
      is_carousel_item: 'true',
      media_type: isVideo ? 'VIDEO' : 'IMAGE'
    };

    if (isVideo) {
      itemPayload.video_url = validateAndFormatVideoUrl(m.media_url);
    } else {
      itemPayload.image_url = m.media_url;
    }

    var createItemUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(cleanAcc) + '/threads?access_token=' + encodeURIComponent(cleanTok);
    var resItem = UrlFetchApp.fetch(createItemUrl, {
      method: 'post',
      payload: itemPayload,
      muteHttpExceptions: true
    });

    var jsonItem = {};
    try {
      jsonItem = JSON.parse(resItem.getContentText() || '{}');
    } catch (e) {
      throw new Error('Respons item karusel tidak valid: ' + resItem.getContentText().substring(0, 80));
    }

    if (resItem.getResponseCode() !== 200 || !jsonItem.id) {
      var errDetail = jsonItem.error && jsonItem.error.message ? jsonItem.error.message : ('HTTP ' + resItem.getResponseCode());
      throw new Error('Gagal membuat container item karusel #' + (i + 1) + ': ' + errDetail);
    }

    waitForContainerReady(jsonItem.id, cleanTok, isVideo);
    childIds.push(jsonItem.id);
  }

  // Buat Kontainer Karusel Utama
  var createCarouselUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(cleanAcc) + '/threads?access_token=' + encodeURIComponent(cleanTok);
  var carouselPayload = {
    access_token: cleanTok,
    media_type: 'CAROUSEL',
    children: childIds.join(',')
  };

  if (text && text.trim().length > 0) {
    carouselPayload.text = text.trim();
  }
  if (replyToId) {
    carouselPayload.reply_to_id = replyToId;
  }

  var resCarousel = UrlFetchApp.fetch(createCarouselUrl, {
    method: 'post',
    payload: carouselPayload,
    muteHttpExceptions: true
  });

  var jsonCarousel = {};
  try {
    jsonCarousel = JSON.parse(resCarousel.getContentText() || '{}');
  } catch (e) {
    throw new Error('Respons container karusel tidak valid: ' + resCarousel.getContentText().substring(0, 80));
  }

  if (resCarousel.getResponseCode() !== 200 || !jsonCarousel.id) {
    var carErr = jsonCarousel.error && jsonCarousel.error.message ? jsonCarousel.error.message : ('HTTP ' + resCarousel.getResponseCode());
    throw new Error('Gagal membuat container karusel Meta: ' + carErr);
  }

  waitForContainerReady(jsonCarousel.id, cleanTok, false);

  // Terbitkan Kontainer Karusel
  var publishUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(cleanAcc) + '/threads_publish?access_token=' + encodeURIComponent(cleanTok);
  var resPub = UrlFetchApp.fetch(publishUrl, {
    method: 'post',
    payload: { creation_id: jsonCarousel.id, access_token: cleanTok },
    muteHttpExceptions: true
  });

  var jsonPub = {};
  try {
    jsonPub = JSON.parse(resPub.getContentText() || '{}');
  } catch (e) {
    throw new Error('Respons publikasi karusel tidak valid: ' + resPub.getContentText().substring(0, 80));
  }

  if (resPub.getResponseCode() !== 200 || !jsonPub.id) {
    var pubErr = jsonPub.error && jsonPub.error.message ? jsonPub.error.message : ('HTTP ' + resPub.getResponseCode());
    throw new Error('Gagal mempublikasikan karusel Meta: ' + pubErr);
  }

  return jsonPub.id;
}

// Helper penerbitan satu node thread ke Meta Threads Graph API
function publishThreadsNode(accountId, token, text, mediaObj, replyToId) {
  var cleanTok = cleanToken(token);
  var cleanAcc = cleanAccountId(accountId);

  if (!cleanTok) {
    throw new Error('Access Token Meta kosong atau tidak valid.');
  }

  // Kirim access_token di query string dan form payload agar gateway & endpoint Meta Threads mengenali token
  var createUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(cleanAcc) + '/threads?access_token=' + encodeURIComponent(cleanTok);
  var payload = { access_token: cleanTok };

  if (text && text.trim().length > 0) {
    payload.text = text.trim();
  }

  if (replyToId) {
    payload.reply_to_id = replyToId;
  }

  var isMedia = false;
  var isVideo = false;
  if (mediaObj && mediaObj.media_url) {
    isMedia = true;
    isVideo = (mediaObj.media_type === 'VIDEO');
    payload.media_type = isVideo ? 'VIDEO' : 'IMAGE';
    if (isVideo) {
      payload.video_url = validateAndFormatVideoUrl(mediaObj.media_url);
    } else {
      payload.image_url = mediaObj.media_url;
    }
  } else {
    payload.media_type = 'TEXT';
    if (!payload.text) {
      throw new Error('Teks thread tidak boleh kosong jika tanpa lampiran media.');
    }
  }

  var res1;
  try {
    res1 = UrlFetchApp.fetch(createUrl, {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true
    });
  } catch (netErr1) {
    throw new Error('Koneksi jaringan gagal saat membuat kontainer Threads: ' + netErr1.message);
  }

  var json1 = {};
  try {
    json1 = JSON.parse(res1.getContentText() || '{}');
  } catch (parseErr1) {
    throw new Error('Respons pembuatan kontainer Meta tidak valid: ' + res1.getContentText().substring(0, 80));
  }

  if (res1.getResponseCode() !== 200 || !json1.id) {
    var msg1 = json1.error && json1.error.message ? json1.error.message : ('HTTP ' + res1.getResponseCode());
    if (isVideo) {
      throw new Error('Gagal membuat kontainer video Meta: ' + msg1 + '. Pastikan video berdurasi maks 5 menit dan dapat diakses publik.');
    }
    throw new Error('Meta API Container Error: ' + msg1);
  }

  if (isMedia) {
    waitForContainerReady(json1.id, cleanTok, isVideo);
  } else {
    Utilities.sleep(1000);
  }

  var publishUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(cleanAcc) + '/threads_publish?access_token=' + encodeURIComponent(cleanTok);
  var res2;
  try {
    res2 = UrlFetchApp.fetch(publishUrl, {
      method: 'post',
      payload: { creation_id: json1.id, access_token: cleanTok },
      muteHttpExceptions: true
    });
  } catch (netErr2) {
    throw new Error('Koneksi jaringan gagal saat publikasi Threads: ' + netErr2.message);
  }

  var json2 = {};
  try {
    json2 = JSON.parse(res2.getContentText() || '{}');
  } catch (parseErr2) {
    throw new Error('Respons publikasi Meta tidak valid: ' + res2.getContentText().substring(0, 80));
  }

  if (res2.getResponseCode() !== 200 || !json2.id) {
    var msg2 = json2.error && json2.error.message ? json2.error.message : ('HTTP ' + res2.getResponseCode());
    throw new Error('Meta API Publish Error: ' + msg2);
  }

  return json2.id;
}

// Helper sanitasi URL affiliate tanpa awalan protokol https:// atau http://
function stripProtocol(url) {
  if (!url) return '';
  return String(url).replace(/^https?:\/\//i, '').trim();
}

// Helper resolusi multi-link affiliate (spesifik produk atau generic) dengan sub_id dinamis
function resolveAffiliateLinks(text, node, prodMap, subId, fallbackProdId) {
  if (!text) return '';
  var res = String(text);

  // 1. Tag spesifik per produk: [LINK_AFFILIATE:PROD_ID]
  res = res.replace(/\[LINK_AFFILIATE:([^\]]+)\]/g, function(match, pid) {
    var pKey = String(pid).trim();
    var p = prodMap[pKey];
    var url = p ? (p.shopee_affiliate_url || '') : '';
    if (!url) url = 'shope.ee/affiliate';
    if (url.indexOf('sub_id=') === -1 && subId) {
      var sep = url.indexOf('?') === -1 ? '?' : '&';
      url = url + sep + 'sub_id=' + subId;
    }
    return stripProtocol(url);
  });

  // 2. Tag generic: [LINK_AFFILIATE]
  if (res.indexOf('[LINK_AFFILIATE]') !== -1) {
    var pids = [];
    if (node && Array.isArray(node.product_ids) && node.product_ids.length > 0) {
      pids = node.product_ids;
    } else if (node && node.product_id) {
      pids = [node.product_id];
    } else if (fallbackProdId) {
      pids = [fallbackProdId];
    }

    var primaryId = pids.length > 0 ? pids[0] : '';
    var pObj = primaryId ? prodMap[primaryId] : null;
    var primaryUrl = pObj ? (pObj.shopee_affiliate_url || 'shope.ee/affiliate') : 'shope.ee/affiliate';
    if (primaryUrl.indexOf('sub_id=') === -1 && subId) {
      var sep2 = primaryUrl.indexOf('?') === -1 ? '?' : '&';
      primaryUrl = primaryUrl + sep2 + 'sub_id=' + subId;
    }
    res = res.replace(/\[LINK_AFFILIATE\]/g, stripProtocol(primaryUrl));
  }

  // 3. Bersihkan awalan protokol https:// atau http:// pada domain Shopee agar Meta Threads tidak memicu link preview card
  res = res.replace(/https?:\/\/(shope\.ee\/|s\.shopee\.co\.id\/|shopee\.co\.id\/)/gi, '$1');

  return res;
}

// Eksekutor rantai multi-thread (root post & sub-threads berurutan) dengan deteksi karusel & multi-produk
function executeThreadChain(accountId, token, nodes, mediaMap, prods, subId, initialReplyToId) {
  var prodMap = {};
  if (prods && Array.isArray(prods)) {
    for (var p = 0; p < prods.length; p++) {
      prodMap[prods[p].product_id] = prods[p];
    }
  }

  var lastThreadId = initialReplyToId || null;
  for (var n = 0; n < nodes.length; n++) {
    var node = nodes[n];
    var textContent = typeof node === 'string' ? node : (node.content || '');
    textContent = resolveAffiliateLinks(textContent, node, prodMap, subId, '');

    // Kumpulkan daftar media (single ID atau multi-media karusel)
    var mIds = [];
    if (node && Array.isArray(node.media_ids)) {
      mIds = node.media_ids;
    } else if (node && node.media_ids) {
      mIds = String(node.media_ids).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    } else if (node && node.media_id) {
      mIds = [String(node.media_id).trim()];
    }

    var mediaItems = [];
    for (var m = 0; m < mIds.length; m++) {
      var itemObj = mediaMap[String(mIds[m]).trim()];
      if (itemObj) mediaItems.push(itemObj);
    }

    if (mediaItems.length > 1) {
      lastThreadId = publishThreadsCarousel(accountId, token, textContent, mediaItems, lastThreadId);
    } else {
      var singleMedia = mediaItems.length === 1 ? mediaItems[0] : null;
      lastThreadId = publishThreadsNode(accountId, token, textContent, singleMedia, lastThreadId);
    }

    if (typeof node === 'object' && node !== null) {
      node.threads_post_id = lastThreadId;
      node.permalink = getThreadPermalink(lastThreadId, token);
      node.status = 'PUBLISHED';
    }

    if (n < nodes.length - 1) {
      Utilities.sleep(1500); // jeda antar sub-thread
    }
  }
  return lastThreadId;
}

// Memproses antrean posting terjadwal yang sudah memasuki waktu tayang
function checkAndPublishQueue() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var qSheet = ss.getSheetByName('Posting_Queue');
  if (!qSheet) return { processed: 0, message: 'Sheet Posting_Queue belum ada' };

  var qData = qSheet.getDataRange().getValues();
  if (qData.length <= 1) return { processed: 0, message: 'Antrean kosong' };

  var accounts = getSheetRecords('Accounts');
  var mediaRecords = getSheetRecords('Media');
  var tokenMap = {};
  for (var a = 0; a < accounts.length; a++) {
    tokenMap[cleanAccountId(accounts[a].account_id)] = cleanToken(accounts[a].access_token);
  }

  var mediaMap = {};
  for (var m = 0; m < mediaRecords.length; m++) {
    mediaMap[String(mediaRecords[m].media_id).trim()] = mediaRecords[m];
  }

  var now = Date.now();
  var processed = 0;
  var batchLimit = 5;

  for (var i = 1; i < qData.length && processed < batchLimit; i++) {
    var status = String(qData[i][6] || '').toLowerCase();
    if (status !== 'scheduled' && status !== 'pending' && status !== 'pending_scheduled') continue;

    // CEGAH POSTING TERBIT SEBELUM WAKTUNYA: Hanya eksekusi jika waktu sistem saat ini sudah melewati jam tayang
    var schedTime = new Date(qData[i][3]).getTime();
    if (isNaN(schedTime) || schedTime > now) continue;

    var postId = qData[i][1];
    var accountId = cleanAccountId(qData[i][2]);
    var token = tokenMap[accountId];

    if (!token && accounts.length === 1) {
      token = cleanToken(accounts[0].access_token);
      accountId = cleanAccountId(accounts[0].account_id);
    }

    if (!token) {
      qSheet.getRange(i + 1, 7).setValue('Failed');
      qSheet.getRange(i + 1, 8).setValue('Access token tidak ditemukan untuk account_id: ' + accountId);
      qSheet.getRange(i + 1, 9).setValue(new Date().toISOString());
      continue;
    }

    try {
      var rawThread = qData[i][4];
      var nodes = [];
      try {
        nodes = typeof rawThread === 'string' ? JSON.parse(rawThread) : rawThread;
      } catch (e) {
        nodes = [{ content: String(rawThread || ''), media_id: '' }];
      }
      if (!Array.isArray(nodes)) nodes = [nodes];

      var prods = getSheetRecords('Products');
      var subId = accountId + '_' + postId;

      // Cek apakah postingan berantai ini dalam mode drip waiting views
      var postRecords = getSheetRecords('Posts');
      var isDripWaiting = false;
      var postThreshold = 500;
      for (var prIdx = 0; prIdx < postRecords.length; prIdx++) {
        if (String(postRecords[prIdx].post_id) === String(postId)) {
          if (postRecords[prIdx].drip_status === 'WAITING_VIEWS') {
            isDripWaiting = true;
          }
          postThreshold = Number(postRecords[prIdx].views_threshold || 500);
          break;
        }
      }

      // Jika drip waiting views, terbitkan HANYA Utas 1 (Root Hook bebas link). Utas 2 menunggu views tercapai.
      var nodesToPublish = (isDripWaiting && nodes.length > 1) ? [nodes[0]] : nodes;
      var lastThreadId = executeThreadChain(accountId, token, nodesToPublish, mediaMap, prods, subId);
      var nowIso = new Date().toISOString();

      if (isDripWaiting && nodes.length > 1) {
        for (var qd = 1; qd < nodes.length; qd++) {
          if (typeof nodes[qd] === 'object' && nodes[qd] !== null) {
            nodes[qd].status = 'WAITING_VIEWS';
          }
        }
      }

      var qRootPermalink = (nodes[0] && nodes[0].permalink) ? nodes[0].permalink : '';

      qSheet.getRange(i + 1, 7).setValue('Posted');
      qSheet.getRange(i + 1, 8).setValue('');
      qSheet.getRange(i + 1, 9).setValue(nowIso);

      if (postId) {
        var postSheet = ss.getSheetByName('Posts');
        if (postSheet) {
          var postData = postSheet.getDataRange().getValues();
          for (var p = 1; p < postData.length; p++) {
            if (String(postData[p][0]) === String(postId)) {
              postSheet.getRange(p + 1, 3).setValue(JSON.stringify(nodes));
              postSheet.getRange(p + 1, 9).setValue('PUBLISHED');
              postSheet.getRange(p + 1, 10).setValue(lastThreadId || '');
              if (qRootPermalink) postSheet.getRange(p + 1, 16).setValue(qRootPermalink);
              break;
            }
          }
        }
      }

      // Jika Auto-Drip on Views, simpan Utas #2 ke Posting_Queue dengan status 'Waiting_Views_Threshold'
      if (isDripWaiting && nodes.length > 1 && lastThreadId) {
        var remainingNodesSched = nodes.slice(1);
        var qIdDrip = 'Q_' + Utilities.getUuid().substring(0, 8);
        qSheet.appendRow([
          qIdDrip,
          postId,
          accountId,
          nowIso,
          JSON.stringify(remainingNodesSched),
          JSON.stringify(extractMediaIds(remainingNodesSched)),
          'Waiting_Views_Threshold',
          '',
          '',
          lastThreadId,
          2
        ]);
      }
      processed++;
    } catch (err) {
      var errDetail = (err && err.message) ? err.message : String(err);
      qSheet.getRange(i + 1, 7).setValue('Failed');
      qSheet.getRange(i + 1, 8).setValue(errDetail);
      qSheet.getRange(i + 1, 9).setValue(new Date().toISOString());

      if (postId) {
        try {
          var pSheet = ss.getSheetByName('Posts');
          if (pSheet) {
            var pData = pSheet.getDataRange().getValues();
            for (var k = 1; k < pData.length; k++) {
              if (String(pData[k][0]) === String(postId)) {
                pSheet.getRange(k + 1, 9).setValue('FAILED');
                break;
              }
            }
          }
        } catch (pErr) {}
      }
    }
  }

  return { processed: processed, timestamp: new Date().toISOString() };
}

// Hapus Produk & seluruh media ulasan terkait dari database serta server Cloudinary
function deleteProduct(payload) {
  var prodId = payload.product_id ? String(payload.product_id).trim() : '';
  if (!prodId) throw new Error('Product ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var prodSheet = ss.getSheetByName('Products');
  if (prodSheet) {
    var pData = prodSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      if (String(pData[i][0]) === prodId) {
        prodSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  var mediaSheet = ss.getSheetByName('Media');
  if (mediaSheet) {
    var mData = mediaSheet.getDataRange().getValues();
    var settings = getDashboardData().settings || {};
    var cloudName = settings.cloudinary_cloud;
    var apiKey = settings.cloudinary_api_key;
    var apiSecret = settings.cloudinary_api_secret;
    var canDeleteCloudinary = cloudName && apiKey && apiSecret;

    for (var m = mData.length - 1; m >= 1; m--) {
      if (String(mData[m][1]) === prodId) {
        var mediaUrl = String(mData[m][3] || '').trim();
        var mediaType = mData[m][4] || 'IMAGE';
        if (canDeleteCloudinary && mediaUrl.indexOf('cloudinary.com') !== -1) {
          var publicId = extractCloudinaryPublicId(mediaUrl);
          if (publicId) {
            try {
              deleteCloudinaryAsset(cloudName, apiKey, apiSecret, publicId, mediaType);
            } catch (cErr) {}
          }
        }
        mediaSheet.deleteRow(m + 1);
      }
    }
  }

  return { success: true, product_id: prodId };
}

// Ekstrak Cloudinary public_id dari URL asset
function extractCloudinaryPublicId(url) {
  if (!url) return '';
  var clean = String(url).split('?')[0];
  var m = clean.match(/\/upload\/(?:[^\/]+\/)*?(?:v\d+\/)?(.+?)(\.[^.\/]+)?$/);
  if (m && m[1]) {
    var pid = m[1];
    pid = pid.replace(/^(?:[a-z]_[a-z0-9_-]+,?)+\//i, '');
    return pid;
  }
  var segments = clean.split('/');
  var last = segments[segments.length - 1];
  return last.replace(/\.[^/.]+$/, '');
}

// Hapus file fisik secara permanen dari server Cloudinary menggunakan Admin API
function deleteCloudinaryAsset(cloudName, apiKey, apiSecret, publicId, mediaType) {
  var cleanCloud = String(cloudName).trim();
  var cleanKey = String(apiKey).trim();
  var cleanSecret = String(apiSecret).trim();
  var cleanId = String(publicId).trim();
  var resourceType = (String(mediaType).toUpperCase() === 'VIDEO') ? 'video' : 'image';

  // 1. Panggil Cloudinary Admin API endpoint dengan HTTP Basic Authentication
  var adminUrl = 'https://api.cloudinary.com/v1_1/' + encodeURIComponent(cleanCloud) + '/resources/' + resourceType + '/upload?public_ids[]=' + encodeURIComponent(cleanId);
  var authHeader = 'Basic ' + Utilities.base64Encode(cleanKey + ':' + cleanSecret);

  var resp = UrlFetchApp.fetch(adminUrl, {
    method: 'delete',
    headers: { 'Authorization': authHeader },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var text = resp.getContentText() || '{}';
  var resJson = {};
  try {
    resJson = JSON.parse(text);
  } catch (e) {}

  // 2. Fallback ke Cloudinary Destroy API jika Admin API mengembalikan status non-200
  if (code !== 200 || !resJson.deleted) {
    var timestamp = Math.floor(Date.now() / 1000).toString();
    var toSign = 'public_id=' + cleanId + '&timestamp=' + timestamp + cleanSecret;
    var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, toSign, Utilities.Charset.UTF_8);
    var signature = '';
    for (var i = 0; i < rawHash.length; i++) {
      var byteVal = rawHash[i];
      if (byteVal < 0) byteVal += 256;
      var hex = byteVal.toString(16);
      if (hex.length === 1) hex = '0' + hex;
      signature += hex;
    }

    var destroyUrl = 'https://api.cloudinary.com/v1_1/' + encodeURIComponent(cleanCloud) + '/' + resourceType + '/destroy';
    var destroyResp = UrlFetchApp.fetch(destroyUrl, {
      method: 'post',
      payload: {
        public_id: cleanId,
        timestamp: timestamp,
        api_key: cleanKey,
        signature: signature
      },
      muteHttpExceptions: true
    });

    var destCode = destroyResp.getResponseCode();
    var destText = destroyResp.getContentText() || '{}';
    var destJson = {};
    try {
      destJson = JSON.parse(destText);
    } catch (e2) {}

    if (destCode !== 200 || (destJson.result !== 'ok' && destJson.result !== 'not found')) {
      var errMsg = (destJson.error && destJson.error.message) ? destJson.error.message : ('HTTP ' + destCode + ': ' + destText);
      throw new Error('Gagal menghapus file dari Cloudinary: ' + errMsg);
    }
  } else {
    var statusVal = resJson.deleted ? resJson.deleted[cleanId] : '';
    if (statusVal && statusVal !== 'deleted' && statusVal !== 'not_found') {
      throw new Error('Cloudinary Admin API gagal menghapus file: status ' + statusVal);
    }
  }

  return true;
}

// Hapus satu item media dari database & server Cloudinary
function deleteMedia(payload) {
  var mediaId = payload.media_id ? String(payload.media_id).trim() : '';
  if (!mediaId) throw new Error('Media ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mediaSheet = ss.getSheetByName('Media');
  if (!mediaSheet) throw new Error('Sheet Media tidak ditemukan');

  var mData = mediaSheet.getDataRange().getValues();
  var targetRow = -1;
  var targetMedia = null;

  for (var i = 1; i < mData.length; i++) {
    if (String(mData[i][0]) === mediaId) {
      targetRow = i + 1;
      targetMedia = {
        media_id: mData[i][0],
        product_id: mData[i][1],
        media_url: mData[i][3],
        media_type: mData[i][4]
      };
      break;
    }
  }

  if (targetRow === -1 || !targetMedia) {
    throw new Error('Media tidak ditemukan di database');
  }

  var mediaUrl = String(targetMedia.media_url || '').trim();
  var publicId = extractCloudinaryPublicId(mediaUrl);

  // Jika URL berasal dari Cloudinary, hapus fisik terlebih dahulu via Admin API
  if (mediaUrl.indexOf('cloudinary.com') !== -1 && publicId) {
    var settings = getDashboardData().settings || {};
    var cloudName = settings.cloudinary_cloud;
    var apiKey = settings.cloudinary_api_key;
    var apiSecret = settings.cloudinary_api_secret;

    if (!cloudName || !apiKey || !apiSecret) {
      throw new Error('Kredensial Cloudinary (Cloud Name, API Key, API Secret) belum lengkap di Settings.');
    }

    deleteCloudinaryAsset(cloudName, apiKey, apiSecret, publicId, targetMedia.media_type);
  }

  mediaSheet.deleteRow(targetRow);
  return { success: true, media_id: mediaId, public_id: publicId };
}

// Hapus seluruh paket sesi ulasan dari database & Cloudinary
function deleteMediaSession(payload) {
  var prodId = payload.product_id ? String(payload.product_id).trim() : '';
  var sessionId = payload.session_id ? String(payload.session_id).trim() : '';
  if (!sessionId) throw new Error('Session ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mediaSheet = ss.getSheetByName('Media');
  if (!mediaSheet) throw new Error('Sheet Media tidak ditemukan');

  var settings = getDashboardData().settings || {};
  var cloudName = settings.cloudinary_cloud;
  var apiKey = settings.cloudinary_api_key;
  var apiSecret = settings.cloudinary_api_secret;
  var canDeleteCloudinary = cloudName && apiKey && apiSecret;

  var mData = mediaSheet.getDataRange().getValues();
  var deletedCount = 0;

  for (var m = mData.length - 1; m >= 1; m--) {
    var rowProd = String(mData[m][1] || '').trim();
    var rowSess = String(mData[m][2] || '').trim();
    if (rowSess === sessionId && (!prodId || rowProd === prodId)) {
      var mediaUrl = String(mData[m][3] || '').trim();
      var mediaType = mData[m][4] || 'IMAGE';
      if (canDeleteCloudinary && mediaUrl.indexOf('cloudinary.com') !== -1) {
        var publicId = extractCloudinaryPublicId(mediaUrl);
        if (publicId) {
          try {
            deleteCloudinaryAsset(cloudName, apiKey, apiSecret, publicId, mediaType);
          } catch(e) {}
        }
      }
      mediaSheet.deleteRow(m + 1);
      deletedCount++;
    }
  }

  return { success: true, session_id: sessionId, deleted_count: deletedCount };
}

// Perbarui postingan di antrean (Posts & Posting_Queue)
function updateQueuePost(payload) {
  var postId = payload.post_id ? String(payload.post_id).trim() : '';
  if (!postId) throw new Error('Post ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var postSheet = ss.getSheetByName('Posts');
  if (!postSheet) throw new Error('Sheet Posts tidak ditemukan');

  var pData = postSheet.getDataRange().getValues();
  var targetRow = -1;
  var currentStatus = '';

  for (var i = 1; i < pData.length; i++) {
    if (String(pData[i][0]) === postId) {
      targetRow = i + 1;
      currentStatus = String(pData[i][8]);
      break;
    }
  }

  if (targetRow === -1) throw new Error('Postingan tidak ditemukan');
  if (currentStatus === 'PUBLISHED') throw new Error('Postingan yang sudah terbit tidak dapat diedit');

  var headers = pData[0];
  var accCol = headers.indexOf('account_id');
  var contentCol = headers.indexOf('thread_content_json');
  var threshCol = headers.indexOf('views_threshold');
  var statusCol = headers.indexOf('status');

  if (payload.account_id && accCol !== -1) postSheet.getRange(targetRow, accCol + 1).setValue(payload.account_id);
  if (payload.thread_content_json && contentCol !== -1) postSheet.getRange(targetRow, contentCol + 1).setValue(payload.thread_content_json);
  if (payload.views_threshold !== undefined && threshCol !== -1) postSheet.getRange(targetRow, threshCol + 1).setValue(Number(payload.views_threshold || 500));
  if (payload.status && statusCol !== -1) postSheet.getRange(targetRow, statusCol + 1).setValue(payload.status);

  var qSheet = ss.getSheetByName('Posting_Queue');
  if (qSheet) {
    var qData = qSheet.getDataRange().getValues();
    var qFound = false;
    for (var q = 1; q < qData.length; q++) {
      if (String(qData[q][1]) === postId) {
        if (payload.account_id) qSheet.getRange(q + 1, 3).setValue(payload.account_id);
        if (payload.scheduled_time) {
          var dt = new Date(payload.scheduled_time);
          if (!isNaN(dt.getTime())) {
            qSheet.getRange(q + 1, 4).setValue(dt.toISOString());
          }
        }
        if (payload.thread_content_json) qSheet.getRange(q + 1, 5).setValue(payload.thread_content_json);
        if (payload.status === 'SCHEDULED') qSheet.getRange(q + 1, 7).setValue('Scheduled');
        if (payload.status === 'DRAFT') qSheet.getRange(q + 1, 7).setValue('Draft');
        qFound = true;
        break;
      }
    }

    if (!qFound && payload.status === 'SCHEDULED') {
      var qId = 'Q_' + Utilities.getUuid().substring(0, 8);
      var schedIso = payload.scheduled_time ? new Date(payload.scheduled_time).toISOString() : new Date().toISOString();
      qSheet.appendRow([
        qId,
        postId,
        payload.account_id || pData[targetRow - 1][accCol],
        schedIso,
        payload.thread_content_json || pData[targetRow - 1][contentCol],
        pData[targetRow - 1][3] || '[]',
        'Scheduled',
        '',
        ''
      ]);
    }
  }

  return { success: true, post_id: postId };
}

// Hapus/batalkan postingan antrean dari Posts & Posting_Queue
function deleteQueuePost(payload) {
  var postId = payload.post_id ? String(payload.post_id).trim() : '';
  if (!postId) throw new Error('Post ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var postSheet = ss.getSheetByName('Posts');
  if (postSheet) {
    var pData = postSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      if (String(pData[i][0]) === postId) {
        postSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  var qSheet = ss.getSheetByName('Posting_Queue');
  if (qSheet) {
    var qData = qSheet.getDataRange().getValues();
    for (var q = qData.length - 1; q >= 1; q--) {
      if (String(qData[q][1]) === postId) {
        qSheet.deleteRow(q + 1);
      }
    }
  }

  return { success: true, post_id: postId };
}

// Putuskan/hapus akun Meta Threads dari sistem
function deleteThreadsAccount(payload) {
  var accId = cleanAccountId(payload.account_id);
  if (!accId) throw new Error('Account ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var accSheet = ss.getSheetByName('Accounts');
  if (!accSheet) throw new Error('Sheet Accounts tidak ditemukan');

  var data = accSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (cleanAccountId(data[i][0]) === accId) {
      accSheet.deleteRow(i + 1);
      return { success: true, account_id: accId };
    }
  }

  throw new Error('Akun tidak ditemukan di database');
}

// Perbarui kategori portofolio akun Threads
function updateAccountCategory(payload) {
  var accId = cleanAccountId(payload.account_id);
  if (!accId) throw new Error('Account ID wajib diisi');
  var category = String(payload.category || 'General').trim();

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  if (!sheet) throw new Error('Sheet Accounts tidak ditemukan');

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var catCol = headers.indexOf('category');
  if (catCol === -1) {
    catCol = headers.length;
    sheet.getRange(1, catCol + 1).setValue('category');
  }

  for (var i = 1; i < data.length; i++) {
    if (cleanAccountId(data[i][0]) === accId) {
      sheet.getRange(i + 1, catCol + 1).setValue(category);
      return { success: true, account_id: accId, category: category };
    }
  }

  throw new Error('Akun tidak ditemukan di database');
}

// Mengambil statistik pemakaian storage & kuota dari Cloudinary Usage Admin API
function getCloudinaryUsage(payload) {
  var settings = getDashboardData().settings || {};
  var cloudName = (payload && payload.cloudinary_cloud) ? payload.cloudinary_cloud : settings.cloudinary_cloud;
  var apiKey = (payload && payload.cloudinary_api_key) ? payload.cloudinary_api_key : settings.cloudinary_api_key;
  var apiSecret = (payload && payload.cloudinary_api_secret) ? payload.cloudinary_api_secret : settings.cloudinary_api_secret;

  cloudName = String(cloudName || '').trim();
  apiKey = String(apiKey || '').trim();
  apiSecret = String(apiSecret || '').trim();

  if (!cloudName || !apiKey || !apiSecret) {
    throw new Error('Kredensial Cloudinary (Cloud Name, API Key, API Secret) belum lengkap di Settings.');
  }

  var url = 'https://api.cloudinary.com/v1_1/' + encodeURIComponent(cloudName) + '/usage';
  var authHeader = 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret);

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': authHeader },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var rawText = resp.getContentText() || '{}';
  var resJson = {};
  try {
    resJson = JSON.parse(rawText);
  } catch (e) {
    throw new Error('Format respons Cloudinary tidak valid (HTTP ' + code + ')');
  }

  if (code !== 200) {
    var errMsg = (resJson && resJson.error && resJson.error.message) ? resJson.error.message : ('HTTP ' + code);
    throw new Error('Gagal mengambil data Cloudinary: ' + errMsg);
  }

  var plan = resJson.plan || 'Free';
  var storageObj = resJson.storage || {};
  var creditsObj = resJson.credits || {};
  var objectsObj = resJson.objects || {};
  var bandwidthObj = resJson.bandwidth || {};

  var usedBytes = Number(storageObj.usage || 0);
  var limitBytes = Number(storageObj.limit || 0);
  var percent = Number(storageObj.used_percent || 0);

  if (limitBytes === 0 && creditsObj.limit) {
    limitBytes = Number(creditsObj.limit) * 1024 * 1024 * 1024;
  }
  if (limitBytes === 0) {
    limitBytes = 25 * 1024 * 1024 * 1024;
  }

  if (percent === 0 && limitBytes > 0) {
    percent = Math.min(100, (usedBytes / limitBytes) * 100);
  }

  return {
    plan: plan,
    used_bytes: usedBytes,
    limit_bytes: limitBytes,
    used_percent: Number(percent.toFixed(1)),
    objects_count: Number(objectsObj.usage || 0),
    bandwidth_bytes: Number(bandwidthObj.usage || 0),
    credits_usage: Number(creditsObj.usage || 0),
    credits_limit: Number(creditsObj.limit || 0),
    credits_percent: Number(creditsObj.used_percent || 0),
    last_updated: resJson.last_updated || new Date().toISOString()
  };
}

// Analisis radar riwayat posting & pencarian rekomendasi slot kosong untuk akun terpilih
function getAccountScheduleRadar(payload) {
  var accIds = [];
  if (payload && Array.isArray(payload.account_ids)) {
    accIds = payload.account_ids;
  } else if (payload && payload.account_id) {
    accIds = [payload.account_id];
  }
  accIds = accIds.map(function(id) { return cleanAccountId(id); }).filter(Boolean);

  var posts = getSheetRecords('Posts');
  var queue = getSheetRecords('Posting_Queue');
  var now = Date.now();

  var lastPostedTime = 0;
  var lastPostedIso = '';
  var nextQueueTime = 0;
  var nextQueueIso = '';
  var allUpcomingTimes = [];

  for (var p = 0; p < posts.length; p++) {
    var row = posts[p];
    var acc = cleanAccountId(row.account_id);
    if (accIds.length === 0 || accIds.indexOf(acc) !== -1) {
      if (row.status === 'PUBLISHED') {
        var t = row.created_at ? new Date(row.created_at).getTime() : 0;
        if (t > lastPostedTime) {
          lastPostedTime = t;
          lastPostedIso = row.created_at;
        }
      }
    }
  }

  for (var q = 0; q < queue.length; q++) {
    var qRow = queue[q];
    var qAcc = cleanAccountId(qRow.account_id);
    if (accIds.length === 0 || accIds.indexOf(qAcc) !== -1) {
      var qStatus = String(qRow.status || '').toLowerCase();
      if (qStatus === 'posted') {
        var execT = qRow.executed_at ? new Date(qRow.executed_at).getTime() : 0;
        if (execT > lastPostedTime) {
          lastPostedTime = execT;
          lastPostedIso = qRow.executed_at;
        }
      } else if (qStatus === 'scheduled' || qStatus === 'pending') {
        var sTime = qRow.scheduled_time ? new Date(qRow.scheduled_time).getTime() : 0;
        if (sTime > now) {
          allUpcomingTimes.push(sTime);
          if (nextQueueTime === 0 || sTime < nextQueueTime) {
            nextQueueTime = sTime;
            nextQueueIso = qRow.scheduled_time;
          }
        }
      }
    }
  }

  allUpcomingTimes.sort(function(a, b) { return a - b; });

  // Algoritma Rekomendasi Slot Kosong (Rentang jam aktif audiens 07:00 - 21:00, jeda aman >= 60 menit)
  var candidate = new Date(now + 20 * 60 * 1000);
  if (candidate.getHours() < 7) {
    candidate.setHours(7, 15, 0, 0);
  } else if (candidate.getHours() >= 21) {
    candidate.setDate(candidate.getDate() + 1);
    candidate.setHours(7, 15, 0, 0);
  }

  var foundSlot = null;
  var minGapMs = 60 * 60 * 1000;
  for (var step = 0; step < 48; step++) {
    var candTime = candidate.getTime();
    var cHour = candidate.getHours();

    if (cHour >= 7 && cHour <= 21) {
      var isConflict = false;
      if (lastPostedTime && Math.abs(candTime - lastPostedTime) < minGapMs) {
        isConflict = true;
      }
      if (!isConflict) {
        for (var u = 0; u < allUpcomingTimes.length; u++) {
          if (Math.abs(candTime - allUpcomingTimes[u]) < minGapMs) {
            isConflict = true;
            break;
          }
        }
      }
      if (!isConflict && candTime > (now + 10 * 60 * 1000)) {
        foundSlot = new Date(candTime);
        break;
      }
    }

    candidate = new Date(candidate.getTime() + 45 * 60 * 1000);
    if (candidate.getHours() > 21) {
      candidate.setDate(candidate.getDate() + 1);
      candidate.setHours(7, 15, 0, 0);
    }
  }

  if (!foundSlot) {
    foundSlot = new Date(now + 90 * 60 * 1000);
  }

  return {
    last_posted_at: lastPostedIso ? serializeForClient(lastPostedIso) : '',
    next_queue_at: nextQueueIso ? serializeForClient(nextQueueIso) : '',
    all_scheduled_times: allUpcomingTimes.map(function(ts) { return new Date(ts).toISOString(); }),
    recommended_slot: foundSlot.toISOString()
  };
}

// Simpan urutan baru media dalam satu sesi ulasan ke Google Sheets
function reorderMediaSession(payload) {
  var sessionId = payload.session_id ? String(payload.session_id).trim() : '';
  var orderedIds = Array.isArray(payload.ordered_media_ids) ? payload.ordered_media_ids : [];
  if (!sessionId || orderedIds.length === 0) throw new Error('Session ID dan daftar media ID wajib diisi');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Media');
  if (!sheet) throw new Error('Sheet Media tidak ditemukan');

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true };

  var headers = data[0];
  var idCol = headers.indexOf('media_id');
  var sessCol = headers.indexOf('session_id');
  if (idCol === -1) idCol = 0;
  if (sessCol === -1) sessCol = 2;

  var sessionRows = [];
  var rowIndices = [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][sessCol]) === sessionId) {
      sessionRows.push({
        data: data[i],
        mediaId: String(data[i][idCol]).trim()
      });
      rowIndices.push(i + 1);
    }
  }

  if (sessionRows.length <= 1) return { success: true };

  sessionRows.sort(function(a, b) {
    var idxA = orderedIds.indexOf(a.mediaId);
    var idxB = orderedIds.indexOf(b.mediaId);
    if (idxA === -1) idxA = 999;
    if (idxB === -1) idxB = 999;
    return idxA - idxB;
  });

  for (var k = 0; k < rowIndices.length; k++) {
    sheet.getRange(rowIndices[k], 1, 1, sessionRows[k].data.length).setValues([sessionRows[k].data]);
  }

  return { success: true, reordered_count: sessionRows.length };
}
