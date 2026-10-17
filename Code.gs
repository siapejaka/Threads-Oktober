const SPREADSHEET_ID = '1Qbpyi39qV4bqyCu9bh4gr0Igy8IbXbKmW-QIJruGv6s';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ThreadAffiliate Engine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var contents = e.postData && e.postData.contents ? e.postData.contents : '{}';
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
    } else if (action === 'saveMediaRecord') {
      result = saveMediaRecord(payload);
    } else if (action === 'saveMediaBatch') {
      result = saveMediaBatch(payload);
    } else if (action === 'generateAiHook') {
      result = generateAiHook(payload);
    } else if (action === 'createPost') {
      result = createPost(payload);
    } else if (action === 'publishPostNow') {
      result = publishPostNow(payload);
    } else if (action === 'checkViewTriggersAndReply') {
      result = checkViewTriggersAndReply(payload);
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
    } else if (action === 'checkAndExecuteQueueAndViews') {
      result = checkAndExecuteQueueAndViews();
    } else if (action === 'setupTimeDrivenTrigger') {
      result = setupTimeDrivenTrigger(payload);
    } else if (action === 'saveHookTemplate') {
      result = saveHookTemplate(payload);
    } else if (action === 'deleteHookTemplate') {
      result = deleteHookTemplate(payload);
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
    'Accounts': ['account_id', 'username', 'profile_pic_url', 'access_token', 'token_expires_at', 'status', 'updated_at'],
    'Products': ['product_id', 'product_name', 'category', 'price', 'product_image_url', 'shopee_raw_url', 'shopee_affiliate_url', 'selling_points', 'created_at'],
    'Media': ['media_id', 'product_id', 'session_id', 'media_url', 'media_type', 'file_hash', 'uploaded_at', 'cooldown_until'],
    'Posts': ['post_id', 'account_id', 'thread_content_json', 'media_ids_json', 'product_id', 'affiliate_sub_id', 'views_threshold', 'current_views', 'status', 'threads_post_id', 'auto_reply_content', 'auto_reply_sent', 'created_at', 'likes_count', 'replies_count', 'permalink', 'drip_status'],
    'Comments': ['comment_id', 'post_id', 'account_id', 'parent_comment_id', 'sender_username', 'content', 'intent_flag', 'status', 'created_at'],
    'Thread_Interactions': ['interaction_id', 'post_id', 'threads_post_id', 'account_id', 'comment_id', 'parent_comment_id', 'sender_username', 'author_type', 'content', 'intent_flag', 'reply_status', 'created_at'],
    'Posting_Queue': ['queue_id', 'post_id', 'account_id', 'scheduled_time', 'thread_content_json', 'media_ids_json', 'status', 'error_log', 'executed_at', 'reply_to_id', 'drip_step'],
    'Hook_Bank': ['hook_id', 'category', 'hook_text', 'formula', 'created_at'],
    'Settings': ['key', 'value']
  };

  for (var tab in schema) {
    var sheet = ss.getSheetByName(tab);
    if (!sheet) {
      sheet = ss.insertSheet(tab);
    }
    var headers = schema[tab];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
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
  if (val instanceof Date) return val.toISOString();
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
  var interactions = getSheetRecords('Thread_Interactions');

  // Jika tabel Thread_Interactions ada data, utamakan untuk inbox
  var dataSource = (interactions.length > 0) ? interactions : rawComments;

  // Bangun struktur pohon komentar berjenjang (nested tree) rekursif
  var commentMap = {};
  for (var c = 0; c < dataSource.length; c++) {
    var item = dataSource[c];
    item.replies = [];
    commentMap[String(item.comment_id || item.interaction_id)] = item;
  }

  var rootComments = [];
  for (var k = 0; k < rawComments.length; k++) {
    var cmtItem = rawComments[k];
    var parentId = cmtItem.parent_comment_id ? String(cmtItem.parent_comment_id).trim() : '';
    if (parentId && commentMap[parentId]) {
      commentMap[parentId].replies.push(cmtItem);
    } else {
      rootComments.push(cmtItem);
    }
  }

  return {
    accounts: accounts,
    products: products,
    media: media,
    posts: posts,
    comments: rootComments,
    interactions: interactions,
    hooks: hooks,
    queue: queue,
    settings: settings
  };
}

function syncThreadsAccount(payload) {
  var token = cleanToken(payload.access_token);
  if (!token) throw new Error('Access Token Meta wajib diisi');

  // Request otentikasi tunggal sesuai spek Meta Threads Graph API
  var url = 'https://graph.threads.net/v1.0/me?fields=id,username,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });
  var status = resp.getResponseCode();
  var resJson = JSON.parse(resp.getContentText() || '{}');

  if (status !== 200 || !resJson.id) {
    var fallbackResp = UrlFetchApp.fetch('https://graph.threads.net/v1.0/me?fields=id,username,threads_profile_picture_url', {
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
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (cleanAccountId(data[i][0]) === cleanAccountId(rawId)) {
      foundRow = i + 1;
      break;
    }
  }

  var nowIso = new Date().toISOString();
  var expiryIso = new Date(Date.now() + 60 * 24 * 3600 * 1000).toISOString(); // 60 hari
  var rowValues = [
    "'" + rawId,
    resJson.username || '',
    resJson.threads_profile_picture_url || '',
    token,
    expiryIso,
    'ACTIVE',
    nowIso
  ];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return {
    account_id: String(resJson.id),
    username: resJson.username || '',
    profile_pic_url: resJson.threads_profile_picture_url || ''
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
  var url = 'https://graph.threads.net/v1.0/me?fields=id,username,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });
  var status = resp.getResponseCode();
  var resJson = JSON.parse(resp.getContentText() || '{}');

  if (status !== 200 || !resJson.id) {
    var fallbackUrl = 'https://graph.threads.net/me?fields=id,username,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
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

  if (targetRow > 0) {
    sheet.getRange(targetRow, tokenCol + 1).setValue(token);
    sheet.getRange(targetRow, statusCol + 1).setValue('ACTIVE');
    sheet.getRange(targetRow, updatedCol + 1).setValue(new Date().toISOString());
  }

  return {
    account_id: String(resJson.id),
    username: resJson.username || '',
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
  var defaultHeaders = ['product_id', 'product_name', 'category', 'price', 'product_image_url', 'shopee_raw_url', 'shopee_affiliate_url', 'selling_points', 'created_at'];

  // Pastikan skema header baru diperbarui jika belum lengkap
  if (headers.indexOf('selling_points') === -1 || headers.indexOf('product_image_url') === -1 || headers.indexOf('shopee_affiliate_url') === -1) {
    headers = defaultHeaders;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var idCol = headers.indexOf('product_id');
  if (idCol === -1) idCol = 0;

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

  if (!productId) {
    productId = 'PROD_' + Utilities.getUuid().substring(0, 8);
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
    created_at: payload.created_at || new Date().toISOString()
  };

  var rowValues = [];
  for (var h = 0; h < headers.length; h++) {
    var k = headers[h];
    rowValues.push(rowObj[k] !== undefined ? rowObj[k] : '');
  }

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { product_id: productId };
}

function saveMediaRecord(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Media');
  var mediaId = 'MED_' + Utilities.getUuid().substring(0, 8);
  var now = new Date();
  var cooldown = new Date(now.getTime() + 14 * 24 * 3600 * 1000); // Cooldown 14 hari

  var row = [
    mediaId,
    payload.product_id || '',
    payload.session_id || ('SESI_' + Date.now()),
    payload.media_url,
    payload.media_type || 'IMAGE',
    payload.file_hash || '',
    now.toISOString(),
    cooldown.toISOString()
  ];
  sheet.appendRow(row);
  return { media_id: mediaId, cooldown_until: cooldown.toISOString() };
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

function generateAiHook(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var settings = getDashboardData().settings;
  var apiKey = settings.gemini_key;
  if (!apiKey) throw new Error('Gemini API Key belum dikonfigurasi di Settings');

  var promptText = 'Buatkan 3 hook kalimat pembuka viral singkat (1-2 kalimat) untuk postingan Meta Threads mempromosikan produk: "' +
    payload.topic + '"' +
    (payload.selling_points ? ' dengan keunggulan/selling points: "' + payload.selling_points + '"' : '') +
    ' dengan gaya bahasa ' + (payload.tone || 'storytelling') +
    '. Format balasan HANYA baris teks biasa per nomor tanpa intro atau outro.';

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
  var requestPayload = {
    contents: [{ parts: [{ text: promptText }] }]
  };

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestPayload),
    muteHttpExceptions: true
  });

  var resJson = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !resJson.candidates) {
    throw new Error('AI Error: ' + (resJson.error ? resJson.error.message : 'Gagal menghasilkan hook'));
  }

  var rawOutput = resJson.candidates[0].content.parts[0].text || '';
  var lines = rawOutput.split('\n')
    .map(function(s) { return s.replace(/^[0-9\.\-\* ]+/, '').trim(); })
    .filter(function(s) { return s.length > 5; });

  return { hooks: lines };
}

function createPost(payload) {
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
  var dripStatus = (isMultiNode && publishMode === 'DRIP_ON_VIEWS') ? 'WAITING_VIEWS' : 'COMPLETED';
  var isReplySent = !isMultiNode;

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

    if (publishMode === 'DRIP_ON_VIEWS' && nodes.length > 1) {
      // Auto-Drip Mode: Terbitkan HANYA Utas 1 (Root Hook) sekarang tanpa link eksternal
      publishedThreadId = executeThreadChain(payload.account_id, token, [nodes[0]], mediaMap);
      dripStatus = 'WAITING_VIEWS';
      isReplySent = false;
    } else {
      // Publish All Mode atau Utas Tunggal: Terbitkan seluruh rantai sekaligus
      publishedThreadId = executeThreadChain(payload.account_id, token, nodes, mediaMap);
      dripStatus = 'COMPLETED';
      isReplySent = true;
    }
    status = 'PUBLISHED';
  }

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
    '',
    dripStatus
  ];
  sheet.appendRow(row);

  if (status === 'SCHEDULED') {
    var qSheet = ss.getSheetByName('Posting_Queue');
    if (qSheet) {
      var qId = 'Q_' + Utilities.getUuid().substring(0, 8);
      var schedIso = payload.scheduled_time ? new Date(payload.scheduled_time).toISOString() : new Date().toISOString();
      qSheet.appendRow([
        qId,
        postId,
        payload.account_id,
        schedIso,
        payload.thread_content_json,
        payload.media_ids_json || '[]',
        'Scheduled',
        '',
        ''
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
        media_ids_json: postData[i][3]
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

  var publishedThreadId = executeThreadChain(targetPost.account_id, token, nodes, mediaMap);

  postSheet.getRange(targetRow, 9).setValue('PUBLISHED');
  postSheet.getRange(targetRow, 10).setValue(publishedThreadId);

  var qSheet = ss.getSheetByName('Posting_Queue');
  if (qSheet) {
    var qData = qSheet.getDataRange().getValues();
    for (var q = 1; q < qData.length; q++) {
      if (String(qData[q][1]) === String(postId)) {
        qSheet.getRange(q + 1, 7).setValue('Posted');
        qSheet.getRange(q + 1, 8).setValue('');
        qSheet.getRange(q + 1, 9).setValue(new Date().toISOString());
      }
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

  // 3. Sinkronkan balasan/komentar dari Threads ke tab Comments
  var commentsSynced = 0;
  var repliesUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(token);
  try {
    var repResp = UrlFetchApp.fetch(repliesUrl, { method: 'get', muteHttpExceptions: true });
    if (repResp.getResponseCode() === 200) {
      var repJson = JSON.parse(repResp.getContentText() || '{}');
      if (repJson.data && Array.isArray(repJson.data)) {
        var cmtSheet = ss.getSheetByName('Comments');
        var existingCmts = getSheetRecords('Comments');
        var cmtIdSet = {};
        for (var c = 0; c < existingCmts.length; c++) {
          cmtIdSet[String(existingCmts[c].comment_id)] = true;
        }
        for (var r = 0; r < repJson.data.length; r++) {
          var repItem = repJson.data[r];
          if (!cmtIdSet[repItem.id]) {
            var intentFlag = /spill|link|beli|toko|harga|order|checkout|beli di mana|buy/i.test(repItem.text || '');
            cmtSheet.appendRow([
              repItem.id,
              postId,
              rowData[accCol],
              '',
              repItem.username || 'user',
              repItem.text || '',
              intentFlag,
              'PUBLISHED',
              repItem.timestamp || new Date().toISOString()
            ]);
            commentsSynced++;
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

  // 5. Jika target views tercapai dan auto-reply belum dikirim, kirimkan balasan otomatis (Sequential Auto-Drip Chain)
  var threshold = Number(rowData[thresholdCol] || 0);
  var autoReplySent = (rowData[replySentCol] === true || String(rowData[replySentCol]).toLowerCase() === 'true');
  var replyTriggered = false;

  if (views >= threshold && !autoReplySent) {
    try {
      // Ambil link Shopee Affiliate terdaftar untuk disisipkan ke [LINK_AFFILIATE]
      var affiliateLink = 'https://shope.ee/affiliate';
      var prodId = rowData[headers.indexOf('product_id')];
      if (prodId) {
        var prods = getSheetRecords('Products');
        for (var pr = 0; pr < prods.length; pr++) {
          if (prods[pr].product_id === prodId) {
            affiliateLink = prods[pr].shopee_affiliate_url || prods[pr].shopee_raw_url || affiliateLink;
            break;
          }
        }
        var sep = affiliateLink.indexOf('?') === -1 ? '?' : '&';
        affiliateLink += sep + 'sub_id=' + (rowData[headers.indexOf('affiliate_sub_id')] || 'sub');
      }

      var nodes = [];
      try {
        var rawNodes = rowData[headers.indexOf('thread_content_json')];
        nodes = typeof rawNodes === 'string' ? JSON.parse(rawNodes) : rawNodes;
      } catch(e) {}

      // Jika postingan memiliki utas berantai (Utas 2, Utas 3, dst.), kirimkan secara berurutan
      if (Array.isArray(nodes) && nodes.length > 1) {
        var mediaRecords = getSheetRecords('Media');
        var mediaMap = {};
        for (var m = 0; m < mediaRecords.length; m++) {
          mediaMap[String(mediaRecords[m].media_id).trim()] = mediaRecords[m];
        }

        var chainParentId = threadsPostId;
        for (var step = 1; step < nodes.length; step++) {
          var subNode = nodes[step];
          var subText = typeof subNode === 'string' ? subNode : (subNode.content || '');
          subText = subText.replace(/\[LINK_AFFILIATE\]/g, affiliateLink);

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
          Utilities.sleep(1500);
        }

        postSheet.getRange(targetRow, replySentCol + 1).setValue(true);
        var dripCol = headers.indexOf('drip_status');
        if (dripCol !== -1) postSheet.getRange(targetRow, dripCol + 1).setValue('COMPLETED');
        replyTriggered = true;
      } else if (rowData[autoReplyContentCol]) {
        // Fallback untuk postingan dengan template teks biasa
        var replyText = String(rowData[autoReplyContentCol]).replace(/\[LINK_AFFILIATE\]/g, affiliateLink);
        publishThreadsNode(rowData[accCol], token, replyText, null, threadsPostId);
        postSheet.getRange(targetRow, replySentCol + 1).setValue(true);
        replyTriggered = true;
      }
    } catch (e) {}
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

function checkViewTriggersAndReply(payload) {
  return syncAllPublishedPosts();
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

  // Peta username akun sendiri untuk klasifikasi author_type
  var accountUserMap = {};
  var tokenMap = {};
  for (var a = 0; a < accounts.length; a++) {
    var accId = cleanAccountId(accounts[a].account_id);
    accountUserMap[accId] = String(accounts[a].username || '').toLowerCase().trim();
    tokenMap[accId] = cleanToken(accounts[a].access_token);
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
    var key = String(existingInteractions[e].comment_id || existingInteractions[e].interaction_id);
    existingIds[key] = true;
  }

  var syncedCount = 0;
  var maxPostsToScan = Math.min(publishedPosts.length, 6);

  for (var i = 0; i < maxPostsToScan; i++) {
    var post = publishedPosts[i];
    var accId = cleanAccountId(post.account_id);
    var token = tokenMap[accId];
    if (!token && accounts.length > 0) token = cleanToken(accounts[0].access_token);
    if (!token) continue;

    var threadsPostId = String(post.threads_post_id).trim();
    var ownUsername = accountUserMap[accId] || '';

    // 1. Coba endpoint /conversation
    var convUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/conversation?fields=id,text,username,timestamp,has_replies&access_token=' + encodeURIComponent(token);
    var rawList = [];

    try {
      var resp = UrlFetchApp.fetch(convUrl, { method: 'get', muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        var json = JSON.parse(resp.getContentText() || '{}');
        if (json.data && Array.isArray(json.data)) {
          rawList = json.data;
        }
      }
    } catch (errConv) {}

    // 2. Fallback ke endpoint /replies jika /conversation kosong
    if (rawList.length === 0) {
      try {
        var repUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(threadsPostId) + '/replies?fields=id,text,username,timestamp&access_token=' + encodeURIComponent(token);
        var resp2 = UrlFetchApp.fetch(repUrl, { method: 'get', muteHttpExceptions: true });
        if (resp2.getResponseCode() === 200) {
          var json2 = JSON.parse(resp2.getContentText() || '{}');
          if (json2.data && Array.isArray(json2.data)) {
            rawList = json2.data;
          }
        }
      } catch (errRep) {}
    }

    // 3. Simpan interaksi ke database Google Sheets
    for (var k = 0; k < rawList.length; k++) {
      var item = rawList[k];
      var commentId = String(item.id || '').trim();
      if (!commentId || existingIds[commentId]) continue;

      var sender = String(item.username || 'user').trim();
      var isOwnAccount = (ownUsername && sender.toLowerCase() === ownUsername);
      var authorType = isOwnAccount ? 'OWNER' : 'AUDIENCE';
      var text = item.text || '';
      var intentFlag = /spill|link|beli|toko|harga|order|checkout|beli di mana|buy|mau/i.test(text);
      var replyStatus = isOwnAccount ? 'OWNER_POST' : (item.has_replies ? 'REPLIED' : 'UNREPLIED');
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
        authorType,
        text,
        intentFlag,
        replyStatus,
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
      }

      existingIds[commentId] = true;
      syncedCount++;
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

  var affiliateLink = 'https://shope.ee/affiliate';
  if (targetPost) {
    for (var r = 0; r < prods.length; r++) {
      if (prods[r].product_id === targetPost.product_id) {
        affiliateLink = prods[r].shopee_affiliate_url || prods[r].shopee_raw_url || affiliateLink;
        break;
      }
    }
    var sep = affiliateLink.indexOf('?') === -1 ? '?' : '&';
    affiliateLink += sep + 'sub_id=' + (targetPost.affiliate_sub_id || 'aff_sub');
  }

  var finalContent = (payload.reply_content || '').replace(/\[LINK_AFFILIATE\]/g, affiliateLink);
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

  return {
    queue_processed: queueProcessed,
    posts_synced: postsSynced,
    timestamp: new Date().toISOString()
  };
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

// Pasang Google Apps Script time-driven trigger (Default: 15 menit)
function setupTimeDrivenTrigger(payload) {
  var interval = (payload && payload.interval) ? Number(payload.interval) : 15;
  if ([1, 5, 10, 15].indexOf(interval) === -1) interval = 15;

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

  return {
    status: 'Trigger background checkAndExecuteQueueAndViews berhasil didaftarkan aktif setiap ' + interval + ' menit!'
  };
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

// Eksekutor rantai multi-thread (root post & sub-threads berurutan) dengan deteksi karusel
function executeThreadChain(accountId, token, nodes, mediaMap) {
  var lastThreadId = null;
  for (var n = 0; n < nodes.length; n++) {
    var node = nodes[n];
    var textContent = typeof node === 'string' ? node : (node.content || '');

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
    if (status !== 'scheduled' && status !== 'pending') continue;

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

      var lastThreadId = executeThreadChain(accountId, token, nodes, mediaMap);
      var nowIso = new Date().toISOString();

      qSheet.getRange(i + 1, 7).setValue('Posted');
      qSheet.getRange(i + 1, 8).setValue('');
      qSheet.getRange(i + 1, 9).setValue(nowIso);

      if (postId) {
        var postSheet = ss.getSheetByName('Posts');
        if (postSheet) {
          var postData = postSheet.getDataRange().getValues();
          for (var p = 1; p < postData.length; p++) {
            if (String(postData[p][0]) === String(postId)) {
              postSheet.getRange(p + 1, 9).setValue('PUBLISHED');
              postSheet.getRange(p + 1, 10).setValue(lastThreadId || '');
              break;
            }
          }
        }
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
