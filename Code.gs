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
    } else if (action === 'syncThreadsAccount') {
      result = syncThreadsAccount(payload);
    } else if (action === 'pingThreadsAccount') {
      result = pingThreadsAccount(payload);
    } else if (action === 'saveProduct') {
      result = saveProduct(payload);
    } else if (action === 'saveMediaRecord') {
      result = saveMediaRecord(payload);
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
    'Products': ['product_id', 'product_name', 'shopee_raw_url', 'category', 'price', 'created_at'],
    'Media': ['media_id', 'product_id', 'session_id', 'media_url', 'media_type', 'file_hash', 'uploaded_at', 'cooldown_until'],
    'Posts': ['post_id', 'account_id', 'thread_content_json', 'media_ids_json', 'product_id', 'affiliate_sub_id', 'views_threshold', 'current_views', 'status', 'threads_post_id', 'auto_reply_content', 'auto_reply_sent', 'created_at', 'likes_count', 'replies_count', 'permalink'],
    'Comments': ['comment_id', 'post_id', 'account_id', 'parent_comment_id', 'sender_username', 'content', 'intent_flag', 'status', 'created_at'],
    'Posting_Queue': ['queue_id', 'post_id', 'account_id', 'scheduled_time', 'thread_content_json', 'media_ids_json', 'status', 'error_log', 'executed_at'],
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

  // Bangun struktur pohon komentar berjenjang (nested tree) rekursif
  var commentMap = {};
  for (var c = 0; c < rawComments.length; c++) {
    var item = rawComments[c];
    item.replies = [];
    commentMap[String(item.comment_id)] = item;
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
  var id = 'PROD_' + Utilities.getUuid().substring(0, 8);
  var row = [
    id,
    payload.product_name || 'Produk Baru',
    payload.shopee_raw_url || '',
    payload.category || 'General',
    payload.price || 0,
    new Date().toISOString()
  ];
  sheet.appendRow(row);
  return { product_id: id };
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

function generateAiHook(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var settings = getDashboardData().settings;
  var apiKey = settings.gemini_key;
  if (!apiKey) throw new Error('Gemini API Key belum dikonfigurasi di Settings');

  var promptText = 'Buatkan 3 hook kalimat pembuka viral singkat (1-2 kalimat) untuk postingan Meta Threads mempromosikan produk: "' +
    payload.topic + '" dengan gaya bahasa ' + (payload.tone || 'storytelling') +
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
  var publishedThreadId = '';

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

    var nodes = [];
    try {
      nodes = typeof payload.thread_content_json === 'string' ? JSON.parse(payload.thread_content_json) : payload.thread_content_json;
    } catch(e) {
      nodes = [{ content: String(payload.thread_content_json || ''), media_id: '' }];
    }

    publishedThreadId = executeThreadChain(payload.account_id, token, nodes, mediaMap);
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
    false,
    new Date().toISOString()
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

  // 5. Jika target views tercapai dan auto-reply belum dikirim, kirimkan balasan otomatis
  var threshold = Number(rowData[thresholdCol] || 0);
  var autoReplySent = (rowData[replySentCol] === true || String(rowData[replySentCol]).toLowerCase() === 'true');
  var replyTriggered = false;

  if (views >= threshold && !autoReplySent && rowData[autoReplyContentCol]) {
    try {
      var replyText = String(rowData[autoReplyContentCol]);
      publishThreadsNode(rowData[accCol], token, replyText, null, threadsPostId);
      postSheet.getRange(targetRow, replySentCol + 1).setValue(true);
      replyTriggered = true;
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

function quickReplyComment(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cmtSheet = ss.getSheetByName('Comments');

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
        affiliateLink = prods[r].shopee_raw_url || affiliateLink;
        break;
      }
    }
    affiliateLink += '?sub_id=' + (targetPost.affiliate_sub_id || 'aff_sub');
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

  cmtSheet.appendRow([
    liveReplyId,
    payload.post_id,
    payload.account_id,
    parentCommentId,
    username,
    finalContent,
    false,
    'PUBLISHED',
    new Date().toISOString()
  ]);

  return { comment_id: liveReplyId, content: finalContent };
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
  var queueResult = checkAndPublishQueue();
  var viewsResult = syncAllPublishedPosts();
  return {
    queue_processed: queueResult.processed,
    posts_synced: viewsResult.updated,
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

// Pasang Google Apps Script time-driven trigger (1 menit / 5 menit)
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

  return {
    status: 'Trigger background aktif otomatis setiap ' + interval + ' menit! Antrean & views dimonitor realtime.'
  };
}

// Helper polling kesiapan kontainer media sebelum dipublikasi
function waitForContainerReady(containerId, token) {
  var cleanTok = cleanToken(token);
  var statusUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(containerId) + '?fields=status,error_message&access_token=' + encodeURIComponent(cleanTok);
  for (var i = 0; i < 8; i++) {
    Utilities.sleep(2000);
    var res = UrlFetchApp.fetch(statusUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText() || '{}');
    var status = json.status || '';
    if (status === 'FINISHED') return true;
    if (status === 'ERROR') throw new Error('Gagal proses media Meta: ' + (json.error_message || 'Format tidak didukung'));
    if (status === 'EXPIRED') throw new Error('Kontainer media Meta kedaluwarsa');
  }
  return true;
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
  if (mediaObj && mediaObj.media_url) {
    isMedia = true;
    payload.media_type = mediaObj.media_type === 'VIDEO' ? 'VIDEO' : 'IMAGE';
    if (payload.media_type === 'VIDEO') {
      payload.video_url = mediaObj.media_url;
    } else {
      payload.image_url = mediaObj.media_url;
    }
  } else {
    payload.media_type = 'TEXT';
    if (!payload.text) {
      throw new Error('Teks thread tidak boleh kosong jika tanpa lampiran media.');
    }
  }

  var res1 = UrlFetchApp.fetch(createUrl, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  var json1 = JSON.parse(res1.getContentText() || '{}');
  if (res1.getResponseCode() !== 200 || !json1.id) {
    throw new Error('Meta API Container Error: ' + (json1.error ? json1.error.message : 'Unknown creation error'));
  }

  if (isMedia) {
    waitForContainerReady(json1.id, cleanTok);
  } else {
    Utilities.sleep(1000);
  }

  var publishUrl = 'https://graph.threads.net/v1.0/' + encodeURIComponent(cleanAcc) + '/threads_publish?access_token=' + encodeURIComponent(cleanTok);
  var res2 = UrlFetchApp.fetch(publishUrl, {
    method: 'post',
    payload: { creation_id: json1.id, access_token: cleanTok },
    muteHttpExceptions: true
  });
  var json2 = JSON.parse(res2.getContentText() || '{}');
  if (res2.getResponseCode() !== 200 || !json2.id) {
    throw new Error('Meta API Publish Error: ' + (json2.error ? json2.error.message : 'Unknown publish error'));
  }

  return json2.id;
}

// Eksekutor rantai multi-thread (root post & sub-threads berurutan)
function executeThreadChain(accountId, token, nodes, mediaMap) {
  var lastThreadId = null;
  for (var n = 0; n < nodes.length; n++) {
    var node = nodes[n];
    var textContent = typeof node === 'string' ? node : (node.content || '');
    var mediaObj = (node.media_id && mediaMap[String(node.media_id).trim()]) ? mediaMap[String(node.media_id).trim()] : null;
    lastThreadId = publishThreadsNode(accountId, token, textContent, mediaObj, lastThreadId);
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
      qSheet.getRange(i + 1, 7).setValue('Failed');
      qSheet.getRange(i + 1, 8).setValue(err.message);
      qSheet.getRange(i + 1, 9).setValue(new Date().toISOString());
    }
  }

  return { processed: processed, timestamp: new Date().toISOString() };
}
