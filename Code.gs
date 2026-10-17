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
    } else if (action === 'saveProduct') {
      result = saveProduct(payload);
    } else if (action === 'saveMediaRecord') {
      result = saveMediaRecord(payload);
    } else if (action === 'generateAiHook') {
      result = generateAiHook(payload);
    } else if (action === 'createPost') {
      result = createPost(payload);
    } else if (action === 'checkViewTriggersAndReply') {
      result = checkViewTriggersAndReply(payload);
    } else if (action === 'quickReplyComment') {
      result = quickReplyComment(payload);
    } else if (action === 'saveSettings') {
      result = saveSettings(payload);
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
    'Posts': ['post_id', 'account_id', 'thread_content_json', 'media_ids_json', 'product_id', 'affiliate_sub_id', 'views_threshold', 'current_views', 'status', 'threads_post_id', 'auto_reply_content', 'auto_reply_sent', 'created_at'],
    'Comments': ['comment_id', 'post_id', 'account_id', 'parent_comment_id', 'sender_username', 'content', 'intent_flag', 'status', 'created_at'],
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
  return { status: 'Database siap digunakan' };
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

  // ponytail: nested tree diproses di server level-1 flat join, upgrade rekursif bila kedalaman balasan > 2
  var rootComments = [];
  var commentMap = {};
  for (var c = 0; c < rawComments.length; c++) {
    rawComments[c].replies = [];
    commentMap[rawComments[c].comment_id] = rawComments[c];
  }
  for (var k = 0; k < rawComments.length; k++) {
    var cmt = rawComments[k];
    if (cmt.parent_comment_id && commentMap[cmt.parent_comment_id]) {
      commentMap[cmt.parent_comment_id].replies.push(cmt);
    } else {
      rootComments.push(cmt);
    }
  }

  return {
    accounts: accounts,
    products: products,
    media: media,
    posts: posts,
    comments: rootComments,
    settings: settings
  };
}

function syncThreadsAccount(payload) {
  var token = payload.access_token;
  if (!token) throw new Error('Access Token Meta wajib diisi');

  var url = 'https://graph.threads.net/v1.0/me?fields=id,username,threads_profile_picture_url&access_token=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var status = resp.getResponseCode();
  var resJson = JSON.parse(resp.getContentText() || '{}');

  if (status !== 200 || !resJson.id) {
    throw new Error('Gagal verifikasi token Threads: ' + (resJson.error ? resJson.error.message : 'Invalid Token'));
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(resJson.id)) {
      foundRow = i + 1;
      break;
    }
  }

  var nowIso = new Date().toISOString();
  var expiryIso = new Date(Date.now() + 60 * 24 * 3600 * 1000).toISOString(); // 60 hari
  var rowValues = [
    String(resJson.id),
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

  var row = [
    postId,
    payload.account_id,
    payload.thread_content_json,
    payload.media_ids_json || '[]',
    payload.product_id || '',
    subId,
    Number(payload.views_threshold || 500),
    0, // current views
    payload.status || 'SCHEDULED',
    'TH_' + Math.floor(100000 + Math.random() * 900000), // mock thread ID
    payload.auto_reply_content || '',
    false, // auto_reply_sent
    new Date().toISOString()
  ];
  sheet.appendRow(row);
  return { post_id: postId, sub_id: subId };
}

function checkViewTriggersAndReply(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Posts');
  var data = sheet.getDataRange().getValues();
  var triggeredCount = 0;

  for (var i = 1; i < data.length; i++) {
    var postId = data[i][0];
    var threshold = Number(data[i][6] || 0);
    var currentViews = Number(data[i][7] || 0);
    var autoReplySent = (data[i][11] === true || String(data[i][11]).toLowerCase() === 'true');

    // Simulasi pertambahan traffic tayangan
    currentViews += Math.floor(150 + Math.random() * 300);
    sheet.getRange(i + 1, 8).setValue(currentViews);

    if (currentViews >= threshold && !autoReplySent) {
      // Trigger reply terkirim
      sheet.getRange(i + 1, 12).setValue(true);
      triggeredCount++;

      // Simulasikan komentar terkirim pada sheet Comments
      var cmtSheet = ss.getSheetByName('Comments');
      var cmtId = 'CMT_' + Utilities.getUuid().substring(0, 8);
      cmtSheet.appendRow([
        cmtId,
        postId,
        data[i][1], // account_id
        '', // root
        'SystemAutoReply',
        data[i][10], // auto_reply_content
        false,
        'PUBLISHED',
        new Date().toISOString()
      ]);
    }
  }

  return { message: 'Pemeriksaan views selesai. ' + triggeredCount + ' auto-reply link berhasil dieksekusi.' };
}

function quickReplyComment(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cmtSheet = ss.getSheetByName('Comments');
  var postSheet = ss.getSheetByName('Posts');
  var prodSheet = ss.getSheetByName('Products');

  // Ambil data sub_id & raw Shopee URL
  var posts = getSheetRecords('Posts');
  var prods = getSheetRecords('Products');
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

  var finalContent = (payload.reply_content || '').replace('[LINK_AFFILIATE]', affiliateLink);
  var newId = 'CMT_' + Utilities.getUuid().substring(0, 8);

  cmtSheet.appendRow([
    newId,
    payload.post_id,
    payload.account_id,
    payload.comment_id, // parent comment ID (nested)
    'AdminAgent',
    finalContent,
    false,
    'PUBLISHED',
    new Date().toISOString()
  ]);

  return { comment_id: newId, content: finalContent };
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
