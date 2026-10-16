const SPREADSHEET_ID = '1Qbpyi39qV4bqyCu9bh4gr0Igy8IbXbKmW-QIJruGv6s';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ThreadAffiliate Engine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var payload = body.payload || {};
    var result;

    if (action === 'getDashboardData') {
      result = getDashboardData();
    } else if (action === 'syncThreadsAccount') {
      result = syncThreadsAccount(payload);
    } else if (action === 'saveProduct') {
      result = saveProduct(payload);
    } else if (action === 'deleteProduct') {
      result = deleteProduct(payload);
    } else if (action === 'uploadMediaToCloudinaryBackend') {
      result = uploadMediaToCloudinaryBackend(payload);
    } else if (action === 'generateHooksWithAI') {
      result = generateHooksWithAI(payload);
    } else if (action === 'savePostToQueue') {
      result = savePostToQueue(payload);
    } else if (action === 'publishQueueItem') {
      result = publishQueueItem(payload);
    } else if (action === 'deleteQueuePost') {
      result = deleteQueuePost(payload);
    } else if (action === 'registerMediaVault') {
      result = registerMediaVault(payload);
    } else if (action === 'deleteMediaItem') {
      result = deleteMediaItem(payload);
    } else if (action === 'checkPostViewsAndTriggerReply') {
      result = checkPostViewsAndTriggerReply();
    } else if (action === 'syncCommentsInbox') {
      result = syncCommentsInbox();
    } else if (action === 'sendCommentQuickReply') {
      result = sendCommentQuickReply(payload);
    } else if (action === 'saveAccount') {
      result = saveAccount(payload);
    } else if (action === 'refreshAccountToken') {
      result = refreshAccountToken(payload);
    } else if (action === 'deleteAccount') {
      result = deleteAccount(payload);
    } else if (action === 'saveAppSettings') {
      result = saveAppSettings(payload);
    } else {
      throw new Error('Unknown action: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// DATABASE INITIALIZATION & SERIALIZATION
// ==========================================

function setupDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var schemas = {
    'Accounts': ['id', 'account_name', 'username', 'threads_user_id', 'access_token', 'profile_pic_url', 'token_expires_at', 'status'],
    'Products': ['product_id', 'product_name', 'category', 'description', 'affiliate_url', 'created_at'],
    'Product_Media': ['media_id', 'product_id', 'review_group', 'media_type', 'cloudinary_url', 'uploaded_at'],
    'MediaVault': ['media_id', 'cloudinary_url', 'media_type', 'md5_hash', 'last_used_at', 'tags'],
    'PostsQueue': ['queue_id', 'account_id', 'root_text', 'media_urls', 'affiliate_url', 'sub_id', 'scheduled_at', 'published_threads_id', 'views_threshold', 'current_views', 'status', 'error_log'],
    'HookSwipeFile': ['hook_id', 'category', 'hook_text', 'performance_score'],
    'InboxCache': ['comment_id', 'post_id', 'account_id', 'parent_comment_id', 'from_user', 'text', 'intent_flag', 'replied_status', 'timestamp'],
    'AppSettings': ['key', 'value']
  };

  Object.keys(schemas).forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = schemas[sheetName];
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  });
}

function serializeForClient(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(function(item) { return serializeForClient(item); });
  }
  if (typeof value === 'object') {
    var res = {};
    Object.keys(value).forEach(function(k) {
      res[k] = serializeForClient(value[k]);
    });
    return res;
  }
  return value;
}

function getSheetDataRows(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) { return []; }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) { return []; }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
      }
      obj[headers[j]] = (val === undefined || val === '') ? null : val;
    }
    obj._rowNumber = i + 2;
    rows.push(obj);
  }
  return rows;
}

// ==========================================
// CORE DATA LOADERS & SETTINGS
// ==========================================

function getDashboardData() {
  var accounts = getSheetDataRows('Accounts');
  var products = getSheetDataRows('Products');
  var productMedia = getSheetDataRows('Product_Media');
  var rawMedia = getSheetDataRows('MediaVault');
  var rawQueue = getSheetDataRows('PostsQueue');
  var rawInbox = getSheetDataRows('InboxCache');
  var rawSettings = getSheetDataRows('AppSettings');

  var settings = {};
  var props = PropertiesService.getScriptProperties();
  ['gemini_api_key', 'cloudinary_cloud_name', 'cloudinary_upload_preset', 'cloudinary_api_key', 'cloudinary_api_secret', 'default_views_threshold'].forEach(function(k) {
    var pVal = props.getProperty(k);
    if (pVal) { settings[k] = pVal; }
  });

  rawSettings.forEach(function(row) {
    if (row.key && !settings[row.key]) { settings[row.key] = row.value; }
  });

  // Calculate Media Cooldown (14 Days)
  var now = new Date().getTime();
  var fourteenDaysMs = 14 * 24 * 60 * 60 * 1000;
  var media = rawMedia.map(function(m) {
    var status = 'FRESH';
    var daysLeft = 0;
    if (m.last_used_at) {
      var lastDate = new Date(m.last_used_at).getTime();
      var diff = now - lastDate;
      if (diff < fourteenDaysMs) {
        status = 'COOLDOWN';
        daysLeft = Math.ceil((fourteenDaysMs - diff) / (24 * 60 * 60 * 1000));
      }
    }
    return {
      media_id: m.media_id,
      cloudinary_url: m.cloudinary_url,
      media_type: m.media_type,
      md5_hash: m.md5_hash,
      last_used_at: m.last_used_at,
      tags: m.tags,
      status: status,
      days_left: daysLeft
    };
  });

  // Account mapping for queue
  var accMap = {};
  accounts.forEach(function(a) { accMap[a.id] = a.account_name; });

  var queue = rawQueue.map(function(q) {
    return {
      queue_id: q.queue_id,
      account_id: q.account_id,
      account_name: accMap[q.account_id] || q.account_id,
      root_text: q.root_text,
      media_urls: q.media_urls,
      affiliate_url: q.affiliate_url,
      sub_id: q.sub_id,
      scheduled_at: q.scheduled_at,
      published_threads_id: q.published_threads_id,
      views_threshold: Number(q.views_threshold) || 500,
      current_views: Number(q.current_views) || 0,
      status: q.status || 'DRAFT',
      error_log: q.error_log
    };
  });

  // Build hierarchical comments tree (Root -> Level 1 -> Level 2)
  var commentMap = {};
  var rootComments = [];
  rawInbox.forEach(function(c) {
    c.replies = [];
    commentMap[c.comment_id] = c;
  });

  rawInbox.forEach(function(c) {
    if (c.parent_comment_id && commentMap[c.parent_comment_id]) {
      commentMap[c.parent_comment_id].replies.push(c);
    } else {
      rootComments.push(c);
    }
  });

  return serializeForClient({
    accounts: accounts,
    media: media,
    queue: queue,
    inbox: rootComments,
    settings: settings
  });
}

function saveAppSettings(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AppSettings');
  if (!sheet) { return false; }

  var props = PropertiesService.getScriptProperties();
  var keys = ['gemini_api_key', 'cloudinary_cloud_name', 'cloudinary_upload_preset', 'default_views_threshold'];

  keys.forEach(function(k) {
    if (payload[k] !== undefined) {
      props.setProperty(k, String(payload[k]));
    }
  });

  // Sync to AppSettings tab
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  var rows = keys.map(function(k) {
    return [k, payload[k] !== undefined ? String(payload[k]) : ''];
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  return true;
}

// ==========================================
// GEMINI 1.5 FLASH AI HOOK STUDIO
// ==========================================

function generateHooksWithAI(payload) {
  var topic = payload.topic || '';
  if (!topic) { throw new Error('Topik atau nama produk wajib diisi.'); }

  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('gemini_api_key');
  if (!apiKey) {
    var sheetSettings = getSheetDataRows('AppSettings');
    var found = sheetSettings.find(function(s) { return s.key === 'gemini_api_key'; });
    apiKey = found ? found.value : '';
  }
  if (!apiKey) {
    throw new Error('Gemini API Key belum dikonfigurasi di Pengaturan.');
  }

  var promptText = 'Buatkan 5 variasi kalimat pembuka (hook) viral gaya Meta Threads berbahasa Indonesia yang emosional, bikin penasaran, dan memicu orang membaca sampai tuntas untuk produk/topik: "' + topic + '".\nFormat output: HANYA berikan daftar bernomor 1 sampai 5 tanpa teks pembuka atau penutup.';

  var endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
  var requestPayload = {
    contents: [{
      parts: [{ text: promptText }]
    }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestPayload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(endpoint, options);
  var json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error('Gemini API Error: ' + json.error.message);
  }

  var rawResponse = json.candidates[0].content.parts[0].text;
  var lines = rawResponse.split('\n');
  var hooks = [];
  for (var i = 0; i < lines.length; i++) {
    var cleanLine = lines[i].replace(/^\d+[\.\)]\s*/, '').trim();
    if (cleanLine.length > 0) {
      hooks.push(cleanLine);
    }
  }

  return { hooks: hooks.slice(0, 5) };
}

// ==========================================
// POST QUEUE & META THREADS PUBLISHING
// ==========================================

function savePostToQueue(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('PostsQueue');
  if (!sheet) { throw new Error('PostsQueue sheet not found'); }

  var queueId = 'PQ-' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
  var subId = payload.account_id + '_' + queueId;

  var mediaJson = payload.media_url ? JSON.stringify([payload.media_url]) : '[]';
  var status = payload.status || 'DRAFT';

  var row = [
    queueId,
    payload.account_id,
    payload.root_text,
    mediaJson,
    payload.affiliate_url || '',
    subId,
    new Date(),
    '',
    payload.views_threshold || 500,
    0,
    status,
    ''
  ];

  sheet.appendRow(row);

  // If directly scheduled or published
  if (status === 'SCHEDULED') {
    publishQueueItem({ queue_id: queueId });
  }

  return { queue_id: queueId, sub_id: subId };
}

function publishQueueItem(payload) {
  var queueId = payload.queue_id;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var queueSheet = ss.getSheetByName('PostsQueue');
  var queueRows = getSheetDataRows('PostsQueue');
  var target = queueRows.find(function(r) { return r.queue_id === queueId; });

  if (!target) { throw new Error('Post queue item not found: ' + queueId); }

  var accounts = getSheetDataRows('Accounts');
  var account = accounts.find(function(a) { return a.id === target.account_id; });
  if (!account || !account.access_token) {
    throw new Error('Akun Threads atau Access Token tidak valid.');
  }

  var threadsUserId = account.threads_user_id;
  var token = account.access_token;
  var rootText = target.root_text;
  var mediaList = [];
  try { mediaList = JSON.parse(target.media_urls || '[]'); } catch(e) {}

  var creationUrl = 'https://graph.threads.net/v1.0/' + threadsUserId + '/threads';
  var postParams = {
    access_token: token,
    text: rootText
  };

  if (mediaList.length > 0) {
    var mediaUrl = mediaList[0];
    postParams.media_type = 'IMAGE';
    postParams.image_url = mediaUrl;

    // Update Media Vault timestamp to start 14-day cooldown
    updateMediaLastUsed(mediaUrl);
  }

  // 1. Create Media Container
  var createRes = UrlFetchApp.fetch(creationUrl, {
    method: 'post',
    payload: postParams,
    muteHttpExceptions: true
  });
  var createJson = JSON.parse(createRes.getContentText());

  if (createJson.error) {
    queueSheet.getRange(target._rowNumber, 11, 1, 2).setValues([['FAILED', createJson.error.message]]);
    throw new Error('Gagal membuat container Threads: ' + createJson.error.message);
  }

  var containerId = createJson.id;

  // Wait 3 seconds for Meta media container to be ready
  Utilities.sleep(3000);

  // 2. Publish Container
  var publishUrl = 'https://graph.threads.net/v1.0/' + threadsUserId + '/threads_publish';
  var pubRes = UrlFetchApp.fetch(publishUrl, {
    method: 'post',
    payload: {
      access_token: token,
      creation_id: containerId
    },
    muteHttpExceptions: true
  });
  var pubJson = JSON.parse(pubRes.getContentText());

  if (pubJson.error) {
    queueSheet.getRange(target._rowNumber, 11, 1, 2).setValues([['FAILED', pubJson.error.message]]);
    throw new Error('Gagal menerbitkan thread: ' + pubJson.error.message);
  }

  var publishedId = pubJson.id;

  // Status becomes WAITING_VIEWS so the view worker monitors impressions
  queueSheet.getRange(target._rowNumber, 8, 1, 4).setValues([[publishedId, target.views_threshold, 0, 'WAITING_VIEWS']]);

  return { success: true, published_threads_id: publishedId };
}

function deleteQueuePost(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('PostsQueue');
  var rows = getSheetDataRows('PostsQueue');
  var target = rows.find(function(r) { return r.queue_id === payload.queue_id; });
  if (target) {
    sheet.deleteRow(target._rowNumber);
    return true;
  }
  return false;
}

// ==========================================
// MEDIA VAULT & COOLDOWN LOGIC
// ==========================================

function registerMediaVault(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('MediaVault');
  if (!sheet) { throw new Error('MediaVault sheet not found'); }

  var mediaId = 'MED-' + Date.now();
  var md5Hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, payload.cloudinary_url)).substring(0, 16);

  var row = [
    mediaId,
    payload.cloudinary_url,
    payload.media_type || 'IMAGE',
    md5Hash,
    '', // last_used_at empty initially (fresh)
    payload.tags || ''
  ];
  sheet.appendRow(row);
  return { media_id: mediaId, md5_hash: md5Hash };
}

function updateMediaLastUsed(mediaUrl) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('MediaVault');
  var rows = getSheetDataRows('MediaVault');
  var found = rows.find(function(r) { return r.cloudinary_url === mediaUrl; });
  if (found) {
    sheet.getRange(found._rowNumber, 5).setValue(new Date());
  }
}

function deleteMediaItem(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('MediaVault');
  var rows = getSheetDataRows('MediaVault');
  var target = rows.find(function(r) { return r.media_id === payload.media_id; });
  if (target) {
    sheet.deleteRow(target._rowNumber);
    return true;
  }
  return false;
}

// ==========================================
// TIME-DRIVEN VIEW POLLING & CONVERSION WORKER
// ==========================================

function checkPostViewsAndTriggerReply() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var queueSheet = ss.getSheetByName('PostsQueue');
  var posts = getSheetDataRows('PostsQueue');
  var accounts = getSheetDataRows('Accounts');
  var accMap = {};
  accounts.forEach(function(a) { accMap[a.id] = a; });

  var waitingPosts = posts.filter(function(p) {
    return p.status === 'WAITING_VIEWS' && p.published_threads_id;
  }).slice(0, 10); // Batch limit 10 to stay strictly under GAS 6-min quota

  var repliedCount = 0;
  var now = new Date().getTime();
  var fortyEightHoursMs = 48 * 60 * 60 * 1000;

  waitingPosts.forEach(function(post) {
    var account = accMap[post.account_id];
    if (!account || !account.access_token) { return; }

    var mediaId = post.published_threads_id;
    var token = account.access_token;
    var currentViews = 0;

    // Fetch Insights from Threads API
    try {
      var insightsUrl = 'https://graph.threads.net/v1.0/' + mediaId + '/insights?metric=views&access_token=' + token;
      var resp = UrlFetchApp.fetch(insightsUrl, { muteHttpExceptions: true });
      var json = JSON.parse(resp.getContentText());
      if (json.data && json.data.length > 0 && json.data[0].values) {
        currentViews = json.data[0].values[0].value || 0;
      }
    } catch (e) {
      currentViews = Number(post.current_views) || 0;
    }

    queueSheet.getRange(post._rowNumber, 10).setValue(currentViews);

    // Condition: Threshold reached OR post is older than 48 hours
    var postScheduledTime = post.scheduled_at ? new Date(post.scheduled_at).getTime() : now;
    var isExpired48h = (now - postScheduledTime) >= fortyEightHoursMs;

    if (currentViews >= (Number(post.views_threshold) || 500) || isExpired48h) {
      if (post.affiliate_url) {
        var replyText = 'Spill link produk yang dipakai di atas ya! Checkout langsung di toko resminya di sini: ' + post.affiliate_url + (post.affiliate_url.indexOf('?') > -1 ? '&' : '?') + 'sub_id=' + post.sub_id;

        // Post self-reply on thread container
        var creationUrl = 'https://graph.threads.net/v1.0/' + account.threads_user_id + '/threads';
        var replyRes = UrlFetchApp.fetch(creationUrl, {
          method: 'post',
          payload: {
            access_token: token,
            reply_to_id: mediaId,
            text: replyText
          },
          muteHttpExceptions: true
        });
        var replyJson = JSON.parse(replyRes.getContentText());

        if (replyJson.id) {
          // Publish self-reply
          UrlFetchApp.fetch('https://graph.threads.net/v1.0/' + account.threads_user_id + '/threads_publish', {
            method: 'post',
            payload: {
              access_token: token,
              creation_id: replyJson.id
            },
            muteHttpExceptions: true
          });

          queueSheet.getRange(post._rowNumber, 11).setValue('REPLIED');
          repliedCount++;
        }
      }
    }
  });

  return { checked_posts: waitingPosts.length, replied_count: repliedCount };
}

// ==========================================
// INBOX & COMMENT CONVERSATION ENGINE
// ==========================================

function syncCommentsInbox() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var inboxSheet = ss.getSheetByName('InboxCache');
  var posts = getSheetDataRows('PostsQueue');
  var accounts = getSheetDataRows('Accounts');
  var accMap = {};
  accounts.forEach(function(a) { accMap[a.id] = a; });

  var hotKeywords = ['spill', 'link', 'beli di mana', 'harga', 'beli', 'rekomen', 'beli dimana'];
  var syncedCount = 0;

  var activePosts = posts.filter(function(p) { return p.published_threads_id; }).slice(0, 5);

  activePosts.forEach(function(post) {
    var account = accMap[post.account_id];
    if (!account || !account.access_token) { return; }

    var url = 'https://graph.threads.net/v1.0/' + post.published_threads_id + '/replies?fields=id,text,timestamp,username,parent_id&access_token=' + account.access_token;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());

    if (json.data && Array.isArray(json.data)) {
      var existingInbox = getSheetDataRows('InboxCache');
      var existingIds = existingInbox.map(function(item) { return item.comment_id; });

      json.data.forEach(function(comment) {
        if (existingIds.indexOf(comment.id) === -1) {
          var textLower = (comment.text || '').toLowerCase();
          var isHot = hotKeywords.some(function(kw) { return textLower.indexOf(kw) > -1; });

          inboxSheet.appendRow([
            comment.id,
            post.published_threads_id,
            account.id,
            comment.parent_id || '',
            comment.username || 'user',
            comment.text || '',
            isHot ? 'HOT_LEAD' : 'NORMAL',
            'PENDING',
            comment.timestamp || Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd HH:mm:ss')
          ]);
          syncedCount++;
        }
      });
    }
  });

  return { synced_count: syncedCount };
}

function sendCommentQuickReply(payload) {
  var commentId = payload.comment_id;
  var replyText = payload.reply_text;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var inboxSheet = ss.getSheetByName('InboxCache');
  var inboxRows = getSheetDataRows('InboxCache');
  var target = inboxRows.find(function(i) { return i.comment_id === commentId; });

  if (!target) { throw new Error('Komentar tidak ditemukan di cache.'); }

  var accounts = getSheetDataRows('Accounts');
  var account = accounts.find(function(a) { return a.id === target.account_id; });
  if (!account || !account.access_token) { throw new Error('Akun tidak terhubung.'); }

  // Create reply container
  var creationUrl = 'https://graph.threads.net/v1.0/' + account.threads_user_id + '/threads';
  var createRes = UrlFetchApp.fetch(creationUrl, {
    method: 'post',
    payload: {
      access_token: account.access_token,
      reply_to_id: commentId,
      text: replyText
    },
    muteHttpExceptions: true
  });
  var createJson = JSON.parse(createRes.getContentText());

  if (createJson.error) {
    throw new Error('Gagal mengirim balasan: ' + createJson.error.message);
  }

  // Publish reply
  UrlFetchApp.fetch('https://graph.threads.net/v1.0/' + account.threads_user_id + '/threads_publish', {
    method: 'post',
    payload: {
      access_token: account.access_token,
      creation_id: createJson.id
    },
    muteHttpExceptions: true
  });

  inboxSheet.getRange(target._rowNumber, 8).setValue('REPLIED');
  return { success: true };
}

// ==========================================
// ACCOUNT MANAGEMENT & TOKEN HEALTH
// ==========================================

function saveAccount(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  if (!sheet) { throw new Error('Accounts sheet not found'); }

  var accId = 'ACC-' + Date.now();
  var expireDate = new Date();
  expireDate.setDate(expireDate.getDate() + 60); // 60 days standard Meta long-lived token

  var row = [
    accId,
    payload.account_name,
    payload.threads_user_id,
    payload.access_token,
    expireDate,
    'ACTIVE'
  ];
  sheet.appendRow(row);
  return { account_id: accId };
}

function refreshAccountToken(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  var rows = getSheetDataRows('Accounts');
  var target = rows.find(function(a) { return a.id === payload.account_id; });

  if (!target) { throw new Error('Akun tidak ditemukan'); }

  // Refresh via Meta Graph API
  var url = 'https://graph.threads.net/refresh_access_token?grant_type=th_refresh_token&access_token=' + target.access_token;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());

  if (json.access_token) {
    var newExpire = new Date();
    newExpire.setDate(newExpire.getDate() + 60);
    sheet.getRange(target._rowNumber, 4, 1, 2).setValues([[json.access_token, newExpire]]);
    return { success: true };
  } else {
    throw new Error('Gagal refresh token: ' + (json.error ? json.error.message : 'Unknown error'));
  }
}

function deleteAccount(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Accounts');
  var rows = getSheetDataRows('Accounts');
  var target = rows.find(function(a) { return a.id === payload.account_id; });
  if (target) {
    sheet.deleteRow(target._rowNumber);
    return true;
  }
  return false;
}
