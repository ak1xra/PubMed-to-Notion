/**************************************
 * PubMed → Notion 論文収集 完全体 GAS
 * - クエリ: スプレッドシート「Queries」シート A2以降
 * - PubMed API: ESearch + EFetch
 * - Notion: Database にページ作成（PMID は number プロパティ）
 * - 環境変数は ScriptProperties から取得
 *   - NOTION_API_KEY
 *   - NOTION_DATABASE_ID
 *   - SPREADSHEET_ID
 **************************************/

const SCRIPT_PROP = PropertiesService.getScriptProperties();

// ===== 設定取得系 =====

function getRequiredProp(key) {
  const value = SCRIPT_PROP.getProperty(key);
  if (!value) {
    throw new Error(
      'Script Property "' +
        key +
        '" が設定されていません。プロジェクト設定 → スクリプトプロパティから登録してください。'
    );
  }
  return value;
}

function getSpreadsheet() {
  const id = getRequiredProp('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getNotionToken() {
  return getRequiredProp('NOTION_API_KEY');
}

function getNotionDatabaseId() {
  return getRequiredProp('NOTION_DATABASE_ID');
}

// ===== エントリーポイント（トリガー用） =====

/**
 * これをトリガーに設定（例: 1日1回など）
 */
function fetchAllPubMedQueries() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Queries');
  if (!sheet) {
    throw new Error('「Queries」シートが見つかりません。シート名を確認してください。');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('Queries シートにクエリがありません（A2以降が空）。処理を終了します。');
    return;
  }

  // A2 以降のクエリを取得
  const range = sheet.getRange(2, 1, lastRow - 1, 1); // row2〜, col1(A)
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const query = (values[i][0] || '').toString().trim();
    if (!query) continue;

    console.log('=== クエリ開始: [' + query + '] ===');

    try {
      const articles = fetchPubMedArticles(query);
      console.log('取得論文数: ' + articles.length);

      articles.forEach(function (article, idx) {
        console.log('  → ' + (idx + 1) + '件目 PMID: ' + article.pmid);
        upsertNotionPageForArticle(article);
      });
    } catch (e) {
      console.error('クエリ [' + query + '] の処理中にエラー: ' + e.message + '\n' + e.stack);
    }
  }

  console.log('=== 全クエリ処理完了 ===');
}

// ===== PubMed 連携 =====

/**
 * クエリから PubMed 論文一覧を取得
 * @param {string} query 
 * @returns {Array<Object>} article objects
 */
function fetchPubMedArticles(query) {
  // 1) PMID を検索
  const pmids = pubmedESearch(query);
  console.log('ESearch ヒット数: ' + pmids.length);

  if (pmids.length === 0) {
    return [];
  }

  // 2) efetch で詳細取得
  const articles = pubmedEFetch(pmids);
  return articles;
}

/**
 * PubMed ESearch で PMIDs を取得
 */
function pubmedESearch(query) {
  const baseUrl = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi';
  const params = {
    db: 'pubmed',
    retmode: 'json',
    sort: 'pub+date',
    retmax: '20', // 取得件数（必要に応じて変更）
    term: query
  };

  const url = baseUrl + '?' + toQueryString(params);
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('PubMed ESearch 失敗: HTTP ' + code + ' / ' + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());
  const idList = (((data || {}).esearchresult || {}).idlist) || [];
  return idList;
}

/**
 * PubMed EFetch で詳細情報を取得（XMLをパース）
 */
function pubmedEFetch(pmids) {
  const baseUrl = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi';
  const params = {
    db: 'pubmed',
    retmode: 'xml',
    id: pmids.join(',')
  };
  const url = baseUrl + '?' + toQueryString(params);

  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('PubMed EFetch 失敗: HTTP ' + code + ' / ' + res.getContentText());
  }

  const xml = res.getContentText();
  const doc = XmlService.parse(xml);
  const root = doc.getRootElement(); // PubmedArticleSet

  const articles = [];
  const pubmedArticles = root.getChildren('PubmedArticle');

  pubmedArticles.forEach(function (pa) {
    try {
      const medlineCitation = pa.getChild('MedlineCitation');
      const article = medlineCitation.getChild('Article');

      // PMID
      const pmid = medlineCitation.getChildText('PMID') || '';

      // Title
      const title = article.getChildText('ArticleTitle') || '[No title]';

      // Abstract
      let abstract = '';
      const abstractElem = article.getChild('Abstract');
      if (abstractElem) {
        const absTexts = abstractElem.getChildren('AbstractText');
        const parts = absTexts.map(function (t) {
          return t.getText();
        });
        abstract = parts.join('\n\n');
      }

      // DOI
      let doi = '';
      const pubmedData = pa.getChild('PubmedData');
      const articleIdList = pubmedData ? pubmedData.getChild('ArticleIdList') : null;
      if (articleIdList) {
        const ids = articleIdList.getChildren('ArticleId');
        ids.forEach(function (idElem) {
          const typeAttr = idElem.getAttribute('IdType');
          if (typeAttr && typeAttr.getValue() === 'doi') {
            doi = idElem.getText();
          }
        });
      }

      // Publication Date（できる範囲で整形）
      const journal = article.getChild('Journal');
      let pubDateStr = '';
      if (journal) {
        const issue = journal.getChild('JournalIssue');
        if (issue) {
          const pubDate = issue.getChild('PubDate');
          if (pubDate) {
            const year = pubDate.getChildText('Year');
            const month = pubDate.getChildText('Month');
            const day = pubDate.getChildText('Day');

            if (year) {
              const mm = month ? monthToNumber(month) : '01';
              const dd = day ? ('0' + day).slice(-2) : '01';
              pubDateStr = year + '-' + mm + '-' + dd;
            }
          }
        }
      }

      const url = 'https://pubmed.ncbi.nlm.nih.gov/' + pmid + '/';

      articles.push({
        pmid: pmid,
        title: title,
        abstract: abstract,
        doi: doi,
        pubDate: pubDateStr || null,
        url: url
      });
    } catch (e) {
      console.error('EFetch 1件パース失敗: ' + e.message);
    }
  });

  return articles;
}

/**
 * 月表現を2桁の数字文字列に変換（Jan, Feb ... 対応）
 */
function monthToNumber(m) {
  if (!m) return '01';
  const map = {
    Jan: '01', Feb: '02', Mar: '03', Apr: '04',
    May: '05', Jun: '06', Jul: '07', Aug: '08',
    Sep: '09', Oct: '10', Nov: '11', Dec: '12'
  };
  if (map[m]) return map[m];

  // すでに数字（例: "3"）の場合も想定
  const n = parseInt(m, 10);
  if (!isNaN(n) && n >= 1 && n <= 12) {
    return ('0' + n).slice(-2);
  }

  return '01';
}

/**
 * クエリパラメータを作るユーティリティ
 */
function toQueryString(params) {
  return Object.keys(params)
    .map(function (k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    })
    .join('&');
}

// ===== Notion 連携 =====

/**
 * 1論文ごとに Notion DB にページ作成 or 既存チェック
 */
function upsertNotionPageForArticle(article) {
  const dbId = getNotionDatabaseId();

  // まず PMID で既存ページを検索（あればスキップ）
  const existing = findNotionPageByPMID(dbId, article.pmid);
  if (existing) {
    console.log('  → 既に同じ PMID が存在するためスキップ: ' + article.pmid);
    return;
  }

  // Notion properties 構築
  const properties = {
    "Title": {
      title: [
        {
          text: {
            content: article.title
          }
        }
      ]
    },
    "PubMed URL": article.url ? { url: article.url } : null,
    "Publication Date": article.pubDate
      ? { date: { start: article.pubDate } }
      : null,
    "DOI URL": article.doi
      ? { url: 'https://doi.org/' + article.doi }
      : null,
    "Abstract": article.abstract
      ? {
          rich_text: [
            {
              text: {
                content: article.abstract
              }
            }
          ]
        }
      : undefined
  };

  // PMID は Notion 側 number プロパティ
  const pmidNumber = parseInt(article.pmid, 10);
  if (!isNaN(pmidNumber)) {
    properties["PMID"] = {
      number: pmidNumber
    };
  }

  const payload = {
    parent: {
      database_id: dbId
    },
    properties: properties
  };

  // プロパティで null になったものを削除（Notion APIは null プロパティで怒ることがある）
  Object.keys(payload.properties).forEach(function (key) {
    if (payload.properties[key] == null) {
      delete payload.properties[key];
    }
  });

  const url = 'https://api.notion.com/v1/pages';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + getNotionToken(),
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();

  console.log('Notion response: ' + body);

  if (code !== 200 && code !== 201) {
    throw new Error('Notion ページ作成失敗: HTTP ' + code + '\n' + body);
  }
}

/**
 * Notion DB を PMID で検索
 * - DB側で「PMID」プロパティが number で存在する想定
 */
function findNotionPageByPMID(databaseId, pmid) {
  const url = 'https://api.notion.com/v1/databases/' + databaseId + '/query';

  const pmidNumber = parseInt(pmid, 10);
  if (isNaN(pmidNumber)) {
    console.warn('PMID が数値に変換できません: ' + pmid);
    return null;
  }

  const payload = {
    filter: {
      property: 'PMID',
      number: {
        equals: pmidNumber
      }
    },
    page_size: 1
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + getNotionToken(),
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code !== 200) {
    console.warn('Notion DB Query 失敗: HTTP ' + code + ' / ' + body);
    return null;
  }

  const data = JSON.parse(body);
  if (data.results && data.results.length > 0) {
    return data.results[0];
  }
  return null;
}
