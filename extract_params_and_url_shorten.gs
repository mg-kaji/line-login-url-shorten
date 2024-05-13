// Spreadsheet、配信管理表、出力CSV用の各Sheetを指定
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("URL作成シート");
const sheet_lr = sheet.getLastRow();
let headers = sheet.getDataRange().getValues()[0];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('短縮URL生成');
  menu.addItem('URL短縮1(SS用)', 'gen_shorten_url_1');
  menu.addItem('URL短縮2(all)', 'gen_shorten_url_2');
  menu.addToUi();
}

function gen_shorten_url_1() {
  set_params('Bearer d75cc81ec24a8ebec2c0f94bca36f1f9be781f8c');
}

function gen_shorten_url_2() {
  set_params('Bearer d7b3df5b156d9d45b803ca5fd04f61ab3da5d45b');  
}

// 列番号を変数として格納
const numCols = {
  origin_url: findHeader('元URL'),
  complete_url: findHeader('完成URL'),
  shorten_url: findHeader('短縮URL'),
  format: findHeader('URL形式'),
  lps: findHeader('LOOPASS LINEログインURL'),
  measure: findHeader('計測パラメータ付きURL')
};

// パラメータをセットして短縮URLを作成
// main関数
function set_params(bearer) {
  for (let i = 2; i <= sheet_lr; i++) {
    // URL, URL形式を取得して全角と空白を削除
    let url = sheet.getRange(i, numCols["origin_url"]).getValue();
    let url_format = sheet.getRange(i, numCols['format']).getValue();
    url = cleanText(url);
    url_format = cleanText(url_format);

    // 計測パラメータがついたURLの取得
    const url_measure = sheet.getRange(i, numCols["measure"]).getValue();
    if (url.length <= 0) {continue;}  // URL欄が空白ならその行をスキップ

    let dict_params = {};
    const param_list = getParamsList(url_format);
    for (param of param_list) {
      const param_value = getParamValue(url, url_format, param);
      const param_name = extParamName(param);
      dict_params[param_name] = param_value;
    }

    if (url_measure) {
      const dict_measure_params = extractParamFromMeasure(url_measure);
      // dict_params = Object.assign(dict_measure_params, dict_params); // 連想配列を結合(重複が存在する場合は元URLの方を優先)
      dict_params = mergePrameters(dict_params, dict_measure_params);
    }
    const attach_url = concatParams(dict_params);

    console.log(dict_params);
    for (const [param_name, param_value] of Object.entries(dict_params)){
      setParamToHeader(param_name);
      const param_header_index = headers.indexOf(param_name)
      sheet.getRange(i, param_header_index+1).setValue(param_value);
    }

    const lps_url = sheet.getRange(i, numCols["lps"]).getValue();
    let url_fn = lps_url + attach_url;

    sheet.getRange(i, numCols["complete_url"]).setValue(url_fn);
    url_shorten = shorten_url(url_fn, bearer);
    sheet.getRange(i, numCols["shorten_url"]).setValue(url_shorten);
  }

  decorateHeader();
  setRuledLine();
}

// 元URLと計測用パラメータのパラメータを結合
// 元URLの値が計測用パラメータの名前に使われていた際は除外
function mergePrameters(origin_params, measure_params){
  const origin_values = Object.values(origin_params);
  const measure_keys = Object.keys(measure_params);

  const duplicateParams = getIsDuplicate(origin_values, measure_keys);
  for (const param_name of duplicateParams) {
    console.log('Duplicate Parameter:', param_name);
    delete measure_params[param_name];
  }
  return Object.assign(measure_params, origin_params);
}

function getIsDuplicate(arr1, arr2) {
  function removeDuplicateValues([...array]) {
    return array.filter((value, index, self) => self.indexOf(value) === index);
  }

  const list_duplicate =  [...arr1, ...arr2].filter(item => arr1.includes(item) && arr2.includes(item));
  return new removeDuplicateValues(list_duplicate);
}

// function debug() {
//   const arr1 = { from: 'taglist_kw_156',
//   utm_source: 'li',
//   utm_medium: 'social',
//   utm_term: 'TL',
//   utm_campaign: '210921'};
//   const arr2 = {a: 'tags',
//   b: '156',
//   c: 'from',
//   d: 'taglist_kw_156' };
//   console.log(mergePrameters(arr2, arr1));
// }

// 連想配列のキーをパラメータ名、値をパラメータの値としてURLに付与する形で成形
function concatParams(dict_param) {
  let full_param = '';
  for (const [param, value] of Object.entries(dict_param)) {
    full_param += modelingParam(param, value);
  }
  return full_param;
}

// URLに使用する形にパラメータを成形
function modelingParam(param, value) {
  let full_param = ''
  full_param = '&' + param + '=' + value;
  return full_param
}

// 計測用パラメータURLからパラメータの部分のみ抽出
function extractParamFromMeasure(url) {
  const url_sep = '?';
  const param_sep = '&';
  const key_val_sep = '=';
  const pattern = /[a-z|A-Z|_|0-9]+\=[a-z|A-Z|0-9|_]+/g;

  const param_url = url.split(url_sep)[1];
  const separated_param = param_url.match(pattern);
  let dict_params = {};
  for (full_param of separated_param) {
    // console.log('Measure param:', full_param);
    const [param_name, value] = full_param.split(key_val_sep);
    dict_params[param_name] = value;
  }
  return dict_params;
}

var zenToHan = function(value) {
    if (!value) return value;
    return String(value).replace(/[！-～]/g, function(all) {
        return String.fromCharCode(all.charCodeAt(0) - 0xFEE0);
    });
};

function cleanText(str) {
  let str_ = str;
  str_ = zenToHan(str_);
  str_ = str_.replace(/\s+/g, '');

  return str_;
}

// URLの中から{xxx}の形式になっている文字列からxxxを抽出
function getParamsList(format) {
  const separated_format = separate(format);
  // const pattern = /^{.+[]}]$/g
  const pattern = /{[a-z|A-Z|0-9|_]+}/g;
  let parameters = [];
  for (element of separated_format) {
    if(element.match(pattern)){
      // const num_letter = element.length;
      // const param = element.substring(1, num_letter-1);  // 正規表現で取ろうとするとうまくいかない
      parameters.push(element);
    }
  }
  return parameters;
}

// ヘッダーに指定のパラメータがなれけば追加
function setParamToHeader(param_name) {
  // const param_name = extractFromParam(param);
  // const param_name = extParamName(param);
  if (findHeader(param_name) === 0) {  // シートにパラメータの名前がないときは追加
    const last_col = sheet.getLastColumn();
    sheet.getRange(1, last_col+1).setValue(param_name);
    // console.log('value', param_name);
  }
  headers = sheet.getDataRange().getValues()[0]; // グローバル変数にしている'header'を最新のものに置換
}

// {aa}の形からaaを取り出す
function extractFromParam(param) {
  const param_pattern = /[^\{][a-z|A-Z|0-9|_]+[^\}]/g;
  const param_name = param.match(param_pattern)[0];  // {aa}の形からaaを取り出す
  return param_name;
}

// {aa}の形からaaを取り出す
// 正規表現だと取れないことがあるので文字列のメソッドで取得
function extParamName(param) {
  const num_letter = param.length;
  return param.substring(1, num_letter-1);
}

// {/ ? & = #}のどれかで分割(URLのパラメータで使ってるものを設定)
function separate(string) {
  const sep = /[^/|?|=|& | \#]+/g;
  const separated = string.match(sep);
  return separated;
}

// URLから指定のパラメータを取得
function getParamValue(url, url_format, param) {
  const splited_url = separate(url);
  const splited_format = separate(url_format);

  const indx = splited_format.indexOf(param);
  const param_value = indx !== -1 ? splited_url[indx] : "";

  return param_value;
};

// 列の探索
function findHeader(value) {
  for (let i = 0; i <= headers.length; i++) {
    if (headers[i] === value) {
      return i + 1;
    }
  }
  return 0;
};

// https://www.example.com/hoge/hoge2/ から example.comだけ取り出す
function getDomain(url) {
  const regex = /[^\/\/]([a-z|A-Z|0-9]+[.|-])+[a-z|A-Z|0-9]+[^\/]/g;
  const full_domain = url.match(regex)[0];

  const domain_pattern = /[a-z|A-Z|0-9|-]+[.][a-z|A-Z|0-9]+$/g;
  const domain = full_domain.match(domain_pattern)[0];
  return domain
}

// bitly APIを叩いて短縮URLを取得
function shorten_url(url_fn, bearer){
  const endpoint = 'https://api-ssl.bitly.com/v4/shorten';
  const headers = {
    "Content-Type": "application/json",
    Authorization: bearer
  };
  const params = {
    long_url: url_fn
  };
  const options = {
    headers: headers,
    method: "POST",
    "payload": JSON.stringify(params)
  };
  const result = UrlFetchApp.fetch(endpoint, options);
  const json = JSON.parse(result.getContentText('utf-8'));
  return json.link;
}

function setRuledLine() {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setBorder(true, true, true, true, true, true);
}

function decorateHeader() {
  const headerRow = 1;
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(headerRow, 1, 1, lastCol);
  range.setFontWeight('bold');
  range.setFontColor('white');
  range.setBackground('#806000');
  range.setFontSize(11);
}