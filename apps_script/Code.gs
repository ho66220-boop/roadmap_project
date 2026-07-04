/**
 * 울산 고교 과목 로드맵 3단계 추천 웹앱 (Google Apps Script)
 *
 * 데이터: 이 스크립트가 바인딩된 스프레드시트의 시트들
 *   고1편제표      — 학교별 2026 신입생(고1) 편제표 (scripts/build_curriculum.py 산출물)
 *   학과추천       — 커리어넷 기반 학과별 일반/진로/융합 추천 과목   → Layer 1
 *   대학트랙       — 대교협 기반 대학·모집단위별 핵심/권장 과목      → Layer 2
 *   대학제시율_학과 / 대학제시율_계열 — 과목별 제시 비율(근거·정렬)
 *   과목마스터     — 공식 과목명 정규화 키
 *
 * 배포: 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일과 index.html을
 *       넣고 "웹 앱"으로 배포한다. (자세한 절차는 저장소 apps_script/README.md)
 */

var SEMESTER_KEYS = ['1-1', '1-2', '2-1', '2-2', '3-1', '3-2'];
var SEMESTER_LABELS = {
  '1-1': '1학년 1학기', '1-2': '1학년 2학기',
  '2-1': '2학년 1학기', '2-2': '2학년 2학기',
  '3-1': '3학년 1학기', '3-2': '3학년 2학기'
};

// 대학트랙 과목명에 등장하는 교과(군) 단위 포괄 표현 — 개별 과목이 아니므로 편제표 대조에서 제외
var GROUP_TERMS = {
  '국어': 1, '수학': 1, '영어': 1, '사회': 1, '과학': 1, '한국사': 1, '체육': 1, '예술': 1,
  '기술가정': 1, '정보': 1, '제2외국어': 1, '한문': 1, '교양': 1, '사회역사도덕포함': 1,
  // 구 교육과정식 대학트랙 표현(2022 개정 개별 과목이 아닌 교과 단위 포괄)
  '역사': 1, '윤리': 1, '지리': 1, '일반사회': 1, '과학교과': 1, '수학1': 1, '수학2': 1,
  '한국사12': 1, '전과목': 1, '과학전과목': 1, '수학전과목': 1,
  '제2외국어관련과목': 1, '제2외국어과목': 1, '제3외국어과목': 1,
  // 원문이 원문자(circled digit) 글리프를 써서 norm()이 로마숫자처럼 변환하지 못하는 실측 표기
  '수학①': 1, '수학②': 1
};

// 구식/포괄 과목 표기 -> 2022 개정 공식 과목명
var L2_SUBJECT_ALIASES = {
  '미적분': ['미적분Ⅰ', '미적분Ⅱ'],
  '물리': ['물리학'],
  '한국지리': ['한국지리 탐구'],
  '동아시아사': ['동아시아 역사 기행'],
  '수학(특히 미적분)': ['미적분Ⅰ', '미적분Ⅱ'],
  '프랑스 회화': ['프랑스어 회화'],
  '도시와 미래탐구': ['도시의 미래 탐구'],
  '과학과제 탐구': ['과학과제 연구'],
  '물리과 에너지': ['역학과 에너지'],
  '독해와 작문': ['영어 독해와 작문']
};

// 목표 학과 축약 표현 보정
var MAJOR_REPLACEMENTS = {
  '컴공': '컴퓨터공학', '컴퓨터': '컴퓨터공학', '소프트웨어': '컴퓨터공학',
  '기계': '기계공학', '생명': '생명공학', '바이오': '생명공학'
};

// ------------------------------------------------------------------ 웹앱 진입점

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('울산 고교 과목 로드맵 상담')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------------------------------------------------------ 유틸

function norm(value) {
  if (value === null || value === undefined) return '';
  var text = String(value).trim().toLowerCase();
  text = text.replace(/Ⅰ/g, '1').replace(/Ⅱ/g, '2');
  text = text.replace(/[\s·ㆍ․,/_()\-\[\]]+/g, '');
  text = text.replace(/대학교/g, '대');
  return text;
}

function toFloat(value) {
  var n = parseFloat(value);
  return isNaN(n) ? 0 : n;
}

function majorBaseKey(text) {
  var value = norm(text);
  var suffixes = ['학과', '전공', '계열', '학부', '과', '부'];
  for (var i = 0; i < suffixes.length; i++) {
    var s = suffixes[i];
    if (value.length > s.length && value.slice(-s.length) === s) {
      value = value.slice(0, -s.length);
      break;
    }
  }
  return value;
}

// ------------------------------------------------------------------ 데이터 로딩

function sheetValues_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + name);
  var values = sheet.getDataRange().getValues();
  values.shift(); // 헤더 제거
  return values;
}

function loadStore_() {
  var store = {
    schools: {},          // 학교 -> [{subject, section, group, type, semesters[], choiceGroup, takeN}]
    deptRecs: {},         // 학과 -> {일반: [], 진로: [], 융합: []}
    deptTrack: {},        // 학과 -> 계열
    uniTracks: [],        // {university, unit, subject, trackType, comment}
    rateByDept: {},       // 학과 -> {과목: 비율}
    rateByTrack: {},      // 계열 -> {과목: 비율}
    subjectMaster: {},    // 정규화키 -> 공식과목명
    allOfferedKeys: {}    // 전 학교(GAS 노출 범위 기준) 개설과목 norm키 집합(공통 미개설 판정용)
  };

  sheetValues_('과목마스터').forEach(function (row) {
    if (!row[0]) return;
    var official = String(row[0]).trim();
    var key = row[1] ? String(row[1]).trim() : norm(official);
    store.subjectMaster[key] = official;
    store.subjectMaster[norm(official)] = official;
  });

  sheetValues_('학과추천').forEach(function (row) {
    if (!row[0] || !row[3]) return;
    var dept = String(row[0]).trim();
    var track = String(row[1] || '').trim();
    var kind = String(row[2] || '일반').trim();
    var subject = String(row[3]).trim();
    if (!store.deptRecs[dept]) store.deptRecs[dept] = {};
    if (!store.deptRecs[dept][kind]) store.deptRecs[dept][kind] = [];
    store.deptRecs[dept][kind].push(subject);
    if (track) store.deptTrack[dept] = track;
  });

  sheetValues_('대학트랙').forEach(function (row) {
    if (!row[0] || !row[2]) return;
    store.uniTracks.push({
      university: String(row[0]).trim(),
      unit: String(row[1] || '').trim(),
      subject: String(row[2]).trim(),
      trackType: String(row[3] || '권장').trim(),
      comment: String(row[5] || '').trim()
    });
  });

  [['대학제시율_학과', 'rateByDept'], ['대학제시율_계열', 'rateByTrack']].forEach(function (pair) {
    sheetValues_(pair[0]).forEach(function (row) {
      if (!row[0] || !row[1]) return;
      var key = String(row[0]).trim();
      if (!store[pair[1]][key]) store[pair[1]][key] = {};
      store[pair[1]][key][String(row[1]).trim()] = toFloat(row[4]);
    });
  });

  sheetValues_('고1편제표').forEach(function (row) {
    if (!row[0]) return;
    var school = String(row[0]).trim();
    var semesters = [];
    for (var i = 0; i < SEMESTER_KEYS.length; i++) {
      if (toFloat(row[10 + i]) > 0) semesters.push(SEMESTER_KEYS[i]);
    }
    var subject = String(row[6] || row[5]).trim();
    if (!store.schools[school]) store.schools[school] = [];
    store.schools[school].push({
      subject: subject,
      section: String(row[2] || '').trim(),
      group: String(row[3] || '').trim(),
      type: String(row[4] || '').trim(),
      semesters: semesters,
      choiceGroup: row[18] ? String(row[18]).trim() : '',
      takeN: row[19] ? parseInt(row[19], 10) || 0 : 0
    });
    store.allOfferedKeys[norm(subject)] = 1;
  });

  return store;
}

// ------------------------------------------------------------------ 공개 API (클라이언트 호출)

function getMeta() {
  var store = loadStore_();
  var universities = {};
  store.uniTracks.forEach(function (t) { universities[t.university] = 1; });
  return {
    schoolCount: Object.keys(store.schools).length,
    deptCount: Object.keys(store.deptRecs).length,
    universityCount: Object.keys(universities).length,
    schools: Object.keys(store.schools).sort(),
    majors: Object.keys(store.deptRecs).sort()
  };
}

function getRecommendation(schoolText, majorText) {
  var store = loadStore_();
  var school = pickSchool_(store, schoolText);
  var major = pickMajor_(store, majorText);

  if (!school) {
    return needMore_('학교명을 알려주면 해당 학교 편제표 안에서 실제 선택 가능한 과목만 골라볼 수 있어요. (지원 학교: '
      + Object.keys(store.schools).sort().join(', ') + ')');
  }
  if (!major) {
    return needMore_('목표 학과나 계열을 알려주면 커리어넷·대교협 기준 추천 과목을 찾아볼게요.');
  }

  var track = store.deptTrack[major] || '';
  var layer1 = layer1Dept_(store, major, track);
  var layer2 = layer2Universities_(store, major);
  var layer3 = layer3School_(store, school, layer1, layer2);

  var result = {
    mode: 'recommendation',
    school: school,
    major: major,
    track: track,
    layer1: layer1,
    layer2: layer2,
    layer3: layer3,
    sources: [
      '학과추천 시트 (커리어넷)',
      '대학트랙 시트 (대교협 시행계획)',
      '대학제시율 시트',
      '고1편제표 시트 / ' + school + ' 2026 신입생 3개년 편제표'
    ],
    caution: '추천은 2026학년도 신입생(현 고1) 편제표와 커리어넷·대교협 공개 자료 기준입니다. '
      + '실제 개설 여부와 선택 규칙(택N)은 학교 안내를 최종 확인하세요.'
  };
  result.answer = composeAnswer_(result);
  return result;
}

// ------------------------------------------------------------------ Layer 1: 커리어넷 학과 추천

function layer1Dept_(store, major, track) {
  var recs = store.deptRecs[major] || {};
  var trackRates = store.rateByTrack[track] || {};
  var categories = [];
  var total = 0;
  ['일반', '진로', '융합'].forEach(function (kind) {
    var seen = {};
    var subjects = [];
    (recs[kind] || []).forEach(function (subject) {
      var key = norm(subject);
      if (seen[key]) return;
      seen[key] = 1;
      subjects.push({ subject: subject, track_rate: trackRates[subject] || null });
    });
    subjects.sort(function (a, b) { return (b.track_rate || 0) - (a.track_rate || 0); });
    categories.push({ kind: kind, subjects: subjects });
    total += subjects.length;
  });
  return { dept: major, track: track, categories: categories, total: total };
}

// ------------------------------------------------------------------ Layer 2: 대교협 대학별 과목

function majorSimilarity_(target, candidate) {
  var t = majorBaseKey(target);
  var c = majorBaseKey(candidate);
  if (!t || !c) return 0;
  if (t === c) return 100;
  if (t.length >= 2 && c.length >= 2 && (c.indexOf(t) >= 0 || t.indexOf(c) >= 0)) return 80;
  var aliases = {
    '컴퓨터공학': ['컴퓨터', '소프트웨어', 'ai', '인공지능', '정보컴퓨터', '데이터'],
    '생명공학': ['생명', '바이오', '의생명', '식품생명', '분자생명'],
    '기계공학': ['기계', '로봇', '자동차', '메카트로닉스']
  };
  var targetNorm = norm(target);
  var candidateNorm = norm(candidate);
  for (var root in aliases) {
    if (targetNorm.indexOf(norm(root)) >= 0) {
      for (var i = 0; i < aliases[root].length; i++) {
        if (candidateNorm.indexOf(norm(aliases[root][i])) >= 0) return 70;
      }
    }
  }
  return 0;
}

/** 대학트랙 과목명을 (개별 공식과목들, 교과군 포괄 표현들, 폐기된 원문들)로 분해한다.
 *
 * 마스터에도 GROUP_TERMS에도 걸리지 않는 잔여는 가짜 과목으로 남기지 않고 폐기한다
 * (감사 결과 마스터-정합 과목은 전부 매칭되므로 안전. 폐기 건은 layer2.dropped_terms로 노출).
 */
function expandTrackSubject_(store, text) {
  var subjects = [];
  var groups = [];
  var dropped = [];
  String(text).split(/\s*(?:또는|,|\/)\s*/).forEach(function (part) {
    part = part.trim();
    if (!part) return;
    var expanded = L2_SUBJECT_ALIASES[part] || [part];
    expanded.forEach(function (item) {
      var key = norm(item);
      var official = store.subjectMaster[key];
      if (official) subjects.push(official);
      else if (GROUP_TERMS[key]) groups.push(part);
      else dropped.push(item);
    });
  });
  return { subjects: subjects, groups: groups, dropped: dropped };
}

function layer2Universities_(store, major) {
  var agg = {};
  var groupMentions = {};
  var droppedCounts = {};
  var matchedUnits = {};

  store.uniTracks.forEach(function (row) {
    if (row.subject.charAt(0) === '(') return; // (계열 공통) 등 메타 행
    if (majorSimilarity_(major, row.unit) < 70) return;
    matchedUnits[row.university + ' ' + row.unit] = 1;

    var expanded = expandTrackSubject_(store, row.subject);
    expanded.groups.forEach(function (group) {
      if (!groupMentions[group]) groupMentions[group] = {};
      groupMentions[group][row.university] = 1;
    });
    expanded.dropped.forEach(function (term) {
      droppedCounts[term] = (droppedCounts[term] || 0) + 1;
    });
    expanded.subjects.forEach(function (subject) {
      if (!agg[subject]) agg[subject] = { subject: subject, universities: {}, types: {}, comment: '' };
      agg[subject].universities[row.university] = 1;
      agg[subject].types[row.trackType] = 1;
      if (row.comment && !agg[subject].comment) agg[subject].comment = row.comment.slice(0, 120);
    });
  });

  var deptRates = store.rateByDept[major] || {};
  var subjects = Object.keys(agg).map(function (name) {
    var item = agg[name];
    var unis = Object.keys(item.universities).sort();
    return {
      subject: name,
      type: item.types['핵심'] ? '핵심' : '권장',
      university_count: unis.length,
      universities: unis.slice(0, 6),
      rate: deptRates[name] || null,
      comment: item.comment
    };
  });
  subjects.sort(function (a, b) {
    var ta = a.type === '핵심' ? 0 : 1;
    var tb = b.type === '핵심' ? 0 : 1;
    return ta !== tb ? ta - tb : b.university_count - a.university_count;
  });

  var groups = Object.keys(groupMentions).map(function (g) {
    return { group: g, university_count: Object.keys(groupMentions[g]).length };
  });
  groups.sort(function (a, b) { return b.university_count - a.university_count; });

  var droppedTerms = Object.keys(droppedCounts).map(function (term) {
    return { term: term, count: droppedCounts[term] };
  });
  droppedTerms.sort(function (a, b) { return b.count - a.count; });

  return {
    unit_count: Object.keys(matchedUnits).length,
    units_sample: Object.keys(matchedUnits).sort().slice(0, 8),
    subjects: subjects.slice(0, 20),
    group_mentions: groups,
    dropped_terms: droppedTerms
  };
}

// ------------------------------------------------------------------ Layer 3: 학교 편제표 대조

function layer3School_(store, school, layer1, layer2) {
  var offerings = store.schools[school] || [];
  var offerMap = {};
  offerings.forEach(function (item) {
    var key = norm(item.subject);
    if (!offerMap[key]) offerMap[key] = [];
    offerMap[key].push(item);
  });

  // 추천 과목 통합: 대학 핵심 > 대학 권장 > 커리어넷 진로 > 일반 > 융합
  var merged = {};
  function add(subject, source, rank, extra) {
    var key = norm(subject);
    if (GROUP_TERMS[key]) return;
    if (!merged[key] || rank < merged[key].rank) {
      var info = { subject: subject, source: source, rank: rank };
      for (var k in extra) info[k] = extra[k];
      merged[key] = info;
    }
  }
  layer2.subjects.forEach(function (item) {
    add(item.subject, '대학 ' + item.type, item.type === '핵심' ? 0 : 1,
      { university_count: item.university_count, rate: item.rate });
  });
  var kindRank = { '진로': 2, '일반': 3, '융합': 4 };
  layer1.categories.forEach(function (category) {
    category.subjects.forEach(function (entry) {
      add(entry.subject, '커리어넷 ' + category.kind, kindRank[category.kind],
        { track_rate: entry.track_rate });
    });
  });

  var infos = Object.keys(merged).map(function (k) { return merged[k]; });
  infos.sort(function (a, b) { return a.rank - b.rank; });

  var available = [];
  var unavailable = [];
  infos.forEach(function (info) {
    var rows = offerMap[norm(info.subject)];
    var item = {
      subject: info.subject, source: info.source,
      university_count: info.university_count, rate: info.rate, track_rate: info.track_rate
    };
    if (rows) {
      var semSet = {};
      var secSet = {};
      rows.forEach(function (r) {
        r.semesters.forEach(function (s) { semSet[s] = 1; });
        if (r.section) secSet[r.section] = 1;
      });
      var semesters = SEMESTER_KEYS.filter(function (s) { return semSet[s]; });
      item.semesters = semesters;
      item.semester_labels = semesters.map(function (s) { return SEMESTER_LABELS[s]; });
      item.sections = Object.keys(secSet).sort();

      // 같은 과목이 여러 행이면 지정(택N 그룹 미소속) 우선(감사상 충돌 0건이라 실질 단일)
      var plainRows = rows.filter(function (r) { return !r.choiceGroup; });
      if (plainRows.length) {
        item.mode = '지정';
        item.take_n = 0;
        item.choice_group = '';
      } else {
        var rep = rows[0];
        item.mode = '선택군';
        item.take_n = rep.takeN;
        item.choice_group = rep.choiceGroup;
      }
      available.push(item);
    } else {
      item.scope = store.allOfferedKeys[norm(info.subject)] ? '' : '공통 미개설';
      unavailable.push(item);
    }
  });

  var bySemester = {};
  available.forEach(function (item) {
    (item.semesters || []).forEach(function (sem) {
      if (!bySemester[sem]) bySemester[sem] = [];
      if (bySemester[sem].length < 8) bySemester[sem].push(item.subject);
    });
  });
  var plan = SEMESTER_KEYS.filter(function (s) { return bySemester[s]; }).map(function (s) {
    return { semester: s, label: SEMESTER_LABELS[s], subjects: bySemester[s] };
  });

  return {
    school: school,
    offering_count: offerings.length,
    available: available.slice(0, 24),
    unavailable: unavailable.slice(0, 12),
    plan: plan
  };
}

// ------------------------------------------------------------------ 매칭

function pickSchool_(store, text) {
  var textNorm = norm(text);
  if (!textNorm) return '';
  var schools = Object.keys(store.schools);
  for (var i = 0; i < schools.length; i++) {
    var s = norm(schools[i]);
    if (textNorm.indexOf(s) >= 0 || s.indexOf(textNorm) >= 0) return schools[i];
  }
  return '';
}

function pickMajor_(store, text) {
  text = String(text || '').trim();
  var matched = matchMajorText_(store, text);
  if (matched) return matched;
  for (var key in MAJOR_REPLACEMENTS) {
    if (text.indexOf(key) >= 0) {
      matched = matchMajorText_(store, MAJOR_REPLACEMENTS[key]);
      if (matched) return matched;
      break;
    }
  }
  return bestMajorByOverlap_(store, text);
}

function matchMajorText_(store, text) {
  var textNorm = norm(text);
  if (!textNorm) return '';
  var base = majorBaseKey(text);
  var depts = Object.keys(store.deptRecs);
  for (var i = 0; i < depts.length; i++) {
    if (norm(depts[i]) === textNorm || majorBaseKey(depts[i]) === base) return depts[i];
  }
  var candidates = [];
  for (var j = 0; j < depts.length; j++) {
    var d = norm(depts[j]);
    if (d.indexOf(textNorm) >= 0 || textNorm.indexOf(d) >= 0) candidates.push(depts[j]);
  }
  if (candidates.length) {
    // '경영학' -> '의료경영학과'가 아니라 '경영학과'를 고르도록 가장 짧은 후보 선택
    candidates.sort(function (a, b) { return norm(a).length - norm(b).length; });
    return candidates[0];
  }
  return '';
}

function bestMajorByOverlap_(store, text) {
  var tokens = String(text).match(/[가-힣A-Za-z0-9]+/g) || [];
  var best = '';
  var bestScore = 0;
  Object.keys(store.deptRecs).forEach(function (dept) {
    var score = 0;
    tokens.forEach(function (token) { if (token && dept.indexOf(token) >= 0) score++; });
    if (score > bestScore) { best = dept; bestScore = score; }
  });
  return best;
}

// ------------------------------------------------------------------ 상담 문장

function composeAnswer_(result) {
  var lines = [];
  var layer1 = result.layer1;
  var layer2 = result.layer2;
  var layer3 = result.layer3;

  var parts = [];
  layer1.categories.forEach(function (category) {
    var names = category.subjects.slice(0, 5).map(function (s) { return s.subject; });
    if (names.length) parts.push(category.kind + ' 선택은 ' + names.join(', '));
  });
  var trackNote = layer1.track ? ' (' + layer1.track + ' 기준)' : '';
  lines.push(parts.length
    ? '학과 추천 과목' + trackNote + '\n커리어넷 기준 ' + result.major + ' 관련 ' + parts.join(' / ')
    : '학과 추천 과목\n' + result.major + '에 대한 커리어넷 추천 데이터를 찾지 못했습니다.');

  if (layer2.subjects.length) {
    var uniLines = ['대학별 강조 과목\n대교협 시행계획 기준 관련 모집단위 ' + layer2.unit_count + '곳을 보면,'];
    var core = layer2.subjects.filter(function (s) { return s.type === '핵심'; }).slice(0, 5);
    var rec = layer2.subjects.filter(function (s) { return s.type === '권장'; }).slice(0, 4);
    if (core.length) {
      uniLines.push('핵심 과목: ' + core.map(function (s) {
        var rate = s.rate ? ', 제시율 ' + Math.round(s.rate) + '%' : '';
        return s.subject + '(' + s.university_count + '개 대학' + rate + ')';
      }).join(', '));
    }
    if (rec.length) {
      uniLines.push('권장 과목: ' + rec.map(function (s) {
        return s.subject + '(' + s.university_count + '개 대학)';
      }).join(', '));
    }
    var groups = layer2.group_mentions.slice(0, 3);
    if (groups.length) {
      uniLines.push('이외에 ' + groups.map(function (g) {
        return g.group + ' 교과 전반(' + g.university_count + '개 대학)';
      }).join(', ') + '을 폭넓게 강조합니다.');
    }
    lines.push(uniLines.join('\n'));
  } else {
    lines.push('대학별 강조 과목\n대교협 자료에서 직접 일치하는 모집단위를 찾지 못했습니다. 학과명을 조금 바꿔 검색해 보세요.');
  }

  if (layer3.available.length) {
    var availLines = [result.school + ' 편제표 확인\n위 추천 과목 중 ' + result.school + '에서 실제 선택할 수 있는 과목은 다음과 같습니다.'];
    layer3.available.slice(0, 6).forEach(function (item) {
      var sems = (item.semester_labels || []).join(', ') || '학기 미확정';
      var modeNote;
      if (item.mode === '선택군') {
        modeNote = item.take_n ? ' (택' + item.take_n + ' 선택군 — 선택 신청 필요)' : ' (선택군 — 선택 신청 필요)';
      } else {
        modeNote = ' (학교지정)';
      }
      availLines.push('- ' + item.subject + ' [' + item.source + ']' + modeNote + ' → ' + sems);
    });
    lines.push(availLines.join('\n'));
  } else {
    lines.push(result.school + ' 편제표 확인\n추천 과목 중 ' + result.school + ' 편제표에서 확인되는 과목이 없습니다. 편제표 데이터를 점검해 주세요.');
  }

  var unavailableCommon = layer3.unavailable.filter(function (i) { return i.scope === '공통 미개설'; });
  var unavailableNormal = layer3.unavailable.filter(function (i) { return i.scope !== '공통 미개설'; });
  if (unavailableNormal.length) {
    var names = unavailableNormal.slice(0, 5).map(function (i) { return i.subject; }).join(', ');
    lines.push('미개설 과목\n' + names + '은(는) 추천 근거에는 있지만 ' + result.school
      + ' 편제표에서 확인되지 않습니다. 공동교육과정이나 온라인학교 개설 여부를 확인해 보세요.');
  }
  if (unavailableCommon.length) {
    var commonNames = unavailableCommon.slice(0, 5).map(function (i) { return i.subject; }).join(', ');
    lines.push('일반고 공통 미개설\n' + commonNames + '은(는) 조사된 모든 학교 편제표에서 확인되지 않습니다. '
      + '일반고 공통 미개설 — 공동교육과정·온라인학교 확인이 필요합니다.');
  }

  lines.push('주의\n' + result.caution);
  return lines.join('\n\n');
}

function needMore_(message) {
  return {
    mode: 'need_more',
    answer: message,
    layer1: { dept: '', track: '', categories: [], total: 0 },
    layer2: { unit_count: 0, units_sample: [], subjects: [], group_mentions: [], dropped_terms: [] },
    layer3: { school: '', offering_count: 0, available: [], unavailable: [], plan: [] },
    sources: [],
    caution: ''
  };
}
