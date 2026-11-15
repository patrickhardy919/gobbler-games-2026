/* ============================================================
   Google Sheets backend for The Great Gobbler Games 2026
   Paste this whole file into Extensions > Apps Script, run setup() once,
   then Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
   ============================================================ */

var TABLES = {
  Players:  ['id', 'name', 'av', 'joined'],
  Ledger:   ['ts', 'pid', 'gid', 'pts', 'note'],
  Plays:    ['pid', 'gid', 'ts', 'status', 'pts', 'st'],
  Games:    ['id', 'day', 'type', 'title', 'emoji', 'blurb', 'cfg'],
  Chat:     ['id', 'ts', 'pid', 'name', 'av', 'text', 'kind'],
  Settings: ['key', 'value']
};
var NUM_COLS = { ts: 1, pts: 1, av: 1, day: 1, joined: 1 };
var JSON_COLS = { st: 1, cfg: 1 };
var TABLE_NAME = { players: 'Players', ledger: 'Ledger', plays: 'Plays', games: 'Games', chat: 'Chat', settings: 'Settings' };

function makeDb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), memo = {};
  function sheet(t) {
    var sh = ss.getSheetByName(TABLE_NAME[t]);
    if (!sh) throw new Error('Missing sheet tab ' + TABLE_NAME[t] + '. Run setup() first.');
    return sh;
  }
  function cell(col, v) {
    if (JSON_COLS[col]) return v == null ? '' : JSON.stringify(v);
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;   // stop typed text from becoming a formula
    return v == null ? '' : v;
  }
  function fromCell(col, v) {
    if (JSON_COLS[col]) { try { return v === '' ? null : JSON.parse(v); } catch (e) { return null; } }
    if (NUM_COLS[col]) return Number(v) || 0;
    return v === '' ? '' : String(v);
  }
  return {
    read: function (t) {
      if (!memo[t]) {
        var cols = TABLES[TABLE_NAME[t]], vals = sheet(t).getDataRange().getValues(), out = [];
        for (var i = 1; i < vals.length; i++) {
          if (vals[i].every(function (x) { return x === ''; })) continue;
          var o = {}; cols.forEach(function (c, k) { o[c] = fromCell(c, vals[i][k]); }); out.push(o);
        }
        memo[t] = JSON.stringify(out);
      }
      return JSON.parse(memo[t]);
    },
    add: function (t, row) {
      var cols = TABLES[TABLE_NAME[t]];
      sheet(t).appendRow(cols.map(function (c) { return cell(c, row[c]); }));
      memo[t] = null;
    },
    save: function (t, rows) {
      var sh = sheet(t), cols = TABLES[TABLE_NAME[t]], n = sh.getLastRow();
      if (n > 1) sh.getRange(2, 1, n - 1, cols.length).clearContent();
      if (rows.length) sh.getRange(2, 1, rows.length, cols.length).setValues(rows.map(function (r) { return cols.map(function (c) { return cell(c, r[c]); }); }));
      memo[t] = null;
    }
  };
}

var READ_ONLY = { sync: 1, state: 1, tv: 1, picks: 1 };

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function run(req) {
  var props = PropertiesService.getScriptProperties();
  var engine = createEngine(makeDb(), { pin: function () { return props.getProperty('HOST_PIN') || '5062'; } });
  if (READ_ONLY[req.a]) return engine.handle(req);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return engine.handle(req); } finally { lock.releaseLock(); }
}
function doGet(e) {
  var req = {}; for (var k in e.parameter) req[k] = e.parameter[k];
  if (!READ_ONLY[req.a]) return respond({ ok: false, error: 'Use POST for that action.' });
  return respond(run(req));
}
function doPost(e) {
  var req; try { req = JSON.parse(e.postData.contents); } catch (x) { return respond({ ok: false, error: 'Bad request.' }); }
  return respond(run(req));
}

/* Run this once from the Apps Script editor. It is safe to run again. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name), cols = TABLES[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#FFC233');
    sh.setFrozenRows(1);
  });
  var lb = ss.getSheetByName('Leaderboard') || ss.insertSheet('Leaderboard', 0);
  lb.clear();
  lb.getRange('A1').setFormula('=QUERY({Players!B2:B, ARRAYFORMULA(SUMIF(Ledger!B2:B, Players!A2:A, Ledger!D2:D))}, "select Col1, Col2 where Col1 is not null order by Col2 desc label Col1 \'Player\', Col2 \'Points\'", 1)');
  lb.getRange('A1:B1').setFontWeight('bold').setBackground('#F26B1D').setFontColor('#ffffff');
  lb.setColumnWidth(1, 220);
  var settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() < 2) settings.getRange(2, 1, 2, 2).setValues([['day', '1'], ['code', '"GOBBLE26"']]);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('HOST_PIN')) props.setProperty('HOST_PIN', '5062');
  createEngine(makeDb(), {}).seed();   // loads the 15 preloaded games into the Games tab (only once)
  var def = ss.getSheetByName('Sheet1'); if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  Logger.log('Setup complete. Now deploy as a web app.');
}


/* ============ game engine (same code the demo uses) ============ */
/* The 15 preloaded games for The Great Gobbler Games 2026 (5 per day).
   Wednesday has no Pick the Winners; Thursday and Friday do. The engine loads these into the Games sheet
   the first time anything writes, so the host only has to tap "Release" each morning.
   Format of a trivia line:  Question | Correct answer | Wrong | Wrong | Wrong   (players see answers shuffled) */
var GG_GAMES = (function () {
  function Q(s) { return s.trim().split('\n').map(function (l) { var p = l.split('|').map(function (x) { return x.trim(); }); return { q: p[0], a: p.slice(1), c: 0 }; }); }
  function M(prefix, list) { return { matches: list.map(function (x, i) { return { id: prefix + (i + 1), a: x[0], b: x[1], w: '' }; }) }; }
  function QA(list) { return list.map(function (x) { return { q: x[0], a: x[1] }; }); }
  function G(id, day, type, title, emoji, blurb, cfg) { return { id: id, day: day, type: type, title: title, emoji: emoji, blurb: blurb, cfg: cfg }; }
  return [
    /* ---------------- WEDNESDAY (day 1) ---------------- */
    G('lib-w1', 1, 'race', 'Live Trivia Race', '📺', 'Shared screen. Race on your phone while the TV shows the action.', { questions: Q(
      'A male turkey is called a what?|Tom|Hen|Poult|Drake\n' +
      'Thanksgiving in the US falls on which Thursday of November?|The fourth|The first|The second|The last day\n' +
      "The red flap of skin under a turkey's chin is called a what?|Wattle|Crest|Gill|Comb\n" +
      'Which spice gives pumpkin pie its signature flavor?|Cinnamon|Basil|Cumin|Paprika\n' +
      'Which dog breed is the star of this tournament?|Dachshund|Beagle|Poodle|Corgi') }),
    G('lib-w2', 1, 'trivia', 'Turkey Day Trivia', '🦃', 'History, parades, and football.', { questions: Q(
      'Which president officially made Thanksgiving a national holiday in 1863?|Abraham Lincoln|George Washington|Teddy Roosevelt|Thomas Jefferson\n' +
      'What was the name of the ship that carried the Pilgrims to America in 1620?|Mayflower|Santa Maria|Nautilus|Endeavour\n' +
      'Which department store holds a famous Thanksgiving Day parade with giant balloons?|Macy\'s|Sears|Nordstrom|Target\n' +
      'The Pilgrims and the Wampanoag shared their famous harvest feast in which year?|1621|1492|1776|1700\n' +
      'Which two cities host the traditional NFL Thanksgiving games every year?|Detroit and Dallas|Chicago and Green Bay|New York and Miami|Denver and Seattle') }),
    G('lib-w3', 1, 'trivia', 'Pie & Plate', '🥧', 'Name that side dish.', { questions: Q(
      'Which pie is the most popular at Thanksgiving tables?|Pumpkin|Cherry|Lemon meringue|Key lime\n' +
      'Cranberries grow on what?|Low vines in bogs|Tall trees|Underground like potatoes|Cactus plants\n' +
      'Green bean casserole is topped with crispy what?|Fried onions|Bacon bits|Croutons|Peanuts\n' +
      'What fluffy white candy tops many sweet potato casseroles?|Marshmallows|Cotton candy|Meringue|Whipped butter\n' +
      'Which pantry staple is most often used to thicken gravy?|Flour|Sugar|Baking soda|Salt') }),
    G('lib-w4', 1, 'twenty', '20 Questions', '🕵️', 'Find the mystery dish before the clock runs out.', {
      secret: 'Pumpkin Pie', seconds: 150,
      choices: ['Pumpkin Pie', 'Apple Pie', 'Pecan Pie', 'Sweet Potato Casserole', 'Cranberry Sauce', 'Mashed Potatoes', 'Green Bean Casserole', 'Cornbread Stuffing'],
      qa: QA([['Is it a dessert?', 'Yes'], ['Is it served cold or at room temperature?', 'Yes'], ['Is it round?', 'Yes'], ['Does it have a crust?', 'Yes'], ['Is it baked in the oven?', 'Yes'], ['Is it orange?', 'Yes'], ['Does it contain nuts?', 'No'], ['Does it contain apples?', 'No'], ['Is it a savory dish?', 'No'], ['Is it a drink?', 'No'], ['Is it often topped with whipped cream?', 'Yes'], ['Is it made from squash or pumpkin?', 'Yes'], ['Is it served before the main meal?', 'No'], ['Is it a side dish?', 'No']]) }),
    G('lib-w5', 1, 'scramble', 'Scramble Dash', '🔤', 'Unscramble five holiday words in 60 seconds.', { seconds: 60, words: [
      { w: 'pumpkin', hint: 'Orange and round, goes in a pie' }, { w: 'pilgrim', hint: 'Sailed on the Mayflower' }, { w: 'cranberry', hint: 'Tart red berry sauce' },
      { w: 'gravy', hint: 'Poured over the potatoes' }, { w: 'harvest', hint: 'What the Pilgrims celebrated' }] }),

    /* ---------------- THURSDAY (day 2) ---------------- */
    G('lib-t1', 2, 'vote', 'Most Likely To', '🙋', 'Shared screen. Vote for the family member who fits.', { prompts: [
      'Most likely to nap before pie', 'Most likely to burn the rolls', 'Most likely to win game night', 'Most likely to take seconds on dessert',
      'Most likely to talk during the football game', 'Most likely to start the leftovers sandwich first'] }),
    G('lib-t2', 2, 'trivia', 'Fun Facts & Pop Culture', '📺', 'Movies, cartoons, and more.', { questions: Q(
      "In the 1987 comedy 'Planes, Trains and Automobiles,' Steve Martin is trying to get home for which holiday?|Thanksgiving|Christmas|Easter|Fourth of July\n" +
      'On Friends, Joey famously puts what on his head at Thanksgiving?|A turkey|A pie|A pumpkin|A pot\n' +
      'In the Peanuts Thanksgiving special, Snoopy serves toast, popcorn, pretzel sticks and what candy?|Jelly beans|Gumdrops|Candy corn|Licorice\n' +
      'Which bird did Benjamin Franklin reportedly wish was the national bird instead of the bald eagle?|The turkey|The owl|The robin|The goose\n' +
      'In which state is Plymouth Rock?|Massachusetts|Virginia|Maine|New York') }),
    G('lib-t3', 2, 'trivia', 'Football & Parade', '🏈', 'Touchdowns and giant balloons.', { questions: Q(
      'Which NFL team hosts a Thanksgiving Day game in Detroit every year?|Lions|Bears|Packers|Vikings\n' +
      "Which classic cartoon cat was among the first giant balloons in the Macy's parade, in 1927?|Felix the Cat|Tom|Garfield|Sylvester\n" +
      'Which team has hosted a Thanksgiving Day game almost every year since 1966?|Dallas Cowboys|New York Giants|Miami Dolphins|Denver Broncos\n' +
      'How many points is a touchdown worth before the extra point?|6|7|3|5\n' +
      'What is the Super Bowl trophy called?|Vince Lombardi Trophy|Heisman Trophy|Stanley Cup|Commissioner\'s Cup') }),
    G('lib-t4', 2, 'warm', 'Word Warm-Up', '🔥', 'Guess the secret word. Hot guesses are close.', {
      secret: 'turkey', seconds: 150,
      words: 'bird gobble chicken thanksgiving feast dinner roast stuffing gravy drumstick feathers wattle poultry duck goose oven carve leftovers wing meat cranberry pilgrim november holiday family table plate potatoes pie harvest farm barn egg nest fly wild hunt tom brine fryer sandwich soup bacon ham beef pork fish salad bread corn autumn fall'.split(' ') }),
    G('lib-t5', 2, 'picks', 'Pick the Winners: Thursday', '🏆', 'Thanksgiving Day football and college hoops. Pick all seven!', M('t', [
      ['Bears', 'Lions'], ['Eagles', 'Cowboys'], ['Bills', 'Chiefs'], ['TCU', 'Texas Tech'],
      ['Michigan State', 'Arkansas'], ['Texas', 'Georgetown'], ['UCLA', "Saint Mary's"]])),

    /* ---------------- FRIDAY (day 3) ---------------- */
    G('lib-f1', 3, 'race', 'Grand Finale Race', '🏁', 'Shared screen. The last race of the tournament.', { questions: Q(
      'What is a baby turkey called?|Poult|Cub|Kit|Chick\n' +
      'How many players does each team have on the field in American football?|11|9|12|10\n' +
      'In Canada, Thanksgiving is celebrated in which month?|October|November|December|September\n' +
      'How many points is a field goal worth?|3|2|6|1\n' +
      'After dinner, two people pull on the wishbone. What do they do?|Make a wish|Count to ten|Sing a song|Bury it') }),
    G('lib-f2', 3, 'trivia', 'Leftovers & Black Friday', '🥪', 'The day after the feast.', { questions: Q(
      'Black Friday falls on the day after Thanksgiving. What day of the week is it?|Friday|Saturday|Monday|Thursday\n' +
      'Which compound in turkey gets the blame for making people sleepy?|Tryptophan|Caffeine|Sodium|Gluten\n' +
      'About how many days are leftovers safe in the fridge, according to food safety guidance?|3 to 4 days|10 days|2 weeks|1 hour\n' +
      'On Friends, what does Ross call his Thanksgiving leftover sandwich?|The Moist Maker|The Big Gobbler|The Turkey Tower|The Day After\n' +
      'What do many people make from a leftover turkey carcass?|Soup stock|Cranberry sauce|Pumpkin pie|Dinner rolls') }),
    G('lib-f3', 3, 'trivia', 'Wiener Dog Trivia', '🌭', 'Everything about our favorite long dogs.', { questions: Q(
      'What nickname do people give dachshunds?|Wiener dog|Mop dog|Torpedo dog|Rope dog\n' +
      'A dachshund named Waldi was the mascot of which Olympic Games?|1972 Munich|1996 Atlanta|1984 Los Angeles|2000 Sydney\n' +
      'Dachshunds were originally bred to hunt which animal?|Badgers|Penguins|Lions|Moose\n' +
      'How many coat types do dachshunds come in?|3|2|5|7\n' +
      'The Oscar Mayer Wienermobile is shaped like a giant what?|Hot dog|Pickle|Dachshund|Corn dog') }),
    G('lib-f4', 3, 'twenty', '20 Questions: Side Dish', '🕵️', 'Find the mystery side dish before the clock runs out.', {
      secret: 'Stuffing', seconds: 150,
      choices: ['Stuffing', 'Mashed Potatoes', 'Cranberry Sauce', 'Dinner Rolls', 'Gravy', 'Green Bean Casserole', 'Pumpkin Pie', 'Apple Pie'],
      qa: QA([['Is it a dessert?', 'No'], ['Is it a side dish?', 'Yes'], ['Is it served hot?', 'Yes'], ['Is it baked?', 'Yes'], ['Is it made with bread?', 'Yes'], ['Is it usually golden brown?', 'Yes'], ['Is it a liquid?', 'No'], ['Does it contain herbs like sage?', 'Yes'], ['Is it made from potatoes?', 'No'], ['Is it red?', 'No'], ['Is it poured over other food?', 'No'], ['Is it sweet?', 'No'], ['Can it be cooked inside the turkey?', 'Yes'], ['Does it usually contain celery or onion?', 'Yes'], ['Is it served cold?', 'No']]) }),
    G('lib-f5', 3, 'picks', 'Pick the Winners: Friday', '🏆', 'Black Friday hoops, football and hockey. Pick all eight!', M('f', [
      ['Virginia', 'Xavier'], ['Kansas', 'Villanova'], ['Houston', 'Oakland'], ['Broncos', 'Steelers'],
      ['NC State (W)', 'Kansas State (W)'], ['LSU (W)', 'Villanova (W)'], ['Vegas Golden Knights', 'San Jose Sharks'], ['Dallas Stars', 'St. Louis Blues']]))
  ];
})();

/* Great Gobbler Games engine.
   All rules, scoring and chat live here. The same code runs in the Google Apps Script backend
   (reading and writing the Google Sheet) and in the in-browser demo (in-memory tables).
   db interface: read(table) -> array of row objects (fresh copy), add(table,row), save(table,rows). */
function createEngine(db, opt) {
  opt = opt || {};
  var now = function () { return Date.now(); };
  var rid = function (p) { return p + Math.random().toString(36).slice(2, 8) + (now() % 1e6).toString(36); };
  var err = function (m) { return { ok: false, error: m }; };
  var shuf = function (a) { a = a.slice(); for (var i = a.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = a[i]; a[i] = a[j]; a[j] = t; } return a; };
  var pinOk = function (p) { return String(p) === String((opt.pin && opt.pin()) || '5062'); };
  var clean = function (s, n) { return String(s == null ? '' : s).replace(/[\u0000-\u001f<>]/g, ' ').trim().slice(0, n || 200); };

  /* ---------- settings ---------- */
  function setting(k, def) {
    var r = db.read('settings').filter(function (x) { return x.key === k; })[0];
    if (!r || r.value === '' || r.value == null) return def;
    try { return JSON.parse(r.value); } catch (e) { return r.value; }
  }
  function setSetting(k, v) {
    var rows = db.read('settings'), s = JSON.stringify(v), r = rows.filter(function (x) { return x.key === k; })[0];
    if (r) r.value = s; else rows.push({ key: k, value: s });
    db.save('settings', rows);
  }
  var curDay = function () { return Number(setting('day', 1)) || 1; };

  /* ---------- scores (the ledger is the single source of truth) ---------- */
  function board() {
    var t = {};
    db.read('ledger').forEach(function (r) { t[r.pid] = (t[r.pid] || 0) + (Number(r.pts) || 0); });
    return db.read('players').map(function (p, i) { return { id: p.id, name: p.name, av: Number(p.av) || 0, score: t[p.id] || 0, o: i }; });
  }
  var ranked = function (b) { return b.slice().sort(function (x, y) { return y.score - x.score || x.o - y.o; }); };
  var rankIn = function (b, pid) { return ranked(b).map(function (p) { return p.id; }).indexOf(pid) + 1; };
  function addPts(pid, gid, pts, note) {
    pts = Math.round(pts); if (!pts) return;
    var before = rankIn(board(), pid);
    db.add('ledger', { ts: now(), pid: pid, gid: gid, pts: pts, note: note || '' });
    var b = board(), after = rankIn(b, pid);
    if (after > 0 && after <= 3 && before > 3) {
      var p = b.filter(function (x) { return x.id === pid; })[0];
      sys(p.name + ' just landed at #' + after + ' on the leaderboard! 🦃');
    }
  }

  /* ---------- chat ---------- */
  function post(row) {
    var c = db.read('chat'), last = c.length ? Number(c[c.length - 1].ts) : 0;
    var ts = Math.max(now(), last + 1);
    db.add('chat', Object.assign({ id: rid('c'), pid: '', name: '', av: 0, text: '', kind: 'user' }, row, { ts: ts }));
  }
  function sys(text) { post({ kind: 'system', name: 'Gobbler HQ', text: text }); }

  /* ---------- lookups ---------- */
  var gameOf = function (id) { return db.read('games').filter(function (g) { return g.id === id; })[0]; };
  var meOf = function (pid) { return db.read('players').filter(function (p) { return p.id === pid; })[0]; };
  var playOf = function (pid, gid) { return db.read('plays').filter(function (p) { return p.pid === pid && p.gid === gid; })[0]; };
  function putPlay(p) {
    var rows = db.read('plays'), i = -1;
    rows.forEach(function (r, k) { if (r.pid === p.pid && r.gid === p.gid) i = k; });
    if (i < 0) rows.push(p); else rows[i] = p;
    db.save('plays', rows);
  }
  var SOLO = ['trivia', 'twenty', 'scramble', 'warm'];

  /* ---------- public state ---------- */
  function livePub(pid) {
    var L = setting('live', null); if (!L) return null;
    var g = gameOf(L.gid); if (!g) return null;
    var c = g.cfg || {};
    var o = { gid: L.gid, kind: L.kind, title: g.title, emoji: g.emoji, qi: L.qi, phase: L.phase, openedAt: L.openedAt,
      total: L.kind === 'race' ? (c.questions || []).length : (c.prompts || []).length,
      answered: Object.keys(L.answers || {}).length, mine: (L.answers && L.answers[pid]) ? L.answers[pid].ch : null, last: L.last || null };
    if (L.qi >= 0) {
      if (L.kind === 'race') { var q = c.questions[L.qi]; o.q = q.q; o.choices = L.perm.map(function (i) { return q.a[i]; }); }
      else o.q = c.prompts[L.qi];
    }
    return o;
  }
  function state(pid, host) {
    var day = curDay(), plays = db.read('plays').filter(function (p) { return p.pid === pid; });
    var gl = db.read('games').filter(function (g) { return host || Number(g.day) <= day; }).map(function (g) {
      var p = plays.filter(function (x) { return x.gid === g.id; })[0];
      var o = { id: g.id, day: Number(g.day), type: g.type, title: g.title, emoji: g.emoji, blurb: g.blurb, status: p ? p.status : 'new', pts: p ? Number(p.pts) || 0 : 0 };
      if (g.type === 'picks') o.n = ((g.cfg && g.cfg.matches) || []).length;
      if (host) o.cfg = g.cfg;
      return o;
    });
    var out = { ok: true, day: day, me: meOf(pid) || null, players: board().map(function (p) { return { id: p.id, name: p.name, av: p.av, score: p.score }; }), games: gl, live: livePub(pid), now: now() };
    if (host) { out.code = setting('code', 'GOBBLE26'); out.plays = db.read('plays').map(function (p) { return { pid: p.pid, gid: p.gid, status: p.status, pts: p.pts }; }); }
    return out;
  }
  function chatSince(since) {
    var c = db.read('chat');
    c = Number(since) > 0 ? c.filter(function (x) { return Number(x.ts) > Number(since); }) : c.slice(-60);
    return c.map(function (x) { return { id: x.id, ts: Number(x.ts), pid: x.pid, name: x.name, av: Number(x.av) || 0, text: x.text, kind: x.kind }; });
  }

  /* ---------- game helpers ---------- */
  function finishPlay(p, g, pts, note) { p.status = 'done'; p.pts = pts; putPlay(p); addPts(p.pid, g.id, pts, note); }
  function expired(p, g) { return now() - Number(p.ts) > ((g.cfg.seconds || 60) * 1000 + 5000); }
  function timeoutFinish(p, g) {
    if (g.type === 'scramble') {
      var st = p.st; finishPlay(p, g, (st.ok || 0) * 50, g.title); return { expired: true, pts: p.pts };
    }
    finishPlay(p, g, 20, g.title); return { expired: true, pts: 20 };
  }

  /* ---------- player actions ---------- */
  var P = {
    state: function (r) { return state(r.pid, false); },
    sync: function (r) { return { ok: true, state: state(r.pid, false), chat: chatSince(r.since) }; },
    tv: function () {
      var b = ranked(board()).map(function (p) { return { id: p.id, name: p.name, av: p.av, score: p.score }; });
      return { ok: true, day: curDay(), players: b, live: livePub(''), code: setting('code', 'GOBBLE26') };
    },
    join: function (r) {
      var code = String(setting('code', 'GOBBLE26')).toUpperCase();
      if (code && clean(r.code, 20).toUpperCase() !== code) return err('That family code is not right.');
      var name = clean(r.name, 14); if (!name) return err('Enter a screen name.');
      var ps = db.read('players');
      if (ps.some(function (p) { return p.name.toLowerCase() === name.toLowerCase(); })) return err('That screen name is taken. Try another.');
      var id = rid('p');
      db.add('players', { id: id, name: name, av: Number(r.av) || 0, joined: now() });
      sys(name + ' joined the games! 🐶');
      return { ok: true, pid: id };
    },
    profile: function (r) {
      var rows = db.read('players'), p = rows.filter(function (x) { return x.id === r.pid; })[0]; if (!p) return err('Join first.');
      var name = clean(r.name, 14) || p.name;
      if (rows.some(function (x) { return x.id !== p.id && x.name.toLowerCase() === name.toLowerCase(); })) return err('That screen name is taken.');
      p.name = name; p.av = Number(r.av) || 0; db.save('players', rows); return { ok: true };
    },
    start: function (r) {
      var g = gameOf(r.gid); if (!g) return err('Game not found.');
      if (!meOf(r.pid)) return err('Join first.');
      if (Number(g.day) > curDay()) return err('This game is not open yet.');
      if (SOLO.indexOf(g.type) < 0) return err('This is not a solo game.');
      if (playOf(r.pid, g.id)) return err('You already played this one. Ask the host to allow a replay.');
      var c = g.cfg, st = {}, out = { ok: true, type: g.type, title: g.title, emoji: g.emoji };
      if (g.type === 'trivia') {
        st.perm = c.questions.map(function (q) { return shuf(q.a.map(function (_, i) { return i; })); });
        out.questions = c.questions.map(function (q, i) { return { q: q.q, a: st.perm[i].map(function (k) { return q.a[k]; }) }; });
        out.seconds = 15;
      } else if (g.type === 'twenty') {
        st = { asked: [], wrong: 0, tried: [] };
        out.questions = c.qa.map(function (x) { return x.q; }); out.choices = shuf(c.choices); out.seconds = c.seconds || 150;
      } else if (g.type === 'scramble') {
        st = { i: 0, ok: 0, words: c.words.map(function (w) {
          var a = w.w.split(''), s = w.w, n = 0; while (s === w.w && n++ < 30) { a = shuf(a); s = a.join(''); } return { scr: s.toUpperCase(), hint: w.hint || '' }; }) };
        out.words = st.words; out.seconds = c.seconds || 60;
      } else if (g.type === 'warm') {
        st = { guesses: [] }; out.seconds = c.seconds || 150;
      }
      db.add('plays', { pid: r.pid, gid: g.id, ts: now(), status: 'started', pts: 0, st: st });
      return out;
    },
    answer: function (r) {
      var g = gameOf(r.gid), p = playOf(r.pid, r.gid); if (!g || !p || p.status !== 'started' || g.type !== 'trivia') return err('No game in progress.');
      var st = p.st, qi = Number(r.qi), q = g.cfg.questions[qi]; if (!q) return err('Bad question.');
      st.ans = st.ans || {}; if (st.ans[qi] != null) return err('Already answered.');
      var choice = Number(r.choice), orig = choice >= 0 ? st.perm[qi][choice] : -1, ok = orig === 0;
      var tl = Math.max(0, Math.min(15, Number(r.tleft) || 0)), pts = ok ? 100 + Math.round(tl / 15 * 50) : 0;
      st.ans[qi] = choice; p.pts = (Number(p.pts) || 0) + pts;
      var done = Object.keys(st.ans).length >= g.cfg.questions.length; if (done) p.status = 'done';
      putPlay(p); if (pts) addPts(r.pid, g.id, pts, g.title + ' Q' + (qi + 1));
      return { ok: true, correct: ok, c: st.perm[qi].indexOf(0), points: pts, done: done, total: p.pts };
    },
    twentyAsk: function (r) {
      var g = gameOf(r.gid), p = playOf(r.pid, r.gid); if (!g || !p || p.status !== 'started' || g.type !== 'twenty') return err('No game in progress.');
      if (expired(p, g)) return Object.assign({ ok: true }, timeoutFinish(p, g));
      var st = p.st, i = Number(r.i), x = g.cfg.qa[i];
      if (!x || st.asked.indexOf(i) >= 0 || st.asked.length >= 20) return err('Cannot ask that.');
      st.asked.push(i); putPlay(p); return { ok: true, a: x.a };
    },
    twentyGuess: function (r) {
      var g = gameOf(r.gid), p = playOf(r.pid, r.gid); if (!g || !p || p.status !== 'started' || g.type !== 'twenty') return err('No game in progress.');
      if (expired(p, g)) return Object.assign({ ok: true }, timeoutFinish(p, g));
      var st = p.st;
      if (String(r.g).toLowerCase() === String(g.cfg.secret).toLowerCase()) {
        var el = Math.floor((now() - Number(p.ts)) / 1000);
        var pts = Math.max(30, 300 - 10 * st.asked.length - 40 * st.wrong - Math.floor(el / 3));
        finishPlay(p, g, pts, g.title); return { ok: true, correct: true, pts: pts, secret: g.cfg.secret };
      }
      st.wrong++; st.tried.push(r.g); putPlay(p); return { ok: true, correct: false };
    },
    scrambleCheck: function (r) {
      var g = gameOf(r.gid), p = playOf(r.pid, r.gid); if (!g || !p || p.status !== 'started' || g.type !== 'scramble') return err('No game in progress.');
      if (expired(p, g)) return Object.assign({ ok: true }, timeoutFinish(p, g));
      var st = p.st, w = g.cfg.words[st.i], ok = false;
      if (r.skip) st.i++;
      else if (String(r.guess).trim().toLowerCase() === String(w.w).toLowerCase()) { st.ok++; st.i++; ok = true; }
      var done = st.i >= g.cfg.words.length, pts = 0;
      if (done) {
        var left = Math.max(0, (g.cfg.seconds || 60) - Math.floor((now() - Number(p.ts)) / 1000));
        pts = st.ok * 50 + (st.ok ? left : 0); finishPlay(p, g, pts, g.title);
      } else putPlay(p);
      return { ok: true, correct: ok, i: st.i, okCount: st.ok, done: done, pts: pts };
    },
    warmGuess: function (r) {
      var g = gameOf(r.gid), p = playOf(r.pid, r.gid); if (!g || !p || p.status !== 'started' || g.type !== 'warm') return err('No game in progress.');
      if (expired(p, g)) return Object.assign({ ok: true }, timeoutFinish(p, g));
      var w = clean(r.w, 30).toLowerCase(); if (!w) return err('Type a word.');
      var st = p.st; if (st.guesses.indexOf(w) >= 0) return { ok: true, dup: true };
      var secret = String(g.cfg.secret).toLowerCase(), list = (g.cfg.words || []).map(function (x) { return String(x).toLowerCase(); });
      var idx = list.indexOf(w), rank = w === secret ? 1 : (idx >= 0 ? idx + 2 : 500);
      st.guesses.push(w);
      if (rank === 1) {
        var el = Math.floor((now() - Number(p.ts)) / 1000);
        var pts = Math.max(60, 400 - st.guesses.length * 15 - Math.floor(el / 3));
        finishPlay(p, g, pts, g.title); return { ok: true, rank: 1, win: true, pts: pts, secret: g.cfg.secret };
      }
      putPlay(p); return { ok: true, rank: rank };
    },
    finish: function (r) {
      var g = gameOf(r.gid), p = playOf(r.pid, r.gid); if (!g || !p) return err('No game.');
      if (p.status === 'done') return { ok: true, pts: p.pts, secret: g.cfg.secret };
      if (g.type === 'trivia') { p.status = 'done'; putPlay(p); return { ok: true, pts: p.pts }; }
      if (now() - Number(p.ts) < ((g.cfg.seconds || 60) - 3) * 1000) return err('The clock is still running.');
      var o = timeoutFinish(p, g); o.secret = g.cfg.secret; return Object.assign({ ok: true }, o);
    },
    picks: function (r) {
      var g = gameOf(r.gid); if (!g || g.type !== 'picks') return err('Game not found.');
      if (Number(g.day) > curDay()) return err('This game is not open yet.');
      var p = playOf(r.pid, g.id);
      return { ok: true, title: g.title, matches: g.cfg.matches, mine: p ? p.st.choices : null, locked: !!p };
    },
    lockPicks: function (r) {
      var g = gameOf(r.gid); if (!g || g.type !== 'picks') return err('Game not found.');
      if (!meOf(r.pid)) return err('Join first.');
      if (Number(g.day) > curDay()) return err('This game is not open yet.');
      var ms = g.cfg.matches || []; if (ms.length < 3) return err('The host is still adding matchups.');
      if (playOf(r.pid, g.id)) return err('Your picks are already locked.');
      var ch = {}; for (var i = 0; i < ms.length; i++) { var c = r.choices && r.choices[ms[i].id]; if (c !== 'a' && c !== 'b') return err('Pick every game.'); ch[ms[i].id] = c; }
      db.add('plays', { pid: r.pid, gid: g.id, ts: now(), status: 'done', pts: 0, st: { choices: ch } });
      return { ok: true };
    },
    liveAnswer: function (r) {
      var L = setting('live', null); if (!L || L.phase !== 'open' || Number(r.qi) !== L.qi) return err('That round is closed.');
      if (!meOf(r.pid)) return err('Join first.');
      L.answers = L.answers || {}; if (L.answers[r.pid]) return err('Already locked in.');
      L.answers[r.pid] = { ch: r.ch, t: now() - L.openedAt }; setSetting('live', L); return { ok: true };
    },
    chat: function (r) {
      var me = meOf(r.pid); if (!me) return err('Join first.');
      var text = clean(r.text, 200); if (!text) return err('Type a message.');
      post({ pid: me.id, name: me.name, av: Number(me.av) || 0, text: text, kind: 'user' }); return { ok: true };
    }
  };

  /* ---------- host actions (every call carries the PIN) ---------- */
  function lines(s) { return String(s || '').split(/\r?\n/).map(function (x) { return x.trim(); }).filter(Boolean); }
  function buildCfg(type, c) {
    c = c || {};
    if (type === 'trivia' || type === 'race') {
      var qs = (c.questions || []).map(function (q) { return { q: clean(q.q, 200), a: (q.a || []).map(function (a) { return clean(a, 80); }), c: 0 }; });
      if (!qs.length) return 'Add at least one question.';
      for (var i = 0; i < qs.length; i++) if (!qs[i].q || qs[i].a.length !== 4 || qs[i].a.some(function (a) { return !a; })) return 'Question ' + (i + 1) + ' needs a question and 4 answers.';
      return { questions: qs };
    }
    if (type === 'twenty') {
      var secret = clean(c.secret, 40), ch = (c.choices || []).map(function (x) { return clean(x, 40); }).filter(Boolean);
      if (ch.map(function (x) { return x.toLowerCase(); }).indexOf(secret.toLowerCase()) < 0) ch.push(secret);
      var qa = (c.qa || []).map(function (x) { return { q: clean(x.q, 120), a: x.a === 'Yes' ? 'Yes' : 'No' }; }).filter(function (x) { return x.q; });
      if (!secret || ch.length < 3 || qa.length < 3) return 'Add the secret, at least 3 choices, and at least 3 questions.';
      return { secret: secret, choices: ch, qa: qa, seconds: Number(c.seconds) || 150 };
    }
    if (type === 'scramble') {
      var ws = (c.words || []).map(function (x) { return { w: clean(x.w, 20).toLowerCase().replace(/[^a-z]/g, ''), hint: clean(x.hint, 80) }; }).filter(function (x) { return x.w.length > 2; });
      if (ws.length < 3) return 'Add at least 3 words.';
      return { words: ws, seconds: Number(c.seconds) || 60 };
    }
    if (type === 'warm') {
      var sec = clean(c.secret, 30).toLowerCase(), wl = (c.words || []).map(function (x) { return clean(x, 30).toLowerCase(); }).filter(function (x) { return x && x !== sec; });
      if (!sec || wl.length < 5) return 'Add the secret word and at least 5 related words, closest first.';
      return { secret: sec, words: wl, seconds: Number(c.seconds) || 150 };
    }
    if (type === 'picks') {
      var ms = (c.matches || []).map(function (m, i) { return { id: m.id || rid('m'), a: clean(m.a, 24), b: clean(m.b, 24), w: m.w === 'a' || m.w === 'b' ? m.w : '' }; }).filter(function (m) { return m.a && m.b; });
      if (ms.length < 3 || ms.length > 8) return 'Pick the Winners needs 3 to 8 matchups.';
      return { matches: ms };
    }
    if (type === 'vote') {
      var ps = (c.prompts || []).map(function (x) { return clean(x, 140); }).filter(Boolean);
      if (!ps.length) return 'Add at least one prompt.'; return { prompts: ps };
    }
    return 'Unknown game type.';
  }
  function clearPlays(gid, pid) {
    db.save('plays', db.read('plays').filter(function (p) { return !(p.gid === gid && (!pid || p.pid === pid)); }));
    db.save('ledger', db.read('ledger').filter(function (l) { return !(l.gid === gid && (!pid || l.pid === pid) && l.note !== 'bonus'); }));
  }
  function liveStep(L) { setSetting('live', L); }
  var H = {
    hostLogin: function () { return { ok: true }; },
    hState: function (r) { return state(r.pid || '', true); },
    hSetDay: function (r) {
      var d = Math.max(1, Math.min(9, Number(r.day) || 1)); setSetting('day', d);
      var n = db.read('games').filter(function (g) { return Number(g.day) === d; }).length;
      sys('Day ' + d + ' is live! ' + n + ' game' + (n === 1 ? '' : 's') + ' just dropped. 🏆'); return { ok: true };
    },
    hNextDay: function () {
      var d = curDay(); if (d >= 3) return err('All three days are already open.');
      setSetting('day', d + 1);
      var n = db.read('games').filter(function (g) { return Number(g.day) === d + 1; }).length;
      sys('Day ' + (d + 1) + ' is live! ' + n + ' new game' + (n === 1 ? '' : 's') + ' just dropped. 🏆'); return { ok: true, day: d + 1 };
    },
    hRestoreGames: function () { var n = loadLibrary(true); return { ok: true, added: n }; },
    hSetCode: function (r) { var c = clean(r.code, 20).toUpperCase(); if (!c) return err('Enter a code.'); setSetting('code', c); return { ok: true }; },
    hAddGame: function (r) {
      var g = r.game || {}, cfg = buildCfg(g.type, g.cfg); if (typeof cfg === 'string') return err(cfg);
      var title = clean(g.title, 40); if (!title) return err('Give the game a title.');
      var id = rid('g');
      db.add('games', { id: id, day: Math.max(1, Math.min(9, Number(g.day) || 1)), type: g.type, title: title, emoji: clean(g.emoji, 4) || '🎲', blurb: clean(g.blurb, 80), cfg: cfg });
      return { ok: true, id: id };
    },
    hUpdateMatches: function (r) {
      var rows = db.read('games'), g = rows.filter(function (x) { return x.id === r.gid; })[0]; if (!g || g.type !== 'picks') return err('Game not found.');
      var old = {}; g.cfg.matches.forEach(function (m) { old[m.id] = m; });
      var cfg = buildCfg('picks', { matches: (r.matches || []).map(function (m) { var o = old[m.id]; return { id: m.id, a: m.a, b: m.b, w: o ? o.w : '' }; }) });
      if (typeof cfg === 'string') return err(cfg);
      if (db.read('plays').some(function (p) { return p.gid === g.id; })) { var ids = cfg.matches.map(function (m) { return m.id; }); if (g.cfg.matches.some(function (m) { return ids.indexOf(m.id) < 0; })) return err('Someone has locked picks. Allow a replay before removing a matchup.'); }
      g.cfg = cfg; db.save('games', rows); return { ok: true };
    },
    hDelGame: function (r) {
      clearPlays(r.gid); db.save('games', db.read('games').filter(function (g) { return g.id !== r.gid; })); return { ok: true };
    },
    hSetResult: function (r) {
      var rows = db.read('games'), g = rows.filter(function (x) { return x.id === r.gid; })[0]; if (!g) return err('Game not found.');
      var m = g.cfg.matches.filter(function (x) { return x.id === r.mid; })[0]; if (!m) return err('Matchup not found.');
      if (r.w !== 'a' && r.w !== 'b') return err('Pick a winner.');
      if (m.w) return err('That result is already final.');
      m.w = r.w; db.save('games', rows);
      var led = db.read('ledger'), won = 0;
      db.read('plays').filter(function (p) { return p.gid === g.id && p.st && p.st.choices; }).forEach(function (p) {
        if (p.st.choices[m.id] === m.w) {
          var note = 'pick:' + m.id;
          if (!led.some(function (l) { return l.pid === p.pid && l.note === note; })) { addPts(p.pid, g.id, 50, note); won++; var pp = playOf(p.pid, g.id); pp.pts = (Number(pp.pts) || 0) + 50; putPlay(pp); }
        }
      });
      sys((m.w === 'a' ? m.a : m.b) + ' won over ' + (m.w === 'a' ? m.b : m.a) + '. ' + won + ' right pick' + (won === 1 ? '' : 's') + '! 🏆');
      return { ok: true, won: won };
    },
    hReplay: function (r) { clearPlays(r.gid, r.pid || ''); return { ok: true }; },
    hBonus: function (r) { if (!meOf(r.pid)) return err('Pick a player.'); addPts(r.pid, 'bonus', Number(r.pts) || 0, 'bonus'); return { ok: true }; },
    hRemovePlayer: function (r) {
      db.save('players', db.read('players').filter(function (p) { return p.id !== r.pid; }));
      db.save('ledger', db.read('ledger').filter(function (l) { return l.pid !== r.pid; }));
      db.save('plays', db.read('plays').filter(function (l) { return l.pid !== r.pid; })); return { ok: true };
    },
    hChatDel: function (r) { db.save('chat', db.read('chat').filter(function (c) { return c.id !== r.id; })); return { ok: true }; },
    hReset: function () {
      db.save('ledger', []); db.save('plays', []); db.save('chat', []); setSetting('live', null);
      var rows = db.read('games'); rows.forEach(function (g) { if (g.type === 'picks') g.cfg.matches.forEach(function (m) { m.w = ''; }); }); db.save('games', rows);
      setSetting('day', 1); return { ok: true };
    },
    /* shared-screen games */
    hLiveStart: function (r) {
      var g = gameOf(r.gid); if (!g || (g.type !== 'race' && g.type !== 'vote')) return err('Not a shared-screen game.');
      liveStep({ gid: g.id, kind: g.type, qi: -1, phase: 'lobby', openedAt: 0, answers: {}, last: null, perm: [0, 1, 2, 3] });
      sys('📺 ' + g.title + ' is starting on the big screen! Open the app and tap the live banner.'); return { ok: true };
    },
    hLiveNext: function () {
      var L = setting('live', null); if (!L) return err('Nothing is live.');
      var g = gameOf(L.gid), total = L.kind === 'race' ? g.cfg.questions.length : g.cfg.prompts.length;
      if (L.qi + 1 >= total) { setSetting('live', null); sys('That is a wrap on ' + g.title + '! Check the leaderboard. 🏆'); return { ok: true, ended: true }; }
      L.qi++; L.phase = 'open'; L.openedAt = now(); L.answers = {}; L.last = null; L.perm = L.kind === 'race' ? shuf([0, 1, 2, 3]) : L.perm;
      liveStep(L); return { ok: true };
    },
    hLiveReveal: function () {
      var L = setting('live', null); if (!L || L.phase !== 'open') return err('Open a question first.');
      var g = gameOf(L.gid), A = L.answers || {}, ids = Object.keys(A), last = { qi: L.qi };
      if (L.kind === 'race') {
        var c = L.perm.indexOf(0), counts = [0, 0, 0, 0];
        ids.forEach(function (pid) {
          var a = A[pid]; if (a.ch >= 0 && a.ch < 4) counts[a.ch]++;
          if (a.ch === c) addPts(pid, g.id, 100 + Math.max(0, Math.round(50 * (1 - a.t / 15000))), g.title + ' Q' + (L.qi + 1));
        });
        last.c = c; last.counts = counts;
      } else {
        var tally = {}; ids.forEach(function (pid) { var t = A[pid].ch; tally[t] = (tally[t] || 0) + 1; });
        var max = 0; Object.keys(tally).forEach(function (k) { if (tally[k] > max) max = tally[k]; });
        var winners = Object.keys(tally).filter(function (k) { return tally[k] === max && max > 0; });
        ids.forEach(function (pid) { if (winners.indexOf(A[pid].ch) >= 0) addPts(pid, g.id, 50, g.title + ' vote'); });
        last.tally = tally; last.winners = winners;
      }
      L.phase = 'reveal'; L.last = last; liveStep(L); return { ok: true };
    },
    hLiveEnd: function () { setSetting('live', null); return { ok: true }; }
  };

  /* The 15 preloaded games (library.js). Loaded once, the first time anything is written. */
  function loadLibrary(force) {
    var have = {}, added = 0;
    db.read('games').forEach(function (g) { have[g.id] = 1; });
    GG_GAMES.forEach(function (g) {
      if (have[g.id]) return;
      db.add('games', { id: g.id, day: g.day, type: g.type, title: g.title, emoji: g.emoji, blurb: g.blurb, cfg: JSON.parse(JSON.stringify(g.cfg)) }); added++;
    });
    setSetting('seeded', 1); return added;
  }
  function ensureSeed() { if (!setting('seeded', 0)) loadLibrary(); }
  var READ_ONLY = { sync: 1, state: 1, tv: 1, picks: 1 };

  function handle(req) {
    try {
      var a = req && req.a; if (!a) return err('No action.');
      if (!READ_ONLY[a] && (H[a] || P[a])) ensureSeed();
      if (H[a]) { if (!pinOk(req.pin)) return err('Wrong host PIN.'); return H[a](req); }
      if (P[a]) return P[a](req);
      return err('Unknown action.');
    } catch (e) { return err('Server error: ' + (e && e.message ? e.message : e)); }
  }
  return { handle: handle, seed: ensureSeed };
}
