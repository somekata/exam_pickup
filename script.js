// ===== DOM参照 =====
const fileInput = document.getElementById("fileInput");
const questionTableBody = document.querySelector("#questionTable tbody");
const selectedListEl = document.getElementById("selectedList");

const statCountEl = document.getElementById("statCount");
const statTotalEl = document.getElementById("statTotal");

const exportBtn = document.getElementById("exportDocx");

// フィルタUI
const filterKeyword = document.getElementById("filterKeyword");
const filterLang = document.getElementById("filterLang");
const filterID = document.getElementById("filterID");
const filterDifficulty = document.getElementById("filterDifficulty");
const filterDomain = document.getElementById("filterDomain");

// 出力オプションUI
const optTitleHeader = document.getElementById("optTitleHeader");
const optShowPageNumber = document.getElementById("optShowPageNumber");
const optPageMargin = document.getElementById("optPageMargin");

const optQuestionNumberStyle = document.getElementById("optQuestionNumberStyle");
const optChoiceLabelStyle = document.getElementById("optChoiceLabelStyle");

const optShowScore = document.getElementById("optShowScore");
const optScorePosition = document.getElementById("optScorePosition");

const optIncludeTitle = document.getElementById("optIncludeTitle");
const optIncludeChoices = document.getElementById("optIncludeChoices");
const optIncludeMeta = document.getElementById("optIncludeMeta");

const optShowAnswer = document.getElementById("optShowAnswer");
const optAnswerColor = document.getElementById("optAnswerColor");
const optAnswerBold = document.getElementById("optAnswerBold");
const optAnswerUnderline = document.getElementById("optAnswerUnderline");

const optFontFamilyQuestion = document.getElementById("optFontFamilyQuestion");
const optFontSizeQuestion = document.getElementById("optFontSizeQuestion");
const optFontFamilyChoices = document.getElementById("optFontFamilyChoices");
const optFontSizeChoices = document.getElementById("optFontSizeChoices");

// ===== データ保持 =====
let allQuestions = [];      // 読み込んだ全問題（統合後）
let filteredQuestions = []; // フィルタ後表示用
let selectedOrdered = [];   // 選択済み問題（並び順）

// ===== イベント登録 =====
fileInput.addEventListener("change", handleFiles);
exportBtn.addEventListener("click", async () => {
  const settings = collectSettingsFromUI();
  await exportDocx(settings);
});

[filterKeyword, filterLang, filterID, filterDifficulty, filterDomain].forEach(el => {
  el.addEventListener("input", () => {
    applyFilter();
    renderTable();
  });
});

// ===== CSV読み込み（複数対応＋重複削除オプション） =====
function handleFiles(event) {
  const files = Array.from(event.target.files || []);
  if (files.length === 0) return;

  let tempAll = [];
  let loaded = 0;

  files.forEach(file => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: function (results) {
        const parsed = results.data.map((row) => ({
          id: safe(row["question_id"]),
          language: safe(row["language"]),
          difficulty: safe(row["difficulty"]),
          domain1: safe(row["domain1"]),
          domain2: safe(row["domain2"]),

          case_id: safe(row["case_id"]),
          title_text: safe(row["title"]),
          case_text: safe(row["case_text"]),
          question_text: safe(row["question_text"]),

          choice_a: safe(row["choice_a"]),
          choice_b: safe(row["choice_b"]),
          choice_c: safe(row["choice_c"]),
          choice_d: safe(row["choice_d"]),
          choice_e: safe(row["choice_e"]),

          correct: safe(row["correct"]),

          selected: false,
          score: 1,
          _uid: `Q_${Math.random().toString(36).slice(2)}`
        }));

        tempAll = tempAll.concat(parsed);
        loaded++;

        if (loaded === files.length) {
          // 重複削除オプション（index.html側で用意済み）
          const removeDup = document.getElementById("optRemoveDuplicate")?.checked;
          if (removeDup) {
            const seen = new Set();
            tempAll = tempAll.filter(q => {
              if (!q.id) return true; // ID空欄は残す
              if (seen.has(q.id)) return false;
              seen.add(q.id);
              return true;
            });
          }

          allQuestions = tempAll;

          applyFilter();
          rebuildSelectedFromAll();
          renderTable();
          renderSelectedList();

          // 読み込んだ全体の集計を表示
          const mapAll = countByLangAndDifficulty(allQuestions);
          const tableSummaryEl = document.getElementById("tableSummary");
          if (tableSummaryEl) {
            tableSummaryEl.innerHTML = renderLangDiffSummary(mapAll);
          }
        }
      }
    });
  });
}

// ===== フィルタ処理 =====
function applyFilter() {
  const kw = filterKeyword.value.trim().toLowerCase();
  const lang = filterLang.value.trim().toLowerCase();
  const qid = filterID.value.trim().toLowerCase();
  const diff = filterDifficulty.value.trim().toLowerCase();
  const dom = filterDomain.value.trim().toLowerCase();

  filteredQuestions = allQuestions.filter(q => {
    if (kw) {
      const haystack = [
        q.title_text,
        q.case_text,
        q.question_text,
        q.choice_a,
        q.choice_b,
        q.choice_c,
        q.choice_d,
        q.choice_e
      ].join(" ").toLowerCase();
      if (!haystack.includes(kw)) return false;
    }

    if (lang && !q.language.toLowerCase().includes(lang)) return false;
    if (qid && !q.id.toLowerCase().includes(qid)) return false;
    if (diff && !q.difficulty.toLowerCase().includes(diff)) return false;

    if (dom) {
      const domHay = (q.domain1 + " " + q.domain2).toLowerCase();
      if (!domHay.includes(dom)) return false;
    }

    return true;
  });

  // フィルタ後の集計表示
  const tableSummaryEl = document.getElementById("tableSummary");
  if (tableSummaryEl) {
    const mapFiltered = countByLangAndDifficulty(filteredQuestions);
    tableSummaryEl.innerHTML = renderLangDiffSummary(mapFiltered);
  }
}

// ===== 選択済みの再構築 =====
function rebuildSelectedFromAll() {
  const oldOrder = selectedOrdered.map(q => q._uid);
  const fresh = allQuestions.filter(q => q.selected);

  const reordered = [];
  oldOrder.forEach(uid => {
    const found = fresh.find(q => q._uid === uid);
    if (found) reordered.push(found);
  });
  fresh.forEach(q => {
    if (!reordered.find(x => x._uid === q._uid)) {
      reordered.push(q);
    }
  });

  selectedOrdered = reordered;
}

// ===== 左側テーブル描画 =====
function renderTable() {
  questionTableBody.innerHTML = "";

  filteredQuestions.forEach(q => {
    const tr = document.createElement("tr");
    if (q.selected) tr.classList.add("selected-row");

    const previewHtml = buildPreviewHTML(q);

    tr.innerHTML = `
      <td>
        <input type="checkbox" data-uid="${q._uid}" ${q.selected ? "checked" : ""}>
      </td>
      <td>${escapeHTML(q.id)}</td>
      <td>${escapeHTML(q.language)}</td>
      <td>${escapeHTML(q.difficulty)}</td>
      <td>${escapeHTML(q.domain1)}</td>
      <td class="question-cell">${previewHtml}</td>
    `;

    questionTableBody.appendChild(tr);
  });

  // 選択トグル
  questionTableBody.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener("change", e => {
      const uid = e.target.getAttribute("data-uid");
      const target = allQuestions.find(x => x._uid === uid);
      if (!target) return;
      target.selected = e.target.checked;

      rebuildSelectedFromAll();
      renderTable();
      renderSelectedList();
    });
  });
}

// ===== タイトル/症例/問題文まとめ =====
function buildPreviewHTML(q) {
  const hasCase = q.case_id && q.case_id.trim() !== "";

  const titlePart = q.title_text
    ? `<div class="title-block">${q.title_text}</div>`
    : "";

  const casePart = hasCase && q.case_text
    ? `<div class="case-block">${q.case_text}</div>`
    : "";

  const qtextPart = q.question_text
    ? `<div class="qtext-block">${q.question_text}</div>`
    : "";

  return titlePart + casePart + qtextPart;
}

// ===== 右側リスト描画 =====
function renderSelectedList() {
  selectedListEl.innerHTML = selectedOrdered.map((q, idx) => {
    const preview = buildPreviewHTML(q);
    const choicesHTML = renderChoicesMini(q);

    return `
      <div class="selected-item" data-uid="${q._uid}">
        <div class="selected-head-row">
          <div class="qhead">${escapeHTML(q.id)}</div>
          <div class="list-controls">
            <button class="reorder-btn" data-dir="-1" data-index="${idx}">↑</button>
            <button class="reorder-btn" data-dir="1" data-index="${idx}">↓</button>
          </div>
        </div>

        <div class="meta">
          [${escapeHTML(q.language)}]
          難易度:${escapeHTML(q.difficulty)}
          ／${escapeHTML(q.domain1)}${q.domain2 ? "・" + escapeHTML(q.domain2) : ""}
        </div>

        ${preview}

        ${choicesHTML}

        <div class="score-row">
          配点:
          <select class="score-select" data-uid="${q._uid}">
            <option value="1" ${q.score == 1 ? "selected":""}>1点</option>
            <option value="2" ${q.score == 2 ? "selected":""}>2点</option>
            <option value="3" ${q.score == 3 ? "selected":""}>3点</option>
            <option value="4" ${q.score == 4 ? "selected":""}>4点</option>
            <option value="5" ${q.score == 5 ? "selected":""}>5点</option>
          </select>
        </div>
      </div>
    `;
  }).join("");

  // 配点変更
  selectedListEl.querySelectorAll(".score-select").forEach(sel => {
    sel.addEventListener("change", e => {
      const uid = e.target.getAttribute("data-uid");
      const newScore = Number(e.target.value);
      const target = allQuestions.find(q => q._uid === uid);
      if (target) {
        target.score = newScore;
      }
      updateStats();
      renderSelectedList();
    });
  });

  // 並び替え(↑↓)
  selectedListEl.querySelectorAll(".reorder-btn").forEach(btn => {
    btn.addEventListener("click", e => {
      const fromIndex = parseInt(btn.getAttribute("data-index"), 10);
      const dir = parseInt(btn.getAttribute("data-dir"), 10);
      const toIndex = fromIndex + dir;
      if (toIndex < 0 || toIndex >= selectedOrdered.length) return;

      const tmp = selectedOrdered[fromIndex];
      selectedOrdered[fromIndex] = selectedOrdered[toIndex];
      selectedOrdered[toIndex] = tmp;

      renderSelectedList();
    });
  });

  updateStats();

  // 選択済みの言語×難易度表示
  const selectedSummaryEl = document.getElementById("selectedSummary");
  if (selectedSummaryEl) {
    const mapSel = countByLangAndDifficulty(selectedOrdered);
    selectedSummaryEl.innerHTML = renderLangDiffSummary(mapSel);
  }
}

// ===== 選択肢UI表示用 =====
function collectChoices(q) {
  return [
    { key: "A", text: q.choice_a },
    { key: "B", text: q.choice_b },
    { key: "C", text: q.choice_c },
    { key: "D", text: q.choice_d },
    { key: "E", text: q.choice_e }
  ].filter(item => item.text && item.text.trim() !== "");
}

function renderChoicesMini(q) {
  const lines = collectChoices(q);
  if (lines.length === 0) return "";

  return `
    <div class="choices">
      ${lines.map(c => `
        <div class="choice-line">
          <span class="choice-label">${c.key}.</span>
          <span class="choice-text">${c.text}</span>
        </div>
      `).join("")}
    </div>
  `;
}

// ===== 集計 =====
function countByLangAndDifficulty(list) {
  const map = {};
  list.forEach(q => {
    const lang = q.language || "不明";
    const diff = q.difficulty || "未設定";
    if (!map[lang]) map[lang] = {};
    map[lang][diff] = (map[lang][diff] || 0) + 1;
  });
  return map;
}

function renderLangDiffSummary(map) {
  const lines = Object.entries(map).map(([lang, diffs]) => {
    const total = Object.values(diffs).reduce((a, b) => a + b, 0);
    const detail = Object.entries(diffs).map(([d, n]) => `レベル${d} ${n}問`).join("　");
    return `${lang} ${total}問（${detail}）`;
  });
  if (lines.length === 0) return "0問";
  return lines.join("<br>");
}

// ===== 統計（右上の数と点数） =====
function updateStats() {
  const count = selectedOrdered.length;
  const total = selectedOrdered.reduce((sum, q) => sum + (q.score || 0), 0);

  statCountEl.textContent = count;
  statTotalEl.textContent = total;
}

// ===== 出力オプション収集 =====
function collectSettingsFromUI() {
  return {
    titleHeader: optTitleHeader.value.trim(),
    showPageNumber: optShowPageNumber.checked,
    pageMargin: optPageMargin.value,

    questionNumberStyle: optQuestionNumberStyle.value,
    choiceLabelStyle: optChoiceLabelStyle.value,

    showScore: optShowScore.checked,
    scorePosition: optScorePosition.value,

    includeTitle: optIncludeTitle.checked,
    includeChoices: optIncludeChoices.checked,
    includeMetaInfo: optIncludeMeta.checked,

    showAnswer: optShowAnswer.checked,
    answerStyle: {
      color: optAnswerColor.value,
      bold: optAnswerBold.checked,
      underline: optAnswerUnderline.checked
    },

    fontFamilyQuestion: optFontFamilyQuestion.value,
    fontSizeQuestionPt: Number(optFontSizeQuestion.value),

    fontFamilyChoices: optFontFamilyChoices.value,
    fontSizeChoicesPt: Number(optFontSizeChoices.value)
  };
}

// ===== Word(docx)出力 =====
async function exportDocx(settings) {
  if (selectedOrdered.length === 0) {
    alert("問題が選択されていません。");
    return;
  }

  const {
    Document, Packer, Paragraph, TextRun,
    Header, Footer, AlignmentType, PageNumber
  } = window.docx;

  const margin = getMarginPreset(settings.pageMargin);

  // ヘッダー
  const headerChildren = [];
  if (settings.titleHeader) {
    headerChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: settings.titleHeader,
            bold: true
          })
        ],
        alignment: AlignmentType.CENTER
      })
    );
  }
  const headerObj = new Header({ children: headerChildren });

  // フッター（ページ番号）※出る環境だけ出す
  let footerObj = new Footer({ children: [] });
  if (settings.showPageNumber) {
    footerObj = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page " }),
            PageNumber.CURRENT,
            new TextRun({ text: " / " }),
            PageNumber.TOTAL_PAGES
          ]
        })
      ]
    });
  }

  // 本文
  const bodyParas = [];
  selectedOrdered.forEach((q, idx) => {
    const theseParas = buildDocxParagraphsForQuestion(q, idx, settings);
    bodyParas.push(...theseParas);

    bodyParas.push(
      new Paragraph({
        children: [new TextRun({ text: "" })],
        spacing: { after: 120 }
      })
    );
  });

  const doc = new Document({
    sections: [
      {
        headers: { default: headerObj },
        footers: { default: footerObj, first: footerObj },
        page: {
          margin: margin
        },
        children: bodyParas
      }
    ]
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "exam_questions.docx";
  a.click();
}

// ===== 1問ぶんの段落を構成 =====
function buildDocxParagraphsForQuestion(q, index, settings) {
  const { Paragraph, TextRun } = window.docx;

  const paras = [];

  // メタ情報を先に出す場合
  if (settings.includeMetaInfo) {
    paras.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `[${q.language}] 難易度:${q.difficulty} / ${q.domain1}${q.domain2 ? "・" + q.domain2 : ""}`,
            size: settings.fontSizeQuestionPt * 2,
            font: settings.fontFamilyQuestion,
            color: "666666"
          })
        ],
        spacing: { after: 80 }
      })
    );
  }

  // タイトル
  if (settings.includeTitle && q.title_text) {
    const titleRuns = htmlToRuns(q.title_text, {
      fontFamily: settings.fontFamilyQuestion,
      fontSizePt: settings.fontSizeQuestionPt
    });
    paras.push(
      new Paragraph({
        children: titleRuns,
        spacing: { after: 80 }
      })
    );
  }

  // 症例
  const hasCase = q.case_id && q.case_id.trim() !== "";
  if (hasCase && q.case_text) {
    const caseRuns = htmlToRuns(q.case_text, {
      fontFamily: settings.fontFamilyQuestion,
      fontSizePt: settings.fontSizeQuestionPt
    });
    paras.push(
      new Paragraph({
        children: caseRuns,
        spacing: { after: 80 }
      })
    );
  }

  // 問題文＋番号
  const bodyRuns = htmlToRuns(q.question_text, {
    fontFamily: settings.fontFamilyQuestion,
    fontSizePt: settings.fontSizeQuestionPt
  });

  const questionHeadRuns = buildQuestionHeaderRuns(q, index, settings);

  // 配点を末尾につけたい場合
  if (settings.showScore && settings.scorePosition === "end") {
    bodyRuns.push(
      new TextRun({
        text: `[${q.score}]`,
        bold: false,
        size: settings.fontSizeQuestionPt * 2,
        font: settings.fontFamilyQuestion
      })
    );
  }

  paras.push(
    new Paragraph({
      children: [...questionHeadRuns, ...bodyRuns],
      spacing: { after: 120 }
    })
  );

  // 選択肢
  if (settings.includeChoices) {
    const choiceList = buildChoiceList(q, settings.choiceLabelStyle);
    choiceList.forEach(choiceObj => {
      const isCorrect = isChoiceCorrect(q.correct, choiceObj.labelOriginal);

      // 本文をrunsに
      const choiceRuns = htmlToRuns(choiceObj.text, {
        fontFamily: settings.fontFamilyChoices,
        fontSizePt: settings.fontSizeChoicesPt
      });

      // ラベル "A." "B." など
      const labelRun = new TextRun({
        text: choiceObj.labelDisplay + " ",
        bold: true,
        size: settings.fontSizeChoicesPt * 2,
        font: settings.fontFamilyChoices
      });

      // 正解を表示したいときだけ★を足す（安全なやり方）
      if (settings.showAnswer && isCorrect) {
        paras.push(
          new Paragraph({
            spacing: { after: 60 },
            children: [
              labelRun,
              ...choiceRuns,
              new TextRun({
                text: " ★正答肢",
                color: settings.answerStyle.color || "FF0000",
                bold: settings.answerStyle.bold || true,
                underline: settings.answerStyle.underline ? {} : undefined,
                size: settings.fontSizeChoicesPt * 2,
                font: settings.fontFamilyChoices
              })
            ]
          })
        );
      } else {
        paras.push(
          new Paragraph({
            spacing: { after: 60 },
            children: [labelRun, ...choiceRuns]
          })
        );
      }
    });
  }

  return paras;
}

// ===== ヘッダ行(Q番号など) =====
function buildQuestionHeaderRuns(q, index, settings) {
  const { TextRun } = window.docx;
  const label = formatQuestionNumber(index, settings.questionNumberStyle);

  let headerText = label;
  if (settings.showScore && settings.scorePosition === "header") {
    headerText += `[${q.score}]`;
  }
  headerText += " ";

  return [
    new TextRun({
      text: headerText,
      bold: true,
      size: settings.fontSizeQuestionPt * 2,
      font: settings.fontFamilyQuestion
    })
  ];
}

// ===== Q番号スタイル =====
function formatQuestionNumber(i, style) {
  switch (style) {
    case "qnum": return `Q${i + 1}`;
    case "numdot": return `${i + 1}.`;
    case "paren": return `(${i + 1})`;
    case "dai": return `第${i + 1}問`;
    case "mondai": return `問題${i + 1}`;
    default: return `Q${i + 1}`;
  }
}

// ===== 選択肢ラベルスタイル =====
function formatChoiceLabel(n, style) {
  if (style === "abc") {
    const code = "a".charCodeAt(0) + n;
    return String.fromCharCode(code);
  } else if (style === "ABC") {
    const code = "A".charCodeAt(0) + n;
    return String.fromCharCode(code);
  } else {
    return String(n + 1);
  }
}

// ===== 選択肢（docx用） =====
function buildChoiceList(q, choiceStyle) {
  const raw = collectChoices(q);
  return raw.map((item, idx) => {
    const labelCore = formatChoiceLabel(idx, choiceStyle);
    return {
      labelDisplay: labelCore + ".",
      labelOriginal: item.key,
      text: item.text
    };
  });
}

// ===== 正解判定 =====
function isChoiceCorrect(correctField, originalKey) {
  const answers = String(correctField || "")
    .split("|")
    .map(s => s.trim().toUpperCase())
    .filter(s => s !== "");
  const target = String(originalKey || "").trim().toUpperCase();
  return answers.includes(target);
}

// ===== ページ余白 =====
function getMarginPreset(name) {
  switch (name) {
    case "narrow":
      return { top: 500, bottom: 500, left: 500, right: 500 };
    case "wide":
      return { top: 1000, bottom: 1000, left: 1000, right: 1000 };
    case "normal":
    default:
      return { top: 800, bottom: 800, left: 800, right: 800 };
  }
}

// ===== HTML文字列 -> TextRun配列 =====
function htmlToRuns(htmlString, baseStyle) {
  const { fontFamily, fontSizePt } = baseStyle;
  const container = document.createElement("div");
  container.innerHTML = htmlString || "";

  function walk(node, inheritedStyle) {
    let styleNow = { ...inheritedStyle };

    if (node.nodeType === Node.ELEMENT_NODE) {
      const tag = node.tagName ? node.tagName.toLowerCase() : "";

      if (tag === "b" || tag === "strong") {
        styleNow.bold = true;
      }
      if (tag === "i" || tag === "em") {
        styleNow.italics = true;
      }
      if (tag === "u") {
        styleNow.underline = true;
      }
      if (tag === "br") {
        return [
          new window.docx.TextRun({
            text: "\n",
            bold: styleNow.bold || false,
            italics: styleNow.italics || false,
            underline: styleNow.underline ? {} : undefined,
            size: fontSizePt * 2,
            font: fontFamily
          })
        ];
      }

      let out = [];
      node.childNodes.forEach(child => {
        out = out.concat(walk(child, styleNow));
      });
      return out;
    }

    if (node.nodeType === Node.TEXT_NODE) {
      const raw = node.nodeValue ?? "";
      if (raw.replace(/\s+/g, "") === "") {
        return [];
      }

      return [
        new window.docx.TextRun({
          text: raw,
          bold: styleNow.bold || false,
          italics: styleNow.italics || false,
          underline: styleNow.underline ? {} : undefined,
          size: fontSizePt * 2,
          font: fontFamily
        })
      ];
    }

    return [];
  }

  let runs = [];
  container.childNodes.forEach(child => {
    runs = runs.concat(walk(child, {
      bold: false,
      italics: false,
      underline: false
    }));
  });

  if (runs.length === 0) {
    runs.push(new window.docx.TextRun({
      text: "",
      size: fontSizePt * 2,
      font: fontFamily
    }));
  }

  return runs;
}

// ===== ユーティリティ =====
function safe(v) {
  if (v === undefined || v === null) return "";
  return String(v);
}

function escapeHTML(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
