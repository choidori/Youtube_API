// ==========================================
// 1. 전역 상태 및 UI 요소 매핑
// ==========================================
let API_KEY = localStorage.getItem("YT_API_KEY") || "";
const API_BASE = "https://www.googleapis.com/youtube/v3";

const UI = {
    modal: document.getElementById('api-key-modal'),
    apiKeyInput: document.getElementById('api-key-input'),
    btnSettings: document.getElementById('api-key-btn'),
    btnSaveKey: document.getElementById('save-key-btn'),
    btnCloseModal: document.getElementById('close-modal-btn'),
    tabs: document.querySelectorAll('.tab-btn'),
    contents: document.querySelectorAll('.tab-content'),
    consoleOutput: document.getElementById('console-output'),

    // Tab 1 UI
    t1OptDate: document.querySelector('input[name="t1-opt"][value="date"]'),
    t1OptRecent: document.querySelector('input[name="t1-opt"][value="recent"]'),
    t1RecentParams: document.getElementById('t1-recent-params'),
    t1DateParams: document.getElementById('t1-date-params'),
    t1RecentN: document.getElementById('t1-recent-n'),
    t1StartDate: document.getElementById('t1-start-date'),
    t1EndDate: document.getElementById('t1-end-date'),
    t1Run: document.getElementById('t1-run'),

    // Tab 2 UI
    t2Run: document.getElementById('t2-run'),

    // Tab 3 UI
    t3OptDate: document.querySelector('input[name="t3-opt"][value="date"]'),
    t3OptRecent: document.querySelector('input[name="t3-opt"][value="recent"]'),
    t3RecentParams: document.getElementById('t3-recent-params'),
    t3DateParams: document.getElementById('t3-date-params'),
    t3RecentN: document.getElementById('t3-recent-n'),
    t3StartDate: document.getElementById('t3-start-date'),
    t3EndDate: document.getElementById('t3-end-date'),
    t3Run: document.getElementById('t3-run')
};

// ==========================================
// 2. 공통 유틸리티 (로깅, 탭 이벤트 등)
// ==========================================
function log(msg, type = "info") {
    const p = document.createElement('div');
    p.textContent = msg;
    if (type === 'error') p.className = 'log-error';
    if (type === 'success') p.className = 'log-success';
    if (type === 'info') p.className = 'log-info';
    UI.consoleOutput.appendChild(p);
    UI.consoleOutput.scrollTop = UI.consoleOutput.scrollHeight;
}

// 폼 동적 추가/삭제 유틸
window.addInput = function (containerId, inputClass, placeholderText) {
    const container = document.getElementById(containerId);
    const div = document.createElement('div');
    div.className = 'input-row';
    div.innerHTML = `
        <input type="text" class="${inputClass}" placeholder="${placeholderText}">
        <button type="button" class="btn-icon btn-remove" onclick="removeInput(this)">✖</button>
    `;
    container.appendChild(div);
    _updateRemoveButtons(container);
};

window.removeInput = function (btn) {
    const container = btn.closest('.dynamic-inputs');
    if (container.children.length > 1) {
        btn.closest('.input-row').remove();
        _updateRemoveButtons(container);
    }
};

function _updateRemoveButtons(container) {
    const btns = container.querySelectorAll('.btn-remove');
    btns.forEach(b => b.disabled = (btns.length === 1));
}

function initUI() {
    // API 팝업 이벤트
    UI.btnSettings.addEventListener('click', () => {
        UI.apiKeyInput.value = API_KEY;
        UI.modal.classList.remove('hidden');
    });
    UI.btnCloseModal.addEventListener('click', () => UI.modal.classList.add('hidden'));
    UI.btnSaveKey.addEventListener('click', () => {
        API_KEY = UI.apiKeyInput.value.trim();
        localStorage.setItem("YT_API_KEY", API_KEY);
        UI.modal.classList.add('hidden');
        log("[시스템] API 키가 저장되었습니다.", "success");
    });

    // 탭 변환 이벤트
    UI.tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            UI.tabs.forEach(t => t.classList.remove('active'));
            UI.contents.forEach(c => c.classList.add('hidden'));
            tab.classList.add('active');
            document.getElementById(tab.dataset.target).classList.remove('hidden');
        });
    });

    // 라디오 버튼(조건부) 이벤트 (Tab 1)
    document.getElementsByName('t1-opt').forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'recent') {
                UI.t1RecentParams.classList.add('open');
                UI.t1DateParams.classList.remove('open');
            } else {
                UI.t1RecentParams.classList.remove('open');
                UI.t1DateParams.classList.add('open');
            }
        });
    });

    // 라디오 버튼(조건부) 이벤트 (Tab 3)
    document.getElementsByName('t3-opt').forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'recent') {
                UI.t3RecentParams.classList.add('open');
                UI.t3DateParams.classList.remove('open');
            } else {
                UI.t3RecentParams.classList.remove('open');
                UI.t3DateParams.classList.add('open');
            }
        });
    });

    // 출력 형식 버튼 이벤트
    document.querySelectorAll('.format-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.dataset.tab;
            document.querySelectorAll(`.format-btn[data-tab="${tab}"]`).forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
        });
    });

    // 기본 안내 표시
    if (!API_KEY) {
        setTimeout(() => UI.modal.classList.remove('hidden'), 500);
    }
}
initUI();

// ==========================================
// 2-1. 출력 형식 선택 헬퍼 및 Word/PPT 저장 함수
// ==========================================
function getSelectedFormat(tabId) {
    const activeBtn = document.querySelector(`.format-btn.active[data-tab="${tabId}"]`);
    return activeBtn ? activeBtn.dataset.format : 'excel';
}

async function saveWord(sheetsObj, filename, insights) {
    try {
        const { Document, Packer, Paragraph, Table, TableRow, TableCell,
                TextRun, HeadingLevel, WidthType, ShadingType, AlignmentType } = docx;
        const sectionChildren = [];

        // 문자열 2차원 배열(첫 행=헤더)로 표 생성하는 로컬 헬퍼
        const makeTable = (rows) => new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: rows.map((cells, ri) => new TableRow({
                tableHeader: ri === 0,
                children: cells.map(c => new TableCell({
                    shading: ri === 0 ? { fill: '000000', type: ShadingType.CLEAR, color: 'auto' } : undefined,
                    children: [new Paragraph({
                        children: [new TextRun({ text: String(c), bold: ri === 0, color: ri === 0 ? 'FFFFFF' : '000000', size: 18 })],
                    })],
                })),
            })),
        });

        // 표지 타이틀
        sectionChildren.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: 'YouTube Stats Analyzer 보고서', bold: true, size: 52 })],
        }));
        sectionChildren.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: '생성일: ' + new Date().toLocaleDateString('ko-KR'), size: 22, color: '888888' })],
        }));
        sectionChildren.push(new Paragraph({ text: '' }));

        // 댓글 인사이트 섹션 (탭2/탭3 전용)
        if (insights) {
            const s = insights.sentiment;
            sectionChildren.push(new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun({ text: '💡 댓글 인사이트 분석', bold: true })],
            }));

            sectionChildren.push(new Paragraph({ children: [new TextRun({ text: '감성 분석', bold: true, size: 26 })] }));
            sectionChildren.push(makeTable([
                ['감성', '댓글 수', '비율(%)'],
                ['😊 긍정', String(s.positive), String(s.posPct)],
                ['😐 중립', String(s.neutral), String(s.neuPct)],
                ['😞 부정', String(s.negative), String(s.negPct)],
                ['합계', String(s.total), '100'],
            ]));
            sectionChildren.push(new Paragraph({ text: '' }));

            // 극성별 단어 리스트 (긍정/부정/중립)
            const sw = insights.sentimentWords;
            const wordTable = (title, list) => {
                sectionChildren.push(new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 26 })] }));
                sectionChildren.push(makeTable([
                    ['순위', '단어', '빈도'],
                    ...(list.length ? list.map((k, i) => [String(i + 1), k.word, String(k.count)]) : [['-', '(없음)', '-']]),
                ]));
                sectionChildren.push(new Paragraph({ text: '' }));
            };
            wordTable('😊 긍정 단어', sw.positive);
            wordTable('😞 부정 단어', sw.negative);
            wordTable('😐 중립 단어', sw.neutral);

            sectionChildren.push(new Paragraph({ children: [new TextRun({ text: '🔑 키워드 TOP 15', bold: true, size: 26 })] }));
            sectionChildren.push(makeTable([
                ['순위', '키워드', '빈도'],
                ...insights.keywords.map((k, i) => [String(i + 1), k.word, String(k.count)]),
            ]));
            sectionChildren.push(new Paragraph({ text: '' }));
        }

        for (const sheetName in sheetsObj) {
            const rows = sheetsObj[sheetName];
            if (!rows || rows.length === 0) continue;
            const cleanName = sheetName.replace(/[📊📈📋💬]/g, '').trim();

            sectionChildren.push(new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun({ text: cleanName, bold: true })],
            }));

            const headers = Object.keys(rows[0]);
            const tableRows = [
                new TableRow({
                    tableHeader: true,
                    children: headers.map(h => new TableCell({
                        shading: { fill: '000000', type: ShadingType.CLEAR, color: 'auto' },
                        children: [new Paragraph({
                            children: [new TextRun({ text: h, bold: true, color: 'FFFFFF', size: 18 })],
                        })],
                    })),
                }),
                ...rows.map(row => new TableRow({
                    children: headers.map(h => new TableCell({
                        children: [new Paragraph({
                            children: [new TextRun({ text: String(row[h] ?? ''), size: 18 })],
                        })],
                    })),
                })),
            ];

            sectionChildren.push(new Table({
                rows: tableRows,
                width: { size: 100, type: WidthType.PERCENTAGE },
            }));
            sectionChildren.push(new Paragraph({ text: '' }));
        }

        const doc = new Document({ sections: [{ children: sectionChildren }] });
        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
    } catch (e) { throw new Error('Word 생성 실패: ' + e.message); }
}

function savePPT(sheetsObj, filename, insights) {
    try {
        const pptx = new PptxGenJS();
        pptx.layout = 'LAYOUT_WIDE';

        // 타이틀 슬라이드
        const titleSlide = pptx.addSlide();
        titleSlide.background = { color: '000000' };
        titleSlide.addText('YOUTUBE STATS ANALYZER', {
            x: 0.5, y: 1.8, w: '90%', h: 1.5,
            fontSize: 40, bold: true, color: 'FFFFFF', align: 'center',
        });
        titleSlide.addText(new Date().toLocaleDateString('ko-KR') + ' 기준 보고서', {
            x: 0.5, y: 3.5, w: '90%', h: 0.8, fontSize: 18, color: 'AAAAAA', align: 'center',
        });

        // 댓글 인사이트 슬라이드 (탭2/탭3 전용)
        if (insights) {
            const s = insights.sentiment;
            const slide = pptx.addSlide();
            slide.background = { color: 'FFFFFF' };
            slide.addText('댓글 인사이트 분석', {
                x: 0, y: 0, w: '100%', h: 0.65, fontSize: 18, bold: true,
                color: 'FFFFFF', fill: { color: '000000' }, valign: 'middle', align: 'center',
            });
            slide.addText([
                { text: '😊 긍정  ', options: { bold: true } }, { text: `${s.positive}건 (${s.posPct}%)\n` },
                { text: '😐 중립  ', options: { bold: true } }, { text: `${s.neutral}건 (${s.neuPct}%)\n` },
                { text: '😞 부정  ', options: { bold: true } }, { text: `${s.negative}건 (${s.negPct}%)` },
            ], { x: 0.5, y: 1.1, w: 5.3, h: 2.5, fontSize: 18, color: '333333', valign: 'top', lineSpacingMultiple: 1.5 });

            const chartData = [{
                name: '빈도',
                labels: insights.keywords.map(k => k.word),
                values: insights.keywords.map(k => k.count),
            }];
            slide.addChart(pptx.ChartType.bar, chartData, {
                x: 6.0, y: 0.9, w: 6.9, h: 6.1, barDir: 'bar', showValue: true,
                showLegend: false, showTitle: false, catAxisLabelFontSize: 9,
                dataLabelFontSize: 9, chartColors: ['000000'],
            });

            // 극성별 단어 리스트 슬라이드 (긍정/부정/중립 3열)
            const sw = insights.sentimentWords;
            const wordSlide = pptx.addSlide();
            wordSlide.background = { color: 'FFFFFF' };
            wordSlide.addText('감성 단어 리스트', {
                x: 0, y: 0, w: '100%', h: 0.65, fontSize: 18, bold: true,
                color: 'FFFFFF', fill: { color: '000000' }, valign: 'middle', align: 'center',
            });
            const cols = [
                { title: '😊 긍정 단어', fill: 'EAF5EA', list: sw.positive },
                { title: '😞 부정 단어', fill: 'FBEAEA', list: sw.negative },
                { title: '😐 중립 단어', fill: 'F0F0F0', list: sw.neutral },
            ];
            cols.forEach((col, ci) => {
                const x = 0.3 + ci * 4.35;
                const header = [{ text: col.title, options: { bold: true, color: 'FFFFFF', fill: '000000', align: 'center', fontSize: 11 } }];
                const body = (col.list.slice(0, 15)).map(k => [
                    { text: `${k.word}  (${k.count})`, options: { fontSize: 10, fill: col.fill, valign: 'middle' } }
                ]);
                if (body.length === 0) body.push([{ text: '(없음)', options: { fontSize: 10, italic: true, color: '888888' } }]);
                wordSlide.addTable([header, ...body], {
                    x, y: 0.85, w: 4.1, colW: [4.1],
                    border: { type: 'solid', pt: 0.3, color: 'DDDDDD' }, rowH: 0.32,
                });
            });
        }

        for (const sheetName in sheetsObj) {
            const rows = sheetsObj[sheetName];
            if (!rows || rows.length === 0) continue;
            const headers = Object.keys(rows[0]);
            const cleanName = sheetName.replace(/[📊📈📋💬]/g, '').trim();
            const chunkSize = 14;

            for (let i = 0; i < rows.length; i += chunkSize) {
                const chunk = rows.slice(i, i + chunkSize);
                const slide = pptx.addSlide();
                slide.background = { color: 'FFFFFF' };

                const pageInfo = rows.length > chunkSize
                    ? ` (${i + 1}~${Math.min(i + chunkSize, rows.length)} / 총 ${rows.length}건)` : '';
                slide.addText(cleanName + pageInfo, {
                    x: 0, y: 0, w: '100%', h: 0.65,
                    fontSize: 15, bold: true, color: 'FFFFFF',
                    fill: { color: '000000' }, valign: 'middle', margin: [0, 0, 0, 12],
                });

                const colW = parseFloat((12.8 / headers.length).toFixed(2));
                const headerRow = headers.map(h => ({
                    text: h,
                    options: { bold: true, fill: '333333', color: 'FFFFFF', fontSize: 9, align: 'center', valign: 'middle' }
                }));
                const dataRows = chunk.map((row, ri) => headers.map(h => ({
                    text: String(row[h] ?? ''),
                    options: { fontSize: 8, valign: 'middle', fill: ri % 2 === 0 ? 'FFFFFF' : 'F5F5F5' }
                })));

                slide.addTable([headerRow, ...dataRows], {
                    x: 0.2, y: 0.78, w: 12.8,
                    colW: Array(headers.length).fill(colW),
                    border: { type: 'solid', pt: 0.3, color: 'DDDDDD' },
                    rowH: 0.28,
                });
            }
        }

        pptx.writeFile({ fileName: filename });
    } catch (e) { throw new Error('PPT 생성 실패: ' + e.message); }
}

async function saveData(sheetsObj, baseName, tabId, commentTexts) {
    const format = getSelectedFormat(tabId);
    const extMap = { word: 'docx', ppt: 'pptx', excel: 'xlsx' };
    const ext = extMap[format] || 'xlsx';

    // 댓글 본문이 있으면(탭2/탭3) 인사이트 분석
    let insights = null;
    if (commentTexts && commentTexts.length > 0) {
        log(`🔍 댓글 인사이트 분석 중... (${commentTexts.length}건)`);
        await ensureSentiDict(); // 감성사전 로드 보장 (실패해도 기본 사전으로 진행)
        log(SENTI_MAP ? `  └ 감성사전 적용 (${SENTI_MAP.size}개 단어)` : '  └ 감성사전 미로드 → 기본 사전 사용');
        insights = analyzeComments(commentTexts);
    }

    log(`📦 ${format.toUpperCase()} 형식으로 파일 생성 중...`);
    if (format === 'word') {
        await saveWord(sheetsObj, `${baseName}.docx`, insights);
    } else if (format === 'ppt') {
        savePPT(sheetsObj, `${baseName}.pptx`, insights);
    } else {
        saveExcel(sheetsObj, `${baseName}.xlsx`, insights);
    }
    return `${baseName}.${ext}`;
}

// ==========================================
// 2-2. 댓글 인사이트 분석 (감성 / 키워드 / 워드클라우드)
//  - 외부 API 없이 순수 JS 사전(lexicon) 기반 경량 분석
// ==========================================
const POSITIVE_WORDS = [
    "좋다", "좋아", "좋은", "좋네", "좋고", "좋아요", "최고", "대박", "멋지", "멋져", "멋진",
    "훌륭", "감동", "감사", "고맙", "사랑", "예쁘", "이쁘", "귀엽", "재밌", "재미있", "재미",
    "웃기", "행복", "최애", "짱", "굿", "완벽", "추천", "응원", "화이팅", "파이팅", "힐링",
    "위로", "잘한", "잘했", "잘하", "인정", "천재", "명작", "띵작", "갓", "존좋", "따봉",
    "흐뭇", "뿌듯", "만족", "신기", "놀라", "대단", "기대", "설레", "즐겁", "즐거", "유익",
    "도움", "감탄", "행복해", "고마워", "좋았", "good", "best", "love", "nice", "perfect"
];
const NEGATIVE_WORDS = [
    "싫다", "싫어", "싫은", "별로", "최악", "나쁘", "나쁜", "짜증", "화나", "화남", "실망",
    "끔찍", "노잼", "재미없", "지루", "답답", "불편", "불쾌", "한심", "어이없", "황당",
    "쓰레기", "망했", "망함", "비추", "욕", "논란", "거슬", "빡치", "빡침", "분노", "슬프",
    "슬픔", "우울", "무섭", "무서", "아쉽", "아쉬", "손절", "거짓", "사기", "억지", "불만",
    "극혐", "혐오", "노답", "실패", "형편없", "같잖", "거지같", "최악이", "싫음", "토나",
    "bad", "worst", "hate", "terrible", "awful"
];
const STOP_WORDS = new Set([
    "그리고", "그래서", "하지만", "근데", "그런데", "정말", "진짜", "너무", "완전", "그냥",
    "이거", "저거", "그거", "여기", "거기", "저기", "우리", "저희", "너희", "약간", "조금",
    "많이", "다들", "모두", "우와", "사람", "사람들", "영상", "댓글", "구독", "채널", "보고",
    "보면", "봤는데", "같아요", "같다", "합니다", "했어요", "해요", "하는", "하고", "한번",
    "이런", "저런", "그런", "어떤", "무슨", "제가", "내가", "나는", "너는", "그게", "이게",
    "저게", "그것", "이것", "저것", "근까", "그거는", "있는", "있어요", "있다", "없다", "없어",
    // 대명사·기능 어절(조사 제거 후에도 남아 점수를 오염시키는 것들)
    "저는", "저도", "저를", "저만", "저랑", "나는", "나도", "나를", "너는", "너도", "우린",
    "그는", "그를", "그가", "네가", "당신", "여러분", "본인", "이건", "그건", "저건",
    "the", "and", "you", "for", "this", "that", "with", "are", "was", "have"
]);
// 흔한 조사 접미사 (긴 것부터 제거)
const JOSA_SUFFIX = [
    "으로부터", "에서는", "으로는", "에게서", "에서", "에게", "한테", "으로", "로서", "로써",
    "보다", "처럼", "만큼", "까지", "부터", "마다", "조차", "마저", "이나", "이라", "라고",
    "은", "는", "이", "가", "을", "를", "에", "의", "도", "로", "와", "과", "만", "랑", "요"
];

function stripJosa(word) {
    for (const j of JOSA_SUFFIX) {
        // 조사 제거 후 어간이 2자 이상 남을 때만 제거
        if (word.length > j.length + 1 && word.endsWith(j)) {
            return word.slice(0, -j.length);
        }
    }
    return word;
}

// --- KNU 한국어 감성사전 (data/senti_dict.json) ---
// 출처: KnuSentiLex (군산대 KNU) — 단어별 극성 점수(-2~+2). 약 1.4만 단어.
let SENTI_MAP = null;          // Map<word, polarity(int)>
let SENTI_DICT_PROMISE = null;
function ensureSentiDict() {
    if (SENTI_MAP) return Promise.resolve(SENTI_MAP);
    if (!SENTI_DICT_PROMISE) {
        SENTI_DICT_PROMISE = fetch('data/senti_dict.json')
            .then(r => r.json())
            .then(arr => {
                const m = new Map();
                for (const e of arr) {
                    if (!e || !e.word) continue;
                    const p = parseInt(e.polarity, 10);
                    if (!Number.isNaN(p)) m.set(e.word, p);
                }
                SENTI_MAP = m;
                return m;
            })
            .catch(() => { SENTI_MAP = null; return null; }); // 실패 시 기본 사전으로 폴백
    }
    return SENTI_DICT_PROMISE;
}

// 부정어: 뒤 단어를 뒤집음(안/못/전혀…) / 앞 단어를 뒤집음(않/없/아니…)
const PRE_NEG = new Set(['안', '못', '전혀', '결코', '하나도', '그닥', 'not', 'never', 'no']);
function isPostNeg(tok) {
    return tok.startsWith('않') || tok.startsWith('없') || tok.startsWith('아니') || tok === '말';
}

// 단일 토큰의 극성 점수: 감성사전 우선, 없으면 자체 긍/부정 어간 매칭(±1)
function lookupScore(tok) {
    if (STOP_WORDS.has(tok)) return 0; // 대명사·기능어는 사전 노이즈를 무시하고 중립
    if (SENTI_MAP) {
        // 정확 일치
        let s = SENTI_MAP.get(tok);
        if (s !== undefined) return s;
        // 어간(접두) 일치: 활용형 보정 ("축하해"→"축하", "감사합니다"→"감사")
        if (/^[가-힣]+$/.test(tok)) {
            for (let len = tok.length - 1; len >= 2; len--) {
                s = SENTI_MAP.get(tok.slice(0, len));
                if (s !== undefined) return s;
            }
        }
    }
    for (const w of POSITIVE_WORDS) if (tok.startsWith(w)) return 1;
    for (const w of NEGATIVE_WORDS) if (tok.startsWith(w)) return -1;
    return 0;
}

// 댓글 1건의 감성 점수(부정어 반전 반영). >0 긍정 / <0 부정 / 0 중립
function commentScore(text) {
    const tokens = String(text).toLowerCase().match(/[가-힣a-z0-9]+/g) || [];
    let score = 0, prev = 0, negNext = false;
    for (const raw of tokens) {
        const tok = /^[가-힣]+$/.test(raw) ? stripJosa(raw) : raw;
        if (!tok) continue;
        if (PRE_NEG.has(tok)) { negNext = true; continue; }
        if (isPostNeg(tok)) { if (prev !== 0) { score -= 2 * prev; prev = 0; } continue; } // 앞 감성어 반전
        let s = lookupScore(tok);
        // 붙어있는 부정 접두사: "안좋아", "못잊어"
        if (s === 0 && (tok.startsWith('안') || tok.startsWith('못'))) {
            const innerS = lookupScore(tok.slice(1));
            if (innerS !== 0) s = -innerS;
        }
        // 한 토큰 안에 부정형 포함: "재미없다", "좋지않아"
        if (s > 0 && (tok.includes('없') || tok.includes('않') || tok.includes('아니'))) s = -s;
        if (s !== 0) {
            if (negNext) { s = -s; negNext = false; }
            score += s;
            prev = s;
        }
    }
    return score;
}

function analyzeComments(texts) {
    let positive = 0, negative = 0, neutral = 0;
    const freq = {};
    // 단어를 극성별로 분류해 빈도 집계
    const posFreq = {}, negFreq = {}, neuFreq = {};

    for (const raw of texts) {
        if (!raw) continue;
        const text = String(raw);

        // 1) 감성(댓글 단위): 감성사전 점수 합산 + 부정어 반전
        const sc = commentScore(text);
        if (sc > 0) positive++;
        else if (sc < 0) negative++;
        else neutral++;

        // 2) 단어 토큰화 → 조사 제거 → 불용어/길이 필터 → 전체 빈도 + 극성별 빈도 집계
        const tokens = text.toLowerCase().match(/[가-힣a-z0-9]+/g) || [];
        for (let tok of tokens) {
            if (/^[가-힣]+$/.test(tok)) tok = stripJosa(tok);
            if (tok.length < 2) continue;
            if (/^\d+$/.test(tok)) continue;
            if (STOP_WORDS.has(tok)) continue;
            freq[tok] = (freq[tok] || 0) + 1;

            const wp = lookupScore(tok); // 단어 자체의 극성
            if (wp > 0) posFreq[tok] = (posFreq[tok] || 0) + 1;
            else if (wp < 0) negFreq[tok] = (negFreq[tok] || 0) + 1;
            else neuFreq[tok] = (neuFreq[tok] || 0) + 1;
        }
    }

    const total = positive + negative + neutral;
    const topN = (obj, n) => Object.entries(obj)
        .sort((a, b) => b[1] - a[1])
        .slice(0, n)
        .map(([word, count]) => ({ word, count }));

    return {
        sentiment: {
            positive, negative, neutral, total,
            posPct: total ? +(positive / total * 100).toFixed(1) : 0,
            negPct: total ? +(negative / total * 100).toFixed(1) : 0,
            neuPct: total ? +(neutral / total * 100).toFixed(1) : 0,
        },
        keywords: topN(freq, 15),
        sentimentWords: {
            positive: topN(posFreq, 20),
            negative: topN(negFreq, 20),
            neutral: topN(neuFreq, 20),
        }
    };
}

// 인사이트를 Excel 시트(객체 배열)로 변환
function buildInsightSheets(insights) {
    const s = insights.sentiment;
    const sentimentRows = [
        { "감성": "😊 긍정", "댓글 수": s.positive, "비율(%)": s.posPct },
        { "감성": "😐 중립", "댓글 수": s.neutral, "비율(%)": s.neuPct },
        { "감성": "😞 부정", "댓글 수": s.negative, "비율(%)": s.negPct },
        { "감성": "합계", "댓글 수": s.total, "비율(%)": 100 },
    ];
    const keywordRows = insights.keywords.map((k, i) => ({ "순위": i + 1, "키워드": k.word, "빈도": k.count }));
    const wordRows = (list) => list.map((k, i) => ({ "순위": i + 1, "단어": k.word, "빈도": k.count }));
    const sw = insights.sentimentWords;
    return {
        "📊 감성 분석": sentimentRows,
        "😊 긍정 단어": wordRows(sw.positive),
        "😞 부정 단어": wordRows(sw.negative),
        "😐 중립 단어": wordRows(sw.neutral),
        "🔑 키워드 TOP15": keywordRows,
    };
}

// ==========================================
// 3. 파싱 스크립트 함수 (Python to JS)
// ==========================================
function extractChannelId(url) {
    url = url.trim();
    if (url.startsWith("UC") && url.length >= 24) return url;
    if (url.includes("youtube.com/channel/")) return url.split("channel/")[1].split("/")[0].split("?")[0];
    if (url.includes("@")) {
        const handle = "@" + url.split("@")[1].split("/")[0].split("?")[0];
        return handle;
    }
    return url; // fallback
}

function extractVideoId(url) {
    url = url.trim();
    if (url.includes("youtu.be/")) return url.split("youtu.be/")[1].split("?")[0];
    if (url.includes("/shorts/")) return url.split("/shorts/")[1].split("/")[0].split("?")[0];
    if (url.includes("/live/")) return url.split("/live/")[1].split("/")[0].split("?")[0];
    if (url.includes("v=")) {
        const params = new URLSearchParams(url.split("?")[1]);
        if (params.has("v")) return params.get("v");
    }
    const match = url.match(/(?:v=|\/)([0-9A-Za-z_-]{11})(?:\?|&|\/|$)/);
    if (match) return match[1];
    return url.length === 11 ? url : null;
}

function parseDurationType(duration) {
    // PT1H2M3S
    const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
    if (match) {
        const h = parseInt(match[1] || 0), m = parseInt(match[2] || 0), s = parseInt(match[3] || 0);
        const total = h * 3600 + m * 60 + s;
        return (total > 0 && total <= 61) ? "쇼츠(Shorts)" : "일반 영상";
    }
    return "일반 영상";
}

function formatDateKST(isoString) {
    if (!isoString) return "";
    const date = new Date(isoString);
    const dtf = new Intl.DateTimeFormat('ko-KR', {
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit',
        hour12: false, timeZone: 'Asia/Seoul'
    });
    return dtf.format(date).replace(/\. /g, '-').replace('.', '');
}

// ==========================================
// 4. 유튜브 API Fetch Wrappers
// ==========================================
async function apiGet(endpoint, params) {
    if (!API_KEY) throw new Error("API KEY가 설정되지 않았습니다.");
    params.key = API_KEY;
    const url = new URL(`${API_BASE}/${endpoint}`);
    Object.keys(params).forEach(k => {
        if (params[k] !== undefined && params[k] !== null) url.searchParams.append(k, params[k]);
    });
    const res = await fetch(url.toString());
    const data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || "API 요청 실패");
    return data;
}

// 핸들로 채널 ID 찾기
async function resolveChannelId(input) {
    const parsed = extractChannelId(input);
    if (parsed.startsWith("UC")) return parsed;
    if (parsed.startsWith("@")) {
        const res = await apiGet("search", { part: "snippet", type: "channel", q: parsed, maxResults: 1 });
        if (res.items && res.items.length > 0) return res.items[0].id.channelId;
    }
    return null;
}

// 채널 내 영상 목록 가져오기 (기간 or 최근N개)
async function getChannelVideos(channelId, mode = "recent", maxResults = 20, start = null, end = null) {
    let videos = [];
    let pageToken = "";

    // 채널의 Upload Playlist ID 가져오기 설정 (대량 수집의 권장 방법)
    const chRes = await apiGet("channels", { part: "contentDetails", id: channelId });
    if (!chRes.items || chRes.items.length === 0) return videos;
    const playlistId = chRes.items[0].contentDetails.relatedPlaylists.uploads;

    // 만약 기간 검색이면 Search API를 써야 함.
    if (mode === "date") {
        while (true) {
            const res = await apiGet("search", {
                part: "snippet", channelId: channelId, type: "video",
                publishedAfter: start ? new Date(start).toISOString() : null,
                publishedBefore: end ? new Date(new Date(end).setHours(23, 59, 59)).toISOString() : null,
                maxResults: 50, pageToken: pageToken || null, order: "date"
            });
            for (let item of res.items) {
                videos.push({ video_id: item.id.videoId, published_at: item.snippet.publishedAt, title: item.snippet.title });
            }
            pageToken = res.nextPageToken;
            if (!pageToken) break;
        }
    } else {
        // Playlist API 최적화 (Quota 저렴)
        let limit = maxResults;
        while (limit > 0) {
            let fetchCount = Math.min(limit, 50);
            const res = await apiGet("playlistItems", {
                part: "snippet", playlistId: playlistId, maxResults: fetchCount, pageToken: pageToken || null
            });
            for (let item of res.items) {
                videos.push({ video_id: item.snippet.resourceId.videoId, published_at: item.snippet.publishedAt, title: item.snippet.title });
            }
            pageToken = res.nextPageToken;
            limit -= res.items.length;
            if (!pageToken || res.items.length === 0) break;
        }
    }
    return videos;
}

// ==========================================
// 5. 엑셀 생성 기능 중심부
// ==========================================
function saveExcel(sheetsObj, filename, insights) {
    const wb = XLSX.utils.book_new();
    // 인사이트가 있으면 감성/키워드 시트를 앞에 삽입 (Excel은 이미지 미지원 → 표로 대체)
    const allSheets = insights ? { ...buildInsightSheets(insights), ...sheetsObj } : sheetsObj;
    for (const sheetName in allSheets) {
        const ws = XLSX.utils.json_to_sheet(allSheets[sheetName]);
        XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31)); // 시트명 길이제한 31자
    }
    XLSX.writeFile(wb, filename);
}

// 탭 1 기능
UI.t1Run.addEventListener('click', async () => {
    try {
        UI.t1Run.disabled = true;
        UI.t1Run.querySelector('.btn-main-text').textContent = "EXTRACTING...";

        const inputNodes = document.querySelectorAll('.t1-channel');
        const inputs = Array.from(inputNodes).map(n => n.value.trim()).filter(v => v);

        if (inputs.length === 0) throw new Error("최소 1개의 채널 주소를 입력하세요.");

        const mode = UI.t1OptRecent.checked ? "recent" : "date";
        let sheetsData = {};
        let summaryData = [];

        for (let chInput of inputs) {
            log(`[채널검석] ${chInput} 아이디 식별 시도...`);
            const chId = await resolveChannelId(chInput);
            if (!chId) { log(`❌ 채널 식별 실패: ${chInput}`, 'error'); continue; }

            // 채널 이름 추출
            const chRes = await apiGet("channels", { part: "snippet", id: chId });
            const chName = chRes.items[0].snippet.title;
            log(`▶ [${chName}] 채널 데이터 수집을 시작합니다.`, 'info');

            let vids = [];
            if (mode === 'recent') {
                vids = await getChannelVideos(chId, 'recent', parseInt(UI.t1RecentN.value));
            } else {
                vids = await getChannelVideos(chId, 'date', 50, UI.t1StartDate.value, UI.t1EndDate.value);
            }

            log(`└ 영상 ${vids.length}개 탐색 완료. 세부 지표 수집을 시작합니다...`);
            let videoStats = [];
            let sumViews = 0, sumComments = 0, sumLikes = 0;

            for (let i = 0; i < vids.length; i++) {
                const vid = vids[i];
                try {
                    const statsRes = await apiGet("videos", { part: "statistics,contentDetails,snippet", id: vid.video_id });
                    if (!statsRes.items.length) continue;

                    const item = statsRes.items[0];
                    const st = item.statistics;
                    const dur = item.contentDetails.duration;
                    const vType = parseDurationType(dur);
                    const vCount = parseInt(st.viewCount || 0);
                    const lCount = parseInt(st.likeCount || 0);
                    const cCount = parseInt(st.commentCount || 0);
                    const er = vCount > 0 ? ((lCount + cCount) / vCount * 100) : 0;

                    sumViews += vCount; sumComments += cCount; sumLikes += lCount;

                    videoStats.push({
                        "제목": item.snippet.title,
                        "영상 종류": vType,
                        "조회수": vCount,
                        "댓글수": cCount,
                        "좋아요": lCount,
                        "참여율(%)": parseFloat(er.toFixed(2)),
                        "업로드일자": formatDateKST(item.snippet.publishedAt),
                        "URL": `https://youtube.com/watch?v=${vid.video_id}`
                    });
                } catch (e) { }
            }

            // 요약
            const avgER = (vids.length > 0 && sumViews > 0) ? ((sumLikes + sumComments) / sumViews * 100) : 0;
            summaryData.push({
                "채널명": chName,
                "집계된 영상수": vids.length,
                "총 조회수": sumViews,
                "평균 조회수": vids.length ? Math.round(sumViews / vids.length) : 0,
                "총 댓글수": sumComments,
                "총 좋아요": sumLikes,
                "평균 채널 참여율(%)": parseFloat(avgER.toFixed(2)),
                "채널 링크": `https://youtube.com/channel/${chId}`
            });

            videoStats.sort((a, b) => b.조회수 - a.조회수);
            const safeName = chName.replace(/[\/*?:<>|]/g, "").substring(0, 15);
            sheetsData[`📈 ${safeName}_영상리스트`] = videoStats;
        }

        if (summaryData.length === 0) throw new Error("추출된 결과가 없습니다.");

        // 병합
        let finalOutput = { "📊 성과 요약": summaryData, ...sheetsData };
        const fname = await saveData(finalOutput, `유튜브통계분석_${new Date().getTime()}`, 't1');
        log(`🎉 다운로드 완료: ${fname}`, "success");

    } catch (err) {
        log(`오류 발생: ${err.message}`, "error");
    } finally {
        UI.t1Run.disabled = false;
        UI.t1Run.querySelector('.btn-main-text').textContent = "EXTRACT DATA";
    }
});

// 탭 2 / 탭 3에서 공통으로 쓰는 개별영상 댓글수집기
async function scrapeVideoComments(videoId, videoTitle, videoType, channelName) {
    let commentsCollected = [];
    let pageToken = "";
    try {
        while (true) {
            const res = await apiGet("commentThreads", {
                part: "snippet,replies", videoId: videoId, maxResults: 100, pageToken: pageToken || null, textFormat: "plainText"
            });

            for (let item of res.items) {
                const top = item.snippet.topLevelComment.snippet;
                const threadId = item.id;

                const baseRow = {
                    "채널명": channelName,
                    "영상 제목": videoTitle,
                    "영상 종류": videoType,
                    "영상 링크": `https://youtube.com/watch?v=${videoId}`,
                    "그룹ID": threadId
                };

                commentsCollected.push({ ...baseRow, "분류": "메인 댓글", "작성자": top.authorDisplayName, "댓글 내용": top.textDisplay, "좋아요 수": top.likeCount, "작성 일자": formatDateKST(top.publishedAt) });

                if (item.snippet.totalReplyCount > 0) {
                    if (item.replies && item.replies.comments.length === item.snippet.totalReplyCount) {
                        for (let rep of item.replies.comments) {
                            const rs = rep.snippet;
                            commentsCollected.push({ ...baseRow, "분류": "답글(대댓글)", "작성자": rs.authorDisplayName, "댓글 내용": rs.textDisplay, "좋아요 수": rs.likeCount, "작성 일자": formatDateKST(rs.publishedAt) });
                        }
                    } else {
                        // 대댓글 페이징
                        let rPage = "";
                        while (true) {
                            const rRes = await apiGet("comments", { part: "snippet", parentId: threadId, maxResults: 100, pageToken: rPage || null, textFormat: "plainText" });
                            for (let rep of rRes.items) {
                                const rs = rep.snippet;
                                commentsCollected.push({ ...baseRow, "분류": "답글(대댓글)", "작성자": rs.authorDisplayName, "댓글 내용": rs.textDisplay, "좋아요 수": rs.likeCount, "작성 일자": formatDateKST(rs.publishedAt) });
                            }
                            rPage = rRes.nextPageToken;
                            if (!rPage) break;
                        }
                    }
                }
            }
            pageToken = res.nextPageToken;
            if (!pageToken) break;
        }
    } catch (e) {
        log(`- 댓글 추출 우회/제한 발견 [${videoId}]: ${e.message}`, "error");
    }
    return commentsCollected;
}

// 탭 2 기능 (영상별 댓글 추출)
UI.t2Run.addEventListener('click', async () => {
    try {
        UI.t2Run.disabled = true;
        UI.t2Run.querySelector('.btn-main-text').textContent = "EXTRACTING...";

        const inputNodes = document.querySelectorAll('.t2-video');
        const inputs = Array.from(inputNodes).map(n => n.value.trim()).filter(v => v);

        if (inputs.length === 0) throw new Error("최소 1개의 영상 주소를 넣어주세요.");

        let allSummaries = [];
        let finalOutput = {};
        let allCommentTexts = [];

        for (let vidUrl of inputs) {
            const vId = extractVideoId(vidUrl);
            if (!vId) { log(`❌ 유효하지 않은 영상: ${vidUrl}`, "error"); continue; }
            log(`▶ 영상 분석 시작: [${vId}]`);

            // 영상 정보
            const vRes = await apiGet("videos", { part: "snippet,statistics,contentDetails", id: vId });
            if (!vRes.items.length) { log(`존재하지 않는 영상: ${vId}`, "error"); continue; }
            const vInfo = vRes.items[0];
            const vTitle = vInfo.snippet.title;
            const cName = vInfo.snippet.channelTitle;
            const viewC = vInfo.statistics.viewCount;
            const likeC = vInfo.statistics.likeCount;
            const commC = vInfo.statistics.commentCount;
            const vType = parseDurationType(vInfo.contentDetails.duration);

            const comments = await scrapeVideoComments(vId, vTitle, vType, cName);

            let mCount = comments.filter(c => c.분류 === "메인 댓글").length;
            let rCount = comments.filter(c => c.분류 === "답글(대댓글)").length;

            allSummaries.push({
                "채널명": cName, "영상 제목": vTitle, "영상 종류": vType, "영상 링크": `https://youtube.com/watch?v=${vId}`,
                "공식 조회수": viewC, "공식 좋아요수": likeC, "공식 댓글수": commC,
                "추출 메인댓글": mCount, "추출 대댓글": rCount, "총 추출수": comments.length
            });

            if (comments.length > 0) {
                const sheetKey = `💬 ${vTitle.replace(/[\/*?:<>|]/g, "").substring(0, 10)}_댓글`;
                finalOutput[sheetKey] = comments;
                allCommentTexts.push(...comments.map(c => c["댓글 내용"]));
            }
        }

        if (allSummaries.length === 0) throw new Error("저장할 데이터가 없습니다.");

        finalOutput = { "📊 댓글 요약": allSummaries, ...finalOutput };
        const fname = await saveData(finalOutput, `유튜브_댓글추출_${new Date().getTime()}`, 't2', allCommentTexts);
        log(`🎉 다운로드 완료: ${fname}`, "success");

    } catch (e) {
        log(e.message, "error");
    } finally {
        UI.t2Run.disabled = false;
        UI.t2Run.querySelector('.btn-main-text').textContent = "EXTRACT COMMENTS";
    }
});

// 탭 3 기능 (채널 기반 영상-댓글 일괄 추출)
UI.t3Run.addEventListener('click', async () => {
    try {
        UI.t3Run.disabled = true;
        UI.t3Run.querySelector('.btn-main-text').textContent = "EXTRACTING ALL...";

        const inputNodes = document.querySelectorAll('.t3-channel');
        const inputs = Array.from(inputNodes).map(n => n.value.trim()).filter(v => v);

        if (inputs.length === 0) throw new Error("최소 1개의 채널 주소를 입력하세요.");

        const mode = UI.t3OptRecent.checked ? "recent" : "date";

        let allSummaries = [];
        let finalChannelsData = {};
        let allCommentTexts = [];

        for (let chInput of inputs) {
            log(`[채널검석] ${chInput} 식별 시도...`);
            const chId = await resolveChannelId(chInput);
            if (!chId) { log(`❌ 채널 식별 오류: ${chInput}`, 'error'); continue; }

            const chRes = await apiGet("channels", { part: "snippet", id: chId });
            const chName = chRes.items[0].snippet.title;
            log(`▶ [${chName}] 수집 시작. 대상 영상들을 분석합니다...`, 'info');

            let vids = [];
            if (mode === 'recent') {
                vids = await getChannelVideos(chId, 'recent', parseInt(UI.t3RecentN.value));
            } else {
                vids = await getChannelVideos(chId, 'date', 50, UI.t3StartDate.value, UI.t3EndDate.value);
            }

            log(`└ [총 ${vids.length}개] 영상의 모든 댓글 추출 시작...`);

            let currentChannelAllComments = [];

            for (let i = 0; i < vids.length; i++) {
                try {
                    const vRes = await apiGet("videos", { part: "snippet,statistics,contentDetails", id: vids[i].video_id });
                    if (!vRes.items.length) continue;
                    const vInfo = vRes.items[0];
                    const vTitle = vInfo.snippet.title;
                    const vType = parseDurationType(vInfo.contentDetails.duration);

                    const comments = await scrapeVideoComments(vInfo.id, vTitle, vType, chName);

                    allSummaries.push({
                        "채널명": chName, "영상 제목": vTitle, "영상 종류": vType, "영상 링크": `https://youtube.com/watch?v=${vInfo.id}`,
                        "공식 조회수": vInfo.statistics.viewCount,
                        "공식 좋아요수": vInfo.statistics.likeCount,
                        "공식 댓글수": vInfo.statistics.commentCount,
                        "최종 수집 댓글수": comments.length
                    });

                    currentChannelAllComments.push(...comments);
                    log(`    ... ${i + 1}/${vids.length} 완료 (${vTitle.substring(0, 10)}.. : ${comments.length}건)`);
                } catch (err) { }
            }

            if (currentChannelAllComments.length > 0) {
                const sheetKey = `💬 ${chName.replace(/[\/*?:<>|]/g, "").substring(0, 15)}_전체댓글`;
                finalChannelsData[sheetKey] = currentChannelAllComments;
                allCommentTexts.push(...currentChannelAllComments.map(c => c["댓글 내용"]));
            }
        }

        if (allSummaries.length === 0) throw new Error("추출 결과가 없습니다.");

        let finalOutput = { "📊 영상별 댓글 요약": allSummaries, ...finalChannelsData };
        const fname = await saveData(finalOutput, `채널기반_대량댓글분석_${new Date().getTime()}`, 't3', allCommentTexts);
        log(`🎉 대규모 다운로드 완료: ${fname}`, "success");

    } catch (e) {
        log(e.message, "error");
    } finally {
        UI.t3Run.disabled = false;
        UI.t3Run.querySelector('.btn-main-text').textContent = "EXTRACT ALL COMMENTS";
    }
});

// 감성사전 백그라운드 사전 로드 (모든 선언이 끝난 뒤 호출)
ensureSentiDict();
