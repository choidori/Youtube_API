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

async function saveWord(sheetsObj, filename) {
    try {
        const { Document, Packer, Paragraph, Table, TableRow, TableCell,
                TextRun, HeadingLevel, WidthType, ShadingType, AlignmentType } = docx;
        const sectionChildren = [];

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

function savePPT(sheetsObj, filename) {
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

async function saveData(sheetsObj, baseName, tabId) {
    const format = getSelectedFormat(tabId);
    const extMap = { word: 'docx', ppt: 'pptx', excel: 'xlsx' };
    const ext = extMap[format] || 'xlsx';
    log(`📦 ${format.toUpperCase()} 형식으로 파일 생성 중...`);
    if (format === 'word') {
        await saveWord(sheetsObj, `${baseName}.docx`);
    } else if (format === 'ppt') {
        savePPT(sheetsObj, `${baseName}.pptx`);
    } else {
        saveExcel(sheetsObj, `${baseName}.xlsx`);
    }
    return `${baseName}.${ext}`;
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
function saveExcel(sheetsObj, filename) {
    const wb = XLSX.utils.book_new();
    for (const sheetName in sheetsObj) {
        const ws = XLSX.utils.json_to_sheet(sheetsObj[sheetName]);
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
            }
        }

        if (allSummaries.length === 0) throw new Error("저장할 데이터가 없습니다.");

        finalOutput = { "📊 댓글 요약": allSummaries, ...finalOutput };
        const fname = await saveData(finalOutput, `유튜브_댓글추출_${new Date().getTime()}`, 't2');
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
            }
        }

        if (allSummaries.length === 0) throw new Error("추출 결과가 없습니다.");

        let finalOutput = { "📊 영상별 댓글 요약": allSummaries, ...finalChannelsData };
        const fname = await saveData(finalOutput, `채널기반_대량댓글분석_${new Date().getTime()}`, 't3');
        log(`🎉 대규모 다운로드 완료: ${fname}`, "success");

    } catch (e) {
        log(e.message, "error");
    } finally {
        UI.t3Run.disabled = false;
        UI.t3Run.querySelector('.btn-main-text').textContent = "EXTRACT ALL COMMENTS";
    }
});
