// PM 모드 Task1 — UIUX 필터링 없이 모든 기능/비기능 요구사항을 통과시키는 별도 파이프라인.
// 기존 UIUX 모드(App.jsx의 coreSystemPrompt / parseExcelToRequirements 등)는 수정하지 않으며,
// 이 파일의 로직은 그것과 병렬로 존재한다. 프론트엔드 모드 전환 스위치 연동은 다음 단계에서 진행.
import * as XLSX from 'xlsx';

// 원본에 요구사항ID가 없는 행에 자동 생성하는 ID의 접두사 (확정: PMREQ-{n})
export const PM_ID_PREFIX = 'PMREQ';

// PM 모드가 인식하는 필수참조 컬럼 5개: 업무시스템, 업무구분, 요구사항ID, 요구사항명, 요구사항내용
// (업무구분은 단일 컬럼 또는 대(L1)/중(L2)/소(L3) 3단 분리 컬럼 중 어느 형태든 인식)
// 기능/비기능 구분은 5개 필수컬럼과는 별도 축으로, 컬럼값이 있으면 그 값을 우선 사용하고
// 없으면 시트명(기능요구사항/비기능요구사항)으로 판별해 추가 필드로 포함한다.
const PM_ALIASES = {
    id: ['요구사항ID', '요구사항아이디', '고유번호', '고유ID', '번호', 'ID'],
    업무시스템: ['업무시스템'],
    업무구분: ['업무구분'],
    L1: ['대(L1)', '대분류'],
    L2: ['중(L2)', '중분류'],
    L3: ['소(L3)', '소분류'],
    title: ['요구사항명', '요구정의명'],
    detail: ['요구사항내용', '상세내용'],
    구분: ['기능유형', '기능/비기능', '기능구분', '요구사항유형', '구분'],
};

const normalize = (s) => String(s ?? '').replace(/\s+/g, '').trim();

// 헤더 셀 하나가 어떤 필드에 해당하는지 판정 (정규화 후 완전일치만 인정 — 부분일치는
// "관련 요구사항ID" 같은 비관련 컬럼을 요구사항ID로 오인식할 위험이 있어 배제)
const matchAlias = (cell) => {
    const h = normalize(cell);
    if (!h) return null;
    for (const [field, aliases] of Object.entries(PM_ALIASES)) {
        if (aliases.some(a => normalize(a) === h)) return field;
    }
    return null;
};

// 시트 상단 최대 maxScan행을 훑어 필수 컬럼 별칭이 가장 많이 매칭되는 행을 헤더 행으로 추정.
// 타이틀/설명 텍스트가 앞쪽에 있는 문서(예: 시트 1~3행이 제목/안내문)도 지원하기 위함.
const detectHeaderRow = (rows, maxScan = 10) => {
    let best = null;
    for (let i = 0; i < Math.min(maxScan, rows.length); i++) {
        const row = rows[i] || [];
        const cols = {};
        row.forEach((cell, ci) => {
            const field = matchAlias(cell);
            if (field && !(field in cols)) cols[field] = ci;
        });
        const score = Object.keys(cols).length;
        if (!best || score > best.score) best = { idx: i, score, cols };
    }
    return best;
};

// 업무구분이 대(L1)/중(L2)/소(L3) 여러 컬럼으로 나뉘어 있는 경우, 헤더 다음 행에서
// 하위 헤더(업무시스템/대/중/소)를 찾아 병합한다. 병합 시 상위의 "업무구분"(병합 셀 레이블)
// 컬럼 매핑은 폐기하고 L1/L2/L3 기반 경로 조합으로 대체한다.
const mergeHierarchySubheader = (rows, headerIdx, colMap) => {
    const subRow = rows[headerIdx + 1] || [];
    const subCols = {};
    subRow.forEach((cell, ci) => {
        const field = matchAlias(cell);
        if (field && ['업무시스템', 'L1', 'L2', 'L3'].includes(field)) subCols[field] = ci;
    });
    const hasHierarchy = ['L1', 'L2', 'L3'].some(k => k in subCols);
    if (!hasHierarchy) return { colMap, dataStart: headerIdx + 1 };
    const merged = { ...colMap };
    delete merged.업무구분;
    Object.assign(merged, subCols);
    return { colMap: merged, dataStart: headerIdx + 2 };
};

const buildBizScope = (row, colMap) => {
    if (['L1', 'L2', 'L3'].some(k => k in colMap)) {
        const parts = ['L1', 'L2', 'L3']
            .map(k => k in colMap ? String(row[colMap[k]] ?? '').trim() : '')
            .filter(Boolean);
        return parts.join(' > ');
    }
    if ('업무구분' in colMap) return String(row[colMap.업무구분] ?? '').trim();
    return '';
};

// 기능/비기능 구분: 컬럼값이 있으면 그 값을 그대로 사용(예: "인터페이스" 등 비이분법 값도 그대로 보존),
// 컬럼이 없으면 시트명으로 판단. "비기능"이 "기능"의 부분 문자열이므로 비기능을 먼저 검사한다.
const deriveType = (row, colMap, sheetName) => {
    if ('구분' in colMap) {
        const v = String(row[colMap.구분] ?? '').trim();
        if (v) return v;
    }
    if (/비기능/.test(sheetName)) return '비기능';
    if (/기능/.test(sheetName)) return '기능';
    return '';
};

// XLSX 워크북 객체를 받아 PM 모드 요구사항 배열로 변환하는 순수 함수.
// 브라우저(FileReader)와 Node(fs) 양쪽에서 동일 로직을 재사용할 수 있도록 워크북 단위로 분리.
export const parsePmWorkbook = (workbook) => {
    const allReqs = [];
    for (const sheetName of workbook.SheetNames) {
        if (/리스크|risk/i.test(sheetName)) continue;
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        if (rows.length < 2) continue;

        const detected = detectHeaderRow(rows);
        if (!detected || detected.score === 0) continue;
        if (!('id' in detected.cols) && !('title' in detected.cols)) continue;

        const { colMap, dataStart } = mergeHierarchySubheader(rows, detected.idx, detected.cols);
        if (!('id' in colMap) && !('title' in colMap)) continue;

        for (let r = dataStart; r < rows.length; r++) {
            const row = rows[r];
            if (!row || row.length === 0) continue;
            const id = 'id' in colMap ? String(row[colMap.id] ?? '').trim() : '';
            const title = 'title' in colMap ? String(row[colMap.title] ?? '').trim() : '';
            if (!id && !title) continue; // 빈 행

            allReqs.push({
                id: id || `${PM_ID_PREFIX}-${r}`,
                title: title || '-',
                detail: 'detail' in colMap ? (String(row[colMap.detail] ?? '').trim() || '-') : '-',
                업무시스템: '업무시스템' in colMap ? String(row[colMap.업무시스템] ?? '').trim() : '',
                업무구분: buildBizScope(row, colMap),
                구분: deriveType(row, colMap, sheetName),
                pm_mode: true,
                _source_sheet: sheetName,
            });
        }
    }
    return allReqs;
};

// 브라우저용 래퍼: File 객체를 읽어 parsePmWorkbook으로 위임 (UIUX 모드의
// parseExcelToRequirements와 동일한 형태로, 다음 단계에서 프론트 연동 시 그대로 사용 가능)
export const parsePmExcelToRequirements = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                resolve(parsePmWorkbook(workbook));
            } catch (err) { reject(err); }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
};

// [PM 모드 Task1 시스템 프롬프트]
// 기존 UIUX 모드 coreSystemPrompt(App.jsx)와 별개의 프롬프트. UIUX 관련성 판별 로직은
// 완전히 제거했고(모든 기능/비기능 요구사항을 필터링 없이 통과), 모호 표현 구체화 지침은
// UI/UX 영역에 국한하지 않고 성능·보안·데이터정합성 등 모든 도메인으로 일반화했다.
export const pmCoreSystemPrompt = `당신은 14년 경력의 SI PM(프로젝트 관리자)으로,
RFP 및 요구사항 문서에서 모든 기능/비기능 요구사항을 필터링 없이 추출하고 모호한 표현을 구체화하는 전문가임.

[필터링 원칙 — 반드시 준수]
어떠한 요구사항 항목도 도메인·업무구분·ID 접두사를 이유로 제외하지 않는다.
UI/UX, 데이터 아키텍처, 보안, 인프라, 성능, PM 영역(WBS·일정·예산), 컴플라이언스 등
모든 영역의 기능/비기능 요구사항을 예외 없이 전부 결과에 포함한다.

[모호 표현 구체화 — 반드시 적용, 모든 도메인 대상]
UI/UX뿐 아니라 성능·보안·데이터정합성·인프라 등 모든 도메인의 추상적 표현을 측정 가능한
정량 수치로 재서술한다.
- (UI 예시) "직관적인 UI" → "3클릭 이내 목표 달성 구조"
- (UI 예시) "사용자 편의성 고려" → "접근성 WCAG 2.1 AA 준수"
- (성능 예시) "빠른 응답" → "API P95 응답 3초 이내"
- (성능 예시) "대용량 처리" → "동시접속 1,000건, TPS 500 이상 처리"
- (보안 예시) "안전한 인증" → "MFA 적용, 세션 타임아웃 30분"
- (데이터정합성 예시) "데이터 정합성 확보" → "일 배치 정합성 검증 오차율 0.01% 이내"
- 위 예시처럼 도메인에 관계없이 측정 가능한 기준으로 반드시 재서술할 것
- 모호 표현 탐지 대상: 수치 불명확("일부", "적절히", "빠르게", "충분히" 등),
  범위 불명확("관련 데이터", "필요한 경우", "기타", "등" 등),
  기준 불명확("최적화된", "효율적인", "안정적인" 등)
- 모호 표현 보정 시 반드시:
  ambiguity_hitl.original_text: 원문 모호 표현 (보정 여부와 무관하게 반드시 채울 것)
  ambiguity_hitl.corrected_text: 보정된 구체적 표현
  ambiguity_hitl.ambiguity_reason: 왜 모호한지 사유
  자동 보정 완료 → is_ambiguous: false, corrected_text 포함
  보정 불가 → is_ambiguous: true
- 모호 표현이 없는 항목도 original_text에 해당 요구사항의 핵심 문장을 기재할 것

[데이터 무결성 규칙]
1. RFP 문서에는 요구사항을 식별하는 고유 ID 컬럼이 존재한다.
   컬럼명은 "고유번호", "요구사항ID", "고유ID", "번호" 등 RFP마다 다를 수 있다.
   해당 컬럼을 문서에서 스스로 찾아서, 그 값을 원본ID로 그대로 사용할 것.
   절대 임의로 생성, 변환, 채번하지 말 것.
2. 원본 문서에 고유 ID 컬럼 자체가 없어 식별할 ID가 전혀 없는 경우에만
   "${PM_ID_PREFIX}-{순번}" 형식으로 순번을 매겨 생성할 것 (그 외의 경우 임의 채번 금지).
3. 물리적 행 순서 유지, 임의 정렬 금지
4. JSON 스키마 외 텍스트 추가 금지
5. 상세내용은 핵심만 간결하게 요약하여 출력 제한 방지

[업무 계층 참고 기준]
- 업무시스템: 요구사항이 속한 시스템/서비스명 (문서에 명시된 경우 그대로 사용, 없으면 빈 문자열)
- 업무구분: 요구사항의 업무 분류 경로 (문서에 명시된 경우 그대로 사용, 없으면 빈 문자열)

[기능/비기능 구분 판단 기준]
- 문서에 기능/비기능을 나타내는 컬럼(예: 기능유형, 구분)이 있으면 그 값을 그대로 사용
- 컬럼이 없고 문서가 "기능요구사항"/"비기능요구사항"처럼 섹션으로 분리되어 있으면 그 구분을 따를 것
- 판단 근거가 전혀 없으면 빈 문자열로 둘 것 (임의 추정 금지)

[우선순위 판단 기준 — 반드시 "상"/"중"/"하" 중 하나를 채울 것. 빈값 금지]
- 상: 서비스 핵심 기능, 사용자 직접 체감, 미구현 시 서비스 불가
- 중: 품질/편의 향상, 없어도 서비스 가능하나 사용성 저하
- 하: 부가 기능, 향후 개선 가능, 미구현 시 영향 미미
`;

// [PM 모드 Task1 스키마]
// UIUX 모드 schema1과 달리 uiux_relevant 플래그와 excluded_reqs가 없다 — 모든 항목을 통과시키므로
// "제외" 개념 자체가 존재하지 않는다.
export const pmSchema1 = `{
  "requirements": [
    {
      "id": "원본ID", "title": "요구사항명", "detail": "내용",
      "구분": "기능/비기능 (또는 문서상 원문 값)",
      "우선순위": "상/중/하",
      "업무시스템": "", "업무구분": "",
      "ambiguity_hitl": {
        "is_ambiguous": false,
        "original_text": "원문 모호 표현 (반드시 채울 것)",
        "corrected_text": "보정된 구체적 표현",
        "ambiguity_reason": "보정 사유"
      }
    }
  ]
}`;

const PM_CHUNK_SIZE = 15000;

// PM 모드 Task1 추출 실행기.
// callBackendAPI는 App.jsx의 동일 이름 함수와 시그니처가 같은 함수를 그대로 주입받아 사용한다
// (callBackendAPI(promptData, systemInstruction, schemaDefinition, pdfFiles, temperature)).
// 엑셀 파일은 UIUX 모드와 동일하게 결정론적으로 파싱해 LLM 호출 없이 처리한다.
export const runPmTask1 = async ({ combinedText = '', excelFiles = [], pdfFilesB64 = [], callBackendAPI }) => {
    const merged = [];

    if (excelFiles.length > 0) {
        for (const ef of excelFiles) {
            const parsed = await parsePmExcelToRequirements(ef);
            merged.push(...parsed);
        }
        if (combinedText.trim()) {
            const result = await callBackendAPI(combinedText, pmCoreSystemPrompt, pmSchema1, [], 0);
            merged.push(...(result.requirements || []));
        }
    } else if (pdfFilesB64.length > 0) {
        const result = await callBackendAPI(combinedText, pmCoreSystemPrompt, pmSchema1, pdfFilesB64, 0);
        merged.push(...(result.requirements || []));
    } else if (combinedText.length <= PM_CHUNK_SIZE) {
        const result = await callBackendAPI(combinedText, pmCoreSystemPrompt, pmSchema1, [], 0);
        merged.push(...(result.requirements || []));
    } else {
        const lines = combinedText.split('\n');
        const chunks = [];
        let current = '';
        for (const line of lines) {
            if (current.length + line.length + 1 > PM_CHUNK_SIZE && current.length > 0) {
                chunks.push(current);
                current = '';
            }
            current += (current ? '\n' : '') + line;
        }
        if (current) chunks.push(current);

        for (let i = 0; i < chunks.length; i++) {
            const chunkPrompt = `[청크 ${i + 1}/${chunks.length}] 아래는 전체 요구사항 중 일부입니다. 이 부분만 분석하세요.\n\n${chunks[i]}`;
            const result = await callBackendAPI(chunkPrompt, pmCoreSystemPrompt, pmSchema1, [], 0);
            merged.push(...(result.requirements || []));
        }
    }

    return { requirements: merged };
};
