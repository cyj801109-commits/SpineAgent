import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';

const MAX_EXCEL_ROWS = 300;
const MAX_COMBINED_CHARS = 80000;

// --- All Icons Defined (inline SVG, same as original) ---
const Layers = ({className, size=20}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>;
const Play = ({size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>;
const RefreshCw = ({className, size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>;
const Plus = ({className, size=20}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>;
const X = ({className, size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>;
const FileSpreadsheet = ({className, size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M8 13h2"></path><path d="M8 17h2"></path><path d="M14 13h2"></path><path d="M14 17h2"></path></svg>;
const ShieldCheck = ({className, size=18}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path><path d="M9 12l2 2 4-4"></path></svg>;
const Maximize2 = ({size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 3 21 3 21 9"></polyline><polyline points="9 21 3 21 3 15"></polyline><line x1="21" y1="3" x2="14" y2="10"></line><line x1="3" y1="21" x2="10" y2="14"></line></svg>;
const Monitor = ({size=20}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect><line x1="8" y1="21" x2="16" y2="21"></line><line x1="12" y1="17" x2="12" y2="21"></line></svg>;
const AlertTriangle = ({className, size=10}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>;
const FileText = ({size=12}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>;
const SettingsIcon = () => <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>;
const LinkIcon = ({size=12}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg>;
const CheckCircle = ({className, size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>;
const Loader2 = ({className, size=32}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><line x1="12" y1="2" x2="12" y2="6"></line><line x1="12" y1="18" x2="12" y2="22"></line><line x1="4.93" y1="4.93" x2="7.76" y2="7.76"></line><line x1="16.24" y1="16.24" x2="19.07" y2="19.07"></line><line x1="2" y1="12" x2="6" y2="12"></line><line x1="18" y1="12" x2="22" y2="12"></line><line x1="4.93" y1="19.07" x2="7.76" y2="16.24"></line><line x1="16.24" y1="7.76" x2="19.07" y2="4.93"></line></svg>;
const AlertOctagon = ({className, size=48}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><polygon points="7.86 2 16.14 2 22 7.86 22 16.14 16.14 22 7.86 22 2 16.14 2 7.86 7.86 2"></polygon><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>;
const Edit2 = ({size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path></svg>;
const Download = ({size=14}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>;

const App = () => {
    const [apiModel, setApiModel] = useState('gemini-2.5-flash');
    const [showSettings, setShowSettings] = useState(false);
    const [inputText, setInputText] = useState('');
    const [isAnalyzing, setIsAnalyzing] = useState(false);
    const [uploadedFiles, setUploadedFiles] = useState([]);

    const [progressStep, setProgressStep] = useState(0);
    const [stepElapsed, setStepElapsed] = useState(0);
    const stepTimerRef = useRef(null);

    // Core Data States
    const [metrics, setMetrics] = useState(null);
    const [rawFunc, setRawFunc] = useState([]);
    const [rawNonFunc, setRawNonFunc] = useState([]);
    const [optimizedReqs, setOptimizedReqs] = useState([]);
    const [conflicts, setConflicts] = useState([]);

    // Error & Warning State
    const [errorMessage, setErrorMessage] = useState(null);
    const [truncationWarning, setTruncationWarning] = useState(null);

    // Tab State
    const [activeTab, setActiveTab] = useState('요구사항정의서');
    const [selectedItem, setSelectedItem] = useState(null);
    const [isFullScreen, setIsFullScreen] = useState(false);

    // Modal Edit States
    const [isCustomEditing, setIsCustomEditing] = useState(false);
    const [customOptionText, setCustomOptionText] = useState('');

    const fileInputRef = useRef(null);

    const startStepTimer = () => {
        if (stepTimerRef.current) clearInterval(stepTimerRef.current);
        setStepElapsed(0);
        stepTimerRef.current = setInterval(() => {
            setStepElapsed(prev => prev + 1);
        }, 1000);
    };

    const processingSteps = [
        { title: "데이터 파싱", desc: "텍스트 및 엑셀 데이터 추출 준비" },
        { title: "관계망 구축 준비", desc: "파싱된 데이터 구조화" },
        { title: "UIUX 관련성 3단계 판단", desc: "데이터 아키텍처 제외 + 화면 설계 대상 확정 (API Task 1)" },
        { title: "UIUX 선별 완료", desc: "UIUX 관련/제외 요구사항 분리 완료" },
        { title: "리소스 최적화 및 모호성 분석", desc: "담당 범위(직접담당·협의필요·인지필요) 분류 + 모호성 구체화 (API Task 2)" },
        { title: "최적화 완료", desc: "표준 양식 매핑 완료" },
        { title: "충돌·전제조건·유사 관계 분석", desc: "충돌 감지 + 선행 요구사항 도출 + 유사 항목 그룹핑 (API Task 3)" },
        { title: "검토 지점 설정 완료", desc: "HITL 의사결정 노드 구성 완료" }
    ];

    useEffect(() => {
        const savedModel = localStorage.getItem('SPINE_API_MODEL');
        if (savedModel) setApiModel(savedModel);
    }, []);

    // 모달 초기화
    useEffect(() => {
        setIsCustomEditing(false);
        setCustomOptionText('');
    }, [selectedItem]);

    const handleSaveSettings = () => {
        localStorage.setItem('SPINE_API_MODEL', apiModel);
        setShowSettings(false);
    };

    const handleFileUpload = (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;
        const newFiles = files.map(file => ({ name: file.name, id: Date.now() + Math.random(), originFile: file }));
        setUploadedFiles(prev => [...prev, ...newFiles]);
        if (fileInputRef.current) fileInputRef.current.value = "";
    };

    const readFileAsBase64 = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result.split(',')[1]);
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    };

    const parseFileContent = async (file) => {
        const ext = file.name.split('.').pop().toLowerCase();
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    if (ext === 'xlsx' || ext === 'xls') {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                        const json = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                        const totalRows = json.length;
                        const rows = json.slice(0, MAX_EXCEL_ROWS);
                        let result = rows.map(row => row.join('\t')).join('\n');
                        if (totalRows > MAX_EXCEL_ROWS) {
                            result += `\n\n... (이하 ${totalRows - MAX_EXCEL_ROWS}행 생략 - 토큰 제한으로 상위 ${MAX_EXCEL_ROWS}행만 분석)`;
                            setTruncationWarning(prev => prev || `엑셀: 상위 ${MAX_EXCEL_ROWS}행만 분석`);
                        }
                        resolve(result);
                    } else {
                        resolve(e.target.result);
                    }
                } catch (err) { reject(err); }
            };
            if (ext === 'xlsx' || ext === 'xls') reader.readAsArrayBuffer(file);
            else reader.readAsText(file, 'UTF-8');
        });
    };

    // 엑셀 내보내기 (SheetJS 활용)
    const exportToExcel = () => {
        const wb = XLSX.utils.book_new();

        if (optimizedReqs.length > 0) {
            const optData = optimizedReqs.map(r => ({
                "NO": r.NO,
                "요구사항ID": r.요구사항ID,
                "업무분류": r.업무분류,
                "업무_대": r.업무_대,
                "기능_중": r.기능_중,
                "구성_소": r.구성_소,
                "요구정의명": r.요구정의명,
                "고객_요구사항_상세_내용": r.고객_요구사항_상세_내용,
                "연관 요구사항 및 HITL 근거": r.related_reqs?.map(rel => `[${rel.relation_type}] ${rel.target_id}`).join(', ') || '',
                "제약사항": r.제약사항 || '',
                "요건_발생일": r.요건_발생일,
                "분류": r.요구사항유형분류?.분류 || '',
                "유형": r.요구사항유형분류?.유형 || '',
                "우선순위": r.우선순위
            }));
            const wsOpt = XLSX.utils.json_to_sheet(optData);
            XLSX.utils.book_append_sheet(wb, wsOpt, "요구사항정의서");
        }

        if (rawFunc.length > 0) {
            const wsFunc = XLSX.utils.json_to_sheet(rawFunc.map(r => ({"원본 ID": r.id, "요구사항명": r.title, "상세내용": r.detail})));
            XLSX.utils.book_append_sheet(wb, wsFunc, "기능");
        }

        if (rawNonFunc.length > 0) {
            const wsNonFunc = XLSX.utils.json_to_sheet(rawNonFunc.map(r => ({"원본 ID": r.id, "요구사항명": r.title, "상세내용": r.detail})));
            XLSX.utils.book_append_sheet(wb, wsNonFunc, "비기능");
        }

        if (conflicts.length > 0) {
            const wsConflicts = XLSX.utils.json_to_sheet(conflicts.map(c => ({"충돌 ID": c.conflict_id, "관련 요구사항": c.involved_req_ids?.join(', '), "충돌 사유": c.conflict_reason})));
            XLSX.utils.book_append_sheet(wb, wsConflicts, "충돌");
        }

        if(wb.SheetNames.length === 0) {
            alert("추출할 분석 데이터가 없습니다.");
            return;
        }

        XLSX.writeFile(wb, "요구사항정의서_결과.xlsx");
    };

    // HITL에서 PM이 채택/수정한 값을 실시간 요구사항정의서(optimizedReqs)에 반영
    const handleHitlDecision = (item, selectedOption, hitlType) => {
        if (hitlType === 'conflict') {
            setConflicts(prev => prev.filter(c => c.conflict_id !== item.conflict_id));
            const newDecisionReq = {
                NO: optimizedReqs.length + 1,
                요구사항ID: `REQ-HITL-${String(optimizedReqs.length + 1).padStart(3, '0')}`,
                업무분류: "의사결정", 업무_대: "정책확정", 기능_중: "HITL", 구성_소: "충돌해결",
                요구정의명: `[확정] ${item.conflict_reason.substring(0, 20)}...`,
                고객_요구사항_상세_내용: `PM 의사결정으로 정책 확정: ${selectedOption}`,
                제약사항: `관련 요건: ${item.involved_req_ids.join(', ')}`,
                요건_발생일: new Date().toISOString().split('T')[0],
                요구사항유형분류: { 분류: "기능", 유형: "정책" },
                우선순위: "상",
                related_reqs: item.involved_req_ids.map(id => ({ target_id: id, relation_type: "통합", reason: "충돌 해결 및 단일화" }))
            };
            setOptimizedReqs(prev => [...prev, newDecisionReq]);
        } else if (hitlType === 'ambiguity') {
            setOptimizedReqs(prev => prev.map(req => {
                if (req.요구사항ID === item.요구사항ID) {
                    return {
                        ...req,
                        고객_요구사항_상세_내용: `[정량화 확정] ${selectedOption}`,
                        ambiguity_hitl: { ...req.ambiguity_hitl, is_ambiguous: false }
                    };
                }
                return req;
            }));
        }
        setSelectedItem(null);
    };

    const callBackendAPI = async (promptData, systemInstruction, schemaDefinition, pdfFiles = []) => {
        let retries = 3;
        let delay = 2000;
        while (retries > 0) {
            try {
                const response = await fetch('/api/analyze', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        model: apiModel,
                        prompt: promptData,
                        systemInstruction: systemInstruction,
                        schema: schemaDefinition,
                        pdf_files: pdfFiles
                    })
                });

                if (!response.ok) {
                    const err = await response.json().catch(() => ({}));
                    throw new Error(err.detail || "API 통신 오류");
                }

                const data = await response.json();
                const text = data.text;

                const firstBrace = text.indexOf('{');
                const lastBrace = text.lastIndexOf('}');

                if (firstBrace !== -1 && lastBrace !== -1) {
                    const resultText = text.substring(firstBrace, lastBrace + 1);
                    try {
                        return JSON.parse(resultText);
                    } catch (parseErr) {
                        throw new Error(`JSON 파싱 실패: ${parseErr.message}`);
                    }
                } else {
                    throw new Error("AI 파이프라인 응답 파싱 실패 (유효한 JSON 아님)");
                }
            } catch (error) {
                retries--;
                if (retries === 0) {
                    if (error.message === "MAX_TOKENS_REACHED" || error.message.includes("JSON 파싱 실패")) {
                        throw new Error("[출력 한도 초과] 분석할 데이터가 너무 많아 AI 응답이 중간에 끊겼습니다. 요구사항을 15~20개씩 나누어서 가동해 주세요.");
                    }
                    const msg = error.message.toLowerCase();
                    if (msg.includes("quota") || msg.includes("429")) throw new Error("[API 할당량 초과] 제공량이 소진되었습니다.");
                    throw error;
                }
                await new Promise(res => setTimeout(res, delay));
                delay *= 1.5;
            }
        }
    };

    const analyzeLogic = async () => {
        setIsAnalyzing(true);
        setErrorMessage(null);
        setMetrics(null);
        setTruncationWarning(null);

        try {
            setProgressStep(0);
            let combinedText = inputText;
            const pdfFilesB64 = [];
            if (uploadedFiles.length > 0) {
                for (const f of uploadedFiles) {
                    if (!f.originFile) continue;
                    const ext = f.name.split('.').pop().toLowerCase();
                    if (ext === 'pdf') {
                        const b64 = await readFileAsBase64(f.originFile);
                        pdfFilesB64.push({ name: f.name, data: b64 });
                    } else {
                        const content = await parseFileContent(f.originFile);
                        combinedText += `\n\n--- [FILE: ${f.name}] ---\n${content}`;
                    }
                }
            }
            if (!combinedText.trim() && pdfFilesB64.length === 0) throw new Error("분석할 요구사항 데이터가 없습니다. 파일을 업로드하거나 텍스트를 입력해 주세요.");
            if (combinedText.length > MAX_COMBINED_CHARS) {
                combinedText = combinedText.substring(0, MAX_COMBINED_CHARS) + '\n\n[입력 데이터가 너무 커서 일부가 잘렸습니다. 파일을 분할하여 업로드해주세요.]';
                setTruncationWarning(prev => prev ? prev + ' / 전체 입력 80,000자 초과로 잘림' : '전체 입력 80,000자 초과로 일부 잘림');
            }
            setProgressStep(1);

            const coreSystemPrompt = `당신은 14년 경력의 SI UIUX 기획자로,
RFP 및 요구사항 문서에서 UIUX 관련 항목을 선별하고 기획자 언어로 구체화하는 전문가임.

[UIUX 관련성 판단 3단계 로직 — 반드시 준수]
STEP 1. Out of Scope 판단: 아래 항목은 UIUX 범위 외로 1차 제외한다.
- PM 영역: 디자인시스템 정의, WBS, 일정, 예산
- 컴플라이언스: 개인정보처리방침, 보안정책
- 인프라: 서버 튜닝, 수수료 계산, 배치
- 데이터 아키텍처: DB 설계, 데이터 표준화, 데이터 모델링, 데이터 스키마·메타 정의,
  ETL 개발·운영, 데이터 수집·연계·마이그레이션·품질검증, 데이터 아키텍처 설계
  ※ 단, 해당 항목이 "화면에서 어떻게 보여줄지(시각화 방식, 컴포넌트 구조)"를 포함하는 경우는
  STEP 2에서 복구 가능

STEP 2. Context Recovery: STEP 1 제외 항목이라도 아래 UI 키워드 포함 시
'Conditional'로 분류하여 STEP 3에서 재판단한다.
- 복구 키워드: 화면, 조회UI, 노출, 버튼, 클릭, 팝업, 모달, 입력폼, 레이아웃

STEP 3. 최종 판단 기준:
"UIUX 기획자가 이 항목을 위해 화면 설계서(와이어프레임·스토리보드)를 직접 작성해야 하는가?"
- 화면 설계서가 산출물 → 포함 (O)
- 데이터 설계서·아키텍처 문서·API 명세서가 산출물 → 제외 (X)
- 화면 설계와 데이터 설계가 혼재 → review_role: "협의필요"로 포함하되 표시

[요구사항 최적화 판단 — 4가지 Candidate 유형]
선별된 UIUX 요구사항에 대해 아래 유형으로 분류하여 related_reqs에 반영한다.
- Merge Candidate: 유사 기능·공통 정책으로 통합 가능한 요구사항 → relation_type: "통합"
- Split Candidate: 서로 다른 화면·담당자가 혼재된 요구사항 → relation_type: "연결"
- Remove Candidate: 완전 중복·UIUX 범위 외 → excluded_reqs로 분류
- Re-scope Candidate: 현재 범위 초과·의존성 미충족 → 제약사항 필드에 명시

[FO/BO 구분 기준]
- FO(Front Office): 일반 사용자가 직접 접근하는 화면
- BO(Back Office): 관리자/운영자가 사용하는 화면
- 판단 불가 시 TBD 표기

[화면 본수 추정 기준]
- 단일 기능 단순 화면: 1본
- 탭/단계 구분 있는 화면: 2~3본
- 목록+상세 구조: 2본
- 판단 불가 시 TBD 표기

[모호 표현 구체화 — 반드시 적용]
- "직관적인 UI" → "3클릭 이내 목표 달성 구조"
- "사용자 편의성 고려" → "접근성 WCAG 2.1 AA 준수"
- "최적화된 폼" → "입력 필드 자동완성, 실시간 유효성 검사"
- "빠른 응답" → "API P95 응답 3초 이내"
- 위 예시처럼 측정 가능한 기준으로 반드시 재서술할 것

[데이터 무결성 규칙]
1. 원본 ID 100% 그대로 추출, 임의 변경 금지
2. 물리적 행 순서 유지, 임의 정렬 금지
3. JSON 스키마 외 텍스트 추가 금지
4. 상세내용은 핵심만 간결하게 요약하여 출력 제한 방지

[review_role 판단 기준]
- 직접담당: 업무분류가 UI/UX이거나 UIUX 기획자가 화면 설계서를 직접 작성하는 항목
- 협의필요: 다른 팀 주담당이지만 화면 영향도 있어 UIUX 검토가 필요한 항목
  (예: 알림 정책 → 알림 UI 설계에 영향, 데이터 시각화 방식 결정 등)
- 인지필요: 화면 직접 관련은 없으나 성능·접근성 기준 등 UIUX가 알아야 할 항목

[relation_type 정의]
- 중복: 내용이 거의 동일하여 하나로 합칠 수 있는 요구사항
- 통합: 방향이 같아 묶어서 처리 가능한 요구사항
- 연결: 서로 다른 화면·담당이지만 연동되는 요구사항
- 충돌: 내용이 서로 모순되거나 구현 방향이 상반되는 요구사항
- 전제조건: 해당 요구사항 구현 전에 반드시 완료되어야 하는 요구사항
- 유사: 내용이 비슷하지만 범위나 대상이 달라 통합 전 검토가 필요한 요구사항

[요구사항 ID 명명 규칙]
구조: REQ-{분류영문}-{Level1(2자리)}-{Level2(3자리)}
- FO-USR(01), FO-ADM(02), BO-ADM(03), BO-SRV(04), CMN(05), NFR(06)
- Level2: 01(로그인), 02(로그아웃), 03(메인화면), 04(통합검색), 05(MySpace),
  06(프로젝트수정), 07(프로젝트정보), 08(게시판), 09(Help/산출물),
  10(코드관리), 11(권한관리), 35(성능), 37(산출물) — 문맥에 맞게 유추 할당`;

            // Task 1: 추출 에이전트
            setProgressStep(2); startStepTimer();
            const schema1 = `{
  "uiux_functional_reqs": [
    { "id": "원본ID", "title": "요구사항명", "detail": "내용", "uiux_relevant": true }
  ],
  "excluded_reqs": [
    { "id": "원본ID", "title": "요구사항명", "exclude_reason": "제외 사유" }
  ]
}`;

            const extractedData = await callBackendAPI(combinedText, coreSystemPrompt, schema1, pdfFilesB64);
            setRawFunc(extractedData.uiux_functional_reqs || []);
            setRawNonFunc(extractedData.excluded_reqs || []);
            setProgressStep(3);

            // Task 2: 최적화 및 구체화 에이전트
            setProgressStep(4); startStepTimer();
            const schema2 = `{
  "optimization_summary": {
    "total_input_count": 0,
    "uiux_selected_count": 0,
    "excluded_count": 0,
    "ambiguity_resolved_count": 0,
    "selection_reason": "선별 판단 근거 요약"
  },
  "optimized_requirements": [
    {
      "NO": 1,
      "요구사항ID": "REQ-FO-USR-01-001",
      "fo_bo": "FO/BO/TBD",
      "review_role": "직접담당/협의필요/인지필요",
      "화면본수": "1본/TBD",
      "업무분류": "대분류",
      "업무_대": "업무명",
      "기능_중": "기능명",
      "구성_소": "상세구성",
      "요구정의명": "명확한 기능명",
      "고객_요구사항_상세_내용": "모호성 제거된 구체적 요건",
      "uiux_relevance_step": "STEP1통과/STEP2복구/STEP3확정",
      "제약사항": "제약사항 또는 Re-scope 사유",
      "요건_발생일": "YYYY-MM-DD",
      "요구사항유형분류": { "분류": "기능/비기능", "유형": "세부유형" },
      "우선순위": "상/중/하",
      "related_reqs": [
        {
          "target_id": "원본ID",
          "relation_type": "중복/통합/연결/충돌/전제조건/유사",
          "reason": "Merge·Split·Remove·Re-scope 판단 근거"
        }
      ],
      "ambiguity_hitl": {
        "is_ambiguous": false,
        "original_text": "원문 그대로 보존",
        "ambiguity_reason": "모호 사유",
        "suggested_options": ["정량 대안1", "정량 대안2"]
      }
    }
  ]
}`;

            const optimizedData = await callBackendAPI(JSON.stringify(extractedData), coreSystemPrompt, schema2);
            setMetrics(optimizedData.optimization_summary);
            setOptimizedReqs(optimizedData.optimized_requirements || []);
            setProgressStep(5);

            // Task 3: 정책 충돌 + 전제조건 + 유사 관계 검증 에이전트
            // [prerequisite_relations 판단 기준]
            // A가 B의 전제조건 → B 구현 전 A 완료 필요. 예: 표준 정의 → 구현 항목들
            // [similar_reqs 판단 기준]
            // 동일 기능명·유사 내용이지만 범위/대상이 달라 통합 여부를 사람이 판단해야 하는 것
            setProgressStep(6); startStepTimer();
            const schema3 = `{
  "conflicts": [
    {
      "conflict_id": "C-001",
      "involved_req_ids": ["충돌 원본 ID 배열"],
      "conflict_reason": "충돌 사유 및 리스크",
      "option_a": { "description": "...", "risk": "..." },
      "option_b": { "description": "...", "risk": "..." },
      "recommendation": "SPINE 추천안 및 근거",
      "impact_scope": "영향받는 연관 요구사항 범위"
    }
  ],
  "prerequisite_relations": [
    {
      "relation_id": "PR-001",
      "prerequisite_id": "선행 완료 필요 요구사항 ID",
      "dependent_ids": ["이 요구사항 구현 전 선행 필요한 요구사항 ID들"],
      "reason": "선행 완료가 필요한 이유"
    }
  ],
  "similar_reqs": [
    {
      "relation_id": "SR-001",
      "req_ids": ["유사 요구사항 ID들"],
      "similarity_summary": "유사한 내용 요약",
      "action": "통합검토필요/분리확인필요"
    }
  ]
}`;

            const conflictData = await callBackendAPI(JSON.stringify(optimizedData.optimized_requirements), coreSystemPrompt, schema3);
            setConflicts(conflictData.conflicts || []);

            if (optimizedData.optimization_summary) {
                setMetrics({
                    ...optimizedData.optimization_summary,
                    conflict_req_count: conflictData.conflicts?.length || 0
                });
            }
            setProgressStep(7);

            setActiveTab('요구사항정의서');
        } catch (e) {
            setErrorMessage(`[Step ${progressStep + 1} 실패] ${e.message}`);
        } finally {
            setIsAnalyzing(false);
            if (stepTimerRef.current) clearInterval(stepTimerRef.current);
            setStepElapsed(0);
        }
    };

    const renderBadges = (relations) => {
        if (!relations || relations.length === 0) return <span className="text-sub/40">-</span>;
        const uniqueRels = [];
        const seenId = new Set();
        relations.forEach(r => {
            if(!seenId.has(r.target_id)){
                uniqueRels.push(r);
                seenId.add(r.target_id);
            }
        });

        return (
            <div className="flex flex-wrap gap-1 justify-center">
                {uniqueRels.map((rel, idx) => (
                    <span key={idx} className="px-2 py-0.5 rounded text-[10px] font-bold flex items-center gap-1 bg-white text-primary border border-borderline">
                        <LinkIcon size={10}/> {rel.relation_type}: {rel.target_id}
                    </span>
                ))}
            </div>
        );
    };

    const renderRawTable = (data) => (
        <table className="w-full text-left text-xs border-collapse min-w-[800px] bg-white text-primary">
            <thead className="sticky top-0 bg-pagebg text-[11px] uppercase tracking-widest text-sub border-b border-borderline z-10 font-bold">
                <tr><th className="p-4 w-32 border-r border-borderline">원본 ID</th><th className="p-4 w-64 border-r border-borderline">요구사항명</th><th className="p-4">상세내용</th></tr>
            </thead>
            <tbody className="divide-y divide-borderline">
                {data.map((r, i) => (
                    <tr key={i} className="hover:bg-pagebg transition-colors">
                        <td className="p-4 font-mono font-bold text-primary border-r border-borderline tracking-wider">{r.id}</td>
                        <td className="p-4 font-bold text-primary border-r border-borderline">{r.title}</td>
                        <td className="p-4 text-primary leading-relaxed">{r.detail}</td>
                    </tr>
                ))}
            </tbody>
        </table>
    );

    const renderConflictsTable = () => (
        <table className="w-full text-left text-xs border-collapse min-w-[800px] bg-white text-primary">
            <thead className="sticky top-0 bg-pagebg shadow-sm z-10 text-[11px] uppercase tracking-widest text-sub border-b border-borderline font-bold">
                <tr><th className="p-4 w-32 border-r border-borderline">충돌 ID</th><th className="p-4 w-48 border-r border-borderline">관련 요구사항</th><th className="p-4 border-r border-borderline">충돌 사유</th><th className="p-4 w-32 text-center">의사결정</th></tr>
            </thead>
            <tbody className="divide-y divide-borderline">
                {conflicts.map((c, i) => (
                    <tr key={i} className="hover:bg-pagebg cursor-pointer transition-colors group" onClick={() => setSelectedItem({type: 'conflict', data: c})}>
                        <td className="p-4 font-bold text-primary tracking-wider border-r border-borderline group-hover:text-accent transition-colors">{c.conflict_id}</td>
                        <td className="p-4 font-mono text-primary text-[10px] font-bold border-r border-borderline">{c.involved_req_ids.join(', ')}</td>
                        <td className="p-4 text-primary border-r border-borderline leading-relaxed">{c.conflict_reason}</td>
                        <td className="p-4 text-center"><span className="bg-rose-500 text-white px-3 py-1 rounded text-[10px] font-bold shadow-sm flex items-center justify-center gap-1"><AlertTriangle size={10}/> HITL 요망</span></td>
                    </tr>
                ))}
            </tbody>
        </table>
    );

    return (
        <div className="min-h-screen flex flex-col w-full font-sans bg-pagebg text-primary">
            <header className="flex items-center justify-between bg-white px-6 py-4 border-b border-borderline shadow-sm w-full z-50 shrink-0">
                <div className="flex items-center gap-4 text-primary">
                    <div className="w-10 h-10 bg-primary rounded flex items-center justify-center text-white"><Layers /></div>
                    <div>
                        <h1 className="text-xl tracking-tight leading-none font-bold">UIUX 요구사항 선별 에이전트</h1>
                        <p className="text-[10px] text-sub tracking-widest uppercase mt-1 font-bold">UIUX 검토 대상 필터링</p>
                    </div>
                </div>
                <div className="flex items-center gap-4">
                    <button onClick={() => window.open('/spine-pipeline.html', '_blank')} className="flex items-center gap-2 bg-white hover:bg-pagebg px-3 py-1.5 rounded border border-borderline text-sub hover:text-primary transition-all active:scale-95">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="6" y1="3" x2="6" y2="15"></line><circle cx="18" cy="6" r="3"></circle><circle cx="6" cy="18" r="3"></circle><path d="M18 9a9 9 0 0 1-9 9"></path></svg>
                        <span className="text-[10px] font-bold uppercase tracking-widest">Pipeline</span>
                    </button>
                    <div className="flex items-center gap-2 bg-pagebg px-3 py-1.5 rounded border border-borderline">
                        <span className="w-2 h-2 rounded-full bg-green-500"></span>
                        <span className="text-[10px] font-bold text-sub uppercase tracking-widest">Engine Ready</span>
                    </div>
                    <button onClick={() => setShowSettings(!showSettings)} className="text-sub hover:text-primary transition-colors active:scale-90"><SettingsIcon /></button>
                </div>
            </header>

            {showSettings && (
                <div className="fixed inset-0 z-[60] bg-primary/40 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in" onClick={() => setShowSettings(false)}>
                    <div className="bg-white rounded-lg shadow-lg p-8 max-w-sm w-full animate-in zoom-in-95 border border-borderline" onClick={e => e.stopPropagation()}>
                        <h2 className="text-xl font-bold text-primary mb-2">에이전트 환경 설정</h2>
                        <p className="text-xs text-sub mb-6 leading-relaxed font-bold tracking-tight uppercase">원본 ID 무결성이 보장된 척추 모델</p>
                        <div className="mb-6">
                            <label className="block text-[11px] font-bold text-sub mb-1 uppercase tracking-widest">Model Selection</label>
                            <select value={apiModel} onChange={e => setApiModel(e.target.value)} className="w-full border border-borderline bg-white rounded p-3 text-sm focus:outline-none focus:ring-1 focus:ring-accent text-primary font-bold transition-all cursor-pointer">
                                <option value="gemini-2.5-flash">Gemini 2.5 Flash (다중호출 최적)</option>
                                <option value="gemini-2.5-pro">Gemini 2.5 Pro (고성능/속도주의)</option>
                            </select>
                        </div>
                        <button onClick={handleSaveSettings} className="w-full bg-accent text-white rounded p-3 text-xs font-bold uppercase tracking-widest hover:bg-opacity-90 transition-all btn-scale">설정 완료</button>
                    </div>
                </div>
            )}

            <main className="flex-1 max-w-[1600px] mx-auto w-full p-6 grid grid-cols-1 lg:grid-cols-12 gap-6 items-start overflow-y-auto min-h-0">
                <div className="lg:col-span-3 space-y-6 lg:sticky lg:top-0">
                    <div className="bg-white p-6 rounded-lg border border-borderline shadow-sm flex flex-col h-[700px]">
                        <button onClick={analyzeLogic} disabled={isAnalyzing} className="w-full py-4 mb-6 bg-accent hover:bg-opacity-90 disabled:bg-borderline disabled:text-sub text-white text-sm font-bold tracking-widest rounded transition-all active:scale-95 flex items-center justify-center gap-2">
                            {isAnalyzing ? <RefreshCw className="animate-spin text-white" size={14} /> : <Play size={14} />} 엔진 가동 (Execute)
                        </button>

                        <div className="mb-6 space-y-3">
                            <div onClick={() => fileInputRef.current.click()} className="border-2 border-dashed border-borderline rounded p-6 flex flex-col items-center justify-center gap-3 hover:bg-pagebg hover:border-accent cursor-pointer transition-all text-center group">
                                <Plus size={24} className="text-sub group-hover:text-accent transition-colors" />
                                <p className="text-[11px] font-bold text-sub uppercase tracking-widest group-hover:text-accent transition-colors">Excel/CSV/PDF 업로드</p>
                                <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" multiple accept=".csv,.xlsx,.xls,.txt,.pdf" />
                            </div>
                            <div className="max-h-[120px] overflow-y-auto space-y-2 pr-2">
                                {uploadedFiles.map(file => (
                                    <div key={file.id} className="flex items-center justify-between p-3 bg-pagebg rounded border border-borderline text-[11px] font-bold text-primary">
                                        <span className="flex items-center gap-2 truncate pr-4"><FileSpreadsheet size={14} className="shrink-0 text-sub"/>{file.name}</span>
                                        <button onClick={() => setUploadedFiles(prev => prev.filter(f => f.id !== file.id))} className="text-sub hover:text-red-500 transition-colors"><X size={14} /></button>
                                    </div>
                                ))}
                            </div>
                        </div>
                        <textarea
                            value={inputText}
                            onChange={(e) => setInputText(e.target.value)}
                            placeholder="분석할 요구사항을 입력하거나 엑셀을 업로드하세요."
                            className="w-full flex-1 p-4 text-sm bg-white border border-borderline rounded focus:ring-1 focus:ring-accent outline-none resize-none font-normal text-primary mb-4"
                        />
                    </div>
                </div>

                <div className="lg:col-span-9 h-full flex flex-col space-y-6 w-full min-w-0">
                    {truncationWarning && !isAnalyzing && (
                        <div className="bg-amber-50 border border-amber-300 text-amber-800 px-4 py-3 rounded-lg flex items-center gap-3 text-sm font-bold animate-in fade-in shrink-0">
                            <AlertTriangle size={16} className="shrink-0 text-amber-500" />
                            <span>파일 크기가 커서 일부만 분석합니다. ({truncationWarning})</span>
                            <button onClick={() => setTruncationWarning(null)} className="ml-auto text-amber-500 hover:text-amber-700"><X size={14}/></button>
                        </div>
                    )}
                    {metrics && !isAnalyzing && !errorMessage && (
                        <div className="bg-white p-6 rounded-lg border border-borderline shadow-sm animate-in fade-in shrink-0">
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-4">
                                <div className="bg-white rounded-lg p-5 border border-borderline shadow-sm flex flex-col justify-center">
                                    <p className="text-[11px] font-bold uppercase text-sub mb-1">UIUX 선별</p>
                                    <p className="text-3xl font-bold text-accent">{metrics.uiux_selected_count}<span className="text-sm ml-1 font-bold text-sub">/{metrics.total_input_count}</span></p>
                                </div>
                                <div className="bg-white rounded-lg p-5 border border-borderline shadow-sm flex flex-col justify-center">
                                    <p className="text-[11px] font-bold text-sub mb-1">모호성 보정</p>
                                    <p className="text-3xl font-bold text-primary">{metrics.ambiguity_resolved_count}<span className="text-sm ml-1 font-bold text-sub">건</span></p>
                                </div>
                                <div className="bg-white rounded-lg p-5 border border-borderline shadow-sm flex flex-col justify-center">
                                    <p className="text-[11px] font-bold text-sub mb-1">정책 충돌</p>
                                    <p className="text-3xl font-bold text-rose-500">{metrics.conflict_req_count}<span className="text-sm ml-1 font-bold text-sub">건</span></p>
                                </div>
                            </div>
                            <div className="bg-pagebg border border-borderline p-4 rounded text-sm text-primary leading-relaxed font-normal">"{metrics.selection_reason}"</div>
                        </div>
                    )}

                    <div className="flex flex-col flex-1 min-h-[400px] relative w-full min-w-0">
                        {!isAnalyzing && metrics && !errorMessage && (
                            <div className="flex gap-0 px-2 shrink-0 overflow-x-auto text-primary z-20 relative">
                                {['UIUX 선별', '제외 항목', '요구사항정의서', '충돌'].map(tab => (
                                    <button key={tab}
                                            onClick={() => setActiveTab(tab)}
                                            className={`tab-btn px-6 py-3 rounded-t text-xs flex items-center gap-2 ${activeTab === tab ? (tab === '충돌' ? 'tab-conflict-active' : 'tab-active') : 'tab-inactive'}`}>
                                        {tab}
                                        <span className="bg-pagebg px-2 py-0.5 rounded-full text-[10px] text-primary border border-borderline font-bold">{tab === 'UIUX 선별' ? rawFunc.length : tab === '제외 항목' ? rawNonFunc.length : tab === '요구사항정의서' ? optimizedReqs.length : conflicts.length}</span>
                                    </button>
                                ))}
                            </div>
                        )}

                        <div className={`bg-white border border-borderline shadow-sm flex flex-col flex-1 overflow-hidden w-full z-10 ${!metrics && !isAnalyzing && !errorMessage ? 'rounded-lg' : 'rounded-b-lg'}`}>
                            {(!isAnalyzing && metrics && !errorMessage) && (
                                <div className="p-4 border-b border-borderline bg-white flex items-center justify-between shrink-0">
                                    <h2 className="text-sm font-bold text-primary flex items-center gap-2 border-l-4 border-accent pl-2 uppercase tracking-wide">결과 테이블</h2>
                                    <div className="flex items-center gap-2">
                                        <button onClick={exportToExcel} className="px-3 py-1.5 bg-accent hover:bg-opacity-90 border border-accent rounded text-[11px] font-bold transition-all text-white flex items-center gap-2 uppercase tracking-wide btn-scale"><Download size={14}/> Excel Export</button>
                                        <button onClick={() => setIsFullScreen(true)} className="px-3 py-1.5 bg-white hover:bg-pagebg border border-borderline rounded text-[11px] font-bold transition-all text-primary flex items-center gap-2 uppercase tracking-wide btn-scale"><Maximize2 size={14}/> Full Screen</button>
                                    </div>
                                </div>
                            )}

                            <div className="flex-1 relative w-full min-h-0 bg-white">
                                <div className="absolute inset-0 overflow-auto">

                                    {isAnalyzing && (
                                        <div className="h-full flex flex-col items-center justify-center py-20 bg-pagebg animate-in fade-in">
                                            <div className="flex items-center gap-3 mb-8"><Loader2 size={32} className="text-accent animate-spin" /><h3 className="text-xl font-bold text-primary tracking-tight uppercase">AI 에이전트 다중 파이프라인 가동 중</h3></div>
                                            <div className="w-full max-w-md bg-white border border-borderline rounded-lg shadow-sm p-6 space-y-3 text-primary">
                                                {processingSteps.map((step, idx) => (
                                                    <div key={idx} className={`flex items-center gap-4 p-3 rounded transition-all duration-300 ${idx === progressStep ? 'bg-pagebg border border-accent/50 scale-105 transform origin-left' : idx < progressStep ? 'bg-white border border-borderline text-primary opacity-60' : 'text-sub opacity-40'}`}>
                                                        <div className={`shrink-0 w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold ${idx === progressStep ? 'bg-accent text-white animate-pulse' : 'bg-pagebg border border-borderline'}`}>{idx < progressStep ? <CheckCircle size={14} className="text-accent"/> : idx + 1}</div>
                                                        <div className="flex-1">
    <div className="flex items-center justify-between gap-2">
        <p className="text-sm font-bold tracking-tight text-left">{step.title}</p>
        {idx === progressStep && (
            <span className={`text-[11px] font-mono shrink-0 ${stepElapsed > 90 ? 'text-rose-500 font-bold' : 'text-accent'}`}>
                {stepElapsed}s{stepElapsed > 90 ? ' ⚠ 응답 지연' : ''}
            </span>
        )}
    </div>
    <p className="text-[11px] opacity-80 text-left mt-0.5">{step.desc}</p>
</div>
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    )}

                                    {errorMessage && (
                                        <div className="h-full flex flex-col items-center justify-center bg-white py-20 text-center px-8 animate-in zoom-in-95">
                                            <AlertOctagon size={64} className="mb-6 text-rose-500" />
                                            <h3 className="text-2xl font-bold text-primary mb-4 tracking-tight uppercase">파이프라인 실행 중단</h3>
                                            <div className="bg-pagebg p-6 rounded border border-borderline max-w-2xl w-full text-left overflow-auto max-h-[30vh] font-mono text-sm"><p className="font-normal text-rose-600 leading-relaxed whitespace-pre-wrap">{errorMessage}</p></div>
                                            <button onClick={() => setErrorMessage(null)} className="mt-8 px-8 py-3 bg-primary text-white rounded font-bold text-sm uppercase tracking-wide btn-scale">닫기 및 재가동</button>
                                        </div>
                                    )}

                                    {!isAnalyzing && !metrics && !errorMessage && (
                                        <div className="h-full flex flex-col items-center justify-center text-sub py-40">
                                            <Layers size={48} className="mb-4 text-borderline" />
                                            <p className="text-lg font-bold tracking-tight text-primary mb-2 uppercase">"요구사항 전체에서 UIUX 담당자의 검토 항목을 추려냅니다."</p>
                                            <p className="text-sm font-normal text-sub max-w-sm text-center">화면 설계에 영향을 주는 모든 요구사항을 담당 범위별로 분류하여 추출합니다.</p>
                                        </div>
                                    )}

                                    {!isAnalyzing && !errorMessage && (activeTab === 'UIUX 선별' || activeTab === '제외 항목') && (activeTab === 'UIUX 선별' ? rawFunc.length > 0 : rawNonFunc.length > 0) && renderRawTable(activeTab === 'UIUX 선별' ? rawFunc : rawNonFunc)}

                                    {!isAnalyzing && !errorMessage && activeTab === '요구사항정의서' && optimizedReqs.length > 0 && (
                                        <table className="w-full text-left text-xs border-collapse min-w-[1200px] bg-white text-primary">
                                            <thead className="sticky top-0 bg-pagebg text-[11px] uppercase tracking-widest text-sub border-b border-borderline z-10 font-bold">
                                                <tr>
                                                    <th className="p-4 w-12 text-center border-r border-borderline">NO</th>
                                                    <th className="p-4 w-36 border-r border-borderline">요구사항ID</th>
                                                    <th className="p-4 border-r border-borderline">FO/BO</th>
                                                    <th className="p-4 border-r border-borderline">화면본수</th>
                                                    <th className="p-4 border-r border-borderline">판단단계</th>
                                                    <th className="p-4 w-20 border-r border-borderline">업무분류</th>
                                                    <th className="p-4 w-24 border-r border-borderline">업무_대</th>
                                                    <th className="p-4 w-24 border-r border-borderline">기능_중</th>
                                                    <th className="p-4 border-r border-borderline w-48 text-primary">요구정의명</th>
                                                    <th className="p-4 border-r border-borderline">상세내용 (구체화)</th>
                                                    <th className="p-4 w-40 text-center">연관 요구사항 / HITL</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-borderline">
                                                {optimizedReqs.map((r, i) => (
                                                    <tr key={i} className="hover:bg-pagebg cursor-pointer group transition-colors" onClick={() => setSelectedItem({type: 'opt', data: r})}>
                                                        <td className="p-4 text-center text-sub border-r border-borderline">{r.NO}</td>
                                                        <td className="p-4 font-bold text-primary tracking-wider border-r border-borderline">{r.요구사항ID}</td>
                                                        <td className="p-4 text-sub border-r border-borderline">{r.fo_bo || 'TBD'}</td>
                                                        <td className="p-4 text-sub border-r border-borderline">{r.화면본수 || 'TBD'}</td>
                                                        <td className="p-4 text-sub border-r border-borderline text-[10px]">{r.uiux_relevance_step || '-'}</td>
                                                        <td className="p-4 text-sub border-r border-borderline">{r.업무분류}</td>
                                                        <td className="p-4 text-sub border-r border-borderline">{r.업무_대}</td>
                                                        <td className="p-4 text-sub border-r border-borderline">{r.기능_중}</td>
                                                        <td className="p-4 font-bold text-primary group-hover:text-accent transition-colors border-r border-borderline">{r.요구정의명}</td>
                                                        <td className="p-4 text-primary leading-relaxed border-r border-borderline">{r.고객_요구사항_상세_내용}</td>
                                                        <td className="p-4 text-center">
                                                            {renderBadges(r.related_reqs)}
                                                            {r.ambiguity_hitl?.is_ambiguous && <div className="mt-2 inline-flex items-center gap-1 bg-white text-accent px-2 py-1 rounded text-[10px] font-bold border border-accent"><AlertTriangle size={10}/> 모호성 HITL</div>}
                                                        </td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}

                                    {!isAnalyzing && !errorMessage && activeTab === '충돌' && conflicts.length > 0 && renderConflictsTable()}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </main>

            {/* --- Detail Modal --- */}
            {selectedItem && (
                <div className="fixed inset-0 z-[90] flex items-center justify-center bg-primary/40 backdrop-blur-sm p-6 animate-in fade-in duration-200" onClick={() => setSelectedItem(null)}>
                    <div className="bg-white rounded-lg shadow-lg w-full max-w-3xl flex flex-col max-h-[90vh] animate-in zoom-in-95 border border-borderline" onClick={e => e.stopPropagation()}>
                        <div className="flex justify-between items-center p-6 border-b border-borderline bg-white rounded-t-lg shrink-0 text-primary">
                            <div className="flex items-center gap-3">
                                {selectedItem.type === 'conflict' || selectedItem.data?.ambiguity_hitl?.is_ambiguous ? <div className="bg-white p-2 rounded text-rose-500 border border-borderline"><AlertTriangle size={20}/></div> : <div className="bg-white p-2 rounded text-primary border border-borderline"><ShieldCheck size={20}/></div>}
                                <div>
                                    <h3 className="text-lg font-bold text-primary uppercase leading-tight">모호성/충돌 결정(HITL)</h3>
                                    <p className="text-[11px] text-sub font-bold uppercase mt-1 tracking-widest">{selectedItem.type === 'conflict' ? selectedItem.data.conflict_id : selectedItem.data.요구사항ID}</p>
                                </div>
                            </div>
                            <button onClick={() => setSelectedItem(null)} className="p-2 bg-white hover:bg-pagebg rounded border border-borderline transition-all text-primary active:scale-90"><X size={16}/></button>
                        </div>
                        <div className="p-8 space-y-6 overflow-y-auto text-primary">
                            {selectedItem.type === 'conflict' ? (
                                <>
                                    <div className="bg-pagebg border border-borderline p-6 rounded text-primary">
                                        <h4 className="font-bold text-primary mb-2 uppercase tracking-wide text-xs">충돌 리스크 분석</h4>
                                        <p className="text-sm text-primary leading-relaxed">{selectedItem.data.conflict_reason}</p>
                                        <div className="mt-4 pt-4 border-t border-borderline flex gap-2 items-center"><span className="text-[11px] font-bold uppercase text-sub">관련 ID:</span><span className="font-mono text-xs font-bold text-primary bg-white px-2 py-1 rounded border border-borderline">{selectedItem.data.involved_req_ids?.join(', ')}</span></div>
                                    </div>
                                    <div className="grid grid-cols-1 gap-3">
                                        {selectedItem.data.suggested_options.map((opt, idx) => (
                                            <button key={idx} onClick={() => handleHitlDecision(selectedItem.data, opt, 'conflict')} className="w-full text-left px-5 py-4 border border-borderline rounded hover:border-accent hover:bg-pagebg text-sm text-primary transition-all group flex justify-between items-center bg-white">
                                                <span className="flex-1 leading-relaxed">{opt}</span>
                                                <div className="bg-accent text-white px-4 py-1.5 rounded text-xs font-bold opacity-0 group-hover:opacity-100 transition-opacity uppercase flex items-center gap-1"><CheckCircle size={12}/> 채택하기</div>
                                            </button>
                                        ))}

                                        {!isCustomEditing ? (
                                            <button onClick={() => setIsCustomEditing(true)} className="w-full text-center px-5 py-4 border border-dashed border-sub/50 rounded hover:border-accent hover:text-accent font-bold text-sm text-sub transition-all flex justify-center items-center gap-2">
                                                <Edit2 size={16}/> 직접 수정하여 확정하기
                                            </button>
                                        ) : (
                                            <div className="flex flex-col gap-3 p-5 border border-accent rounded bg-pagebg">
                                                <textarea
                                                    value={customOptionText}
                                                    onChange={(e) => setCustomOptionText(e.target.value)}
                                                    placeholder="PM 의사결정 내용을 직접 입력하세요..."
                                                    className="w-full p-4 text-sm bg-white border border-borderline rounded focus:ring-1 focus:ring-accent outline-none resize-none text-primary min-h-[80px]"
                                                />
                                                <div className="flex justify-end gap-2">
                                                    <button onClick={() => setIsCustomEditing(false)} className="px-5 py-2 text-xs font-bold text-sub hover:text-primary transition-colors">취소</button>
                                                    <button onClick={() => {
                                                        if(customOptionText.trim()) handleHitlDecision(selectedItem.data, customOptionText.trim(), 'conflict');
                                                    }} className="px-5 py-2 bg-accent text-white rounded text-xs font-bold hover:bg-opacity-90 transition-colors flex items-center gap-1">
                                                        <CheckCircle size={14}/> 확정 완료
                                                    </button>
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                </>
                            ) : (
                                <>
                                    <div className="grid grid-cols-3 gap-4">
                                        <div className="bg-pagebg p-4 rounded border border-borderline"><span className="text-[11px] font-bold text-sub block mb-1 uppercase tracking-wide">Hierarchy</span><span className="font-bold text-primary text-sm">{selectedItem.data.업무분류} &gt; {selectedItem.data.업무_대}</span></div>
                                        <div className="bg-pagebg p-4 rounded border border-borderline"><span className="text-[11px] font-bold text-sub block mb-1 uppercase tracking-wide">Priority</span><span className="font-bold text-accent text-sm">{selectedItem.data.우선순위}</span></div>
                                        <div className="bg-pagebg p-4 rounded border border-borderline"><span className="text-[11px] font-bold text-sub block mb-1 uppercase tracking-wide">Constraints</span><span className="font-bold text-primary text-sm">{selectedItem.data.제약사항 || 'N/A'}</span></div>
                                    </div>
                                    <div className="text-primary">
                                        <h4 className="font-bold text-primary text-lg mb-2">{selectedItem.data.요구정의명}</h4>
                                        <div className="p-6 bg-pagebg border border-borderline rounded text-sm text-primary leading-relaxed">"{selectedItem.data.고객_요구사항_상세_내용}"</div>
                                    </div>

                                    {selectedItem.data.related_reqs && selectedItem.data.related_reqs.length > 0 && (
                                        <div className="bg-white border border-borderline p-6 rounded mt-6 text-primary">
                                            <h4 className="font-bold text-primary tracking-wide mb-4 text-xs uppercase flex items-center gap-2"><Layers size={14} className="text-sub" /> 연관 요구사항 분석</h4>
                                            <div className="flex flex-wrap gap-2">
                                                {selectedItem.data.related_reqs.map((rel, idx) => (
                                                    <span key={idx} className="bg-pagebg text-primary px-3 py-1.5 rounded text-[11px] font-bold border border-borderline"><span className="text-sub mr-1">[{rel.relation_type}]</span> {rel.target_id}</span>
                                                ))}
                                            </div>
                                        </div>
                                    )}

                                    {selectedItem.data.ambiguity_hitl?.is_ambiguous && (
                                        <div className="bg-white border border-accent p-6 rounded mt-6 text-primary shadow-sm">
                                            <h4 className="font-bold text-primary text-base tracking-wide mb-2 text-accent uppercase flex items-center gap-2"><AlertTriangle size={16}/> Ambiguity Resolution</h4>
                                            <p className="text-sm text-primary leading-relaxed mb-6">"{selectedItem.data.ambiguity_hitl.ambiguity_reason}"</p>
                                            <div className="grid grid-cols-1 gap-3">
                                                {selectedItem.data.ambiguity_hitl.suggested_options.map((opt, idx) => (
                                                    <button key={idx} onClick={() => handleHitlDecision(selectedItem.data, opt, 'ambiguity')} className="w-full text-left px-5 py-4 border border-borderline rounded hover:border-accent hover:bg-pagebg text-sm text-primary transition-all group flex justify-between items-center bg-white">
                                                        <span className="flex-1 leading-relaxed">{opt}</span>
                                                        <div className="bg-accent text-white px-4 py-1.5 rounded text-xs font-bold opacity-0 group-hover:opacity-100 transition-opacity uppercase flex items-center gap-1"><CheckCircle size={12}/> 확정하기</div>
                                                    </button>
                                                ))}

                                                {!isCustomEditing ? (
                                                    <button onClick={() => setIsCustomEditing(true)} className="w-full text-center px-5 py-4 border border-dashed border-sub/50 rounded hover:border-accent hover:text-accent font-bold text-sm text-sub transition-all flex justify-center items-center gap-2">
                                                        <Edit2 size={16}/> 직접 수정하여 확정하기
                                                    </button>
                                                ) : (
                                                    <div className="flex flex-col gap-3 p-5 border border-accent rounded bg-pagebg">
                                                        <textarea
                                                            value={customOptionText}
                                                            onChange={(e) => setCustomOptionText(e.target.value)}
                                                            placeholder="PM 의사결정 내용을 직접 입력하세요..."
                                                            className="w-full p-4 text-sm bg-white border border-borderline rounded focus:ring-1 focus:ring-accent outline-none resize-none text-primary min-h-[80px]"
                                                        />
                                                        <div className="flex justify-end gap-2">
                                                            <button onClick={() => setIsCustomEditing(false)} className="px-5 py-2 text-xs font-bold text-sub hover:text-primary transition-colors">취소</button>
                                                            <button onClick={() => {
                                                                if(customOptionText.trim()) handleHitlDecision(selectedItem.data, customOptionText.trim(), 'ambiguity');
                                                            }} className="px-5 py-2 bg-accent text-white rounded text-xs font-bold hover:bg-opacity-90 transition-colors flex items-center gap-1">
                                                                <CheckCircle size={14}/> 확정 완료
                                                            </button>
                                                        </div>
                                                    </div>
                                                )}
                                            </div>
                                        </div>
                                    )}
                                </>
                            )}
                        </div>
                        <div className="p-4 bg-pagebg border-t border-borderline text-center rounded-b-lg"><p className="text-[10px] font-bold text-sub tracking-widest uppercase">UIUX 요구사항 선별 에이전트 Pipeline Framework v1.0</p></div>
                    </div>
                </div>
            )}

            {/* --- Full Screen Modal --- */}
            {isFullScreen && (
                <div className="fixed inset-0 z-[80] bg-primary/40 backdrop-blur-sm p-6 flex flex-col animate-in fade-in duration-200 overflow-hidden">
                    <div className="bg-white rounded-lg shadow-lg flex-1 flex flex-col overflow-hidden w-full max-w-full border border-borderline text-primary">
                        <div className="px-8 pt-6 border-b border-borderline bg-white flex flex-col shrink-0 gap-6">
                            <div className="flex justify-between items-center">
                                <div className="flex items-center gap-4">
                                    <div className="p-3 bg-primary rounded text-white shadow-sm border border-borderline"><Monitor size={24}/></div>
                                    <h2 className="text-xl text-primary uppercase tracking-wide font-bold">결과 테이블 전체 보기</h2>
                                </div>
                                <div className="flex items-center gap-3">
                                    <button onClick={exportToExcel} className="px-4 py-2 bg-accent hover:bg-opacity-90 rounded text-xs font-bold transition-all text-white flex items-center gap-2 uppercase tracking-wide btn-scale"><Download size={14}/> Excel Export</button>
                                    <button onClick={() => setIsFullScreen(false)} className="p-3 bg-white hover:bg-pagebg rounded border border-borderline transition-all text-primary active:scale-90"><X size={20}/></button>
                                </div>
                            </div>
                            <div className="flex gap-0 relative z-20">
                                {['UIUX 선별', '제외 항목', '요구사항정의서', '충돌'].map(tab => (
                                    <button key={tab} onClick={() => setActiveTab(tab)} className={`tab-btn px-8 py-3 rounded-t text-sm font-bold transition-all ${activeTab === tab ? (tab === '충돌' ? 'tab-conflict-active' : 'tab-active') : 'tab-inactive'}`}>{tab}</button>
                                ))}
                            </div>
                        </div>
                        <div className="flex-1 overflow-auto p-6 bg-pagebg z-10">
                            {activeTab === '요구사항정의서' ? (
                                <table className="w-full text-left text-[11px] border-collapse min-w-[2400px] shadow-sm rounded border border-borderline bg-white">
                                    <thead className="bg-pagebg text-sub sticky top-0 z-10 text-[11px] uppercase tracking-widest font-bold">
                                        <tr>
                                            <th className="p-4 w-12 text-center border-r border-borderline">NO</th>
                                            <th className="p-4 w-40 border-r border-borderline">요구사항ID</th>
                                            <th className="p-4 border-r border-borderline">FO/BO</th>
                                            <th className="p-4 border-r border-borderline">화면본수</th>
                                            <th className="p-4 border-r border-borderline">판단단계</th>
                                            <th className="p-4 border-r border-borderline">업무분류</th>
                                            <th className="p-4 border-r border-borderline">업무_대</th>
                                            <th className="p-4 border-r border-borderline">기능_중</th>
                                            <th className="p-4 border-r border-borderline">구성_소</th>
                                            <th className="p-4 min-w-[200px] border-r border-borderline">요구정의명</th>
                                            <th className="p-4 border-r border-borderline min-w-[400px]">고객_요구사항_상세_내용</th>
                                            <th className="p-4 border-r border-borderline min-w-[200px]">연관 요구사항 및 HITL 근거</th>
                                            <th className="p-4 border-r border-borderline min-w-[150px]">제약사항</th>
                                            <th className="p-4 border-r border-borderline">요건_발생일</th>
                                            <th className="p-4 border-r border-borderline text-center">분류/유형</th>
                                            <th className="p-4 text-center">우선순위</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-borderline font-normal text-primary">
                                        {optimizedReqs.map((r, i) => (
                                            <tr key={i} className="hover:bg-pagebg cursor-pointer transition-all duration-200" onClick={() => { setSelectedItem({type: 'opt', data: r}); }}>
                                                <td className="p-4 text-center text-sub border-r border-borderline">{r.NO}</td>
                                                <td className="p-4 font-bold text-primary border-r border-borderline tracking-wider">{r.요구사항ID}</td>
                                                <td className="p-4 border-r border-borderline">{r.fo_bo || 'TBD'}</td>
                                                <td className="p-4 border-r border-borderline">{r.화면본수 || 'TBD'}</td>
                                                <td className="p-4 border-r border-borderline text-[10px] text-sub">{r.uiux_relevance_step || '-'}</td>
                                                <td className="p-4 border-r border-borderline">{r.업무분류}</td>
                                                <td className="p-4 border-r border-borderline">{r.업무_대}</td>
                                                <td className="p-4 border-r border-borderline">{r.기능_중}</td>
                                                <td className="p-4 border-r border-borderline">{r.구성_소}</td>
                                                <td className="p-4 font-bold text-primary border-r border-borderline leading-tight">{r.요구정의명}</td>
                                                <td className="p-4 leading-relaxed border-r border-borderline">{r.고객_요구사항_상세_내용}</td>
                                                <td className="p-4 border-r border-borderline">
                                                    {r.related_reqs?.map((rel, idx) => (
                                                        <div key={idx} className="mb-1 text-[11px]"><span className="text-sub font-bold">[{rel.relation_type}]</span> {rel.target_id}</div>
                                                    ))}
                                                </td>
                                                <td className="p-4 text-primary border-r border-borderline">{r.제약사항 || '-'}</td>
                                                <td className="p-4 font-mono text-sub border-r border-borderline">{r.요건_발생일}</td>
                                                <td className="p-4 text-center border-r border-borderline">{r.요구사항유형분류?.분류}/{r.요구사항유형분류?.유형}</td>
                                                <td className="p-4 text-center font-bold text-accent">{r.우선순위}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            ) : (
                                <div className="rounded border border-borderline shadow-sm bg-white">
                                    {activeTab === '충돌' ? renderConflictsTable() : renderRawTable(activeTab === 'UIUX 선별' ? rawFunc : rawNonFunc)}
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

export default App;
