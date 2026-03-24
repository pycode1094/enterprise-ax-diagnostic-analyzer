/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  Radar, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Cell
} from 'recharts';
import { 
  Activity, 
  ChevronRight, 
  ClipboardCheck, 
  Building2, 
  Users, 
  Briefcase, 
  TrendingUp, 
  Award,
  Info,
  RefreshCcw,
  LayoutDashboard,
  Download,
  FileText,
  GraduationCap
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { saveAs } from 'file-saver';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel, 
  AlignmentType, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType,
  BorderStyle
} from 'docx';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- OpenRouter API (direct client call) ---

const OPENROUTER_MODEL = 'openrouter/free';

async function callOpenRouter(prompt: string, apiKey: string): Promise<string> {
  const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model: OPENROUTER_MODEL,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `API error: ${res.status}`);
  }

  const data = await res.json();
  return data.choices?.[0]?.message?.content || '';
}

// --- Types ---

type AXLevel = 'L1' | 'L2' | 'L3' | 'L4' | 'L5';

interface DeptScore {
  name: string;
  score: number;
}

interface DiagnosticInput {
  companyName: string;
  industry: string;
  employeeCount: string;
  execScore: number;
  deptScores: DeptScore[];
  individualScore: number;
}

interface DiagnosticResult {
  companyName: string;
  overallScore: number;
  overallLevel: AXLevel;
  overallLevelName: string;
  execLevel: AXLevel;
  deptLevels: Record<string, AXLevel>;
  individualLevel: AXLevel;
  chartData: { name: string; score: number; fullMark: number }[];
}

interface AIReport {
  overall: string;
  exec: string;
  dept: string;
  indiv: string;
  summary: string;
}

// --- Constants ---

const LEVEL_MAP: Record<AXLevel, { name: string; color: string; range: string }> = {
  L1: { name: '인식 단계 (Awareness)', color: 'bg-slate-400', range: '1.0~1.9' },
  L2: { name: '실험 단계 (Experimentation)', color: 'bg-blue-400', range: '2.0~2.9' },
  L3: { name: '적용 단계 (Application)', color: 'bg-indigo-400', range: '3.0~3.9' },
  L4: { name: '확산 단계 (Diffusion)', color: 'bg-violet-400', range: '4.0~4.4' },
  L5: { name: '최적화 단계 (Optimization)', color: 'bg-emerald-500', range: '4.5~5.0' },
};

const getLevel = (score: number): AXLevel => {
  if (score >= 4.5) return 'L5';
  if (score >= 4.0) return 'L4';
  if (score >= 3.0) return 'L3';
  if (score >= 2.0) return 'L2';
  return 'L1';
};

// --- Components ---

const LevelBadge = ({ level }: { level: AXLevel }) => {
  const info = LEVEL_MAP[level];
  return (
    <span className={cn(
      "px-2 py-1 rounded-md text-xs font-bold text-white shadow-sm",
      info.color
    )}>
      {level}
    </span>
  );
};

export default function App() {
  const [step, setStep] = useState<'input' | 'result'>('input');
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('openrouter_api_key') || '');
  const [aiReport, setAiReport] = useState<AIReport | null>(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);
  const [formData, setFormData] = useState<DiagnosticInput>({
    companyName: '',
    industry: '',
    employeeCount: '',
    execScore: 3,
    deptScores: [
      { name: '생산', score: 3 },
      { name: '품질', score: 3 },
      { name: '영업', score: 3 },
      { name: '관리', score: 3 },
    ],
    individualScore: 3,
  });

  const handleDeptScoreChange = (index: number, value: number) => {
    const newDeptScores = [...formData.deptScores];
    newDeptScores[index].score = value;
    setFormData({ ...formData, deptScores: newDeptScores });
  };

  const diagnosticResult = useMemo((): DiagnosticResult | null => {
    if (step !== 'result') return null;

    const deptAvg = formData.deptScores.reduce((acc, curr) => acc + curr.score, 0) / formData.deptScores.length;
    const overallScore = Number(((formData.execScore + deptAvg + formData.individualScore) / 3).toFixed(2));
    
    const deptLevels: Record<string, AXLevel> = {};
    formData.deptScores.forEach(dept => {
      deptLevels[dept.name] = getLevel(dept.score);
    });

    const overallLevel = getLevel(overallScore);

    const radarData = [
      { name: '경영진 의지', score: formData.execScore, fullMark: 5 },
      ...formData.deptScores.map(d => ({ name: d.name, score: d.score, fullMark: 5 })),
      { name: '개인 역량', score: formData.individualScore, fullMark: 5 },
    ];

    return {
      companyName: formData.companyName,
      overallScore,
      overallLevel,
      overallLevelName: LEVEL_MAP[overallLevel].name,
      execLevel: getLevel(formData.execScore),
      deptLevels,
      individualLevel: getLevel(formData.individualScore),
      chartData: radarData,
    };
  }, [formData, step]);

  const runDiagnostic = async () => {
    localStorage.setItem('openrouter_api_key', apiKey);
    setStep('result');
    setAiLoading(true);
    setAiError(null);
    setAiReport(null);

    const deptAvg = formData.deptScores.reduce((acc, curr) => acc + curr.score, 0) / formData.deptScores.length;
    const overallScore = ((formData.execScore + deptAvg + formData.individualScore) / 3).toFixed(2);
    const overallLevel = getLevel(Number(overallScore));

    const prompt = `You are a professional enterprise AI transformation (AX) consultant. Analyze the following company's AX diagnostic data and write a detailed consulting report. ALL output text MUST be written in Korean.

## Company Information
- Company Name: ${formData.companyName}
- Industry: ${formData.industry}
- Number of Employees: ${formData.employeeCount}

## Diagnostic Scores (out of 5.0)
- Executive Commitment: ${formData.execScore} (${getLevel(formData.execScore)})
- Department Readiness: ${formData.deptScores.map(d => `${d.name} ${d.score}`).join(', ')} (Average ${deptAvg.toFixed(1)})
- Individual AX Competency: ${formData.individualScore} (${getLevel(formData.individualScore)})
- Overall Score: ${overallScore} (${overallLevel} - ${LEVEL_MAP[overallLevel].name})

## AX Level Criteria
- L1 Awareness (1.0-1.9): Beginning to recognize AI adoption needs
- L2 Experimentation (2.0-2.9): Some pilot attempts underway
- L3 Application (3.0-3.9): AI applied to specific tasks
- L4 Expansion (4.0-4.4): Organization-wide AI adoption
- L5 Optimization (4.5-5.0): AI as core competitive advantage

## Instructions
Respond ONLY with a JSON object. No other text before or after the JSON.
Each field must contain specific, actionable consulting advice tailored to the company's industry and size. Write at least 3 sentences per field. ALL content must be in Korean.

{
  "overall": "Comprehensive diagnosis analysis (evaluate the company's current AX level and present key challenges) - write in Korean",
  "exec": "Executive commitment analysis (assess leadership's AX drive and suggest improvements) - write in Korean",
  "dept": "Department readiness analysis (current AI utilization per department, gaps between departments, improvement plans) - write in Korean",
  "indiv": "Individual competency analysis (employee AI skill levels, training and growth strategies) - write in Korean",
  "summary": "Key summary (core message in 3 lines or less) - write in Korean"
}`;

    try {
      const response = await callOpenRouter(prompt, apiKey);
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error('AI 응답에서 JSON을 찾을 수 없습니다.');
      const parsed: AIReport = JSON.parse(jsonMatch[0]);
      setAiReport(parsed);
    } catch (err) {
      console.error('AI Report generation failed:', err);
      setAiError(err instanceof Error ? err.message : 'AI 보고서 생성에 실패했습니다.');
    } finally {
      setAiLoading(false);
    }
  };

  const getConsultingText = (level: AXLevel, area: 'exec' | 'dept' | 'indiv' | 'overall') => {
    const texts: Record<AXLevel, Record<typeof area, string>> = {
      L1: {
        overall: "AI 도입 필요성에 대한 인식은 형성되었으나, 구체적인 전략이나 실행 계획이 부재한 초기 단계입니다. 조직 내 AI 리터러시 제고를 위한 기초 교육과 비전 수립이 시급합니다.",
        exec: "경영진의 AX 추진 의지가 초기 단계에 머물러 있어 전사적 동력 확보가 어려운 상태입니다. AI를 단순 기술이 아닌 경영 전략의 핵심으로 인식하고 명확한 로드맵을 제시해야 합니다.",
        dept: "부서별 AI 활용이 거의 전무하거나 산발적인 관심 수준에 그치고 있습니다. 부서별 특화된 AI 유즈케이스를 발굴하고 파일럿 프로젝트를 통한 체감 효과 창출이 필요합니다.",
        indiv: "구성원들의 AI 도구 활용 능력 및 데이터 이해도가 낮아 디지털 전환에 대한 막연한 불안감이 존재합니다. 전사적 AI 기초 교육 프로그램 도입을 통해 심리적 장벽을 해소해야 합니다."
      },
      L2: {
        overall: "AI 활용이 일부 시도되고 있으나 조직 차원의 실행 구조가 부족한 상태입니다. 성공 사례(Quick-Win) 발굴을 통해 전사적 공감대를 형성하고 표준 가이드라인을 수립해야 합니다.",
        exec: "AX에 대한 관심은 높으나 구체적인 자원 배분과 조직 개편이 미흡하여 실행력이 떨어지는 상태입니다. 전담 조직(COE) 구성과 예산 확보를 통해 추진력을 강화해야 합니다.",
        dept: "일부 부서에서 실험적인 AI 도입이 이루어지고 있으나 타 부서와의 협업이나 데이터 공유가 단절되어 있습니다. 부서 간 데이터 연계 체계를 마련하고 성공 사례를 공유하는 워크숍이 필요합니다.",
        indiv: "일부 얼리어답터를 중심으로 AI 활용이 시도되나 대다수 인원은 기존 업무 방식을 고수하고 있습니다. 실무 중심의 AI 활용 가이드를 배포하고 우수 활용 사례에 대한 보상 체계를 마련해야 합니다."
      },
      L3: {
        overall: "특정 업무에서 AI 활용이 시작되었으나 조직 확산을 위한 체계가 필요합니다. 부서 간 데이터 사일로를 해소하고 전사적 AX 거버넌스를 강화하여 통합된 시너지를 창출해야 합니다.",
        exec: "경영진이 AX의 중요성을 인지하고 적극 지원하고 있으나, 성과 측정 지표(KPI)와의 연계가 부족합니다. AI 도입 성과를 정량화하여 비즈니스 가치와 연결하는 성과 관리 체계 도입이 필요합니다.",
        dept: "핵심 부서를 중심으로 AI가 업무 프로세스에 적용되었으나 데이터 품질 관리 및 표준화가 미흡합니다. 전사적 데이터 거버넌스 수립을 통해 데이터 신뢰성을 확보하고 활용 범위를 넓혀야 합니다.",
        indiv: "실무에서 AI 도구를 활용하기 시작했으나 고도화된 분석이나 문제 해결 능력은 부족한 상태입니다. 직무별 심화 AI 역량 강화 교육을 통해 AI를 도구가 아닌 파트너로 활용하도록 유도해야 합니다."
      },
      L4: {
        overall: "조직 전반에 걸쳐 AI 활용이 확산되어 업무 효율성이 증대되고 있으나, 데이터 기반 의사결정 문화의 완전한 정착이 과제입니다. AI 모델 고도화와 함께 운영 안정성을 확보해야 합니다.",
        exec: "AX가 경영 전략의 핵심으로 자리 잡았으나 급변하는 AI 기술 트렌드에 대한 유연한 대응 체계가 보완되어야 합니다. 외부 파트너십 강화 및 오픈 이노베이션을 통해 기술 우위를 지속적으로 확보해야 합니다.",
        dept: "대부분의 부서에서 AI를 활용하여 성과를 창출하고 있으나 부서 간 최적화 수준의 편차가 존재합니다. 부서 간 베스트 프랙티스를 상향 평준화하고 전사적 최적화 관점의 AI 통합 운영이 필요합니다.",
        indiv: "구성원 대다수가 AI를 업무에 능숙하게 활용하고 있으나 AI 윤리 및 보안 의식은 추가 보완이 필요합니다. AI 거버넌스 내 보안 및 윤리 가이드라인을 강화하여 안정적인 AI 활용 문화를 정착시켜야 합니다."
      },
      L5: {
        overall: "AI가 기업의 핵심 경쟁력으로 내재화되어 실시간 최적화가 이루어지고 있는 선도적 단계입니다. 지속적인 기술 혁신과 함께 AI 생태계 리딩을 위한 전략적 투자가 지속되어야 합니다.",
        exec: "최고 수준의 AX 리더십을 보유하고 있으나 미래 불확실성에 대비한 AI 기반 시나리오 경영을 더욱 강화해야 합니다. 글로벌 수준의 AI 인재 확보 및 유지를 위한 차별화된 인재 경영 전략이 필요합니다.",
        dept: "전 부서가 AI 기반으로 유기적으로 연결되어 데이터 기반의 실시간 의사결정이 이루어지고 있습니다. AI 모델의 자가 학습 및 고도화 체계를 구축하여 기술적 초격차를 유지해야 합니다.",
        indiv: "모든 구성원이 AI 전문가 수준의 역량을 보유하고 창의적인 업무에 집중하고 있습니다. 구성원들이 AI와 함께 새로운 비즈니스 모델을 창출할 수 있도록 사내 벤처 및 혁신 문화를 장려해야 합니다."
      }
    };
    return texts[level][area];
  };

  const [viewMode, setViewMode] = useState<'dashboard' | 'report' | 'strategy'>('dashboard');

  const getExecutionStrategy = (level: AXLevel) => {
    const strategies: Record<AXLevel, {
      directions: string[];
      steps: { title: string; period: string; content: string }[];
      courses: { category: string; content: string }[];
      pocs: { area: string; task: string }[];
      jobTraining: {
        category: string;
        annualPlan: { quarter: string; title: string }[];
      }[];
    }> = {
      L1: {
        directions: [
          "전사적 AI 리터러시 제고 및 심리적 장벽 해소",
          "경영진 중심의 AX 비전 및 추진 로드맵 수립",
          "현업 중심의 작고 확실한 성공 사례(Quick-Win) 발굴"
        ],
        steps: [
          { title: "기반 조성 및 교육", period: "1~3개월", content: "전 임직원 대상 AI 기초 교육 및 AX 비전 선포식 진행. 부서별 AX 챔피언 선발." },
          { title: "아이디어 발굴 및 PoC", period: "3~6개월", content: "부서별 업무 효율화 아이디어 공모전 실시. 1~2개 핵심 과제 대상 PoC(개념 증명) 수행." },
          { title: "초기 확산 및 평가", period: "6~12개월", content: "PoC 결과 공유회 및 성과 분석. 우수 부서 포상 및 차년도 확산 계획 수립." }
        ],
        courses: [
          { category: "기초", content: "생성형 AI의 이해와 업무 활용 기초 (ChatGPT, Claude 등)" },
          { category: "실무", content: "AI 프롬프트 엔지니어링 및 업무 자동화 입문" },
          { category: "고급", content: "AX 리더십 및 디지털 전환 전략 수립 과정" }
        ],
        pocs: [
          { area: "사무/관리", task: "회의록 자동 요약 및 이메일 초안 작성 자동화" },
          { area: "영업/마케팅", task: "시장 트렌드 분석 및 경쟁사 모니터링 자동화" }
        ],
        jobTraining: [
          {
            category: "경영진",
            annualPlan: [
              { quarter: "1Q", title: "AX 비전 수립 및 리더십 전략" },
              { quarter: "2Q", title: "글로벌 AI 트렌드 및 산업 변화" },
              { quarter: "3Q", title: "AI 거버넌스 및 리스크 관리" },
              { quarter: "4Q", title: "데이터 기반 의사결정 체계" }
            ]
          },
          {
            category: "생산/공정",
            annualPlan: [
              { quarter: "1Q", title: "스마트 팩토리 및 AI 기초" },
              { quarter: "2Q", title: "AI 공정 자동화 우수 사례" },
              { quarter: "3Q", title: "현장 데이터 수집 및 관리" },
              { quarter: "4Q", title: "현장 AI 도구 활용 실무" }
            ]
          },
          {
            category: "품질/R&D",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 품질 관리 입문" },
              { quarter: "2Q", title: "R&D 효율화를 위한 AI 도구" },
              { quarter: "3Q", title: "데이터 분석 및 시각화 기초" },
              { quarter: "4Q", title: "AI 협업 플랫폼 활용" }
            ]
          },
          {
            category: "영업/마케팅",
            annualPlan: [
              { quarter: "1Q", title: "생성형 AI 마케팅 기초" },
              { quarter: "2Q", title: "AI 고객 분석 및 타겟팅" },
              { quarter: "3Q", title: "프롬프트 엔지니어링 실무" },
              { quarter: "4Q", title: "AI 기반 콘텐츠 제작" }
            ]
          },
          {
            category: "경영지원/인사",
            annualPlan: [
              { quarter: "1Q", title: "AI 업무 자동화 및 리터러시" },
              { quarter: "2Q", title: "디지털 워크플레이스 구축" },
              { quarter: "3Q", title: "AI 보안 및 정보 보호" },
              { quarter: "4Q", title: "AI 기반 인재 관리 기초" }
            ]
          }
        ]
      },
      L2: {
        directions: [
          "부서별 산발적 실험을 전사적 표준 체계로 통합",
          "AX 전담 조직(Task Force) 구성 및 자원 집중",
          "데이터 수집 및 관리 표준 가이드라인 수립"
        ],
        steps: [
          { title: "역량 강화 및 표준화", period: "1~3개월", content: "부서별 AX 챔피언 심화 교육. 전사 공통 AI 활용 가이드라인 및 보안 수칙 수립." },
          { title: "다각적 PoC 수행", period: "3~6개월", content: "생산, 품질, 영업 등 주요 밸류체인별 3~4개 PoC 동시 추진. 데이터 정제 작업 병행." },
          { title: "성과 검증 및 인프라", period: "6~12개월", content: "PoC 성과 정량화 및 비즈니스 가치 검증. 전사 AI 플랫폼(SaaS 등) 도입 검토." }
        ],
        courses: [
          { category: "기초", content: "데이터 기반 의사결정의 기초와 AI 활용" },
          { category: "실무", content: "노코드(No-code) 도구를 활용한 업무 앱 구축" },
          { category: "고급", content: "데이터 거버넌스 및 AI 보안 관리 실무" }
        ],
        pocs: [
          { area: "생산/공정", task: "생산일보 자동 기록 및 단순 공정 모니터링" },
          { area: "품질/검사", task: "과거 불량 사례 검색 및 유사 사례 추천 시스템" }
        ],
        jobTraining: [
          {
            category: "경영진",
            annualPlan: [
              { quarter: "1Q", title: "AX 전담 조직 및 거버넌스 구축" },
              { quarter: "2Q", title: "AI 투자 수익률(ROI) 분석" },
              { quarter: "3Q", title: "디지털 전환 성과 관리" },
              { quarter: "4Q", title: "AI 기반 미래 전략 기획" }
            ]
          },
          {
            category: "생산/공정",
            annualPlan: [
              { quarter: "1Q", title: "공정 데이터 표준화 실무" },
              { quarter: "2Q", title: "AI 기반 설비 예지 보전" },
              { quarter: "3Q", title: "스마트 팩토리 고도화 전략" },
              { quarter: "4Q", title: "현장 AI 챔피언 양성 과정" }
            ]
          },
          {
            category: "품질/R&D",
            annualPlan: [
              { quarter: "1Q", title: "실험 계획법 및 AI 분석" },
              { quarter: "2Q", title: "AI 기반 불량 원인 분석" },
              { quarter: "3Q", title: "데이터 레이크 구축 기초" },
              { quarter: "4Q", title: "R&D 데이터 자산화 전략" }
            ]
          },
          {
            category: "영업/마케팅",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 고객 여정 분석" },
              { quarter: "2Q", title: "개인화 마케팅 자동화" },
              { quarter: "3Q", title: "AI 영업 지원 도구 활용" },
              { quarter: "4Q", title: "데이터 기반 매출 예측" }
            ]
          },
          {
            category: "경영지원/인사",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 채용 및 평가 기초" },
              { quarter: "2Q", title: "사내 지식 베이스 AI 구축" },
              { quarter: "3Q", title: "AI 보안 가이드라인 수립" },
              { quarter: "4Q", title: "디지털 변화 관리(Change Management)" }
            ]
          }
        ]
      },
      L3: {
        directions: [
          "부서 간 데이터 사일로 해소 및 통합 데이터 레이크 구축",
          "AI 성과 관리 체계(KPI) 도입 및 비즈니스 가치 연결",
          "실무자 중심의 자율적 AI 혁신 문화 정착"
        ],
        steps: [
          { title: "데이터 통합 및 고도화", period: "1~3개월", content: "전사 통합 데이터 플랫폼 구축 착수. 부서별 데이터 표준화 및 클렌징 가속화." },
          { title: "전사 확산 및 내재화", period: "3~6개월", content: "성공한 PoC의 전사 확대 적용. 부서별 AI 활용 성과 지표 관리 및 피드백." },
          { title: "지능형 프로세스 구축", period: "6~12개월", content: "주요 의사결정 프로세스에 AI 분석 결과 반영. AI 기반 예측 모델링 고도화." }
        ],
        courses: [
          { category: "기초", content: "AI 윤리와 책임 있는 AI 활용 (Responsible AI)" },
          { category: "실무", content: "파이썬 기반 데이터 분석 및 머신러닝 기초" },
          { category: "고급", content: "AI 비즈니스 모델 혁신 및 전략 기획" }
        ],
        pocs: [
          { area: "영업/마케팅", task: "고객 이탈 예측 및 맞춤형 프로모션 추천" },
          { area: "품질/R&D", task: "AI 기반 최적 배합비/공정 조건 도출" }
        ],
        jobTraining: [
          {
            category: "경영진",
            annualPlan: [
              { quarter: "1Q", title: "전사 데이터 통합 리더십" },
              { quarter: "2Q", title: "AI 기반 비즈니스 모델 혁신" },
              { quarter: "3Q", title: "AI 성과 지표(KPI) 고도화" },
              { quarter: "4Q", title: "글로벌 AI 생태계 협력" }
            ]
          },
          {
            category: "생산/공정",
            annualPlan: [
              { quarter: "1Q", title: "지능형 생산 계획 최적화" },
              { quarter: "2Q", title: "AI 기반 에너지 효율 관리" },
              { quarter: "3Q", title: "디지털 트윈 구축 및 운영" },
              { quarter: "4Q", title: "공정 지능화 전문가 과정" }
            ]
          },
          {
            category: "품질/R&D",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 신소재/제품 설계" },
              { quarter: "2Q", title: "데이터 기반 품질 예측 모델" },
              { quarter: "3Q", title: "고급 통계 및 머신러닝" },
              { quarter: "4Q", title: "R&D 지식 그래프 구축" }
            ]
          },
          {
            category: "영업/마케팅",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 수요 예측 고도화" },
              { quarter: "2Q", title: "초개인화 고객 경험 설계" },
              { quarter: "3Q", title: "AI 챗봇 및 상담 자동화" },
              { quarter: "4Q", title: "데이터 기반 가격 최적화" }
            ]
          },
          {
            category: "경영지원/인사",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 인재 육성 전략" },
              { quarter: "2Q", title: "지능형 재무/회계 관리" },
              { quarter: "3Q", title: "AI 법률 및 규제 대응" },
              { quarter: "4Q", title: "전사 AI 문화 확산 전략" }
            ]
          }
        ]
      },
      L4: {
        directions: [
          "실시간 데이터 기반의 지능형 자율 경영 체계 구축",
          "AI 모델의 지속적 학습 및 고도화(MLOps) 체계 마련",
          "AI 기반의 신규 비즈니스 모델 및 수익원 창출"
        ],
        steps: [
          { title: "운영 최적화 및 안정화", period: "1~3개월", content: "전사 AI 운영 현황 모니터링 및 성능 최적화. MLOps 도입을 통한 모델 관리 자동화." },
          { title: "비즈니스 모델 혁신", period: "3~6개월", content: "AI 기반 신규 서비스/제품 기획 및 시장 검증. 고객 경험(CX) 혁신 프로젝트 추진." },
          { title: "글로벌 경쟁력 확보", period: "6~12개월", content: "글로벌 수준의 AI 거버넌스 확립. 외부 생태계(스타트업 등)와의 오픈 이노베이션 강화." }
        ],
        courses: [
          { category: "기초", content: "최신 AI 트렌드와 산업별 혁신 사례 세미나" },
          { category: "실무", content: "LLM(거대언어모델) 파인튜닝 및 사내 지식 베이스 구축" },
          { category: "고급", content: "AI 최고 책임자(CAIO) 역량 강화 과정" }
        ],
        pocs: [
          { area: "전략/기획", task: "실시간 시장 지표 기반 수요 예측 및 재고 최적화" },
          { area: "고객 서비스", task: "멀티모달 AI 기반 지능형 고객 상담 센터 구축" }
        ],
        jobTraining: [
          {
            category: "경영진",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 시나리오 경영" },
              { quarter: "2Q", title: "글로벌 AI 파트너십 전략" },
              { quarter: "3Q", title: "AI 윤리 및 ESG 경영" },
              { quarter: "4Q", title: "미래 기술 로드맵 고도화" }
            ]
          },
          {
            category: "생산/공정",
            annualPlan: [
              { quarter: "1Q", title: "자율 주행 물류 및 로봇" },
              { quarter: "2Q", title: "실시간 공정 자율 최적화" },
              { quarter: "3Q", title: "MLOps 기반 공정 관리" },
              { quarter: "4Q", title: "스마트 팩토리 마스터 과정" }
            ]
          },
          {
            category: "품질/R&D",
            annualPlan: [
              { quarter: "1Q", title: "생성형 AI 기반 가상 설계" },
              { quarter: "2Q", title: "AI 기반 지식 재산권 전략" },
              { quarter: "3Q", title: "고급 딥러닝 및 강화학습" },
              { quarter: "4Q", title: "R&D 디지털 트랜스포메이션" }
            ]
          },
          {
            category: "영업/마케팅",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 옴니채널 마케팅" },
              { quarter: "2Q", title: "실시간 고객 감성 분석" },
              { quarter: "3Q", title: "AI 기반 신시장 발굴" },
              { quarter: "4Q", title: "데이터 기반 브랜드 전략" }
            ]
          },
          {
            category: "경영지원/인사",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 조직 문화 진단" },
              { quarter: "2Q", title: "지능형 리스크 관리 시스템" },
              { quarter: "3Q", title: "AI 기반 전략적 인력 계획" },
              { quarter: "4Q", title: "디지털 경영 지원 고도화" }
            ]
          }
        ]
      },
      L5: {
        directions: [
          "AI First 문화의 완전한 정착 및 창의적 혁신 가속화",
          "글로벌 AI 생태계 리딩 및 기술 표준 주도",
          "지속 가능한 AI 성장을 위한 ESG 경영 연계"
        ],
        steps: [
          { title: "초격차 기술 확보", period: "1~3개월", content: "차세대 AI 기술(Agentic AI 등) 선제적 도입 및 연구. 사내 AI 연구소 역량 강화." },
          { title: "생태계 확장", period: "3~6개월", content: "플랫폼 비즈 전환 및 외부 파트너십 확장. AI 기술 수출 및 글로벌 브랜딩." },
          { title: "미래 가치 창출", period: "6~12개월", content: "AI 기반의 완전한 비즈니스 트랜스포메이션 달성. 사회적 가치 창출을 위한 AI 프로젝트." }
        ],
        courses: [
          { category: "기초", content: "미래 사회와 AI의 역할 (인문학적 통찰)" },
          { category: "실무", content: "최첨단 AI 알고리즘 및 아키텍처 설계" },
          { category: "고급", content: "글로벌 AI 거버넌스 및 표준화 리더십" }
        ],
        pocs: [
          { area: "신사업", task: "AI 기반 자율 운영 공장(Lights-out Factory) 모델 구축" },
          { area: "R&D", task: "생성형 AI 기반 신소재/신제품 가상 설계 및 시뮬레이션" }
        ],
        jobTraining: [
          {
            category: "경영진",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 초격차 경영 전략" },
              { quarter: "2Q", title: "글로벌 AI 표준 주도 전략" },
              { quarter: "3Q", title: "AI 기반 사회적 가치 창출" },
              { quarter: "4Q", title: "미래 인류와 AI 공존 전략" }
            ]
          },
          {
            category: "생산/공정",
            annualPlan: [
              { quarter: "1Q", title: "완전 자율 생산 시스템 구축" },
              { quarter: "2Q", title: "AI 기반 글로벌 공급망 리딩" },
              { quarter: "3Q", title: "차세대 제조 AI 연구" },
              { quarter: "4Q", title: "제조 지능화 글로벌 리더" }
            ]
          },
          {
            category: "품질/R&D",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 자율 연구 시스템" },
              { quarter: "2Q", title: "차세대 AI 알고리즘 개발" },
              { quarter: "3Q", title: "AI 기반 혁신 제품 창출" },
              { quarter: "4Q", title: "글로벌 R&D 생태계 주도" }
            ]
          },
          {
            category: "영업/마케팅",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 글로벌 시장 선점" },
              { quarter: "2Q", title: "초지능형 고객 경험 창출" },
              { quarter: "3Q", title: "AI 기반 브랜드 가치 극대화" },
              { quarter: "4Q", title: "미래 마케팅 패러다임 주도" }
            ]
          },
          {
            category: "경영지원/인사",
            annualPlan: [
              { quarter: "1Q", title: "AI 기반 자율 조직 운영" },
              { quarter: "2Q", title: "지능형 글로벌 경영 지원" },
              { quarter: "3Q", title: "AI 기반 미래 인재 경영" },
              { quarter: "4Q", title: "디지털 트랜스포메이션 완성" }
            ]
          }
        ]
      }
    };
    return strategies[level];
  };

  const downloadWordReport = async () => {
    if (!diagnosticResult) return;

    const strategy = getExecutionStrategy(diagnosticResult.overallLevel);
    const deptAvg = (formData.deptScores.reduce((a, b) => a + b.score, 0) / formData.deptScores.length).toFixed(1);
    const dateStr = new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' });

    const thinBorder = {
      top: { style: BorderStyle.SINGLE, size: 1, color: "d1d5db" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "d1d5db" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "d1d5db" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "d1d5db" },
    };
    const noBorder = {
      top: { style: BorderStyle.NONE, size: 0, color: "ffffff" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "ffffff" },
      left: { style: BorderStyle.NONE, size: 0, color: "ffffff" },
      right: { style: BorderStyle.NONE, size: 0, color: "ffffff" },
    };

    const sectionTitle = (num: string, title: string) => new Paragraph({
      children: [
        new TextRun({ text: `${num}. `, bold: true, size: 24, color: "4f46e5" }),
        new TextRun({ text: title, bold: true, size: 24, color: "1e293b" }),
      ],
      spacing: { before: 500, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "e2e8f0", space: 8 } },
    });

    const subHeading = (text: string) => new Paragraph({
      children: [new TextRun({ text, bold: true, size: 20, color: "334155" })],
      spacing: { before: 300, after: 120 },
    });

    const bodyText = (text: string) => new Paragraph({
      children: [new TextRun({ text, size: 20, color: "475569" })],
      spacing: { after: 160, line: 360 },
    });

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "맑은 고딕", size: 20 },
          },
        },
      },
      sections: [{
        properties: {
          page: {
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
          },
        },
        children: [
          // ── Cover Section ──
          new Paragraph({ text: "", spacing: { after: 600 } }),
          new Paragraph({
            children: [new TextRun({ text: "AX DIAGNOSTIC REPORT", bold: true, size: 16, color: "4f46e5", font: "맑은 고딕" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "기업 AI 전환 수준 진단 보고서", bold: true, size: 36, color: "0f172a" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [new TextRun({ text: diagnosticResult.companyName, bold: true, size: 28, color: "4f46e5" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Cover info table
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [2800, 6200],
            rows: [
              ...([
                ["기업명", diagnosticResult.companyName],
                ["업종", formData.industry],
                ["종업원 수", formData.employeeCount],
                ["진단일", dateStr],
                ["종합 레벨", `${diagnosticResult.overallLevel} — ${diagnosticResult.overallLevelName}`],
              ] as [string, string][]).map(([label, value]) =>
                new TableRow({
                  children: [
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, color: "64748b" })], spacing: { before: 60, after: 60 }, indent: { left: 120 } })],
                      width: { size: 2800, type: WidthType.DXA },
                      shading: { fill: "f8fafc" },
                      borders: thinBorder,
                    }),
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: value, size: 18, color: "1e293b" })], spacing: { before: 60, after: 60 }, indent: { left: 120 } })],
                      width: { size: 6200, type: WidthType.DXA },
                      borders: thinBorder,
                    }),
                  ],
                })
              ),
            ],
          }),
          new Paragraph({ text: "", spacing: { after: 200 } }),

          // Score summary cards
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [2250, 2250, 2250, 2250],
            rows: [
              new TableRow({
                children: [
                  ...([
                    { label: "종합 점수", value: `${diagnosticResult.overallScore}`, level: diagnosticResult.overallLevel, bg: "eef2ff" },
                    { label: "경영진 의지", value: `${formData.execScore}`, level: diagnosticResult.execLevel, bg: "fffbeb" },
                    { label: "부서 평균", value: deptAvg, level: getLevel(Number(deptAvg)), bg: "f0fdf4" },
                    { label: "개인 역량", value: `${formData.individualScore}`, level: diagnosticResult.individualLevel, bg: "ecfdf5" },
                  ]).map(card =>
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: card.label, size: 14, color: "94a3b8", bold: true })],
                          alignment: AlignmentType.CENTER,
                          spacing: { before: 120 },
                        }),
                        new Paragraph({
                          children: [new TextRun({ text: card.value, size: 32, bold: true, color: "1e293b" })],
                          alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({
                          children: [new TextRun({ text: `${card.level} — ${LEVEL_MAP[card.level].name.split(' ')[0]}`, size: 14, bold: true, color: "4f46e5" })],
                          alignment: AlignmentType.CENTER,
                          spacing: { after: 120 },
                        }),
                      ],
                      width: { size: 2250, type: WidthType.DXA },
                      shading: { fill: card.bg },
                      borders: thinBorder,
                    })
                  ),
                ],
              }),
            ],
          }),

          // ── 01. 종합 진단 ──
          sectionTitle("01", "종합 진단 결과"),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [9000],
            rows: [new TableRow({
              children: [new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: `${diagnosticResult.overallLevel}  `, bold: true, size: 28, color: "4f46e5" }),
                      new TextRun({ text: diagnosticResult.overallLevelName, bold: true, size: 22, color: "1e293b" }),
                      new TextRun({ text: `    종합 ${diagnosticResult.overallScore} / 5.0`, size: 20, color: "64748b" }),
                    ],
                    spacing: { before: 160, after: 160 },
                    indent: { left: 200 },
                  }),
                ],
                shading: { fill: "f1f5f9" },
                borders: thinBorder,
              })],
            })],
          }),
          new Paragraph({ text: "", spacing: { after: 120 } }),
          bodyText(aiReport?.overall || getConsultingText(diagnosticResult.overallLevel, 'overall')),
          ...(aiReport?.summary ? [
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              columnWidths: [9000],
              rows: [new TableRow({
                children: [new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: "KEY INSIGHT", bold: true, size: 14, color: "4f46e5" })],
                      spacing: { before: 120 },
                      indent: { left: 200 },
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: aiReport.summary, bold: true, size: 20, color: "1e293b" })],
                      spacing: { before: 80, after: 120 },
                      indent: { left: 200, right: 200 },
                    }),
                  ],
                  shading: { fill: "eef2ff" },
                  borders: thinBorder,
                })],
              })],
            }),
          ] : []),

          // ── 02. 경영진 ──
          sectionTitle("02", "경영진 의지 분석"),
          new Paragraph({
            children: [
              new TextRun({ text: `현재 수준: ${formData.execScore}점`, bold: true, size: 20, color: "1e293b" }),
              new TextRun({ text: `  (${diagnosticResult.execLevel} — ${LEVEL_MAP[diagnosticResult.execLevel].name})`, size: 18, color: "64748b" }),
            ],
            spacing: { after: 160 },
          }),
          bodyText(aiReport?.exec || getConsultingText(diagnosticResult.execLevel, 'exec')),

          // ── 03. 부서별 ──
          sectionTitle("03", "부서별 AX 준비도 분석"),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [2250, 2250, 2250, 2250],
            rows: [
              new TableRow({
                children: ["부서명", "점수", "레벨", "단계"].map(h =>
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18, color: "ffffff" })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 } })],
                    shading: { fill: "334155" },
                    borders: thinBorder,
                  })
                ),
              }),
              ...formData.deptScores.map((dept, i) =>
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: dept.name, bold: true, size: 18 })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })], shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" }, borders: thinBorder }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${dept.score}`, bold: true, size: 18, color: "4f46e5" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })], shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" }, borders: thinBorder }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: getLevel(dept.score), bold: true, size: 18, color: "4f46e5" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })], shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" }, borders: thinBorder }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: LEVEL_MAP[getLevel(dept.score)].name.split(' ')[0], size: 16, color: "64748b" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })], shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" }, borders: thinBorder }),
                  ],
                })
              ),
            ],
          }),
          new Paragraph({ text: "", spacing: { after: 160 } }),
          bodyText(aiReport?.dept || getConsultingText(diagnosticResult.overallLevel, 'dept')),

          // ── 04. 개인 역량 ──
          sectionTitle("04", "개인 AX 역량 분석"),
          new Paragraph({
            children: [
              new TextRun({ text: `현재 수준: ${formData.individualScore}점`, bold: true, size: 20, color: "1e293b" }),
              new TextRun({ text: `  (${diagnosticResult.individualLevel} — ${LEVEL_MAP[diagnosticResult.individualLevel].name})`, size: 18, color: "64748b" }),
            ],
            spacing: { after: 160 },
          }),
          bodyText(aiReport?.indiv || getConsultingText(diagnosticResult.individualLevel, 'indiv')),

          // ── 05. 실행 전략 ──
          sectionTitle("05", "맞춤형 실행 전략"),

          subHeading("5-1. 핵심 개선 방향"),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [3000, 3000, 3000],
            rows: [new TableRow({
              children: strategy.directions.map((dir, i) =>
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: `0${i + 1}`, bold: true, size: 28, color: "4f46e5" })], spacing: { before: 120 }, indent: { left: 120 } }),
                    new Paragraph({ children: [new TextRun({ text: dir, bold: true, size: 18, color: "1e293b" })], spacing: { after: 120 }, indent: { left: 120, right: 120 } }),
                  ],
                  shading: { fill: "f8fafc" },
                  borders: thinBorder,
                })
              ),
            })],
          }),

          subHeading("5-2. 단계별 실행 로드맵"),
          ...strategy.steps.map(s => new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [1600, 7400],
            rows: [new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: s.period, bold: true, size: 16, color: "ffffff" })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 } })],
                  width: { size: 1600, type: WidthType.DXA },
                  shading: { fill: "1e293b" },
                  borders: thinBorder,
                  verticalAlign: "center" as const,
                }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: s.title, bold: true, size: 20, color: "1e293b" })], spacing: { before: 80 }, indent: { left: 160 } }),
                    new Paragraph({ children: [new TextRun({ text: s.content, size: 18, color: "64748b" })], spacing: { after: 80 }, indent: { left: 160, right: 120 } }),
                  ],
                  shading: { fill: "ffffff" },
                  borders: thinBorder,
                }),
              ],
            })],
          })),

          subHeading("5-3. 추천 교육과정"),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [1400, 7600],
            rows: strategy.courses.map(c =>
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: c.category.toUpperCase(), bold: true, size: 16, color: "4f46e5" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })],
                    width: { size: 1400, type: WidthType.DXA },
                    shading: { fill: "eef2ff" },
                    borders: thinBorder,
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: c.content, size: 18 })], spacing: { before: 60, after: 60 }, indent: { left: 120 } })],
                    borders: thinBorder,
                  }),
                ],
              })
            ),
          }),

          subHeading("5-4. 직군별 연간 교육 계획"),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [1800, 1800, 1800, 1800, 1800],
            rows: [
              new TableRow({
                children: ["직군", "1Q", "2Q", "3Q", "4Q"].map(h =>
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 16, color: "ffffff" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })],
                    shading: { fill: "334155" },
                    borders: thinBorder,
                  })
                ),
              }),
              ...strategy.jobTraining.map((job, i) =>
                new TableRow({
                  children: [
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: job.category, bold: true, size: 16 })], alignment: AlignmentType.CENTER, spacing: { before: 50, after: 50 } })],
                      shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" },
                      borders: thinBorder,
                    }),
                    ...job.annualPlan.map(plan =>
                      new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: plan.title, size: 15, color: "475569" })], alignment: AlignmentType.CENTER, spacing: { before: 50, after: 50 } })],
                        shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" },
                        borders: thinBorder,
                      })
                    ),
                  ],
                })
              ),
            ],
          }),

          subHeading("5-5. 추천 PoC 과제"),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [2250, 6750],
            rows: [
              new TableRow({
                children: ["영역", "과제"].map(h =>
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 16, color: "ffffff" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })],
                    shading: { fill: "334155" },
                    borders: thinBorder,
                  })
                ),
              }),
              ...strategy.pocs.map((poc, i) =>
                new TableRow({
                  children: [
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: poc.area, bold: true, size: 18, color: "4f46e5" })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })],
                      width: { size: 2250, type: WidthType.DXA },
                      shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" },
                      borders: thinBorder,
                    }),
                    new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: poc.task, size: 18 })], spacing: { before: 60, after: 60 }, indent: { left: 120 } })],
                      shading: { fill: i % 2 === 0 ? "f8fafc" : "ffffff" },
                      borders: thinBorder,
                    }),
                  ],
                })
              ),
            ],
          }),

          // ── Footer ──
          new Paragraph({ text: "", spacing: { after: 600 } }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [9000],
            rows: [new TableRow({
              children: [new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: aiReport ? "AI 기반 맞춤형 컨설팅 보고서" : "자동 생성 분석 보고서", size: 16, color: "94a3b8" }),
                      new TextRun({ text: `  |  ${dateStr}  |  AX Diagnostic Analyzer`, size: 16, color: "94a3b8" }),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120 },
                  }),
                ],
                shading: { fill: "f8fafc" },
                borders: noBorder,
              })],
            })],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `AX_Report_${diagnosticResult.companyName}_${new Date().toISOString().slice(0,10)}.docx`);
  };

  const resetForm = () => {
    setStep('input');
    setViewMode('dashboard');
  };

  return (
    <div className="min-h-screen p-4 md:p-8 lg:p-12">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <header className="mb-10 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-11 h-11 bg-gradient-to-br from-indigo-600 to-indigo-700 rounded-2xl flex items-center justify-center text-white shadow-lg shadow-indigo-200">
              <Activity size={22} />
            </div>
            <div>
              <h1 className="text-xl font-extrabold text-slate-900 tracking-tight">AX Diagnostic Analyzer</h1>
              <p className="text-slate-400 text-xs font-medium mt-0.5">Enterprise AI Transformation Level Diagnostic</p>
            </div>
          </div>
          {step === 'result' && (
            <div className="flex items-center gap-3">
              <div className="flex bg-slate-100 p-1 rounded-xl">
                {(['dashboard', 'report', 'strategy'] as const).map((mode) => (
                  <button
                    key={mode}
                    onClick={() => setViewMode(mode)}
                    className={cn(
                      "px-4 py-2 text-xs font-semibold rounded-lg transition-all",
                      viewMode === mode
                        ? "bg-white text-indigo-600 shadow-sm"
                        : "text-slate-500 hover:text-slate-700"
                    )}
                  >
                    {mode === 'dashboard' ? 'Dashboard' : mode === 'report' ? 'AI Report' : 'Strategy'}
                  </button>
                ))}
              </div>
              <button
                onClick={resetForm}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all"
              >
                <RefreshCcw size={14} />
                새 진단
              </button>
            </div>
          )}
        </header>

        <AnimatePresence mode="wait">
          {step === 'input' ? (
            <motion.div
              key="input-step"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -16 }}
              className="max-w-3xl mx-auto space-y-6"
            >
              {/* Hero */}
              <div className="text-center mb-8">
                <h2 className="text-3xl font-extrabold text-slate-900 mb-2">기업 AX 수준 진단</h2>
                <p className="text-slate-500 text-sm">AI 전환 준비도를 진단하고, AI 기반 맞춤형 컨설팅 보고서를 받아보세요.</p>
              </div>

              <section className="glass-card p-6 md:p-8">
                <h2 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-5">API 설정</h2>
                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-slate-600">OpenRouter API Key</label>
                  <input
                    type="password"
                    value={apiKey}
                    onChange={(e) => setApiKey(e.target.value)}
                    placeholder="sk-or-v1-..."
                    className="input-field font-mono text-sm"
                  />
                  <p className="text-[11px] text-slate-400">
                    <a href="https://openrouter.ai/keys" target="_blank" rel="noopener noreferrer" className="text-indigo-500 hover:underline">openrouter.ai/keys</a>에서 무료로 발급받으세요. 키는 브라우저에만 저장됩니다.
                  </p>
                </div>
              </section>

              <section className="glass-card p-6 md:p-8">
                <h2 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-5">기업 정보</h2>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-600">기업명</label>
                    <input
                      type="text"
                      value={formData.companyName}
                      onChange={(e) => setFormData({ ...formData, companyName: e.target.value })}
                      placeholder="(주)에이아이테크"
                      className="input-field"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-600">업종</label>
                    <input
                      type="text"
                      value={formData.industry}
                      onChange={(e) => setFormData({ ...formData, industry: e.target.value })}
                      placeholder="제조업 (자동차 부품)"
                      className="input-field"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-600">종업원 수</label>
                    <select
                      value={formData.employeeCount}
                      onChange={(e) => setFormData({ ...formData, employeeCount: e.target.value })}
                      className="input-field bg-white"
                    >
                      <option value="">선택하세요</option>
                      <option value="50인 미만">50인 미만</option>
                      <option value="50~100인">50~100인</option>
                      <option value="100~300인">100~300인</option>
                      <option value="300인 이상">300인 이상</option>
                    </select>
                  </div>
                </div>
              </section>

              <section className="glass-card p-6 md:p-8">
                <h2 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-6">AX 준비도 평가</h2>

                <div className="space-y-8">
                  {/* 경영진 */}
                  <div className="p-5 bg-amber-50/50 rounded-xl border border-amber-100/60">
                    <div className="flex justify-between items-center mb-3">
                      <label className="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <Award size={16} className="text-amber-500" />
                        경영진 추진 의지
                      </label>
                      <span className="text-lg font-black text-indigo-600 font-mono">{formData.execScore}</span>
                    </div>
                    <input
                      type="range" min="1" max="5" step="0.5"
                      value={formData.execScore}
                      onChange={(e) => setFormData({ ...formData, execScore: parseFloat(e.target.value) })}
                      className="w-full accent-indigo-600"
                    />
                    <p className="text-[11px] text-slate-400 mt-2">AI 도입에 대한 관심도와 자원 배분 의지</p>
                  </div>

                  {/* 부서별 */}
                  <div>
                    <label className="text-sm font-bold text-slate-800 flex items-center gap-2 mb-4">
                      <Briefcase size={16} className="text-indigo-500" />
                      부서별 준비도
                    </label>
                    <div className="grid grid-cols-2 gap-4">
                      {formData.deptScores.map((dept, idx) => (
                        <div key={dept.name} className="p-4 bg-slate-50/80 rounded-xl border border-slate-100">
                          <div className="flex justify-between items-center mb-2">
                            <span className="text-xs font-bold text-slate-600">{dept.name}</span>
                            <span className="text-sm font-black text-indigo-600 font-mono">{dept.score}</span>
                          </div>
                          <input
                            type="range" min="1" max="5" step="0.5"
                            value={dept.score}
                            onChange={(e) => handleDeptScoreChange(idx, parseFloat(e.target.value))}
                            className="w-full accent-indigo-600"
                          />
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* 개인 역량 */}
                  <div className="p-5 bg-emerald-50/50 rounded-xl border border-emerald-100/60">
                    <div className="flex justify-between items-center mb-3">
                      <label className="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <Users size={16} className="text-emerald-500" />
                        개인별 AX 역량 (평균)
                      </label>
                      <span className="text-lg font-black text-indigo-600 font-mono">{formData.individualScore}</span>
                    </div>
                    <input
                      type="range" min="1" max="5" step="0.5"
                      value={formData.individualScore}
                      onChange={(e) => setFormData({ ...formData, individualScore: parseFloat(e.target.value) })}
                      className="w-full accent-indigo-600"
                    />
                    <p className="text-[11px] text-slate-400 mt-2">AI 도구 활용 능력과 데이터 리터러시 수준</p>
                  </div>
                </div>

                <button
                  onClick={runDiagnostic}
                  disabled={!apiKey || !formData.companyName || !formData.industry || aiLoading}
                  className="w-full mt-8 py-3.5 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-700 hover:to-indigo-800 disabled:from-slate-300 disabled:to-slate-300 text-white font-bold text-sm rounded-xl shadow-lg shadow-indigo-200/50 transition-all flex items-center justify-center gap-2 group"
                >
                  {aiLoading ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                      AI 분석 중...
                    </>
                  ) : (
                    <>
                      AI 진단 시작
                      <ChevronRight size={18} className="group-hover:translate-x-0.5 transition-transform" />
                    </>
                  )}
                </button>
              </section>

              {/* Level Reference */}
              <div className="flex items-center justify-center gap-6 py-4">
                {Object.entries(LEVEL_MAP).map(([level, info]) => (
                  <div key={level} className="flex items-center gap-1.5 text-[11px]">
                    <div className={cn("w-2 h-2 rounded-full", info.color)} />
                    <span className="font-bold text-slate-500">{level}</span>
                    <span className="text-slate-400">{info.range}</span>
                  </div>
                ))}
              </div>
            </motion.div>
          ) : (
            <motion.div
              key="result-step"
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6"
            >
              {viewMode === 'dashboard' ? (
                <>
                  {/* Score Overview */}
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                    <div className="glass-card p-6 text-center md:row-span-2">
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Overall</p>
                      <div className={cn(
                        "w-24 h-24 mx-auto rounded-2xl flex flex-col items-center justify-center text-white shadow-lg mb-3",
                        LEVEL_MAP[diagnosticResult?.overallLevel || 'L1'].color
                      )}>
                        <span className="text-3xl font-black">{diagnosticResult?.overallLevel}</span>
                      </div>
                      <h2 className="text-sm font-bold text-slate-900 mb-0.5">{diagnosticResult?.overallLevelName}</h2>
                      <p className="text-indigo-600 font-black text-lg font-mono">{diagnosticResult?.overallScore}<span className="text-slate-400 text-xs font-medium"> / 5.0</span></p>
                    </div>
                    {[
                      { label: '경영진 의지', score: formData.execScore, level: diagnosticResult?.execLevel || 'L1', color: 'border-l-amber-400' },
                      { label: '부서 평균', score: Number((formData.deptScores.reduce((a, b) => a + b.score, 0) / formData.deptScores.length).toFixed(1)), level: getLevel(formData.deptScores.reduce((a, b) => a + b.score, 0) / formData.deptScores.length), color: 'border-l-indigo-400' },
                      { label: '개인 역량', score: formData.individualScore, level: diagnosticResult?.individualLevel || 'L1', color: 'border-l-emerald-400' },
                    ].map((item) => (
                      <div key={item.label} className={cn("glass-card p-4 border-l-[3px]", item.color)}>
                        <span className="text-[10px] font-bold text-slate-400 uppercase">{item.label}</span>
                        <div className="flex items-end justify-between mt-1">
                          <span className="text-xl font-black text-slate-900 font-mono">{item.score}</span>
                          <LevelBadge level={item.level} />
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
                    {/* Radar Chart */}
                    <div className="lg:col-span-2 glass-card p-6">
                      <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">Competency Radar</h3>
                      <div className="h-64 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <RadarChart cx="50%" cy="50%" outerRadius="75%" data={diagnosticResult?.chartData}>
                            <PolarGrid stroke="#e2e8f0" strokeDasharray="3 3" />
                            <PolarAngleAxis dataKey="name" tick={{ fill: '#94a3b8', fontSize: 11, fontWeight: 600 }} />
                            <PolarRadiusAxis tick={false} axisLine={false} />
                            <Radar name="Score" dataKey="score" stroke="#4f46e5" fill="#4f46e5" fillOpacity={0.15} strokeWidth={2} />
                          </RadarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    {/* AI Analysis */}
                    <div className="lg:col-span-3 space-y-4">
                      <div className="glass-card p-6 bg-gradient-to-br from-indigo-50/80 to-white border-indigo-100/60">
                        <h3 className="text-xs font-bold text-indigo-500 uppercase tracking-widest mb-3 flex items-center gap-2">
                          AI 종합 진단
                          {aiLoading && <span className="text-indigo-400 animate-pulse">분석 중...</span>}
                        </h3>
                        <p className="text-sm text-slate-700 leading-relaxed whitespace-pre-line">
                          {aiReport?.overall || getConsultingText(diagnosticResult?.overallLevel || 'L1', 'overall')}
                        </p>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="glass-card p-5">
                          <h4 className="text-[10px] font-bold text-amber-500 uppercase tracking-widest mb-2">경영진 분석</h4>
                          <p className="text-xs text-slate-600 leading-relaxed whitespace-pre-line">
                            {aiReport?.exec || getConsultingText(diagnosticResult?.execLevel || 'L1', 'exec')}
                          </p>
                        </div>
                        <div className="glass-card p-5">
                          <h4 className="text-[10px] font-bold text-indigo-500 uppercase tracking-widest mb-2">부서별 분석</h4>
                          <p className="text-xs text-slate-600 leading-relaxed whitespace-pre-line">
                            {aiReport?.dept || getConsultingText(diagnosticResult?.overallLevel || 'L1', 'dept')}
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              ) : viewMode === 'report' ? (
                <motion.div
                  initial={{ opacity: 0, y: 16 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="max-w-4xl mx-auto"
                >
                  <div className="glass-card overflow-hidden">
                    {/* Report Header */}
                    <div className="bg-gradient-to-r from-slate-900 to-slate-800 p-8 text-white">
                      <div className="flex justify-between items-start">
                        <div>
                          <p className="text-slate-400 text-[10px] font-bold uppercase tracking-[0.2em] mb-2">AX Diagnostic Report</p>
                          <h2 className="text-2xl font-extrabold mb-3">{diagnosticResult?.companyName} AX 진단 보고서</h2>
                          <div className="flex items-center gap-4 text-xs text-slate-400">
                            <span>업종: {formData.industry}</span>
                            <span>규모: {formData.employeeCount}</span>
                            <span>진단일: {new Date().toLocaleDateString('ko-KR')}</span>
                          </div>
                        </div>
                        <button
                          onClick={downloadWordReport}
                          className="flex items-center gap-2 px-4 py-2.5 bg-white/10 hover:bg-white/20 text-white text-xs font-bold rounded-lg backdrop-blur transition-all"
                        >
                          <Download size={14} />
                          Word 다운로드
                        </button>
                      </div>
                      {/* Score Summary Bar */}
                      <div className="mt-6 grid grid-cols-4 gap-3">
                        {[
                          { label: '종합', score: diagnosticResult?.overallScore, level: diagnosticResult?.overallLevel },
                          { label: '경영진', score: formData.execScore, level: diagnosticResult?.execLevel },
                          { label: '부서 평균', score: Number((formData.deptScores.reduce((a, b) => a + b.score, 0) / formData.deptScores.length).toFixed(1)), level: getLevel(formData.deptScores.reduce((a, b) => a + b.score, 0) / formData.deptScores.length) },
                          { label: '개인', score: formData.individualScore, level: diagnosticResult?.individualLevel },
                        ].map((item) => (
                          <div key={item.label} className="bg-white/5 rounded-lg p-3 text-center">
                            <p className="text-[10px] text-slate-400 font-medium mb-1">{item.label}</p>
                            <p className="text-lg font-black font-mono">{item.score}</p>
                            <LevelBadge level={item.level || 'L1'} />
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Report Body */}
                    <div className="p-8 space-y-10">
                      {aiLoading && (
                        <div className="flex flex-col items-center justify-center py-12">
                          <div className="w-10 h-10 border-[3px] border-indigo-100 border-t-indigo-600 rounded-full animate-spin mb-4" />
                          <p className="text-slate-400 text-sm font-semibold">AI가 보고서를 작성하고 있습니다...</p>
                        </div>
                      )}

                      {aiError && (
                        <div className="p-5 bg-red-50 border border-red-100 rounded-xl flex items-center justify-between">
                          <div>
                            <p className="text-red-600 font-bold text-sm">AI 보고서 생성 실패</p>
                            <p className="text-red-400 text-xs mt-0.5">{aiError}</p>
                          </div>
                          <button onClick={runDiagnostic} className="px-4 py-2 bg-red-600 text-white rounded-lg text-xs font-bold hover:bg-red-700 transition-all">
                            재시도
                          </button>
                        </div>
                      )}

                      {/* 01. 종합 */}
                      <section>
                        <div className="flex items-center gap-3 mb-4">
                          <span className="w-7 h-7 rounded-lg bg-indigo-600 text-white flex items-center justify-center text-[11px] font-black">01</span>
                          <h3 className="text-lg font-extrabold text-slate-900">종합 진단 결과</h3>
                        </div>
                        <div className="p-6 bg-slate-50 rounded-xl">
                          <div className="flex items-center gap-3 mb-4">
                            <LevelBadge level={diagnosticResult?.overallLevel || 'L1'} />
                            <span className="text-lg font-bold text-slate-900">{diagnosticResult?.overallLevelName}</span>
                          </div>
                          <p className="text-sm text-slate-700 leading-[1.8] whitespace-pre-line">
                            {aiReport?.overall || getConsultingText(diagnosticResult?.overallLevel || 'L1', 'overall')}
                          </p>
                          {aiReport?.summary && (
                            <div className="mt-5 p-4 bg-indigo-50 rounded-lg border border-indigo-100/60">
                              <p className="text-xs font-bold text-indigo-600 uppercase tracking-widest mb-1">Key Insight</p>
                              <p className="text-sm text-indigo-700 font-semibold leading-relaxed whitespace-pre-line">{aiReport.summary}</p>
                            </div>
                          )}
                        </div>
                      </section>

                      {/* 02. 경영진 */}
                      <section>
                        <div className="flex items-center gap-3 mb-4">
                          <span className="w-7 h-7 rounded-lg bg-indigo-600 text-white flex items-center justify-center text-[11px] font-black">02</span>
                          <h3 className="text-lg font-extrabold text-slate-900">경영진 의지 분석</h3>
                          <div className="ml-auto flex items-center gap-2">
                            <span className="text-sm font-mono font-bold text-slate-500">{formData.execScore}점</span>
                            <LevelBadge level={diagnosticResult?.execLevel || 'L1'} />
                          </div>
                        </div>
                        <p className="text-sm text-slate-600 leading-[1.8] p-6 bg-white rounded-xl border border-slate-100 whitespace-pre-line">
                          {aiReport?.exec || getConsultingText(diagnosticResult?.execLevel || 'L1', 'exec')}
                        </p>
                      </section>

                      {/* 03. 부서별 */}
                      <section>
                        <div className="flex items-center gap-3 mb-4">
                          <span className="w-7 h-7 rounded-lg bg-indigo-600 text-white flex items-center justify-center text-[11px] font-black">03</span>
                          <h3 className="text-lg font-extrabold text-slate-900">부서별 준비도 분석</h3>
                        </div>
                        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                          {formData.deptScores.map(dept => (
                            <div key={dept.name} className="p-3 bg-slate-50 rounded-lg text-center">
                              <p className="text-xs font-bold text-slate-500 mb-1">{dept.name}</p>
                              <p className="text-lg font-black text-slate-900 font-mono">{dept.score}</p>
                              <LevelBadge level={getLevel(dept.score)} />
                            </div>
                          ))}
                        </div>
                        <p className="text-sm text-slate-600 leading-[1.8] p-6 bg-white rounded-xl border border-slate-100 whitespace-pre-line">
                          {aiReport?.dept || getConsultingText(diagnosticResult?.overallLevel || 'L1', 'dept')}
                        </p>
                      </section>

                      {/* 04. 개인 */}
                      <section>
                        <div className="flex items-center gap-3 mb-4">
                          <span className="w-7 h-7 rounded-lg bg-indigo-600 text-white flex items-center justify-center text-[11px] font-black">04</span>
                          <h3 className="text-lg font-extrabold text-slate-900">개인 AX 역량 분석</h3>
                          <div className="ml-auto flex items-center gap-2">
                            <span className="text-sm font-mono font-bold text-slate-500">{formData.individualScore}점</span>
                            <LevelBadge level={diagnosticResult?.individualLevel || 'L1'} />
                          </div>
                        </div>
                        <p className="text-sm text-slate-600 leading-[1.8] p-6 bg-white rounded-xl border border-slate-100 whitespace-pre-line">
                          {aiReport?.indiv || getConsultingText(diagnosticResult?.individualLevel || 'L1', 'indiv')}
                        </p>
                      </section>

                      {/* Footer */}
                      <div className="pt-8 border-t border-slate-100 text-center">
                        <p className="text-[11px] text-slate-400">
                          {aiReport ? 'AI 기반 맞춤형 컨설팅 보고서 | Powered by OpenRouter' : '자동 생성 분석 보고서'} | {new Date().toLocaleDateString('ko-KR')}
                        </p>
                      </div>
                    </div>
                  </div>
                </motion.div>
              ) : (
                <motion.div
                  initial={{ opacity: 0, y: 16 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="space-y-6"
                >
                  {/* Directions */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {getExecutionStrategy(diagnosticResult?.overallLevel || 'L1').directions.map((dir, idx) => (
                      <div key={idx} className="glass-card p-5 border-t-[3px] border-t-indigo-500">
                        <span className="text-indigo-500 font-black text-xl font-mono block mb-2">0{idx + 1}</span>
                        <p className="text-sm font-bold text-slate-800 leading-snug">{dir}</p>
                      </div>
                    ))}
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    {/* Roadmap */}
                    <div className="lg:col-span-2 glass-card p-6">
                      <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-6">Implementation Roadmap</h3>
                      <div className="relative pl-6 border-l-2 border-indigo-100 space-y-8">
                        {getExecutionStrategy(diagnosticResult?.overallLevel || 'L1').steps.map((s, idx) => (
                          <div key={idx} className="relative">
                            <div className="absolute -left-[29px] top-0.5 w-3 h-3 rounded-full bg-indigo-600 ring-4 ring-white" />
                            <div className="flex items-center gap-3 mb-1.5">
                              <span className="px-2.5 py-1 bg-slate-900 text-white text-[10px] font-bold rounded-md">{s.period}</span>
                              <h4 className="text-sm font-bold text-slate-900">{s.title}</h4>
                            </div>
                            <p className="text-xs text-slate-500 leading-relaxed">{s.content}</p>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Education & PoC */}
                    <div className="space-y-6">
                      <div className="glass-card p-5 bg-slate-900 text-white border-none">
                        <h3 className="text-[10px] font-bold uppercase tracking-widest text-indigo-400 mb-4">Recommended Courses</h3>
                        <div className="space-y-3">
                          {getExecutionStrategy(diagnosticResult?.overallLevel || 'L1').courses.map((course, idx) => (
                            <div key={idx} className="p-3 bg-white/10 rounded-lg">
                              <span className="text-[9px] font-bold uppercase tracking-widest text-indigo-400 block mb-0.5">{course.category}</span>
                              <p className="text-xs font-medium">{course.content}</p>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div className="glass-card p-5">
                        <h3 className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-4">PoC Tasks</h3>
                        <div className="space-y-3">
                          {getExecutionStrategy(diagnosticResult?.overallLevel || 'L1').pocs.map((poc, idx) => (
                            <div key={idx} className="p-3 bg-slate-50 rounded-lg border border-slate-100">
                              <span className="text-[10px] font-bold text-indigo-600 block mb-0.5">{poc.area}</span>
                              <p className="text-xs font-bold text-slate-800">{poc.task}</p>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Job Training */}
                  <div className="glass-card p-6">
                    <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-5">Annual Training Plan by Role</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      {getExecutionStrategy(diagnosticResult?.overallLevel || 'L1').jobTraining.map((job, idx) => (
                        <div key={idx} className="p-4 bg-slate-50 rounded-xl">
                          <h4 className="text-xs font-bold text-slate-900 mb-3 flex items-center gap-2">
                            <div className="w-1.5 h-1.5 rounded-full bg-indigo-500" />
                            {job.category}
                          </h4>
                          <div className="space-y-2">
                            {job.annualPlan.map((plan, pIdx) => (
                              <div key={pIdx} className="flex items-start gap-2">
                                <span className="flex-shrink-0 px-1.5 py-0.5 bg-indigo-100 text-indigo-600 text-[9px] font-black rounded">{plan.quarter}</span>
                                <p className="text-[11px] text-slate-600 leading-tight">{plan.title}</p>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Consultant Note */}
                  <div className="p-5 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 text-white">
                    <p className="text-xs font-medium leading-relaxed opacity-90">
                      <span className="font-bold">Insight</span> &mdash; AX는 기술의 도입이 아닌 문화의 변화입니다. 교육을 통해 인식을 바꾸고, PoC로 가치를 증명하며, 전사로 확산하는 선순환 구조를 만드십시오.
                    </p>
                  </div>
                </motion.div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}
