import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, RefreshCw, Sliders, TrendingUp, AlertCircle, CheckCircle2 } from 'lucide-react';

export default function App() {
  // 1. 기본 수수료 및 프로모션 설정 상태
  const [fees, setFees] = useState({
    musinsa_kr: 0.28,       // 무신사 국내 (28%)
    musinsa_global: 0.30,   // 무신사 글로벌 (30%)
    zozotown: 0.38,         // 조조타운 (38%)
  });
  
  const [promoConfig, setPromoConfig] = useState({
    exhibitionDiscount: 15, // 기획전 기본 할인율 (%)
    onlineCoupon: 10,       // 온라인 플랫폼 쿠폰율 (%)
    crmCoupon: 10,          // CRM 타겟 쿠폰율 (%)
    brandCrmShare: 0.5,     // CRM 쿠폰 브랜드사 부담 비율 (0.5 = 50% 반반)
    targetMargin: 20,       // 목표 마진율 (%)
    minMargin: 5,           // 최저 허용 마진율 (%)
  });

  // 2. 데이터 상태 관리
  const [masterCosts, setMasterCosts] = useState([]);
  const [salesData, setSalesData] = useState([]);
  const [analysisResults, setAnalysisResults] = useState([]);
  const [activeTab, setActiveTab] = useState('dashboard');

  // 원가 마스터 파일 업로드 (Master Cost)
  const handleCostUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const data = XLSX.utils.sheet_to_json(wb.Sheets[wsname]);
      setMasterCosts(data);
      alert(`원가 데이터 ${data.length}개 상품 로드 완료!`);
    };
    reader.readAsBinaryString(file);
  };

  // 주문 내역 파일 업로드 (Raw Sales)
  const handleSalesUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const data = XLSX.utils.sheet_to_json(wb.Sheets[wsname]);
      setSalesData(data);
      alert(`주문 내역 ${data.length}건 로드 완료!`);
    };
    reader.readAsBinaryString(file);
  };

  // 🚀 시뮬레이션 및 마진 분석 연산 실행
  const runAnalysis = () => {
    if (masterCosts.length === 0) {
      alert("먼저 [원가 데이터] 엑셀 파일을 업로드해주세요.");
      return;
    }

    // 1. 원가 데이터 맵 매핑 (SKU 기준)
    const costMap = {};
    masterCosts.forEach(item => {
      const sku = String(item.상품코드 || item.SKU || '').trim();
      if (sku) {
        costMap[sku] = {
          name: item.상품명 || '상품명 없음',
          tagPrice: Number(item.정가 || item.TAG || 0),
          cogs: Number(item.제조원가 || item.원가 || item.COGS || 0),
          stock: Number(item.현재재고 || item.재고 || 0),
        };
      }
    });

    // 2. 주문 데이터 집계 (SKU + 채널)
    const salesMap = {};
    salesData.forEach(row => {
      const channel = String(row.채널명 || row.채널 || '무신사국내').trim();
      const sku = String(row.상품코드 || row.SKU || '').trim();
      const qty = Number(row.판매수량 || row.수량 || 1);
      const paidAmt = Number(row.실제판매가 || row.결제금액 || 0);

      if (!sku) return;
      const key = `${sku}_${channel}`;
      if (!salesMap[key]) {
        salesMap[key] = { sku, channel, totalQty: 0, totalPaid: 0 };
      }
      salesMap[key].totalQty += qty;
      salesMap[key].totalPaid += paidAmt;
    });

    // 3. 수식 계산 루프
    const results = [];
    const targetM = promoConfig.targetMargin / 100;
    const minM = promoConfig.minMargin / 100;
    const exRate = promoConfig.exhibitionDiscount / 100;
    const onlineRate = promoConfig.onlineCoupon / 100;
    const crmRate = promoConfig.crmCoupon / 100;
    const shareRate = promoConfig.brandCrmShare;

    const keys = Object.keys(salesMap).length > 0 
      ? Object.keys(salesMap) 
      : Object.keys(costMap).map(sku => `${sku}_무신사국내`);

    keys.forEach(key => {
      let sku, channel, totalQty = 0;
      if (Object.keys(salesMap).length > 0) {
        sku = salesMap[key].sku;
        channel = salesMap[key].channel;
        totalQty = salesMap[key].totalQty;
      } else {
        sku = key.split('_')[0];
        channel = '무신사국내';
        totalQty = 0;
      }

      const costInfo = costMap[sku] || { name: '미등록 상품', tagPrice: 100000, cogs: 30000, stock: 10 };
      const tag = costInfo.tagPrice;
      const cogs = costInfo.cogs;
      const stock = costInfo.stock;

      // 채널 수수료 결정
      let feeRate = fees.musinsa_kr;
      if (channel.includes('글로벌')) feeRate = fees.musinsa_global;
      else if (channel.includes('조조')) feeRate = fees.zozotown;

      // --- [프로모션 단계별 가격 계산] ---
      // Step 1: 기획전 할인가
      const exhibitionPrice = tag * (1 - exRate);
      // Step 2: 온라인 쿠폰 적용가
      const onlineCouponPrice = exhibitionPrice * (1 - onlineRate);
      // Step 3: CRM 쿠폰 적용가 및 브랜드 부담액 산출
      const crmDiscountAmount = onlineCouponPrice * crmRate;
      const brandCrmBurden = crmDiscountAmount * shareRate; // 브랜드 부담금 (50%)
      const finalConsumerPrice = onlineCouponPrice - crmDiscountAmount; // 최종 소비자 결제가

      // Step 4: 채널 수수료 및 브랜드 최종 정산액(순매출)
      const platformFee = finalConsumerPrice * feeRate;
      const brandNetRevenue = finalConsumerPrice - platformFee - brandCrmBurden;

      // Step 5: 예상 마진율 (%)
      const expectedMarginRate = finalConsumerPrice > 0 ? ((brandNetRevenue - cogs) / finalConsumerPrice) : 0;

      // Step 6: 목표 마진율 확보를 위한 추천 가격/할인율 역산
      const denom = (1 - feeRate - targetM);
      const recommendedFinalPrice = denom > 0 ? (cogs + brandCrmBurden) / denom : finalConsumerPrice;
      const recommendedDiscountRate = tag > 0 ? Math.max(0, ((tag - recommendedFinalPrice) / tag) * 100) : 0;

      // Step 7: 재고 소진 우선순위 판정
      let priority = '3. 정상 유지 / 관망';
      let priorityClass = 'bg-gray-100 text-gray-700';

      if (expectedMarginRate >= minM) {
        if (stock >= 15 && totalQty <= 5) {
          priority = '🔥 1순위 (적극 소진)';
          priorityClass = 'bg-red-100 text-red-700 font-bold';
        } else if (stock >= 10) {
          priority = '⚡ 2순위 (프로모션 조절)';
          priorityClass = 'bg-amber-100 text-amber-800 font-semibold';
        }
      } else {
        priority = '⚠️ 4순위 (마진 경고)';
        priorityClass = 'bg-blue-100 text-blue-700 font-semibold';
      }

      results.push({
        sku,
        name: costInfo.name,
        channel,
        stock,
        totalQty,
        tag,
        cogs,
        finalConsumerPrice: Math.round(finalConsumerPrice),
        brandNetRevenue: Math.round(brandNetRevenue),
        expectedMarginRate: (expectedMarginRate * 100).toFixed(1),
        recommendedDiscountRate: recommendedDiscountRate.toFixed(1),
        recommendedFinalPrice: Math.round(recommendedFinalPrice),
        priority,
        priorityClass,
      });
    });

    setAnalysisResults(results);
    setActiveTab('dashboard');
    alert('프로모션 시뮬레이션 및 마진 분석이 완료되었습니다!');
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans p-6">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* 상단 헤더 */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-xl shadow-sm border border-gray-100 gap-4">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-black flex items-center gap-2">
              <TrendingUp className="w-7 h-7 text-indigo-600" /> 패션 MD 통합 마진 & 프로모션 시뮬레이터
            </h1>
            <p className="text-sm text-gray-500 mt-1">무신사(국내/글로벌) · 조조타운 기획전/쿠폰/마진 자동 연산 시스템</p>
          </div>
          <button 
            onClick={runAnalysis}
            className="
